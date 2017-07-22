using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel;
using System.Threading;

using Microsoft.Crm.Sdk.Messages;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;

namespace LeadToOpp_Socials
{
    public class CopySocialActivitiesPlugin:IPlugin
    {

        public void Execute(IServiceProvider serviceProvider) {

            // Obtain the execution context from the service provider.
            IPluginExecutionContext context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Get a reference to the Organization service.
            IOrganizationServiceFactory factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            //passing null will act as system
            IOrganizationService service = factory.CreateOrganizationService(null);

            if(context.InputParameters != null)
            {
                Entity oppImage = context.PostEntityImages["OpportunityImage"];
             
                // Check if an Originating Lead exists
                if (oppImage.Attributes.ContainsKey("originatingleadid"))
                {
                    // Retrieving Object Type Code for notes
                    RetrieveEntityRequest requestEnt = new RetrieveEntityRequest();
                    requestEnt.LogicalName = oppImage.LogicalName;
                    requestEnt.EntityFilters = EntityFilters.All;
                    RetrieveEntityResponse responseEnt = (RetrieveEntityResponse)service.Execute(requestEnt);
                    EntityMetadata metadataEnt = (EntityMetadata)responseEnt.EntityMetadata;
                    int? oppObjectTypeCode = metadataEnt.ObjectTypeCode;

                    EntityReference originatingLeadRef = (EntityReference)oppImage.Attributes["originatingleadid"];
                    Guid originatingLeadId = originatingLeadRef.Id;

                    // For each Note on Originating Lead, create a copy on Opportunity
                    DuplicateNotesOnLead(originatingLeadId, oppImage, oppObjectTypeCode, service, factory);

                    // For each Post on Originating Lead, create a copy on Opportunity
                    DuplicatePostsOnOpportunity(originatingLeadId, oppImage, service, factory);
                    
                }
            }
        }

        // User Actions
        private Guid CreateEntityAsUser(Entity newEntity, EntityReference creatingUser, IOrganizationServiceFactory _factory)
        {
            Guid newRecordId = Guid.Empty;

            IOrganizationService oservice = _factory.CreateOrganizationService(creatingUser.Id);
            try
            {
                newRecordId = oservice.Create(newEntity);
            }
            catch (Exception e)
            {
                string errorMsg = string.Format("Error creating record. Please check permissions for {0}.", creatingUser.Name);
                throw new InvalidPluginExecutionException(errorMsg);
            }
            return newRecordId;
        }
        private bool CanAppendAsUser(EntityReference newRecord, EntityReference creatingUser, IOrganizationServiceFactory _factory)
        {
            IOrganizationService oservice = _factory.CreateOrganizationService(creatingUser.Id);
            bool permissionGranted = false;
            Entity pulledRecord = null;

            try
            {
                pulledRecord = oservice.Retrieve(newRecord.LogicalName, newRecord.Id, new ColumnSet("name"));
                if(pulledRecord != null)
                {
                    permissionGranted = true;
                }    
            }
            catch(Exception)
            {
                return permissionGranted;
            }

            return permissionGranted;
            
        }

        // Duplicate Notes
        private void DuplicateNotesOnLead(Guid _originatingLeadId, Entity _oppImage, int? _oppObjectTypeCode, IOrganizationService _service, IOrganizationServiceFactory _factory)
        {
            QueryExpression qeNotes = new QueryExpression();
            qeNotes.EntityName = "annotation";
            qeNotes.ColumnSet = new ColumnSet(true);
            qeNotes.Criteria = new FilterExpression();
            qeNotes.Criteria.FilterOperator = LogicalOperator.And;
            qeNotes.Criteria.AddCondition(new ConditionExpression
            {
                AttributeName = "objectid",
                Operator = ConditionOperator.Equal,
                Values = { _originatingLeadId }
            });

            // Collection of Matching Records
            EntityCollection notesCollection = _service.RetrieveMultiple(qeNotes);

            foreach (Entity noteRecord in notesCollection.Entities.OrderByDescending(y => y.Attributes["createdon"]))
            {
                // Reset Id and Regarding
                noteRecord.Attributes.Remove("annotationid");
                noteRecord.Id = Guid.Empty;
                noteRecord.Attributes["objectid"] = _oppImage.ToEntityReference();
                noteRecord.Attributes["objecttypecode"] = _oppObjectTypeCode;

                // Set System Fields
                EntityReference creatingUser = (EntityReference)noteRecord.Attributes["createdby"];

                // Create Note
                if (CanAppendAsUser(_oppImage.ToEntityReference(),creatingUser,_factory))
                {
                    CreateEntityAsUser(noteRecord, creatingUser, _factory);
                }
                else
                {
                    _service.Create(noteRecord);
                }
                
            }
        }

        // Duplicate Posts
        private void DuplicatePostsOnOpportunity(Guid _originatingLeadId, Entity _oppImage, IOrganizationService _service, IOrganizationServiceFactory _factory)
        {
            // Get Manual Posts related to the Originating Lead

            QueryExpression qePosts = new QueryExpression();
            qePosts.EntityName = "post";
            qePosts.ColumnSet = new ColumnSet(true);
            qePosts.Criteria = new FilterExpression();
            qePosts.Criteria.FilterOperator = LogicalOperator.And;
            qePosts.Criteria.AddCondition(new ConditionExpression
            {
                AttributeName = "regardingobjectid",
                Operator = ConditionOperator.Equal,
                Values = { _originatingLeadId }
            });

            EntityCollection postsCollection = _service.RetrieveMultiple(qePosts);
            OptionSetValue manualValue = new OptionSetValue(2);
            
            // Run through all posts and act on User Posts

            foreach (Entity postRecord in postsCollection.Entities.OrderByDescending(y => y.Attributes["createdon"]))
            {
                if (((OptionSetValue)postRecord.Attributes["source"]).Value == manualValue.Value)
                {
                    // Get Post Replies

                    QueryExpression qePostComments = new QueryExpression();
                    qePostComments.EntityName = "postcomment";
                    qePostComments.ColumnSet = new ColumnSet(true);
                    qePostComments.Criteria = new FilterExpression();
                    qePostComments.Criteria.FilterOperator = LogicalOperator.And;
                    qePostComments.Criteria.AddCondition(new ConditionExpression
                    {
                        AttributeName = "postid",
                        Operator = ConditionOperator.Equal,
                        Values = { postRecord.Id }
                    });

                    EntityCollection postComments = _service.RetrieveMultiple(qePostComments);

                    // Get Post Likes

                    QueryExpression qePostLikes = new QueryExpression();
                    qePostLikes.EntityName = "postlike";
                    qePostLikes.ColumnSet = new ColumnSet(true);
                    qePostLikes.Criteria = new FilterExpression();
                    qePostLikes.Criteria.FilterOperator = LogicalOperator.And;
                    qePostLikes.Criteria.AddCondition(new ConditionExpression
                    {
                        AttributeName = "postid",
                        Operator = ConditionOperator.Equal,
                        Values = { postRecord.Id }
                    });

                    EntityCollection postLikes = _service.RetrieveMultiple(qePostLikes);

                    // Reset Post Id and Regarding

                    postRecord.Attributes.Remove("postid");
                    postRecord.Id = Guid.Empty;
                    postRecord.Attributes["regardingobjectid"] = _oppImage.ToEntityReference();

                    // Set System Fields
                    EntityReference creatingUser = (EntityReference)postRecord.Attributes["createdby"];

                    // Create Post
                    Guid duplicatePostId = Guid.Empty;

                    if (CanAppendAsUser(_oppImage.ToEntityReference(), creatingUser, _factory))
                    {
                        duplicatePostId = CreateEntityAsUser(postRecord, creatingUser, _factory);
                    }
                    else
                    {
                        duplicatePostId = _service.Create(postRecord);
                    }
                    
                    
                    foreach (Entity postComment in postComments.Entities.OrderByDescending(y => y.Attributes["createdon"]))
                    {
                        postComment.Attributes.Remove("postcommentid");
                        postComment.Id = Guid.Empty;

                        postComment.Attributes["postid"] = new EntityReference("post", duplicatePostId);

                        EntityReference postingUser = (EntityReference)postComment.Attributes["createdby"];

                        // Create Post Comment
                        if (CanAppendAsUser(_oppImage.ToEntityReference(), postingUser, _factory))
                        {
                            CreateEntityAsUser(postComment, postingUser, _factory);
                        }
                        else
                        {
                            _service.Create(postComment);
                        }
                    }

                    foreach (Entity postLike in postLikes.Entities.OrderByDescending(y => y.Attributes["createdon"]))
                    {
                        postLike.Attributes.Remove("postlikeid");
                        postLike.Id = Guid.Empty;

                        postLike.Attributes["postid"] = new EntityReference("post", duplicatePostId);

                        EntityReference likingUser = (EntityReference)postLike.Attributes["createdby"];

                        // Create Post Like
                        if (CanAppendAsUser(_oppImage.ToEntityReference(), likingUser, _factory))
                        {
                            CreateEntityAsUser(postLike, likingUser, _factory);
                        }
                    }

                }
                
            }
        }
    }
}

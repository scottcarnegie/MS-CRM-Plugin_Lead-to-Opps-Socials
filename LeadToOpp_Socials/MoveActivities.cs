using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;

namespace LeadToOpp_Socials
{
    class MoveActivities
    {
        // Move Activities (This is done OOTB by the system)     
        private void MovePhoneCallsFromLead(Guid _originatingLeadId, Entity _oppImage, IOrganizationService _service, IOrganizationServiceFactory _factory)
        {
            // Get phone call records

            QueryExpression qeActivities = new QueryExpression();
            qeActivities.EntityName = "phonecall";
            qeActivities.ColumnSet = new ColumnSet(true);
            qeActivities.Criteria = new FilterExpression();
            qeActivities.Criteria.FilterOperator = LogicalOperator.And;
            qeActivities.Criteria.AddCondition(new ConditionExpression
            {
                AttributeName = "regardingobjectid",
                Operator = ConditionOperator.Equal,
                Values = { _originatingLeadId }
            });

            EntityCollection phoneCallCollection = _service.RetrieveMultiple(qeActivities);

            foreach (Entity phoneCall in phoneCallCollection.Entities)
            {
                // Create user context for service to act as user.
                EntityReference owningUser = (EntityReference)phoneCall.Attributes["ownerid"];

                // Check state code: if inactive, must reopen
                /*
                    State Code: 0 Open          Status: 1 Open
                    State Code: 1 Completed     Status: 2 Made
                    State Code: 1 Completed     Status: 4 Received
                    State Code: 2 Canceled      Status: 3 Canceled
                 */

                // Get original codes.

                int startStateCode = ((OptionSetValue)phoneCall["statecode"]).Value;
                int startStatus = ((OptionSetValue)phoneCall["statuscode"]).Value;

                // If Status is not Open..

                if (startStatus != 1)
                {
                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        // Open it.
                        EntityMoniker = new EntityReference
                        {
                            Id = phoneCall.Id,
                            LogicalName = phoneCall.LogicalName
                        },
                        State = new OptionSetValue(0),
                        Status = new OptionSetValue(1)
                    };

                    try
                    {
                        //Change status to open as System
                        _service.Execute(setStateRequest);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }

                    // Update fields.
                    phoneCall.Attributes["regardingobjectid"] = _oppImage.ToEntityReference();
                    phoneCall.Attributes["statuscode"] = new OptionSetValue(1);

                    try
                    {
                        // Update as Owning User
                        UpdateEntityAsUser(phoneCall, owningUser, _factory);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }

                    // Return status to closed
                    setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = phoneCall.Id,
                            LogicalName = phoneCall.LogicalName
                        },
                        State = new OptionSetValue(startStateCode),
                        Status = new OptionSetValue(startStatus)
                    };

                    try
                    {
                        // Execute State Change as Owning User
                        ExecuteRequestAsUser(setStateRequest, owningUser, _factory);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }
                }
                else
                {
                    phoneCall.Attributes["regardingobjectid"] = _oppImage.ToEntityReference();
                    try
                    {
                        UpdateEntityAsUser(phoneCall, owningUser, _factory);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }
                }

            }
        }

        private void UpdateEntityAsUser(Entity newEntity, EntityReference updateUser, IOrganizationService _defaultService, IOrganizationServiceFactory _factory)
        {
            IOrganizationService oservice = _factory.CreateOrganizationService(updateUser.Id);
            try
            {
                oservice.Update(newEntity);
            }
            catch (Exception)
            {
                _defaultService.Update(newEntity);
                //string errorMsg = string.Format("Error creating record. Please check permissions for {0}.", updateUser.Name);
                //throw new InvalidPluginExecutionException(errorMsg);
            }

        }
        private void ExecuteRequestAsUser(OrganizationRequest _orgRequest, EntityReference executionUser, IOrganizationService _defaultService, IOrganizationServiceFactory _factory)
        {
            IOrganizationService oservice = _factory.CreateOrganizationService(executionUser.Id);
            try
            {
                oservice.Execute(_orgRequest);
            }
            catch (Exception)
            {
                _defaultService.Execute(_orgRequest);
                //string errorMsg = string.Format("Error creating record. Please check permissions for {0}.", executionUser.Name);
                //throw new InvalidPluginExecutionException(errorMsg);
            }

        }

        private void MoveEmailsFromLead(Guid _originatingLeadId, Entity _oppImage, IOrganizationService _service, IOrganizationServiceFactory _factory)
        {
            // Get email records

            QueryExpression qeActivities = new QueryExpression();
            qeActivities.EntityName = "email";
            qeActivities.ColumnSet = new ColumnSet(true);
            qeActivities.Criteria = new FilterExpression();
            qeActivities.Criteria.FilterOperator = LogicalOperator.And;
            qeActivities.Criteria.AddCondition(new ConditionExpression
            {
                AttributeName = "regardingobjectid",
                Operator = ConditionOperator.Equal,
                Values = { _originatingLeadId }
            });

            EntityCollection emailCollection = _service.RetrieveMultiple(qeActivities);

            foreach (Entity email in emailCollection.Entities)
            {
                /*
                    State Code: 0 Open	    Status: 1 Draft
                    State Code: 0 Open	    Status: 8 Failed
                    State Code: 1 Completed	Status: 2 Completed
                    State Code: 1 Completed	Status: 3 Sent
                    State Code: 1 Completed	Status: 4 Received
                    State Code: 1 Completed	Status: 6 Pending Send
                    State Code: 1 Completed	Status: 7 Sending
                    State Code: 2 Canceled	Status: 5 Canceled
                 */

                // Update fields Regarding.
                email.Attributes["regardingobjectid"] = _oppImage.ToEntityReference();

                try
                {
                    // Update as System
                    _service.Update(email);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error: " + ex);
                }

            }
        }
        private void MoveTasksFromLead(Guid _originatingLeadId, Entity _oppImage, IOrganizationService _service, IOrganizationServiceFactory _factory)
        {
            // Get tasks records

            QueryExpression qeActivities = new QueryExpression();
            qeActivities.EntityName = "task";
            qeActivities.ColumnSet = new ColumnSet(true);
            qeActivities.Criteria = new FilterExpression();
            qeActivities.Criteria.FilterOperator = LogicalOperator.And;
            qeActivities.Criteria.AddCondition(new ConditionExpression
            {
                AttributeName = "regardingobjectid",
                Operator = ConditionOperator.Equal,
                Values = { _originatingLeadId }
            });

            EntityCollection taskCollection = _service.RetrieveMultiple(qeActivities);

            foreach (Entity taskActivity in taskCollection.Entities)
            {
                // Create user context for service to act as user.
                EntityReference owningUser = (EntityReference)taskActivity.Attributes["ownerid"];

                // Check state code: if inactive, must reopen
                /*
                    State Code: 0 Open      Status: 2 Not Started
                    State Code: 0 Open      Status: 3 In Progress
                    State Code: 0 Open      Status: 4 Waiting on someone else
                    State Code: 0 Open      Status: 7 Deferred
                    State Code: 1 Completed Status:	5 Completed
                    State Code: 2 Canceled	Status: 6 Canceled
                 */

                // Get original codes.

                int startStateCode = ((OptionSetValue)taskActivity["statecode"]).Value;
                int startStatus = ((OptionSetValue)taskActivity["statuscode"]).Value;

                // If Status is not Open..

                if (startStatus == 5 || startStatus == 6)
                {
                    SetStateRequest setStateRequest = new SetStateRequest()
                    {
                        // Open it.
                        EntityMoniker = new EntityReference
                        {
                            Id = taskActivity.Id,
                            LogicalName = taskActivity.LogicalName
                        },
                        State = new OptionSetValue(0),
                        Status = new OptionSetValue(7)
                    };

                    try
                    {
                        //Change status to open as System
                        _service.Execute(setStateRequest);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }

                    // Update fields.
                    taskActivity.Attributes["regardingobjectid"] = _oppImage.ToEntityReference();
                    taskActivity.Attributes["statuscode"] = new OptionSetValue(7);

                    try
                    {
                        // Update as Owning User
                        UpdateEntityAsUser(taskActivity, owningUser, _factory);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }

                    // Return status to closed
                    setStateRequest = new SetStateRequest()
                    {
                        EntityMoniker = new EntityReference
                        {
                            Id = taskActivity.Id,
                            LogicalName = taskActivity.LogicalName
                        },
                        State = new OptionSetValue(startStateCode),
                        Status = new OptionSetValue(startStatus)
                    };

                    try
                    {
                        // Execute State Change as Owning User
                        ExecuteRequestAsUser(setStateRequest, owningUser, _factory);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }
                }
                else
                {
                    taskActivity.Attributes["regardingobjectid"] = _oppImage.ToEntityReference();
                    try
                    {
                        UpdateEntityAsUser(taskActivity, owningUser, _factory);
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Error: " + ex);
                    }
                }

            }
        }
        private void MoveAppointmentsFromLead(Guid _originatingLeadId, Entity _oppImage, IOrganizationService _service, IOrganizationServiceFactory _factory)
        {
            // Get appointment records

            QueryExpression qeActivities = new QueryExpression();
            qeActivities.EntityName = "appointment";
            qeActivities.ColumnSet = new ColumnSet(true);
            qeActivities.Criteria = new FilterExpression();
            qeActivities.Criteria.FilterOperator = LogicalOperator.And;
            qeActivities.Criteria.AddCondition(new ConditionExpression
            {
                AttributeName = "regardingobjectid",
                Operator = ConditionOperator.Equal,
                Values = { _originatingLeadId }
            });

            EntityCollection apptCollection = _service.RetrieveMultiple(qeActivities);

            foreach (Entity appointment in apptCollection.Entities)
            {
                // Create user context for service to act as user.
                EntityReference owningUser = (EntityReference)appointment.Attributes["ownerid"];

                // Check state code: if inactive, must reopen
                /*
                    State Code: 0 Open	    Status: 1 Free
	                State Code: 0 Open      Status: 2 Tentative
                    State Code: 1 Completed	Status: 3 Completed
                    State Code: 2 Canceled	Status: 4 Canceled
                    State Code: 3 Scheduled	Status: 5 Busy
	                State Code: 3 Scheduled Status: 6 Out of Office
                 */

                // Get original codes.

                int startStateCode = ((OptionSetValue)appointment["statecode"]).Value;
                int startStatus = ((OptionSetValue)appointment["statuscode"]).Value;

                appointment.Attributes["regardingobjectid"] = _oppImage.ToEntityReference();
                try
                {
                    UpdateEntityAsUser(appointment, owningUser, _factory);
                }
                catch (Exception ex)
                {
                    throw new Exception("Error: " + ex);
                }
            }
        }

        // User Actions
        private Guid CreateEntityAsUser(Entity newEntity, EntityReference creatingUser, IOrganizationServiceFactory _factory)
        {
            Guid newRecordId = Guid.Empty;
            //passing current user
            IOrganizationService oservice = _factory.CreateOrganizationService(creatingUser.Id);
            try
            {
                newRecordId = oservice.Create(newEntity);
            }
            catch (Exception)
            {
                string errorMsg = string.Format("Error creating record. Please check permissions for {0}.", creatingUser.Name);
                throw new InvalidPluginExecutionException(errorMsg);
            }
            return newRecordId;
        }
        private void UpdateEntityAsUser(Entity newEntity, EntityReference updateUser, IOrganizationServiceFactory _factory)
        {
            IOrganizationService oservice = _factory.CreateOrganizationService(updateUser.Id);
            try
            {
                oservice.Update(newEntity);
            }
            catch (Exception)
            {
                string errorMsg = string.Format("Error creating record. Please check permissions for {0}.", updateUser.Name);
                throw new InvalidPluginExecutionException(errorMsg);
            }

        }
        private void ExecuteRequestAsUser(OrganizationRequest _orgRequest, EntityReference executionUser, IOrganizationServiceFactory _factory)
        {
            IOrganizationService oservice = _factory.CreateOrganizationService(executionUser.Id);
            try
            {
                oservice.Execute(_orgRequest);
            }
            catch (Exception)
            {
                string errorMsg = string.Format("Error creating record. Please check permissions for {0}.", executionUser.Name);
                throw new InvalidPluginExecutionException(errorMsg);
            }

        }

    }
}

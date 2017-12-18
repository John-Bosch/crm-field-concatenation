using FieldConcatenation.plugins.CRM;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FieldConcatenation.plugins
{
    public class ConfigurationChanges : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            var tracer = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            var factory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = factory.CreateOrganizationService(context.UserId);

            try
            {
                switch (context.MessageName.ToLowerInvariant())
                {
                    case "create":
                        {
                            var entity = context.InputParameters["Target"] as Entity;
                            var entityLogicalName = entity.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_EntityLogicalName);
                            var fieldName = entity.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_FieldName);
                            var formatString = entity.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_FormatString);

                            ValidateInformation(formatString, entityLogicalName, service);

                            // Create the create step
                            CreateMessageProcessingtStep(
                                service,
                                tracer,
                                entityLogicalName,
                                CrmData.SdkMessageName.Create,
                                null,
                                CrmData.FieldConcatenationPluginAssemblyName,
                                CrmData.FieldConcatenationPluginTypeName);

                            var stepId = CreateMessageProcessingtStep(
                                    service,
                                    tracer,
                                    entityLogicalName,
                                    CrmData.SdkMessageName.Update,
                                    null,
                                    CrmData.FieldConcatenationPluginAssemblyName,
                                    CrmData.FieldConcatenationPluginTypeName);

                            CreateMessageProcessingStepImage(
                                service,
                                tracer,
                                stepId,
                                "PreUpdateImage",
                                CrmData.SdkImageType.PreImage,
                                CrmData.SdkMessageName.Update,
                                string.Empty);

                            break;
                        }
                    case "update":
                        {
                            var postUpdateImage = context.PostEntityImages["PostUpdateImage"];
                            var entityLogicalName = postUpdateImage.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_EntityLogicalName);
                            var formatString = postUpdateImage.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_FormatString);

                            ValidateInformation(formatString, entityLogicalName, service);
                            
                            CreateMessageProcessingtStep(
                                service,
                                tracer,
                                entityLogicalName,
                                CrmData.SdkMessageName.Create,
                                null,
                                CrmData.FieldConcatenationPluginAssemblyName,
                                CrmData.FieldConcatenationPluginTypeName);

                            var stepId = CreateMessageProcessingtStep(
                                    service,
                                    tracer,
                                    entityLogicalName,
                                    CrmData.SdkMessageName.Update,
                                    null,
                                    CrmData.FieldConcatenationPluginAssemblyName,
                                    CrmData.FieldConcatenationPluginTypeName);

                            CreateMessageProcessingStepImage(
                                service,
                                tracer,
                                stepId,
                                "PreUpdateImage",
                                CrmData.SdkImageType.PreImage,
                                CrmData.SdkMessageName.Update,
                                string.Empty);

                            break;
                        }
                    case "delete":
                        {
                            var entityReference = context.InputParameters["Target"] as EntityReference;
                            var preDeleteImage = context.PreEntityImages.Contains("PreDeleteImage") ? context.PreEntityImages["PreDeleteImage"] : null;
                            var entityLogicalName = preDeleteImage.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_EntityLogicalName);

                            if (!OtherConfigurationsExist(service, entityReference, entityLogicalName))
                            {
                                tracer.Trace("Deleting Number Generator 'create' step for enttity [{0}].", entityLogicalName);

                                DeleteMessageProcessingSteps(service, entityLogicalName, CrmData.SdkMessageName.Create, CrmData.FieldConcatenationPluginAssemblyName, CrmData.FieldConcatenationPluginTypeName);
                            }
                            else
                            {
                                tracer.Trace("Other Number Generator configurations exists for entity [{0}], so leave the Number Generator 'create' step in place.", entityLogicalName);
                            }

                            if (!OtherUpdateConfigurationsExist(service, entityReference, entityLogicalName))
                            {
                                tracer.Trace("Deleting Number Generator 'update' step for entity [{0}].", entityLogicalName);

                                DeleteMessageProcessingSteps(service, entityLogicalName, CrmData.SdkMessageName.Update, CrmData.FieldConcatenationPluginAssemblyName, CrmData.FieldConcatenationPluginTypeName);
                            }

                            break;
                        }
                    default:
                        throw new InvalidOperationException(string.Format("Message [{0}] is not supported.", context.MessageName));
                }
            }
            catch (Exception e)
            {
                throw new InvalidPluginExecutionException(e.Message);
            }
        }

        private void DeleteMessageProcessingSteps(
            IOrganizationService service, 
            string entityLogicalName, 
            CrmData.SdkMessageName messageName, 
            string assemblyName, 
            string pluginName)
        {
            var existingSteps = GetStepReferences(service, entityLogicalName, messageName, assemblyName, pluginName);

            foreach (var step in existingSteps)
            {
                service.Delete(CrmData.SdkMessageProccessingStepEntityName, step.Id);
            }
        }

        private Guid CreateMessageProcessingtStep(
            IOrganizationService service,
            ITracingService tracer,
            string entityLogicalName,
            CrmData.SdkMessageName messageName,
            IEnumerable<string> attributeNames,
            string assemblyName,
            string pluginName)
        {
            var messageReference = GetMessageReference(service, messageName);
            var existingSteps = GetStepReferences(service, entityLogicalName, messageName, assemblyName, pluginName);

            if (existingSteps.Any())
            {
                tracer.Trace("A '{0}' step already exists for entity [{1}].", messageName.ToString(), entityLogicalName);
                return existingSteps.First().Id;
            }
            else
            {
                tracer.Trace("Creating a '{0}' step for the entity [{1}].", messageName.ToString(), entityLogicalName);

                // Register a "create" step under the Number Generator plugin for the requested entity type
                var step = new Entity(CrmData.SdkMessageProccessingStepEntityName);
                step[CrmData.SdkMessageProccessingStep_Name] = string.Format("{0}: {1} of {2}", pluginName, messageName, entityLogicalName);
                step[CrmData.SdkMessageProccessingStep_Mode] = new OptionSetValue((int)CrmData.CrmPluginStepMode.Synchronous);
                step[CrmData.SdkMessageProccessingStep_PluginTypeId] = GetPluginTypeReference(service, assemblyName, pluginName);
                step[CrmData.SdkMessageProccessingStep_Rank] = 1;
                step[CrmData.SdkMessageProccessingStep_SdkMessageId] = messageReference;
                step[CrmData.SdkMessageProccessingStep_Stage] = new OptionSetValue((int)CrmData.CrmPluginStepStage.PreOperation);
                step[CrmData.SdkMessageProccessingStep_SupportedDeployment] = new OptionSetValue((int)CrmData.CrmPluginStepDeployment.ServerOnly);
                step[CrmData.SdkMessageProccessingStep_SdkMessageFilterId] = GetMessageFilterReference(service, messageReference, entityLogicalName);

                if (attributeNames != null && attributeNames.Any())
                {
                    step[CrmData.SdkMessageProccessingStep_FilteringAttributes] = String.Join(",", attributeNames);
                }

                var stepId = service.Create(step);
                return stepId;
            }
        }

        private void CreateMessageProcessingStepImage(
            IOrganizationService service,
            ITracingService tracer,
            Guid stepId,
            string imageName,
            CrmData.SdkImageType imageType,
            CrmData.SdkMessageName messageName,
            string attributes)
        {
            var existingImages = GetStepImageReferences(service, stepId, imageName, imageType);

            if (existingImages.Any())
            {
                tracer.Trace("A step image named '{0}' of type '{1}' already exists.", imageName, imageType.ToString());
            }
            else
            {
                var image = new Entity(CrmData.SdkMessageProcessingStepImageEntityName);
                image[CrmData.SdkMessageProcessingStepImage_SdkMessageProcessingStepId] = new EntityReference(CrmData.SdkMessageProccessingStepEntityName, stepId);
                image[CrmData.SdkMessageProcessingStepImage_Name] = imageName;
                image[CrmData.SdkMessageProcessingStepImage_EntityAlias] = imageName;
                image[CrmData.SdkMessageProcessingStepImage_ImageType] = new OptionSetValue((int)imageType);
                image[CrmData.SdkMessageProcessingStepImage_MessagePropertyName] = CrmData.GetMessageProperyName(messageName);

                if (!string.IsNullOrEmpty(attributes))
                {
                    image[CrmData.SdkMessageProcessingStepImage_Attributes] = attributes;
                }

                service.Create(image);
            }
        }

        private IEnumerable<EntityReference> GetStepImageReferences(
            IOrganizationService service,
            Guid stepId,
            string imageName,
            CrmData.SdkImageType imageType)
        {
            var query = new QueryExpression
            {
                EntityName = CrmData.SdkMessageProcessingStepImageEntityName
            };

            query.Criteria.AddCondition(
                CrmData.SdkMessageProcessingStepImage_SdkMessageProcessingStepId,
                ConditionOperator.Equal,
                stepId);

            query.Criteria.AddCondition(
                CrmData.SdkMessageProcessingStepImage_Name,
                ConditionOperator.Equal,
                imageName);

            query.Criteria.AddCondition(
                CrmData.SdkMessageProcessingStepImage_ImageType,
                ConditionOperator.Equal,
                (int)imageType);

            var steps = service.RetrieveMultiple(query);
            return steps.Entities.Select(e => e.ToEntityReference());
        }

        private EntityReference GetMessageReference(IOrganizationService service, CrmData.SdkMessageName messageName)
        {
            var query = new QueryExpression
            {
                EntityName = CrmData.SdkMessageEntityName,
                ColumnSet = new ColumnSet(CrmData.SdkMessage_Id)
            };

            query.Criteria.AddCondition(
                CrmData.SdkMessage_Name,
                ConditionOperator.Equal,
                messageName.ToString());

            var sdkMessages = service.RetrieveMultiple(query);
            var sdkMessage = sdkMessages.Entities.First();

            return sdkMessage.ToEntityReference();
        }

        private IEnumerable<EntityReference> GetStepReferences(
            IOrganizationService service,
            string entityLogicalName,
            CrmData.SdkMessageName messageName,
            string assemblyName,
            string pluginName)
        {
            var messageReference = GetMessageReference(service, messageName);

            // Find message processing steps
            var query = new QueryExpression
            {
                EntityName = CrmData.SdkMessageProccessingStepEntityName,
                ColumnSet = new ColumnSet(CrmData.SdkMessageProccessingStep_Id)
            };

            // That are for the Number Generator plugin
            query.Criteria.AddCondition(
                CrmData.SdkMessageProccessingStep_PluginTypeId,
                ConditionOperator.Equal,
                GetPluginTypeReference(service, assemblyName, pluginName).Id);

            // And of the correct message type
            query.Criteria.AddCondition(
                CrmData.SdkMessageProccessingStep_SdkMessageId,
                ConditionOperator.Equal,
                messageReference.Id);

            // With a message filter
            var messageFilterLink = query.AddLink(
                CrmData.SdkMessageFilterEntityName,
                CrmData.SdkMessageProccessingStep_SdkMessageFilterId,
                CrmData.SdkMessageFilter_Id);

            // That is for a create message
            messageFilterLink.LinkCriteria.AddCondition(
                CrmData.SdkMessageFilter_SdkMessageId,
                ConditionOperator.Equal,
                messageReference.Id);

            // And to our entity
            messageFilterLink.LinkCriteria.AddCondition(
                CrmData.SdkMessageFilter_PrimaryObjectTypeCode,
                ConditionOperator.Equal,
                entityLogicalName);

            var steps = service.RetrieveMultiple(query);
            return steps.Entities.Select(e => e.ToEntityReference());
        }

        private EntityReference GetPluginTypeReference(IOrganizationService service, string assemblyName, string pluginName)
        {
            var query = new QueryExpression
            {
                EntityName = CrmData.PluginTypeEntityName,
                ColumnSet = new ColumnSet(CrmData.PluginType_Id)
            };

            query.Criteria.AddCondition(
                CrmData.PluginType_TypeName,
                ConditionOperator.Equal,
                pluginName);

            var assemblyLink = query.AddLink(
                    CrmData.PluginAssemblyEntityName,
                    CrmData.PluginType_PluginAssemblyId,
                    CrmData.PluginAssembly_Id);

            assemblyLink.LinkCriteria.AddCondition(
                CrmData.PluginAssembly_Name,
                ConditionOperator.Equal,
                assemblyName);

            var pluginTypes = service.RetrieveMultiple(query);
            var pluginType = pluginTypes.Entities.First();

            return pluginType.ToEntityReference();
        }

        private EntityReference GetMessageFilterReference(IOrganizationService service, EntityReference messageReference, string entityLogicalName)
        {
            var query = new QueryExpression
            {
                EntityName = CrmData.SdkMessageFilterEntityName,
                ColumnSet = new ColumnSet(CrmData.SdkMessageFilter_Id)
            };

            query.Criteria.AddCondition(
                CrmData.SdkMessageFilter_SdkMessageId,
                ConditionOperator.Equal,
                messageReference.Id);

            query.Criteria.AddCondition(
                CrmData.SdkMessageFilter_PrimaryObjectTypeCode,
                ConditionOperator.Equal,
                entityLogicalName);

            var messageFilters = service.RetrieveMultiple(query);
            var messageFilter = messageFilters.Entities.First();

            return messageFilter.ToEntityReference();
        }

        private bool OtherConfigurationsExist(IOrganizationService service, EntityReference thisConfigurationReference, string entityLogicalName)
        {
            var query = new QueryExpression
            {
                EntityName = CrmData.FieldConcatenationConfigurationEntityLogicalName,
                ColumnSet = new ColumnSet(CrmData.FieldConcatenationConfiguration_Id)
            };

            query.Criteria.AddCondition(
                CrmData.FieldConcatenationConfiguration_Id,
                ConditionOperator.NotEqual,
                thisConfigurationReference.Id);

            query.Criteria.AddCondition(
                CrmData.FieldConcatenationConfiguration_EntityLogicalName,
                ConditionOperator.Equal,
                entityLogicalName);

            var configurations = service.RetrieveMultiple(query);
            return configurations.Entities.Any();
        }

        private bool OtherUpdateConfigurationsExist(IOrganizationService service, EntityReference thisConfigurationReference, string entityLogicalName)
        {
            var query = new QueryExpression
            {
                EntityName = CrmData.FieldConcatenationConfigurationEntityLogicalName,
                ColumnSet = new ColumnSet(CrmData.FieldConcatenationConfiguration_Id)
            };

            query.Criteria.AddCondition(
                CrmData.FieldConcatenationConfiguration_Id,
                ConditionOperator.NotEqual,
                thisConfigurationReference.Id);

            query.Criteria.AddCondition(
                CrmData.FieldConcatenationConfiguration_EntityLogicalName,
                ConditionOperator.Equal,
                entityLogicalName);

            var configurations = service.RetrieveMultiple(query);
            return configurations.Entities.Any();
        }

        private void ValidateInformation(string formatString, string entityLogicalName, IOrganizationService service)
        {
            var request = new RetrieveEntityRequest
            {
                LogicalName = entityLogicalName,
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes
            };

            var response = (RetrieveEntityResponse)service.Execute(request);
            var tokens = FormatStringHelper.TokeniseString(response.EntityMetadata.Attributes, formatString);

            var unknownFields = tokens.Where(t => t.Status == AttributeTokenStatus.UnknownAttribute).Select(t => t.Name);
            if (unknownFields.Any())
            {
                throw new InvalidOperationException(string.Format("Fields [{0}] do not exist on entity [{1}].", string.Join(", ", unknownFields), entityLogicalName));
            }

            var unsupportedFields = tokens.Where(t => t.Status == AttributeTokenStatus.UnsupportedType);
            if (unsupportedFields.Any())
            {
                throw new InvalidOperationException(string.Format("The following fields have unsupported types: [{0}].", string.Join(", ", unsupportedFields.Select(t => string.Format("{0}: {1}", t.Name, t.AttributeMetadata.AttributeType.ToString())))));
            }
        }
    }
}

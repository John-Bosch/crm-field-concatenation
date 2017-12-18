using FieldConcatenation.plugins.CRM;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;

namespace FieldConcatenation.plugins
{
    public class ConcatenateFields : IPlugin
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

                            var configurations = GetConfigurations(service, entity.LogicalName);

                            if (!configurations.Any())
                            {
                                tracer.Trace("Found no matching configurations.");
                            }

                            foreach (var configuration in configurations)
                            {
                                ProcessConfiguration(service, entity, null, configuration, tracer);
                            }

                            break;
                        }
                    case "update":
                        {
                            var entity = context.InputParameters["Target"] as Entity;
                            var preUpdateImage = context.PreEntityImages["PreUpdateImage"];

                            var configurations = GetConfigurations(service, entity.LogicalName);

                            if (!configurations.Any())
                            {
                                tracer.Trace("Found no matching configurations.");
                            }

                            foreach (var configuration in configurations)
                            {
                                var fieldName = configuration.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_FieldName);
                                ProcessConfiguration(service, entity, preUpdateImage, configuration, tracer);
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

        private IEnumerable<Entity> GetConfigurations(IOrganizationService service, string entityLogicalName)
        {
            var query = new QueryExpression
            {
                EntityName = CrmData.FieldConcatenationConfigurationEntityLogicalName,
                ColumnSet = new ColumnSet(true)
            };

            query.Criteria.AddCondition(
                CrmData.FieldConcatenationConfiguration_EntityLogicalName,
                ConditionOperator.Equal,
                entityLogicalName);

            var configurations = service.RetrieveMultiple(query);
            return configurations.Entities;
        }

        private void ProcessConfiguration(IOrganizationService service, Entity changeEntity, Entity preChangeEntity, Entity configuration, ITracingService tracer)
        {
            var formatString = configuration.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_FormatString);
            var fieldName = configuration.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_FieldName);

            var request = new RetrieveEntityRequest
            {
                LogicalName = changeEntity.LogicalName,
                EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Attributes
            };

            var response = (RetrieveEntityResponse)service.Execute(request);
            var tokens = FormatStringHelper.TokeniseString(response.EntityMetadata.Attributes, formatString);

            // Check that something we're interested in changed
            var changedAttributes = changeEntity.Attributes.Select(a => a.Key);

            if (tokens.Any(t => changedAttributes.Contains(t.Name)))
            {
                var result = FormatStringHelper.FormatString(formatString, tokens, changeEntity, preChangeEntity);
                changeEntity.Attributes[fieldName] = result;
            }
        }
    }
}

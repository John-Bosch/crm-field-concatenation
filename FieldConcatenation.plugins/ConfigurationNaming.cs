using Veritec.Crm.FieldConcatenation.plugins.CRM;
using Microsoft.Xrm.Sdk;
using System;

namespace Veritec.Crm.FieldConcatenation.plugins
{
    public class ConfigurationNaming : IPlugin
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

                            var logicalName = entity.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_EntityLogicalName);
                            var fieldName = entity.GetAttributeValue<string>(CrmData.FieldConcatenationConfiguration_FieldName);

                            var name = string.Format("{0} - {1}", logicalName, fieldName);
                            entity[CrmData.FieldConcatenationConfiguration_Name] = name;

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
    }
}

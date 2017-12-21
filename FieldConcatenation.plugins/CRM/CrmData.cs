namespace Veritec.Crm.FieldConcatenation.plugins.CRM
{
    internal static class CrmData
    {
        public const string FieldConcatenationConfigurationEntityLogicalName = "ver_fieldconcatenationconfiguration";
        public const string FieldConcatenationConfiguration_Id = "ver_fieldconcatenationconfigurationid";
        public const string FieldConcatenationConfiguration_Name = "ver_name";
        public const string FieldConcatenationConfiguration_EntityLogicalName = "ver_entitylogicalname";
        public const string FieldConcatenationConfiguration_FieldName = "ver_fieldname";
        public const string FieldConcatenationConfiguration_FormatString = "ver_formatstring";

        public const string FieldConcatenationPluginAssemblyName = "Veritec.Crm.FieldConcatenation.Plugins";
        public const string FieldConcatenationPluginTypeName = "Veritec.Crm.FieldConcatenation.plugins.ConcatenateFields";

        public const string SdkMessageProccessingStepEntityName = "sdkmessageprocessingstep";
        public const string SdkMessageProccessingStep_Id = "sdkmessageprocessingstepid";
        public const string SdkMessageProccessingStep_Name = "name";
        public const string SdkMessageProccessingStep_Mode = "mode";
        public const string SdkMessageProccessingStep_Rank = "rank";
        public const string SdkMessageProccessingStep_Stage = "stage";
        public const string SdkMessageProccessingStep_SupportedDeployment = "supporteddeployment";
        public const string SdkMessageProccessingStep_InvocationSource = "invocationsource";
        public const string SdkMessageProccessingStep_PluginTypeId = "plugintypeid";
        public const string SdkMessageProccessingStep_SdkMessageId = "sdkmessageid";
        public const string SdkMessageProccessingStep_SdkMessageFilterId = "sdkmessagefilterid";
        public const string SdkMessageProccessingStep_FilteringAttributes = "filteringattributes";

        public const string SdkMessageProcessingStepImageEntityName = "sdkmessageprocessingstepimage";
        public const string SdkMessageProcessingStepImage_Name = "name";
        public const string SdkMessageProcessingStepImage_EntityAlias = "entityalias";
        public const string SdkMessageProcessingStepImage_ImageType = "imagetype";
        public const string SdkMessageProcessingStepImage_SdkMessageProcessingStepId = "sdkmessageprocessingstepid";
        public const string SdkMessageProcessingStepImage_MessagePropertyName = "messagepropertyname";
        public const string SdkMessageProcessingStepImage_Attributes = "attributes";

        public const string PluginAssemblyEntityName = "pluginassembly";
        public const string PluginAssembly_Id = "pluginassemblyid";
        public const string PluginAssembly_Name = "name";

        public const string PluginTypeEntityName = "plugintype";
        public const string PluginType_Id = "plugintypeid";
        public const string PluginType_TypeName = "typename";
        public const string PluginType_PluginAssemblyId = "pluginassemblyid";

        public const string SdkMessageEntityName = "sdkmessage";
        public const string SdkMessage_Id = "sdkmessageid";
        public const string SdkMessage_Name = "name";

        public const string SdkMessageFilterEntityName = "sdkmessagefilter";
        public const string SdkMessageFilter_Id = "sdkmessagefilterid";
        public const string SdkMessageFilter_SdkMessageId = "sdkmessageid";
        public const string SdkMessageFilter_PrimaryObjectTypeCode = "primaryobjecttypecode";

        public enum CrmPluginStepDeployment
        {
            ServerOnly = 0,
            OfflineOnly = 1,
            Both = 2
        }

        public enum CrmPluginStepMode
        {
            Asynchronous = 1,
            Synchronous = 0
        }

        public enum CrmPluginStepStage
        {
            PreValidation = 10,
            PreOperation = 20,
            PostOperation = 40,
        }

        public enum SdkMessageName
        {
            Create,
            Update,
            Delete,
            Retrieve,
            Assign,
            GrantAccess,
            ModifyAccess,
            RetrieveMultiple,
            RetrievePrincipalAccess,
            RetrieveSharedPrincipalsAndAccess,
            RevokeAccess,
            SetState,
            SetStateDynamicEntity,
        }

        public enum SdkImageType
        {
            PreImage,
            PostImage,
            Both
        }

        public enum SequenceType
        {
            Sequential = 281930000,
            Random = 281930001,
            Guid = 281930002,
            Base36 = 281930003
        }

        public static string GetMessageProperyName(SdkMessageName messageName)
        {
            switch (messageName)
            {
                case SdkMessageName.Assign:
                    return "Target";
                case SdkMessageName.Create:
                    return "Id";
                case SdkMessageName.Delete:
                    return "Target";
                case SdkMessageName.SetState:
                    return "EntityMoniker";
                case SdkMessageName.SetStateDynamicEntity:
                    return "EntityMoniker";
                case SdkMessageName.Update:
                    return "Target";
                default:
                    return string.Empty;
            }
        }
    }
}

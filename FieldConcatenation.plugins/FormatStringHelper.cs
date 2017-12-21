using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Veritec.Crm.FieldConcatenation.plugins
{
    internal static class FormatStringHelper
    {
        private const string tokenRegex = @"{(?:([a-zA-Z0-9_]+)(?::([a-zA-Z0-9_]+)){0,1})}";

        public static readonly List<AttributeTypeCode> SupportedAttributeTypes = new List<AttributeTypeCode>
        {
            AttributeTypeCode.Boolean,
            AttributeTypeCode.BigInt,
            AttributeTypeCode.DateTime,
            AttributeTypeCode.Decimal,
            AttributeTypeCode.Double,
            AttributeTypeCode.EntityName,
            AttributeTypeCode.Integer,
            AttributeTypeCode.Lookup,
            AttributeTypeCode.Memo,
            AttributeTypeCode.Money,
            AttributeTypeCode.Picklist,
            AttributeTypeCode.String,
            AttributeTypeCode.Uniqueidentifier
        };

        private static Func<Match, string> SubstitutionKey = (m) => m.Groups[0].Value;
        private static Func<Match, string> AttributeName = (m) => m.Groups[1].Value;
        private static Func<Match, string> AttributeFormatString = (m) => m.Groups[2].Success ? m.Groups[2].Value : null;


        public static IEnumerable<AttributeToken> TokeniseString(IEnumerable<AttributeMetadata> attributeMetadatas, string formatString)
        {
            var tokens = new HashSet<AttributeToken>();

            var regex = new Regex(tokenRegex);
            var match = regex.Match(formatString);

            while (match.Success)
            {
                var substitutionKey = SubstitutionKey(match);
                var attributeName = AttributeName(match);
                var attributeFormatString = AttributeFormatString(match);

                var token = new AttributeToken
                {
                    Name = attributeName,
                    Format = attributeFormatString
                };

                var attributeMetadata = attributeMetadatas.SingleOrDefault(m => m.LogicalName == token.Name);

                if (attributeMetadata != null)
                {
                    token.AttributeMetadata = attributeMetadata;
                    token.Status = attributeMetadata.AttributeType.HasValue && SupportedAttributeTypes.Contains(attributeMetadata.AttributeType.Value)
                        ? AttributeTokenStatus.Valid
                        : AttributeTokenStatus.UnsupportedType;
                }
                else
                {
                    token.Status = AttributeTokenStatus.UnknownAttribute;
                }

                tokens.Add(token);

                match = match.NextMatch();
            }

            return tokens;
        }

        public static string FormatString(IOrganizationService service, string formatString, IEnumerable<AttributeToken> tokens, Entity changeEntity, Entity preChangeEntity)
        {
            var regex = new Regex(tokenRegex);

            var result = regex.Replace(formatString, match =>
            {
                var token = tokens.Single(t => t.Name == AttributeName(match));
                return GetFormattedAttribute(service, token, changeEntity, preChangeEntity);
            });

            return result;
        }

        private static string GetFormattedAttribute(IOrganizationService service, AttributeToken token, Entity changeEntity, Entity preChangeEntity)
        {
            switch (token.AttributeMetadata.AttributeType)
            {
                case AttributeTypeCode.Boolean:
                    {
                        var booleanMetadata = (BooleanAttributeMetadata)token.AttributeMetadata;
                        var value = GetAttributeValue<bool?>(token.Name, changeEntity, preChangeEntity);
                        return value.HasValue ? GetBooleanString(booleanMetadata, value.Value) : "<<null>>";
                    }
                case AttributeTypeCode.BigInt:
                    {
                        var value = GetAttributeValue<long?>(token.Name, changeEntity, preChangeEntity);
                        return value.HasValue
                            ? (string.IsNullOrEmpty(token.Format) ? value.Value.ToString() : value.Value.ToString(token.Format))
                            : "<<null>>";
                    }
                case AttributeTypeCode.DateTime:
                    {
                        var value = GetAttributeValue<DateTime?>(token.Name, changeEntity, preChangeEntity);
                        return value.HasValue
                            ? (string.IsNullOrEmpty(token.Format) ? value.Value.ToString() : value.Value.ToString(token.Format))
                            : "<<null>>";
                    }
                case AttributeTypeCode.Decimal:
                    {
                        var value = GetAttributeValue<decimal?>(token.Name, changeEntity, preChangeEntity);
                        return value.HasValue
                            ? (string.IsNullOrEmpty(token.Format) ? value.Value.ToString() : value.Value.ToString(token.Format))
                            : "<<null>>";
                    }
                case AttributeTypeCode.Double:
                    {
                        var value = GetAttributeValue<double?>(token.Name, changeEntity, preChangeEntity);
                        return value.HasValue
                            ? (string.IsNullOrEmpty(token.Format) ? value.Value.ToString() : value.Value.ToString(token.Format))
                            : "<<null>>";
                    }
                case AttributeTypeCode.EntityName:
                    {
                        var value = GetAttributeValue<string>(token.Name, changeEntity, preChangeEntity);
                        return value != null ? value : "<<null>>";
                    }
                case AttributeTypeCode.Integer:
                    {
                        var value = GetAttributeValue<int?>(token.Name, changeEntity, preChangeEntity);
                        return value.HasValue
                            ? (string.IsNullOrEmpty(token.Format) ? value.Value.ToString() : value.Value.ToString(token.Format))
                            : "<<null>>";
                    }
                case AttributeTypeCode.Lookup:
                    {
                        var value = GetAttributeValue<EntityReference>(token.Name, changeEntity, preChangeEntity);
                        return value != null
                            ? (value.Name != null ? value.Name : GetEntityPrimaryNameAttributeValue(service, value))
                            : "<<null>>";
                    }
                case AttributeTypeCode.Memo:
                    {
                        var value = GetAttributeValue<string>(token.Name, changeEntity, preChangeEntity);
                        return value != null ? value : "<<null>>";
                    }
                case AttributeTypeCode.Money:
                    {
                        var value = GetAttributeValue<Money>(token.Name, changeEntity, preChangeEntity);
                        return value != null
                            ? (string.IsNullOrEmpty(token.Format) ? value.Value.ToString() : value.Value.ToString(token.Format))
                            : "<<null>>";
                    }
                case AttributeTypeCode.Picklist:
                    {
                        var pickListMetadata = (PicklistAttributeMetadata)token.AttributeMetadata;
                        var value = GetAttributeValue<OptionSetValue>(token.Name, changeEntity, preChangeEntity);
                        return value != null
                            ? GetPicklistString(pickListMetadata, value)
                            : "<<null>>";
                    }
                case AttributeTypeCode.String:
                    {
                        var value = GetAttributeValue<string>(token.Name, changeEntity, preChangeEntity);
                        return value != null ? value : "<<null>>";
                    }
                case AttributeTypeCode.Uniqueidentifier:
                    {
                        var value = GetAttributeValue<Guid?>(token.Name, changeEntity, preChangeEntity);
                        return value.HasValue ? value.ToString() : "<<null>>";
                    }
                default:
                    return "<<unsupported type>>";
            }
        }

        private static T GetAttributeValue<T>(string attributeName, Entity changeEntity, Entity preChangeEntity)
        {
            return changeEntity.Contains(attributeName)
                ? changeEntity.GetAttributeValue<T>(attributeName)
                : (preChangeEntity != null
                    ? preChangeEntity.GetAttributeValue<T>(attributeName)
                    : default(T));
        }

        private static string GetBooleanString(BooleanAttributeMetadata metadata, bool value)
        {
            return value
                ? metadata.OptionSet.TrueOption.Label.UserLocalizedLabel.Label
                : metadata.OptionSet.FalseOption.Label.UserLocalizedLabel.Label;
        }

        private static string GetPicklistString(PicklistAttributeMetadata metadata, OptionSetValue value)
        {
            int? rawValue = value != null
                   ? value.Value
                   : metadata.DefaultFormValue;

            return rawValue.HasValue
                ? metadata.OptionSet.Options.Single(o => o.Value.Value == rawValue.Value).Label.UserLocalizedLabel.Label
                : "<<null>>";
        }

        private static string GetEntityPrimaryNameAttributeValue(IOrganizationService service, EntityReference entityReference)
        {
            var metadataRequest = new RetrieveEntityRequest
            {
                EntityFilters = EntityFilters.Entity,
                LogicalName = entityReference.LogicalName
            };

            var metadataResponse = (RetrieveEntityResponse)service.Execute(metadataRequest);
            var metadata = metadataResponse.EntityMetadata;

            var otherEntity = service.Retrieve(entityReference.LogicalName, entityReference.Id, new ColumnSet(metadata.PrimaryNameAttribute));
            return GetAttributeValue<string>(metadata.PrimaryNameAttribute, otherEntity, null);
        }
    }
}

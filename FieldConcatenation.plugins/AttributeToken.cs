using Microsoft.Xrm.Sdk.Metadata;
using System;

namespace Veritec.Crm.FieldConcatenation.plugins
{
    internal class AttributeToken : IEquatable<AttributeToken>
    {
        public string Name { get; set; }

        public string Format { get; set; }

        public AttributeMetadata AttributeMetadata { get; set; }

        public AttributeTokenStatus Status { get; set; }

        public bool Equals(AttributeToken other)
        {
            if (other == null)
            {
                return false;
            }

            return Name == other.Name && Format == other.Format;
        }

        public override bool Equals(Object obj)
        {
            if (obj == null)
            {
                return false;
            }

            var other = obj as AttributeToken;

            if (other == null)
            {
                return false;
            }

            return Equals(other);
        }

        public override int GetHashCode()
        {
            return Format != null
                ? Name.GetHashCode() ^ Format.GetHashCode()
                : Name.GetHashCode();
        }
    }
}

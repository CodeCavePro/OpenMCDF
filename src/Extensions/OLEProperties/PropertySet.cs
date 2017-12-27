using System.Collections.Generic;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    public class PropertySet
    {
        public uint Size { get; set; }

        public uint NumProperties { get; set; }

        public List<PropertyIdentifierAndOffset> PropertyIdentifierAndOffsets { get; set; } = new List<PropertyIdentifierAndOffset>();

        public List<ITypedPropertyValue> Properties { get; set; } = new List<ITypedPropertyValue>();
    }
}
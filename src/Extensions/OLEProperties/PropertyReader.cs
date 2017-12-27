using System.Collections.Generic;
using System.IO;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    public enum Behavior
    {
        CaseSensitive,
        CaseInsensitive
    }

    public class PropertyContext
    {
        public int CodePage { get; set; }
        public Behavior Behavior { get; set; }
        public uint Locale { get; set; }
    }

    public enum PropertyDimensions
    {
        IsScalar,
        IsVector,
        IsArray
    }

    public class PropertyReader
    {
        private readonly PropertyContext _ctx = new PropertyContext();
        private readonly PropertyFactory _factory;

        public PropertyReader()
        {
            _factory = new PropertyFactory();
        }

        public List<ITypedPropertyValue> ReadProperty(PropertyIdentifiersSummaryInfo propertyIdentifier,
            BinaryReader br)
        {
            var res = new List<ITypedPropertyValue>();
            var dim = PropertyDimensions.IsScalar;

            var pVal = br.ReadUInt16();

            var vType = (VtPropertyType) (pVal & 0x00FF);

            if ((pVal & 0x1000) != 0)
                dim = PropertyDimensions.IsVector;
            else if ((pVal & 0x2000) != 0)
                dim = PropertyDimensions.IsArray;

            var isVariant = ((pVal & 0x00FF) == 0x000C);

            br.ReadUInt16(); // Ushort Padding

            // ReSharper disable once SwitchStatementMissingSomeCases
            switch (dim)
            {
                case PropertyDimensions.IsVector:

                    var vectorHeader = _factory.NewProperty(VtPropertyType.VtVectorHeader, _ctx);
                    vectorHeader.Read(br);

                    var nItems = (uint) vectorHeader.PropertyValue;

                    for (var i = 0; i < nItems; i++)
                    {
                        VtPropertyType vTypeItem;

                        if (isVariant)
                        {
                            var pValItem = br.ReadUInt16();
                            vTypeItem = (VtPropertyType) (pValItem & 0x00FF);
                            br.ReadUInt16(); // Ushort Padding
                        }
                        else
                        {
                            vTypeItem = vType;
                        }

                        var p = _factory.NewProperty(vTypeItem, _ctx);

                        p.Read(br);
                        res.Add(p);
                    }

                    break;

                //Scalar property
                default:
                    var pr = _factory.NewProperty(vType, _ctx);

                    pr.Read(br);

                    if (propertyIdentifier == PropertyIdentifiersSummaryInfo.CodePageString)
                    {
                        _ctx.CodePage = (short) pr.PropertyValue;
                    }

                    res.Add(pr);
                    break;
            }

            return res;
        }
    }
}
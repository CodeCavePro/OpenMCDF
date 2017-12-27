using System;
using System.IO;
using System.Text;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    internal class PropertyFactory
    {
        public ITypedPropertyValue NewProperty(VtPropertyType vType, PropertyContext ctx)
        {
            ITypedPropertyValue pr;

            switch (vType)
            {
                case VtPropertyType.VtI2:
                    pr = new VtI2Property(vType);
                    break;
                case VtPropertyType.VtI4:
                    pr = new VtI4Property(vType);
                    break;
                case VtPropertyType.VtR4:
                    pr = new VtR4Property(vType);
                    break;
                case VtPropertyType.VtLpstr:
                    pr = new VtLpstrProperty(vType, ctx.CodePage);
                    break;
                case VtPropertyType.VtFiletime:
                    pr = new VtFiletimeProperty(vType);
                    break;
                case VtPropertyType.VtDecimal:
                    pr = new VtDecimalProperty(vType);
                    break;
                case VtPropertyType.VtBool:
                    pr = new VtBoolProperty(vType);
                    break;
                case VtPropertyType.VtVectorHeader:
                    pr = new VtVectorHeader(vType);
                    break;
                case VtPropertyType.VtEmpty:
                    pr = new VtEmptyProperty(vType);
                    break;
                default:
                    throw new Exception("Unrecognized property type");
            }

            return pr;
        }


        #region Property implementations

        internal class VtEmptyProperty : TypedPropertyValue
        {
            public VtEmptyProperty(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = null;
            }

            public override void Write(BinaryWriter bw)
            {
            }
        }

        internal class VtI2Property : TypedPropertyValue
        {
            public VtI2Property(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = br.ReadInt16();
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write((short) propertyValue);
            }
        }

        internal class VtI4Property : TypedPropertyValue
        {
            public VtI4Property(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = br.ReadInt32();
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write((int) propertyValue);
            }
        }

        internal class VtR4Property : TypedPropertyValue
        {
            public VtR4Property(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = br.ReadSingle();
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write((float) propertyValue);
            }
        }

        internal class VtR8Property : TypedPropertyValue
        {
            public VtR8Property(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = br.ReadDouble();
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write((double) propertyValue);
            }
        }

        internal class VtCyProperty : TypedPropertyValue
        {
            public VtCyProperty(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = br.ReadInt64() / 10000;
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write((long) propertyValue * 10000);
            }
        }

        internal class VtDateProperty : TypedPropertyValue
        {
            public VtDateProperty(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                var temp = br.ReadDouble();

#if NETSTANDARD1_6
                propertyValue = DateTimeExtensions.FromOADate(temp);
#else
                propertyValue = DateTime.FromOADate(temp);
#endif
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write(((DateTime) propertyValue).ToOADate());
            }
        }

        internal class VtLpstrProperty : TypedPropertyValue
        {
            private uint _size;
            private byte[] _data;
            private readonly int _codePage;

            public VtLpstrProperty(VtPropertyType vType, int codePage) : base(vType)
            {
                _codePage = codePage;
            }

            public override void Read(BinaryReader br)
            {
                _size = br.ReadUInt32();
                _data = br.ReadBytes((int) _size);
                propertyValue = Encoding.GetEncoding(_codePage).GetString(_data);
                var m = (int) _size % 4;
                br.ReadBytes(m); // padding
            }

            public override void Write(BinaryWriter bw)
            {
                _data = Encoding.GetEncoding(_codePage).GetBytes((string) propertyValue);
                _size = (uint) _data.Length;
                var m = (int) _size % 4;
                bw.Write(_data);
                for (var i = 0; i < m; i++) // padding
                    bw.Write(0);
            }
        }

        internal class VtFiletimeProperty : TypedPropertyValue
        {
            public VtFiletimeProperty(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                var tmp = br.ReadInt64();
                propertyValue = DateTime.FromFileTime(tmp);
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write(((DateTime) propertyValue).ToFileTime());
            }
        }

        internal class VtDecimalProperty : TypedPropertyValue
        {
            public VtDecimalProperty(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                decimal d;

                br.ReadInt16(); // wReserved
                var scale = br.ReadByte();
                var sign = br.ReadByte();

                var u = br.ReadUInt32();
                d = Convert.ToDecimal(Math.Pow(2, 64)) * u;
                d += br.ReadUInt64();

                if (sign != 0)
                    d = -d;
                d /= (10 << scale);

                propertyValue = d;
            }

            public override void Write(BinaryWriter bw)
            {
                bw.Write((short) propertyValue);
            }
        }

        internal class VtBoolProperty : TypedPropertyValue
        {
            public VtBoolProperty(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = br.ReadUInt16() == 0xFFFF;
                //br.ReadUInt16();//padding
            }
        }

        internal class VtVectorHeader : TypedPropertyValue
        {
            public VtVectorHeader(VtPropertyType vType) : base(vType)
            {
            }

            public override void Read(BinaryReader br)
            {
                propertyValue = br.ReadUInt32();
            }
        }

        #endregion
    }
}
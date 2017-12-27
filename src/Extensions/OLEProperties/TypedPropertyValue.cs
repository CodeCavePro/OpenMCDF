using System;
using OpenMcdf.Extensions.OLEProperties.Interfaces;

namespace OpenMcdf.Extensions.OLEProperties
{
    public class TypedPropertyValue : ITypedPropertyValue
    {
        public VtPropertyType VtType
        {
            get;
            //set { _VTType = value; }
        }

        protected object propertyValue;

        public TypedPropertyValue(VtPropertyType vtType)
        {
            VtType = vtType;
        }

        public virtual object PropertyValue
        {
            get => propertyValue;

            set => propertyValue = value;
        }


        public bool IsArray
        {
            get => throw new NotImplementedException();

            set => throw new NotImplementedException();
        }

        public bool IsVector
        {
            get => throw new NotImplementedException();

            set => throw new NotImplementedException();
        }

        public virtual void Read(System.IO.BinaryReader br)
        {
        }

        public virtual void Write(System.IO.BinaryWriter bw)
        {
        }
    }
}
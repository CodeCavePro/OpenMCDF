using System;

namespace OpenMcdf.Extensions.OLEProperties
{
    public class PropertySetStream
    {
        public ushort ByteOrder { get; set; }
        public ushort Version { get; set; }
        public uint SystemIdentifier { get; set; }
        public Guid CLSID { get; set; }
        public uint NumPropertySets { get; set; }
        public Guid Fmtid0 { get; set; }
        public uint Offset0 { get; set; }
        public Guid Fmtid1 { get; set; }
        public uint Offset1 { get; set; }
        public PropertySet PropertySet0 { get; set; }
        public PropertySet PropertySet1 { get; set; }

        public void Read(System.IO.BinaryReader br)
        {
            ByteOrder = br.ReadUInt16();
            Version = br.ReadUInt16();
            SystemIdentifier = br.ReadUInt32();
            CLSID = new Guid(br.ReadBytes(16));
            NumPropertySets = br.ReadUInt32();
            Fmtid0 = new Guid(br.ReadBytes(16));
            Offset0 = br.ReadUInt32();

            if (NumPropertySets == 2)
            {
                Fmtid1 = new Guid(br.ReadBytes(16));
                Offset1 = br.ReadUInt32();
            }

            PropertySet0 = new PropertySet
            {
                Size = br.ReadUInt32(),
                NumProperties = br.ReadUInt32()
            };

            // Read property offsets
            for (var i = 0; i < PropertySet0.NumProperties; i++)
            {
                var pio = new PropertyIdentifierAndOffset();
                pio.PropertyIdentifier = (PropertyIdentifiersSummaryInfo) br.ReadUInt32();
                pio.Offset = br.ReadUInt32();
                PropertySet0.PropertyIdentifierAndOffsets.Add(pio);
            }

            // Read properties
            var pr = new PropertyReader();
            for (var i = 0; i < PropertySet0.NumProperties; i++)
            {
                br.BaseStream.Seek(Offset0 + PropertySet0.PropertyIdentifierAndOffsets[i].Offset,
                    System.IO.SeekOrigin.Begin);
                PropertySet0.Properties.AddRange(
                    pr.ReadProperty(PropertySet0.PropertyIdentifierAndOffsets[i].PropertyIdentifier, br));
            }
        }

        public void Write(System.IO.BinaryWriter bw)
        {
            throw new NotImplementedException();
        }

        //        private void LoadFromStream(Stream inStream)
        //        {
        //            BinaryReader br = new BinaryReader(inStream);
        //            PropertySetStream psStream = new PropertySetStream();
        //            psStream.Read(br);
        //            br.Close();

        //            propertySets.Clear();

        //            if (psStream.NumPropertySets == 1)
        //            {
        //                propertySets.Add(psStream.PropertySet0);
        //            }
        //            else
        //            {
        //                propertySets.Add(psStream.PropertySet0);
        //                propertySets.Add(psStream.PropertySet1);
        //            }

        //            return;
        //        }
    }
}
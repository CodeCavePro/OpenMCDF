namespace OpenMcdf.Extensions.OLEProperties
{
    public enum VtPropertyType : ushort
    {
        VtEmpty = 0x0000,
        VtNull = 0x0001,
        VtI2 = 0x0002,
        VtI4 = 0x0003,
        VtR4 = 0x0004,
        VtR8 = 0x0005,
        VtCy = 0x0006,
        VtDate = 0x0007,
        VtBstr = 0x0008,
        VtError = 0x000A,
        VtBool = 0x000B,
        VtDecimal = 0x000E,
        VtI1 = 0x0010,
        VtUi1 = 0x0011,
        VtUi2 = 0x0012,
        VtUi4 = 0x0013,
        VtI8 = 0x0014, // MUST be an 8-byte signed integer. 
        VtUi8 = 0x0015, // MUST be an 8-byte unsigned integer. 
        VtInt = 0x0016, // MUST be a 4-byte signed integer. 
        VtUint = 0x0017, // MUST be a 4-byte unsigned integer. 
        VtLpstr = 0x001E, // MUST be a CodePageString. 
        VtLpwstr = 0x001F, // MUST be a UnicodeString. 
        VtFiletime = 0x0040, // MUST be a FILETIME (Packet Version). 
        VtBlob = 0x0041, // MUST be a BLOB. 

        VtStream =
            0x0042, // MUST be an IndirectPropertyName. The storage representing the (non-simple) property set MUST have a stream element with this name. 

        VtStorage =
            0x0043, // MUST be an IndirectPropertyName. The storage representing the (non-simple) property set MUST have a storage element with this name. 

        VtStreamedObject =
            0x0044, // MUST be an IndirectPropertyName. The storage representing the (non-simple) property set MUST have a stream element with this name. 

        VtStoredObject =
            0x0045, // MUST be an IndirectPropertyName. The storage representing the (non-simple) property set MUST have a storage element with this name. 
        VtBlobObject = 0x0046, //MUST be a BLOB. 
        VtCF = 0x0047, //MUST be a ClipboardData. 
        VtCLSID = 0x0048, //MUST be a GUID (Packet Version)
        VtVersionedStream = 0x0049, //MUST be a Versioned Stream, NOT allowed in simple property
        VtVectorHeader = 0x1000, //--- NOT NORMATIVE
        VtArrayHeader = 0x2000, //--- NOT NORMATIVE
    }
}
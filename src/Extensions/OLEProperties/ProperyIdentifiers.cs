namespace OpenMcdf.Extensions.OLEProperties
{
    public enum PropertyIdentifiersSummaryInfo : uint
    {
        CodePageString = 0x00000001,
        PidsiTitle = 0x00000002,
        PidsiSubject = 0x00000003,
        PidsiAuthor = 0x00000004,
        PidsiKeywords = 0x00000005,
        PidsiComments = 0x00000006,
        PidsiTemplate = 0x00000007,
        PidsiLastauthor = 0x00000008,
        PidsiRevnumber = 0x00000009,
        PidsiAppname = 0x00000012,
        PidsiEdittime = 0x0000000A,
        PidsiLastprinted = 0x0000000B,
        PidsiCreateDtm = 0x0000000C,
        PidsiLastsaveDtm = 0x0000000D,
        PidsiPagecount = 0x0000000E,
        PidsiWordcount = 0x0000000F,
        PidsiCharcount = 0x00000010,
        PidsiDocSecurity = 0x00000013
    }

    public enum PropertyIdentifiersDocumentSummaryInfo : uint
    {
        CodePageString = 0x00000001,
        PiddsiCategory = 0x00000002, //Category VT_LPSTR
        PiddsiPresformat = 0x00000003, //PresentationTarget	VT_LPSTR
        PiddsiBytecount = 0x00000004, //Bytes   	VT_I4
        PiddsiLinecount = 0x00000005, // Lines   	VT_I4
        PiddsiParcount = 0x00000006, // Paragraphs 	VT_I4
        PiddsiSlidecount = 0x00000007, // Slides 	VT_I4
        PiddsiNotecount = 0x00000008, // Notes  	VT_I4
        PiddsiHiddencount = 0x00000009, // HiddenSlides   	VT_I4
        PiddsiMmclipcount = 0x0000000A, // MMClips	VT_I4
        PiddsiScale = 0x0000000B, //ScaleCrop  VT_BOOL
        PiddsiHeadingpair = 0x0000000C, // HeadingPairs VT_VARIANT | VT_VECTOR
        PiddsiDocparts = 0x0000000D, //TitlesofParts   	VT_VECTOR | VT_LPSTR
        PiddsiManager = 0x0000000E, //	  Manager VT_LPSTR
        PiddsiCompany = 0x0000000F, // Company	VT_LPSTR
        PiddsiLinksdirty = 0x00000010, //LinksUpToDate   	VT_BOOL
    }

    public static class Extensions
    {
        public static string GetDescription(this PropertyIdentifiersSummaryInfo identifier)
        {
            switch (identifier)
            {
                case PropertyIdentifiersSummaryInfo.CodePageString:
                    return "CodePage";
                case PropertyIdentifiersSummaryInfo.PidsiTitle:
                    return "Title";
                case PropertyIdentifiersSummaryInfo.PidsiSubject:
                    return "Subject";
                case PropertyIdentifiersSummaryInfo.PidsiAuthor:
                    return "Author";
                case PropertyIdentifiersSummaryInfo.PidsiLastauthor:
                    return "Last Author";
                case PropertyIdentifiersSummaryInfo.PidsiAppname:
                    return "Application Name";
                case PropertyIdentifiersSummaryInfo.PidsiCreateDtm:
                    return "Create Time";
                case PropertyIdentifiersSummaryInfo.PidsiLastsaveDtm:
                    return "Last Modified Time";
                case PropertyIdentifiersSummaryInfo.PidsiKeywords:
                    return "Keywords";
                case PropertyIdentifiersSummaryInfo.PidsiDocSecurity:
                    return "Document Security";
                default: return string.Empty;
            }
        }

        public static string GetDescription(this PropertyIdentifiersDocumentSummaryInfo identifier)
        {
            switch (identifier)
            {
                case PropertyIdentifiersDocumentSummaryInfo.CodePageString:
                    return "CodePage";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiCategory:
                    return "Category";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiCompany:
                    return "Company";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiDocparts:
                    return "Titles of Parts";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiHeadingpair:
                    return "Heading Pairs";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiHiddencount:
                    return "Hidden Slides";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiLinecount:
                    return "Line Count";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiLinksdirty:
                    return "Links up to date";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiManager:
                    return "Manager";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiMmclipcount:
                    return "MMClips";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiNotecount:
                    return "Notes";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiParcount:
                    return "Paragraphs";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiPresformat:
                    return "Presenteation Target";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiScale:
                    return "Scale";
                case PropertyIdentifiersDocumentSummaryInfo.PiddsiSlidecount:
                    return "Slides";
                default: return string.Empty;
            }
        }
    }
}
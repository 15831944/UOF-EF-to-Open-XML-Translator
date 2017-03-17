using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Act.UofTranslator.TranslatorMgr
{
    /// <summary>
    /// OOX file type
    /// </summary>
    /// <author>linyaohu</author>
    public enum MSDocType : int
    {
        StrictWord,
        StrictExcel,
        StrictPowerpoint,
        TransitionalWord,
        TransitionalExcel,
        TransitionalPowerpoint,
        Unknown
    }

    class TranslatorMgrConstants
    {
        public const  string WordDocument_xml = "word/document.xml";
        public const string PresentationDocument_xml = "ppt/presentation.xml";
        public const string SpreadsheetDocument_xml = "xl/workbook.xml";

        public const string XMLNS_W = "http://purl.oclc.org/ooxml/wordprocessingml/main";
        public const string XMLNS_P = "http://purl.oclc.org/ooxml/presentationml/main";
        public const string XMLNS_WS = "http://purl.oclc.org/ooxml/spreadsheetml/main";

    }
}

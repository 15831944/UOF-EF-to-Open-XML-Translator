using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Act.UofTranslator.UofZipUtils;
using System.Xml;
using Act.UofTranslator.UofTranslatorLib;
using Act.UofTranslator.UofTranslatorStrictLib;

namespace Act.UofTranslator.TranslatorMgr
{
    public class TranslatorMgr
    {
        /// <summary>
        /// Check the file type of input file and initialize
        /// </summary>
        /// <param name="fileName">string of input file name</param>
        /// <returns>return Transitional IUOFTranslator</returns>
        public static UofTranslatorLib.IUOFTranslator CheckFileTypeAndInitializeTransitional(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            if (File.Exists(fileName))
            {
                switch (fi.Extension.ToLower())
                {
                    // OOX file
                    case ".pptx": return new UofTranslatorLib.PresentationTranslator();
                    case ".docx": return new UofTranslatorLib.WordTranslator();
                    case ".xlsx": return new UofTranslatorLib.SpreadsheetTranslator();

                    // uof file
                    case ".uot": return new UofTranslatorLib.WordTranslator();
                    case ".uop": return new UofTranslatorLib.PresentationTranslator();
                    case ".uos": return new UofTranslatorLib.SpreadsheetTranslator();
                    default: throw new Exception("not an microsoft office 2010/2013 file or a UOF 2.0 file");
                }
            }
            else { throw new FileNotFoundException(); }
        }

        /// <summary>
        /// Check the file type of input file and initialize
        /// </summary>
        /// <param name="fileName">string of input file name</param>
        /// <returns>return Strict IUOFTranslator</returns>
        public static UofTranslatorStrictLib.IUOFTranslator CheckFileTypeAndInitializeStrict(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            if (File.Exists(fileName))
            {
                switch (fi.Extension.ToLower())
                {
                    // OOX file
                    case ".pptx": return new UofTranslatorStrictLib.PresentationTranslator();
                    case ".docx": return new UofTranslatorStrictLib.WordTranslator();
                    case ".xlsx": return new UofTranslatorStrictLib.SpreadsheetTranslator();

                    // uof file
                    case ".uot": return new UofTranslatorStrictLib.WordTranslator();
                    case ".uop": return new UofTranslatorStrictLib.PresentationTranslator();
                    case ".uos": return new UofTranslatorStrictLib.SpreadsheetTranslator();
                    default: throw new Exception("not an microsoft office 2010/2013 file or a UOF 2.0 file");
                }
            }
            else { throw new FileNotFoundException(); }
        }

        ///// <summary>
        ///// Initialize according to the input file type
        ///// </summary>
        ///// <param name="inputFileType">DocType of input file</param>
        ///// <returns>return IUOFTranslator</returns>
        //public static IUOFTranslator CheckFileType(DocType inputFileType)
        //{
        //    switch (inputFileType)
        //    {
        //        case DocType.Word: return new WordTranslator();
        //        case DocType.Powerpoint: return new PresentationTranslator();
        //        case DocType.Excel: return new SpreadsheetTranslator();
        //        default: throw new Exception("not an office 2010 file");
        //    }
        //}

        /// <summary>
        /// check Microsoft file type
        /// </summary>
        /// <param name="srcFileName">source file name</param>
        /// <returns>document type</returns>
        public static MSDocType CheckMSFileType(string srcFileName)
        {
            FileInfo fi = new FileInfo(srcFileName);
            if (File.Exists(srcFileName))
            {
                XmlReader source = null;
                ZipReader archive = ZipFactory.OpenArchive(srcFileName);

                // get the main entry xml file
                string entry = string.Empty;
                switch (fi.Extension.ToLower())
                {
                    case ".docx": entry = TranslatorMgrConstants.WordDocument_xml; break;
                    case ".pptx": entry = TranslatorMgrConstants.PresentationDocument_xml; break;
                    case ".xlsx": entry = TranslatorMgrConstants.SpreadsheetDocument_xml; break;
                }
                source = XmlReader.Create(archive.GetEntry(entry));

                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(source);

                // namespace of strict document
                XmlNamespaceManager nm=new XmlNamespaceManager(xdoc.NameTable);
                nm.AddNamespace("w",TranslatorMgrConstants.XMLNS_W);
                nm.AddNamespace("p",TranslatorMgrConstants.XMLNS_P);
                nm.AddNamespace("ws",TranslatorMgrConstants.XMLNS_WS);

                switch (fi.Extension.ToLower())
                {
                    case ".docx":
                        {
                            // @w:conformance
                            if (xdoc.SelectSingleNode("w:document", nm) != null)
                            {
                                return MSDocType.StrictWord;
                            }
                            else
                            {
                                return MSDocType.TransitionalWord;
                            }
                        }
                    case ".pptx":
                        {
                            if (xdoc.SelectSingleNode("p:presentation", nm) != null)
                            {
                                return MSDocType.StrictPowerpoint;
                            }
                            else
                            {
                                return MSDocType.TransitionalPowerpoint;
                            }
                        }
                    case ".xlsx":
                        {
                            if (xdoc.SelectSingleNode("ws:workbook", nm) != null)
                            {
                                return MSDocType.StrictExcel;
                            }
                            else
                            {
                                return MSDocType.TransitionalExcel;
                            }
                        }
                    default: return MSDocType.Unknown;
                }

            }
            else { return MSDocType.Unknown; }
        }
    }

}

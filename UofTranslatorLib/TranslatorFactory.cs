using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace Act.UofTranslator.UofTranslatorLib
{

    /// <summary>
    /// OOX file type
    /// </summary>
    /// <author>linyaohu</author>
    public enum DocType : int
    {
        Word,
        Excel,
        Powerpoint,
        Unknown
    }

    /// <summary>
    /// Translator factory which can judeg the input file type and initialize
    /// </summary>
    /// <author>linyaohu</author>
    public class TranslatorFactory
    {
        /// <summary>
        /// Check the file type of input file and initialize
        /// </summary>
        /// <param name="fileName">string of input file name</param>
        /// <returns>return IUOFTranslator</returns>
        public static IUOFTranslator CheckFileTypeAndInitialize(string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            if (File.Exists(fileName))
            {
                switch (fi.Extension.ToLower())
                {
                    // OOX file
                    case ".pptx": return new PresentationTranslator();
                    case ".docx": return new WordTranslator();
                    case ".xlsx": return new SpreadsheetTranslator();

                    // uof file
                    case ".uot": return new WordTranslator();
                    case ".uop": return new PresentationTranslator();
                    case ".uos": return new SpreadsheetTranslator();
                    default: throw new Exception("not an office 2010 file or a UOF 2.0 file");
                }
            }
            else { throw new FileNotFoundException(); }
        }

        /// <summary>
        /// Initialize according to the input file type
        /// </summary>
        /// <param name="inputFileType">DocType of input file</param>
        /// <returns>return IUOFTranslator</returns>
        public static IUOFTranslator CheckFileType(DocType inputFileType)
        {
            switch (inputFileType)
            {
                case DocType.Word: return new WordTranslator();
                case DocType.Powerpoint: return new PresentationTranslator();
                case DocType.Excel: return new SpreadsheetTranslator();
                default: throw new Exception("not an office 2010 file");
            }
        }

        /// <summary>
        /// check uof file type
        /// </summary>
        /// <param name="srcFileName">source file name</param>
        /// <returns>document type</returns>
        public static DocType CheckUOFFileType(string srcFileName)
        {
            FileInfo fi = new FileInfo(srcFileName);
            if (File.Exists(srcFileName))
            {
                switch (fi.Extension.ToLower())
                {
                    case ".uot": return DocType.Word;
                    case ".uop": return DocType.Powerpoint;
                    case ".uos": return DocType.Excel;
                    default: return DocType.Unknown;
                }
            }
            else { return DocType.Unknown; }
        }
    }
}

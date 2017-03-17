using System;
using System.IO;
using Extensibility;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Resources;
using Microsoft.Office.Core;
using System.Xml;
using System.Globalization;
using System.Threading;
using MSword = Microsoft.Office.Interop.Word;
using UofAddinLib;
using stdole;
using Act.UofTranslator.UofTranslatorLib;
using Act.UofTranslator.UofTranslatorStrictLib;

namespace UofTranslatorShell
{

    /// <summary>
    ///     Module ID:      
    ///     Depiction:          UofTranslator Shell
    ///     Author:             朴鑫 SYBASE v-xipia 
    ///     Create Date:        2009-02-05
    ///     Modified by:        
    ///     Last Modified Date: 2009-06-04
    /// </summary>
    class UofTranslatorShellOperation
    {
        /// <summary>
        ///     The Application(Microsoft Word)
        /// </summary>
      //  private MSword.Application applicationObject;

        /// <summary>
        ///     The Culture that this application Use
        /// </summary>
       // private CultureInfo applicationCulture;

        /// <summary>
        ///     The ResourceManager for the Translator
        /// </summary>
       // private ResourceManager translatorResMan;

        static void Main(string[] args)
        {
            
            string FileNameToOpen = args[0];
            string OpenFileType = string.Empty;
            string TempFile = string.Empty;
            System.IO.FileInfo FileInfoToOpen = new System.IO.FileInfo(FileNameToOpen);


            try
            {
                try
                {
                    string srcPath = args[0];
                    System.IO.FileInfo fi = new System.IO.FileInfo(srcPath);


                   // System.Windows.Forms.MessageBox.Show("srcPath=" + srcPath+"\n"+"FileNameToOpen="+FileNameToOpen);
                    
                    OpenFileType = getFileTypeFromPath(srcPath);
                    OpenFileByOffice(OpenFileType, FileNameToOpen);


                }
                catch (System.IO.IOException e)
                {
                    TempFile = getTempFilePath(FileNameToOpen);
                    File.Copy(FileNameToOpen, TempFile);

                    System.IO.FileInfo FileInfoTempFile = new System.IO.FileInfo(TempFile);
                    FileInfoTempFile.IsReadOnly = true;

                    OpenFileType = getFileTypeFromPath(TempFile);
                    OpenFileByOffice(OpenFileType, FileNameToOpen);

                    throw e;
                }



            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
                System.Windows.Forms.MessageBox.Show("This file is not a standard UOF file.");

            }
        }
            
                     
           
            
   

        //get temp file name according to the file path: XXX\filename(i).uof
        private static string getTempFilePath(string filePath)
        {
            string ResultTempFilePath;

            string TempFilePath = System.IO.Path.GetTempPath();
            System.IO.FileInfo FileInfoFilePath = new System.IO.FileInfo(filePath);
            //get the file name without extension
            string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath);


            int i = 0;
            //create a temp file in the temp folder
            while (File.Exists(ResultTempFilePath = TempFilePath + fileName + "(" + i + ")" + FileInfoFilePath.Extension))
            {
                i++;
            }
            //File.Copy(@"D:\testdfd.docx", ResultTempFilePath, true);

            return ResultTempFilePath;

        }

        private static void OpenFileByOffice(string OpenFileType, string FileNameToOpen)
        {

            string procStartInfoArguments = "\"" + FileNameToOpen + "\"";

            UofAddinLib.ConfigManager cfg = new UofAddinLib.ConfigManager(System.IO.Path.GetDirectoryName(typeof(MainDialog).Assembly.Location) + @"\conf\config.xml");
            cfg.LoadConfig();

            //try
            //{
                switch (OpenFileType.ToLower())
                {
                    case "docx":
                   // case "vnd.uof.text":
                    case "uot":
                        {
                          //  System.Windows.Forms.MessageBox.Show("im here");

                            string tmpFile = Path.GetTempFileName();
                            string tmpResult = tmpFile.Substring(0, tmpFile.LastIndexOf(".")) + ".docx";

                            //System.Windows.Forms.MessageBox.Show("tmpFile=" + tmpResult);
                            if (cfg.IsUofToTransitioanlOOX)
                            {
                                Act.UofTranslator.UofTranslatorLib.IUOFTranslator trans = new Act.UofTranslator.UofTranslatorLib.WordTranslator();
                                trans.UofToOox(FileNameToOpen, tmpResult);
                            }
                            else
                            {
                                Act.UofTranslator.UofTranslatorStrictLib.IUOFTranslator trans = new Act.UofTranslator.UofTranslatorStrictLib.WordTranslator();
                                trans.UofToOox(FileNameToOpen, tmpResult);
                            }
                            System.Diagnostics.Process proc = new System.Diagnostics.Process();
                            proc.EnableRaisingEvents = false;
                            proc.StartInfo.FileName = "WINWORD";
                            // proc.StartInfo.FileName = procStartInfoArguments;
                            proc.StartInfo.Arguments = "\"" + tmpResult + "\"";
                            proc.Start();
                          //  System.Diagnostics.Process.Start(typeof(UofTranslatorShell.UofTranslatorShellOperation).Assembly.Location, "\"" + FileNameToOpen + "\"");

                            //OpenByWord(FileNameToOpen);

                            break;
                        }
                    case "xlsx":
                   // case "vnd.uof.spreadsheet":
                    case "uos":
                        {
                            string tmpFile = Path.GetTempFileName();
                            string tmpResult = tmpFile.Substring(0, tmpFile.LastIndexOf("."))+".xlsx";

                           // System.Windows.Forms.MessageBox.Show("tmpFile="+tmpResult);
                            System.Diagnostics.Process proc2 = new System.Diagnostics.Process();
                            proc2.EnableRaisingEvents = false;
                            proc2.StartInfo.FileName = "EXCEL";

                            if (cfg.IsUofToTransitioanlOOX)
                            {
                                Act.UofTranslator.UofTranslatorLib.IUOFTranslator trans = new Act.UofTranslator.UofTranslatorLib.SpreadsheetTranslator();
                                trans.UofToOox(FileNameToOpen, tmpResult);
                            }
                            else
                            {
                                Act.UofTranslator.UofTranslatorStrictLib.IUOFTranslator trans = new Act.UofTranslator.UofTranslatorStrictLib.SpreadsheetTranslator();
                                trans.UofToOox(FileNameToOpen, tmpResult);
                            }

                            proc2.StartInfo.Arguments = "\"" + tmpResult + "\"";
                            //proc2.StartInfo.Arguments = procStartInfoArguments;
                            proc2.Start();
                            break;
                        }
                    case "pptx":
                   // case "vnd.uof.presentation":
                    case "uop":
                        {
                            
                            System.Diagnostics.Process proc3 = new System.Diagnostics.Process();
                            proc3.EnableRaisingEvents = false;
                            proc3.StartInfo.FileName = "POWERPNT";
                            string tmpFile = Path.GetTempFileName();
                            string tmpResult = tmpFile.Substring(0, tmpFile.LastIndexOf(".")) + ".pptx";

                            if (cfg.IsUofToTransitioanlOOX)
                            {
                                Act.UofTranslator.UofTranslatorLib.IUOFTranslator trans = new Act.UofTranslator.UofTranslatorLib.PresentationTranslator();
                                trans.UofToOox(FileNameToOpen, tmpResult);
                            }
                            else
                            {
                                Act.UofTranslator.UofTranslatorStrictLib.IUOFTranslator trans = new Act.UofTranslator.UofTranslatorStrictLib.PresentationTranslator();
                                trans.UofToOox(FileNameToOpen, tmpResult);
                            }
                            //System.Windows.Forms.MessageBox.Show("tmpFile=" + tmpResult);
                           
                            proc3.StartInfo.Arguments = "\"" + tmpResult + "\"";
                            proc3.Start();

                            break;
                        }

                    default:
                        {
                            System.Windows.Forms.MessageBox.Show(OpenFileType);
                            break;
                        }

                }
        }


        private static string getFileTypeFromPath(string fileName)
        {
            System.IO.FileInfo fi = new System.IO.FileInfo(fileName);
            if (File.Exists(fileName))
            {
                switch (fi.Extension.ToLower())
                {
                    //// OOX file
                    case ".pptx": return "pptx";
                    case ".docx": return "docx";
                    case ".xlsx": return "xlsx";

                    // uof file
                    //case ".uof":
                    //    {
                    //        string doctype;
                    //        XmlDocument document = new XmlDocument();
                    //        document.Load(fileName);
                    //        XmlNamespaceManager nm = new XmlNamespaceManager(document.NameTable);
                    //        nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
                    //        try
                    //        {
                    //            doctype = document.DocumentElement.Attributes.GetNamedItem("uof:mimetype").Value.ToString();
                    //        }
                    //        catch (NullReferenceException e)
                    //        {
                    //            throw new NullReferenceException("not a standard .uof file");
                    //        }
                    //        switch (document.DocumentElement.Attributes.GetNamedItem("uof:mimetype").Value.ToString())
                    //        {
                    //            case "vnd.uof.presentation": return "vnd.uof.presentation";
                    //            case "vnd.uof.spreadsheet": return "vnd.uof.spreadsheet";
                    //            case "vnd.uof.text": return "vnd.uof.text";
                    //            default: throw new Exception("not a standard .uof file");
                    //        }
                    //    }
                    //default: return "vnd.uof.text";
                        
                        // UOF2.0
                    case ".uot": return "uot";
                    case ".uop": return "uop";
                    case ".uos": return "uos";

                    default: throw new Exception("not an office 2010/2013 file or a UOF2.0 file");
                }
            }
            else { throw new FileNotFoundException(); }
        }

        //private static void OpenByWord(string FileName)
        //{
        //    Microsoft.Office.Interop.Word.Application myWordApp = new Microsoft.Office.Interop.Word.ApplicationClass();
        //    Microsoft.Office.Interop.Word.Documents Documentstest = myWordApp.Documents;
        //    //Console.WriteLine("Documentstest.Count: " + Documentstest.Count);
        //    //Microsoft.Office.Interop.Word._Document tempDoc;
        //    myWordApp.Visible = true;


        //    Object filename = FileName;
        //    Object ConfirmConversions = false;
        //    //Object ReadOnly = true;
        //    Object ReadOnly = System.Type.Missing;
        //    //Object AddToRecentFiles = false;
        //    Object AddToRecentFiles = System.Type.Missing;
        //    Object PasswordDocument = System.Type.Missing;
        //    Object PasswordTemplate = System.Type.Missing;
        //    Object Revert = System.Type.Missing;
        //    Object WritePasswordDocument = System.Type.Missing;
        //    Object WritePasswordTemplate = System.Type.Missing;
        //    Object Format = System.Type.Missing;
        //    Object Encoding = System.Type.Missing;
        //    Object Visible = System.Type.Missing;
        //    Object OpenAndRepair = System.Type.Missing;
        //    Object DocumentDirection = System.Type.Missing;
        //    Object NoEncodingDialog = System.Type.Missing;
        //    Object XMLTransform = System.Type.Missing;


        //    Word.Document wordDoc = myWordApp.Documents.Open(ref filename, ref ConfirmConversions,
        //        ref ReadOnly, ref AddToRecentFiles, ref PasswordDocument, ref PasswordTemplate,
        //        ref Revert, ref WritePasswordDocument, ref WritePasswordTemplate, ref Format,
        //        ref Encoding, ref Visible, ref OpenAndRepair, ref DocumentDirection,
        //        ref NoEncodingDialog, ref XMLTransform);
        //    wordDoc.Activate();

        //}



        //private static string getFileTypeFromPath(string fileName)
        //{

        //    XmlReader rdr = XmlReader.Create(fileName);
        //    bool FindUofFormat = false;
        //    while (rdr.Read())
        //    {
        //        if (rdr.NodeType == XmlNodeType.Element)
        //        {
        //            Console.WriteLine(rdr.Name);

        //            if (rdr.Name == "uof:文字处理" || rdr.Name == "UOF:文字处理")
        //            {
        //                //Console.WriteLine("找到了文字处理");
        //                if (rdr.GetAttribute(0).ToString() == "u0047")
        //                {
        //                    //Console.WriteLine("Got it");
        //                    FindUofFormat = true;
        //                    return "vnd.uof.text";
        //                }
        //            }
        //            if (rdr.Name == "uof:演示文稿" || rdr.Name == "UOF:演示文稿")
        //            {

        //                //Console.WriteLine("找到了演示文稿");
        //                //FindUofFormat = true;
        //                if (rdr.GetAttribute(0).ToString() == "u0048")
        //                {

        //                    //Console.WriteLine("Got it");
        //                    FindUofFormat = true;
        //                    return "vnd.uof.presentation";
        //                }

        //                //Console.WriteLine("Got it");
        //            }

        //            if (rdr.Name == "uof:电子表格" || rdr.Name == "UOF:电子表格")
        //            {
        //                //Console.WriteLine("找到了文字处理");
        //                //Console.WriteLine(rdr.GetAttribute(0));
        //                //Console.WriteLine("找到了电子表格");
        //                if (rdr.GetAttribute(0).ToString() == "u0049")
        //                {

        //                    //Console.WriteLine("Got it");
        //                    FindUofFormat = true;
        //                    return "vnd.uof.spreadsheet";
        //                }
        //                //Console.WriteLine("Got it");

        //            }


        //        }

        //    }
        //    return "not a uof file";
        //    //Console.WriteLine(FindUofFormat);
        //    //Console.ReadLine();

        //}



    }
}

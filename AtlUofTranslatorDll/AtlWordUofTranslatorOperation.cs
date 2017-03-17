using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Act.UofTranslator.UofZipUtils;
using Act.UofTranslator.UofTranslatorLib;
using System.IO;

namespace AtlUofTranslatorDll
{

    /// <summary>
    ///     Module ID:      
    ///     Depiction:          与 C# UOF 衔接工程
    ///     Author:             朴鑫 SYBASE v-xipia 
    ///     Create Date:        2009-02-05
    ///     Modified by:        
    ///     Last Modified Date: 2009-05-25
    /// </summary>
    public class AtlWordUofTranslatorOperation
    {
        public void SimpleAtlWordUofTranslatorImport(string bstrSourcePath, string bstrDestPath)
        {
            //openFileDialog1.Filter = "(办公文档*.pptx;*.xlsx;*.docx;*.uof;*.xml)|*.pptx;*.xlsx;*.docx;*.uof;*.xml";
            //openFileDialog1.ShowDialog();
            //FileInfo fi = new FileInfo(openFileDialog1.FileName);
            //string outputFileName = Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
            //IUOFTranslator trans = TranslatorFactory.CheckFileTypeAndInitialize(openFileDialog1.FileName);
            //richTextBox1.Text = "";
            //if (fi.Extension.ToLower() == ".uof")
            //{
            //    switch (trans.GetType().Name)
            //    {
            //        case "PresentationTranslator": outputFileName += ".pptx"; break;
            //        case "WordTranslator": outputFileName += ".docx"; break;
            //        case "SpreadsheetTranslator": outputFileName += ".xlsx"; break;
            //        default: break;
            //    }
            //    trans.UofToOox(openFileDialog1.FileName, @"c:\" + outputFileName);
            //}
            //else
            //    trans.OoxToUof(openFileDialog1.FileName, @"c:\" + outputFileName + ".uof");
            //richTextBox1.Text += "转换完成,请在C盘根目录下查看转换后文件!\n";
            try
            {




                IUOFTranslator trans = new WordTranslator();
                trans.UofToOox(bstrSourcePath, bstrDestPath);



                //FileInfo fi = new FileInfo(bstrSourcePath);
                //IUOFTranslator trans = TranslatorFactory.CheckFileTypeAndInitialize(bstrSourcePath);
                //if (fi.Extension.ToLower() == ".uof")
                //{
                //    switch (trans.GetType().Name)
                //    {
                //        //case "PresentationTranslator": outputFileName += ".pptx"; break;
                //        case "WordTranslator":
                //            //outputFileName += ".docx"; 
                //            break;
                //        //case "SpreadsheetTranslator": outputFileName += ".xlsx"; break;
                //        default: break;
                //    }
                //    trans.UofToOox(bstrSourcePath, bstrDestPath);
                //}
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }

        }

        public void SimpleAtlWordUofTranslatorExport(string bstrSourcePath, string bstrDestPath)
        {


            try
            {
                //FileInfo fi = new FileInfo(bstrSourcePath);
                //IUOFTranslator trans = TranslatorFactory.CheckFileTypeAndInitialize(bstrSourcePath);
                IUOFTranslator trans = new WordTranslator();
                trans.OoxToUof(bstrSourcePath, bstrDestPath);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }
 
        }
    }
}

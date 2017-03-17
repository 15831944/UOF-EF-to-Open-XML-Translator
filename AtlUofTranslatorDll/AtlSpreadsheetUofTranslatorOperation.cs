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
    public class AtlSpreadsheetUofTranslatorOperation
    {
        public void SimpleAtlSpreadsheetUofTranslatorImport(string bstrSourcePath, string bstrDestPath)
        {
            try
            {
                //FileInfo FileInfoTempFile = new FileInfo(bstrSourcePath);
                //bool isReadOnly = FileInfoTempFile.IsReadOnly;
                //FileInfoTempFile.IsReadOnly = true;


                IUOFTranslator trans = new SpreadsheetTranslator();
                trans.UofToOox(bstrSourcePath, bstrDestPath);

                //FileInfoTempFile.IsReadOnly = isReadOnly;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
                //System.Windows.Forms.MessageBox.Show("here");
            }
            

        }

        public void SimpleAtlSpreadsheetUofTranslatorExport(string bstrSourcePath, string bstrDestPath)
        {

            try
            {
               // System.Windows.Forms.MessageBox.Show("source=" + bstrSourcePath);
               // System.Windows.Forms.MessageBox.Show("dest=" + bstrDestPath);

                //FileInfo fi = new FileInfo(bstrSourcePath);
                //IUOFTranslator trans = TranslatorFactory.CheckFileTypeAndInitialize(bstrSourcePath);
                IUOFTranslator trans = new SpreadsheetTranslator();

               
                trans.OoxToUof(bstrSourcePath, bstrDestPath);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }

        }
    }
}
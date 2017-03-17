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
    public class AtlPresentationUofTranslatorOperation
    {
        public void SimpleAtlPresentationUofTranslatorImport(string bstrSourcePath, string bstrDestPath)
        {
            try
            {
                //FileInfo FileInfoTempFile = new FileInfo(bstrSourcePath);
                //bool isReadOnly = FileInfoTempFile.IsReadOnly;
                //FileInfoTempFile.IsReadOnly = true;
                //System.Windows.Forms.MessageBox.Show(bstrSourcePath);
                //System.Windows.Forms.MessageBox.Show("Read-only source");


                IUOFTranslator trans = new PresentationTranslator();
                trans.UofToOox(bstrSourcePath, bstrDestPath);



            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString());
            }


        }

        public void SimpleAtlPresentationUofTranslatorExport(string bstrSourcePath, string bstrDestPath)
        {

            try
            {
                //System.Windows.Forms.MessageBox.Show("bstrSourcePath: " + bstrSourcePath);
                string inputFilePath = preprocessOfficeInputFileByAppName(bstrSourcePath);
                //System.Windows.Forms.MessageBox.Show("inputFilePath: " + inputFilePath);


                object tempFile = System.IO.Path.GetTempFileName();
                string destTempFile = (string)tempFile;

                IUOFTranslator trans = new PresentationTranslator();
                trans.OoxToUof(inputFilePath, destTempFile);
                //System.Windows.Forms.MessageBox.Show("bstrDestPath: " + bstrDestPath);

                System.IO.File.Copy(destTempFile, bstrDestPath, true);




                //IUOFTranslator trans = new PresentationTranslator();
                //trans.OoxToUof(inputFilePath, bstrDestPath);
                //System.Windows.Forms.MessageBox.Show("bstrDestPath: " + bstrDestPath);
            }
            catch (Exception e)
            {
                //System.Windows.Forms.MessageBox.Show(e.ToString());
            }


        }

        private string preprocessOfficeInputFileByAppName(string bstrSourcePath)
        {
            object tempFile = System.IO.Path.GetTempFileName();
            //copy current file to temp file
            System.IO.File.Copy(bstrSourcePath, (string)tempFile, true);
            return (string)tempFile;

        }

    }
}
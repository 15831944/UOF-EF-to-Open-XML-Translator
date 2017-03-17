using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Act.UofTranslator.UofZipUtils;
using System.IO;
using System.Xml.XPath;
using System.Xml.Xsl;
using log4net;
using System.Web;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// The first step post process in OOX->UOF of SpreadSheet
    /// </summary>
    /// <author>linyaohu</author>
    class OoxToUofPostProcessorOneSpreadsheet : AbstractProcessor
    {

        private static ILog logger = LogManager.GetLogger(typeof(OoxToUofPostProcessorOneSpreadsheet).FullName);

        public override bool transform()
        {
            bool result = true;
            XmlTextWriter resultWriter = null;
            const int size = 3072;
            byte[] buffer = new byte[size];
            string picBase64String = string.Empty;
            BinaryReader breader = null;
            string picPath = Path.GetDirectoryName(originalFile) + "\\" +"XLSX"+ "\\"+"xl" + "\\" + "media"+"\\";
            //string picPath = Path.GetDirectoryName(originalFile);
            if (Directory.Exists(picPath))
            {
                try
                {
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(inputFile);
                    XmlNamespaceManager nm = new XmlNamespaceManager(xmlDoc.NameTable);
                    nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
                    XmlNode root = xmlDoc.SelectSingleNode("//uof:UOF//uof:对象集", nm);
                    XmlNodeList additionalNodeLists=root.SelectNodes("uof:其他对象",nm);
                    foreach (XmlNode addtionalNode in additionalNodeLists)
                    {
                        string picNameEx;
                        if (addtionalNode.Attributes.GetNamedItem("uof:公共类型") != null)
                        {
                            picNameEx = addtionalNode.Attributes.GetNamedItem("uof:公共类型").Value;
                        }
                        else if(addtionalNode.Attributes.GetNamedItem("uof:私有类型")!=null)
                        {
                            picNameEx = addtionalNode.Attributes.GetNamedItem("uof:私有类型").Value;
                                
                        }
                        else
                        {
                            picNameEx = "jpeg";
                        }
                        string picName = picPath + addtionalNode.Attributes.GetNamedItem("uof:标识符").Value + "." + picNameEx;
                        breader = new BinaryReader(File.Open(picName, FileMode.Open));
                        while (breader.BaseStream.Position < breader.BaseStream.Length)
                        {
                            buffer = breader.ReadBytes(size);
                            picBase64String += Convert.ToBase64String(buffer);
                        }
                        addtionalNode.SelectSingleNode("uof:数据", nm).InnerText = picBase64String;
                        picBase64String = string.Empty;
                        picName = string.Empty;
                        breader.Close();
                    }
                    resultWriter = new XmlTextWriter(outputFile, System.Text.Encoding.UTF8);
                    xmlDoc.Save(resultWriter);
                    resultWriter.Close();
                 }
                catch (Exception e)
                {
                    result = false;
                    logger.Error(e.Message);
                    logger.Error(e.StackTrace);
                    throw new Exception("Fail in post process of SpreadSheet");
                }
                finally
                {
                    if (resultWriter != null)
                        resultWriter.Close();
                    if (breader != null)
                        breader.Close();
                }
            }
            else
            {
                File.Copy(inputFile, outputFile, true);

            }
            return result;
        }

    }
}
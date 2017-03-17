using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using Act.UofTranslator.UofZipUtils;
using System.IO;
using System.Xml.XPath;
using System.Xml.Xsl;
using log4net;
using System.Reflection;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// Pretreatment of OOX->UOF in Presentation
    /// </summary>
    /// <author>huashuran&linyaohu</author>
    class OOXToUOFPreProcessorOnePresentation : AbstractProcessor
    {
        private static ILog logger = LogManager.GetLogger(typeof(OOXToUOFPreProcessorOnePresentation).FullName);

        XmlNamespaceManager nm = null;
        XmlDocument xmlDoc;
        XmlTextWriter resultWriter = null;

        public override bool transform()
        {
            FileStream fs = null;
            XmlTextReader txtReader = null;
            string extractPath = Path.GetDirectoryName(outputFile)+"\\";
            ZipReader archive = ZipFactory.OpenArchive(inputFile);
            archive.ExtractOfficeDocument(inputFile, extractPath);
            string prefix = this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + ".Powerpoint.oox2uof";//路径
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.StartsWith(prefix))
                {
                    string filename = name.Substring(prefix.Length + 1);
                    FileStream writer = new FileStream(extractPath + filename, FileMode.Create);
                    Stream baseStream = asm.GetManifestResourceStream(name);
                    int Length = 10240;
                    Byte[] buffer = new Byte[Length];
                    int bytesRead = baseStream.Read(buffer, 0, Length);
                    while (bytesRead > 0)
                    {
                        writer.Write(buffer, 0, bytesRead);
                        bytesRead = baseStream.Read(buffer, 0, Length);
                    }
                    baseStream.Close();
                    writer.Close();
                }
            }
            bool isSuccess = true;
            try
            {

                // add theme rels
                AddThemeRels(extractPath);

                File.Copy(extractPath + "pre1.xsl", extractPath + @"\ppt\pre1.xsl", true);
                XPathDocument doc = new XPathDocument(extractPath + @"\ppt\presentation.xml");
                XslCompiledTransform transFrom = new XslCompiledTransform();
                XsltSettings setting = new XsltSettings(true, false);
                XmlUrlResolver xur = new XmlUrlResolver();
                transFrom.Load(extractPath + @"\ppt\pre1.xsl", setting, xur);             
                XPathNavigator nav = ((IXPathNavigable)doc).CreateNavigator();
                fs = new FileStream(extractPath + "pre1tmp.xml", FileMode.Create);
                transFrom.Transform(nav, null, fs);
                fs.Close();
                doc = new XPathDocument(extractPath + "pre1tmp.xml");
                nav = ((IXPathNavigable)doc).CreateNavigator();
                fs = new FileStream(extractPath + "pre2tmp.xml", FileMode.Create);
                Assembly ass = Assembly.Load("ppt_oox2uof");//使用预编译后的xslt
                Type t = ass.GetType("pre2");
                transFrom.Load(t);

                transFrom.Transform(nav, null, fs);
                fs.Close();
                xmlDoc = new XmlDocument();
                txtReader = new XmlTextReader(extractPath + "pre2tmp.xml");
                xmlDoc.Load(txtReader);
                txtReader.Close();
                setNSManager();//
                string currentPath = "//p:sp";
                purifyMethod(currentPath);//

                //2011-01-12罗文甜：阴影预处理
                XmlNodeList shapeShadeList = xmlDoc.SelectNodes("//a:outerShdw", nm);
                if (shapeShadeList != null)
                {
                    foreach (XmlNode shapeShade in shapeShadeList)
                    {
                        double dist,dir;
                        if (((XmlElement)shapeShade).HasAttribute("dist"))
                        {
                            dist = Convert.ToDouble(shapeShade.Attributes.GetNamedItem("dist").Value) / 12700;
                        }
                        else
                        {
                            dist=0.0;
                        }
                        if (((XmlElement)shapeShade).HasAttribute("dir"))
                        {
                            dir = (Convert.ToDouble(shapeShade.Attributes.GetNamedItem("dir").Value) * 3.1415) / (180 * 60000);
                        }
                        else
                        {
                            dir = 0.0;
                        }
                        double xValue = dist * Math.Cos(dir);
                        double yValue = dist * Math.Sin(dir);
                        XmlElement x = xmlDoc.CreateElement("x");
                        x.InnerText = Convert.ToString(xValue);
                        shapeShade.AppendChild(x);
                        XmlElement y = xmlDoc.CreateElement("y");
                        y.InnerText = Convert.ToString(yValue);
                        shapeShade.AppendChild(y);
                    }
                }
                resultWriter = new XmlTextWriter(outputFile, System.Text.Encoding.UTF8);
                xmlDoc.Save(resultWriter);
                resultWriter.Close();

                string tmpFile= extractPath + "tmpID.xml";
                ChangeIDVal(outputFile,tmpFile);//

           
            }
            catch (Exception e)
            {
                logger.Error("Fail in OoxToUofPreProcessorOnePresentation: " + e.Message);
                logger.Error(e.StackTrace);
                isSuccess = false;
                throw new Exception("Fail in OoxToUofPreProcessorOnePresentation");
            }
            finally
            {
                if (resultWriter != null)
                    resultWriter.Close();
                if (fs != null)
                    fs.Close();
                if (txtReader != null)
                    txtReader.Close();
            }

            return isSuccess;
        }

        /// <summary>
        /// add theme relation ship contentent to theme file
        /// </summary>
        /// <param name="dir"></param>
        private static void AddThemeRels(string dir)
        {
            string themeRels=dir+Path.AltDirectorySeparatorChar+@"ppt\theme\_rels";
            string theme=dir+Path.AltDirectorySeparatorChar+@"ppt\theme\theme1.xml";
            if (Directory.Exists(themeRels))
            {
                string[] rels = Directory.GetFiles(themeRels);

                if (rels.Length > 0)
                {
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(rels[0]);
                    XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
                    nm.AddNamespace("r", TranslatorConstants.XMLNS_REL);

                    XmlDocument themeDoc = new XmlDocument();
                    themeDoc.Load(theme);
                    XmlNode relNode = themeDoc.ImportNode(xdoc.LastChild, true);

                    XmlNamespaceManager nmt = new XmlNamespaceManager(themeDoc.NameTable);
                    nmt.AddNamespace("a", TranslatorConstants.XMLNS_A);
                    themeDoc.SelectSingleNode("a:theme", nmt).AppendChild(relNode);
                    themeDoc.Save(theme);

                }

            }
        }

        private bool ChangeIDVal(string inputFile, string outputFile)
        {
            xmlDoc = new XmlDocument();
            if (inputFile.Equals(string.Empty))
            {
                return false;
            }
            xmlDoc.Load(inputFile);
            nm = new XmlNamespaceManager(xmlDoc.NameTable);
            XmlNodeList ids = xmlDoc.SelectNodes("//*[@id !='']", nm);
            Dictionary<string,int> idCollect = new Dictionary<string,int>();
            idCollect.Clear();
            int idcurr = 0;
            try
            {
                foreach (XmlNode id in ids)
                {
                    idcurr++;
                    string s = id.Attributes.GetNamedItem("id").Value;
                    if (!idCollect.ContainsKey(s))
                    {
                        idCollect.Add(s, idcurr);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                int c;
                idCollect.TryGetValue("ID0EJC",out c);
                c = 0;
            }
            return true;
        }

        #region 预处理添加的一些操作

        private void setNSManager()
        {
            XmlNameTable table = xmlDoc.NameTable;
            nm = new XmlNamespaceManager(table);
            nm.AddNamespace("字", "http://schemas.uof.org/cn/2003/uof-wordproc");
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("图", "http://schemas.uof.org/cn/2003/graph");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            nm.AddNamespace("演", "http://schemas.uof.org/cn/2003/uof-slideshow");
            nm.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nm.AddNamespace("app", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
            nm.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            nm.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
            nm.AddNamespace("dcmitype", "http://purl.org/dc/dcmitype/");
            nm.AddNamespace("dcterms", "http://purl.org/dc/terms/");
            nm.AddNamespace("fo", "http://www.w3.org/1999/XSL/Format");
            nm.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            nm.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            nm.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

        }


        private void purifyMethod(string path)
        {
            try
            {
                XmlNodeList eleList = xmlDoc.SelectNodes(path, nm);
                // MessageBox.Show(eleList.Count.ToString());           
                for (int i = 0; i < eleList.Count; i++)
                {
                    dealMethod((XmlElement)eleList[i]);
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
            }
            return;
        }


        private void dealMethod(XmlElement currentEle)
        {
            string hID = gethID(currentEle);
            string id = getID(currentEle);
            XmlElement referenceEle = null;
            if (hID != "EOF")
            {
                try
                {
                    referenceEle = (XmlElement)xmlDoc.SelectSingleNode("//p:sp[@id='" + hID + "']", nm);
                    if (referenceEle != null)
                    {
                        compareEqualNode(currentEle, referenceEle);
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
            return;
        }

        private void compareEqualNode(XmlElement curNode, XmlElement refNode)
        {
            //if (curNode.Name != "a:p" && curNode.Name != "a:lstStyle")
            if (true)
            {
                XmlAttributeCollection curAttrList = curNode.Attributes;
                XmlAttributeCollection refAttrList = refNode.Attributes;
                for (int p = 0; p < refAttrList.Count; p++)
                {
                    string refAttrName = refAttrList[p].Name;
                    int q = 0;
                    while (q < curAttrList.Count)
                    {
                        string curAttrName = curAttrList[q].Name;
                        if (curAttrName == refAttrName)
                        {
                            break;
                        }
                        q++;
                    }
                    if (q == curAttrList.Count)
                    {   
                        XmlAttribute refAttr = (XmlAttribute)refAttrList[p].Clone();
                        curAttrList.Append(refAttr);
                    }
                }
                if (curNode.ChildNodes.Count == 0 && refNode.ChildNodes.Count != 0)
                {
                    for (int m = 0; m < refNode.ChildNodes.Count; m++)
                    {
                        if (refNode.ChildNodes[m].Name != "p:txBody" && refNode.ChildNodes[m].Name != "a:p" && refNode.ChildNodes[m].Name != "a:lstStyle")
                        {
                            curNode.AppendChild((XmlNode)refNode.ChildNodes[m].Clone());
                        }
                    }
                }

                if (curNode.ChildNodes.Count != 0 && refNode.ChildNodes.Count != 0)
                {
                    for (int m = 0; m < refNode.ChildNodes.Count; m++)
                    {
                        if (refNode.ChildNodes[m].Name != "a:p" && refNode.ChildNodes[m].Name != "a:lstStyle")
                        {
                            int n = 0;
                            while (n < curNode.ChildNodes.Count)
                            {
                                if (refNode.ChildNodes[m].Name == curNode.ChildNodes[n].Name)
                                {
                                    compareEqualNode((XmlElement)curNode.ChildNodes[n], (XmlElement)refNode.ChildNodes[m]);
                                    break;
                                }
                                n++;


                            }
                            if (n == curNode.ChildNodes.Count)
                            {
                                if (refNode.ChildNodes[m].Name != "p:txBody")
                                {
                                    curNode.AppendChild((XmlNode)refNode.ChildNodes[m].Clone());
                                }
                            }
                        }
                    }
                }

            }
        }

        private string gethID(XmlElement currentEle)
        {
            string referenceId = "";
            if (currentEle.Attributes["hID"] != null)
            {
                referenceId = currentEle.Attributes["hID"].Value;
                return referenceId;
            }
            else
            {
                return "EOF";
            }
        }


        private string getID(XmlElement currentEle)
        {
            string Id = "";
            if (currentEle.Attributes["id"] != null)
            {
                Id = currentEle.Attributes["id"].Value;
                return Id;
            }
            else
            {
                return "EOF";
            }
        }

        #endregion


    }
}

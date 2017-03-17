/*
 * Copyright (c) 2006, Beihang University, China
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 *     * Redistributions of source code must retain the above copyright
 *       notice, this list of conditions and the following disclaimer.
 *     * Redistributions in binary form must reproduce the above copyright
 *       notice, this list of conditions and the following disclaimer in the
 *       documentation and/or other materials provided with the distribution.
 *     * Neither the name of Clever Age, nor the names of its contributors may
 *       be used to endorse or promote products derived from this software
 *       without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE REGENTS AND CONTRIBUTORS ``AS IS'' AND ANY
 * EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE REGENTS AND CONTRIBUTORS BE LIABLE FOR ANY
 * DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using log4net;
using System.IO;
using System.Text.RegularExpressions;
using System.Reflection;
using System.Xml.XPath;
using System.Xml.Xsl;

namespace Act.UofTranslator.UofTranslatorStrictLib
{
    /// <summary>
    /// The first step pre process
    /// </summary>
    /// <author>fangchunyan, linwei</author>
    class UofToOoxPreProcessorOneWord : AbstractProcessor
    {
        private XmlNamespaceManager namespaceManager=null;

        private static ILog logger = LogManager.GetLogger(typeof(UofToOoxPreProcessorOneWord).FullName);

        public override bool transform()
        {
            FileStream fs = null;
            //XmlUrlResolver resourceResolver = null;
            XmlReader xr = null;

            string extractPath = Path.GetDirectoryName(outputFile) + Path.AltDirectorySeparatorChar;
            string prefix = this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + TranslatorConstants.UOFToOOX_WORD_LOCATION;
            string uof2ooxpre = extractPath + "uof2ooxpre.xml";
            string extend = extractPath + "extend.xml";
            string extendPre = extractPath + "extendPre.xml";
            string tmpSect1 = extractPath + "tmpSect1.xml";
            string tmpSect2 = extractPath + "tmpSect2.xml";
            string tblVtcl = extractPath + "tblVertical.xml";
            string picture_xml = "data";

            try
            {
                UOFToOoxTableProcessing tableProcessing2 = new UOFToOoxTableProcessing(extractPath + "content.xml");
                tableProcessing2.Processing(tblVtcl);


                // copy XSLT templates
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

                //resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(), prefix);
                //xr = XmlReader.Create(((ResourceResolver)resourceResolver).GetInnerStream("uof2oopre.xslt"));
                xr = XmlReader.Create(extractPath + "uof2ooxpre.xsl");

                // XPathDocument doc = new XPathDocument(extractPath + @"content.xml");
                XPathDocument doc = new XPathDocument(tblVtcl);

                XslCompiledTransform transFrom = new XslCompiledTransform();
                XsltSettings setting = new XsltSettings(true, false);
                XmlUrlResolver xur = new XmlUrlResolver();
                transFrom.Load(xr, setting, xur);
                XPathNavigator nav = ((IXPathNavigable)doc).CreateNavigator();
                fs = new FileStream(outputFile, FileMode.Create);
                // fs = new FileStream(uof2ooxpre, FileMode.Create);
                transFrom.Transform(nav, null, fs);
                fs.Close();

                if (File.Exists(extend))
                {
                    XmlDocument xdoc1 = new XmlDocument();
                    xdoc1.Load(outputFile);
                    XmlNamespaceManager nm = new XmlNamespaceManager(xdoc1.NameTable);
                    nm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);

                    XmlDocument xdoc2 = new XmlDocument();
                    xdoc2.Load(extend);

                    XmlNode extendNode = xdoc1.CreateElement("uof", "��չ��", TranslatorConstants.XMLNS_UOF);
                    extendNode.InnerXml = xdoc2.LastChild.InnerXml;
                    xdoc1.SelectSingleNode("//uof:UOF", nm).AppendChild(extendNode);

                    xdoc1.Save(extendPre);
                }
                else
                {
                    extendPre = outputFile;
                }

                if (!Directory.Exists(extractPath + picture_xml))
                {
                    Directory.CreateDirectory(extractPath + picture_xml);
                }
                string[] patternFiles = Directory.GetFiles(extractPath, "Image*", SearchOption.TopDirectoryOnly);
                foreach (string patternFile in patternFiles)
                {
                    File.Copy(patternFile, extractPath + picture_xml + Path.AltDirectorySeparatorChar + patternFile.Substring(patternFile.LastIndexOf("/") + 1), true);
                }



                if (Directory.Exists(extractPath + picture_xml))
                {
                    string tmpPic = extractPath + Path.AltDirectorySeparatorChar + "tmpPic.xml";
                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(extendPre);

                    XmlNameTable nt = xdoc.NameTable;
                    XmlNamespaceManager nm = new XmlNamespaceManager(nt);
                    nm.AddNamespace("w", TranslatorConstants.XMLNS_W);
                    nm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);



                    xdoc = PicPretreatment(xdoc, "uof:UOF", extractPath + picture_xml, nm);
                    xdoc.Save(tmpPic);
                    // OutputFilename = tmpPic;
                    extendPre = tmpPic;

                }



                //�ڶ���Ԥ����
                XPathDocument xpdoc = UOFTranslator.GetXPathDoc(TranslatorConstants.UOFToOOX_PRETREAT_STEP2_XSL, TranslatorConstants.UOFToOOX_WORD_LOCATION);
                XslCompiledTransform xslt = new XslCompiledTransform();
                xslt.Load(xpdoc);
                XmlWriter xw = new XmlTextWriter(tmpSect1, Encoding.UTF8);//sndtemDoc.xml
                xslt.Transform(extendPre, xw);//tmpDoc--����Ԥ����2.xsl--sndtemDoc.xml
                if (xw != null)
                {
                    xw.Close();
                }

                //������Ԥ����
                xpdoc = UOFTranslator.GetXPathDoc(TranslatorConstants.UOFToOOX_PRETREAT_SETP3_XSL, TranslatorConstants.UOFToOOX_WORD_LOCATION);
                xslt.Load(xpdoc);
                xw = new XmlTextWriter(tmpSect2, Encoding.UTF8);
                xslt.Transform(tmpSect1, xw);
                if (xw != null)
                {

                    xw.Close();
                }

                XmlDocument shdXdoc = new XmlDocument();
                shdXdoc.Load(tmpSect2);
                XmlNameTable nameTable = shdXdoc.NameTable;
                XmlNamespaceManager snm = new XmlNamespaceManager(nameTable);
                snm.AddNamespace("ͼ", TranslatorConstants.XMLNS_UOFGRAPH);
                snm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);


                XmlNodeList shapeShades = shdXdoc.SelectNodes("//ͼ:��Ӱ_8051/uof:ƫ����_C61B", snm);
                if (shapeShades != null)
                {
                    foreach (XmlNode shapeShade in shapeShades)
                    {
                        if (((XmlElement)shapeShade).HasAttribute("x_C606") && ((XmlElement)shapeShade).HasAttribute("y_C607"))
                        {
                            if (shapeShade.Attributes.GetNamedItem("x_C606").Value != "NaN" && shapeShade.Attributes.GetNamedItem("y_C607").Value != "NaN")
                            {
                                double x = Convert.ToDouble(shapeShade.Attributes.GetNamedItem("x_C606").Value);
                                double y = Convert.ToDouble(shapeShade.Attributes.GetNamedItem("y_C607").Value);
                                double angValueTe = Math.Atan2(Math.Abs(y), Math.Abs(x)) * 180 / 3.1415926;//x,yλ�õ������޸���Ӱ-�Ƕ�ת��BUG
                                double distanceValue = Math.Sqrt((x * x) + (y * y)) * 12700;
                                double angValue = 0;
                                if (x >= 0 && y >= 0)
                                {
                                    angValue = angValueTe * 60000;
                                }
                                if (x >= 0 && y < 0)
                                {
                                    angValue = (angValueTe + 270) * 60000;
                                }
                                if (x < 0 && y >= 0)
                                {
                                    angValue = (angValueTe + 90) * 60000;
                                }
                                if (x < 0 && y < 0)
                                {
                                    angValue = (angValueTe + 180) * 60000;
                                }
                                XmlElement ang = shdXdoc.CreateElement("dir");
                                ang.InnerText = Convert.ToString(angValue);
                                shapeShade.AppendChild(ang);
                                XmlElement distance = shdXdoc.CreateElement("dist");
                                distance.InnerText = Convert.ToString(distanceValue);
                                shapeShade.AppendChild(distance);
                            }
                        }

                    }
                }

                shdXdoc.Save(outputFile);
                //��ʽ
                if (File.Exists(extractPath + "equations.xml"))
                {
                    string tmpEqu = extractPath + Path.AltDirectorySeparatorChar + "tmpEqu.xml";

                    string equXML = extractPath + "equations.xml";
                    XmlDocument equDoc = new XmlDocument();
                    equDoc.Load(equXML);

                    XmlDocument xdoc = new XmlDocument();
                    xdoc.Load(outputFile);

                    XmlNode equ = xdoc.ImportNode(equDoc.LastChild, true);
                    xdoc.LastChild.AppendChild(equ);
                    xdoc.Save(tmpEqu);

                    outputFile = tmpEqu;
                }

                //outputFile = tmpSect2;
                //  preMethod(uof2ooxpre);
                return true;


            }
            catch (Exception ex)
            {
                logger.Error("Fail in Uof2.0 to OOX pretreatment1: " + ex.Message);
                logger.Error(ex.StackTrace);
                throw new Exception("Fail in Uof2.0 to OOX pretreatment1 of WORD");
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }


        //public override bool transform()
        //{
        //   // string tempfile = Path.GetTempPath() + "tempfile.xml";
        //    Guid gWordPath = Guid.NewGuid();
        //    string preWordPath = Path.GetTempPath() + gWordPath.ToString() + "\\";
        //    if(!Directory.Exists(preWordPath))
        //        Directory.CreateDirectory(preWordPath);
        //    string tempfile=preWordPath+"tempfile.xml";
        //    string preOutputFileName = preWordPath + "tmpDoc1.xml";
        //    StreamReader sr = new StreamReader(inputFile);
        //    StreamWriter sw = new StreamWriter(tempfile);
        //    string strTemp = "";
        //    while ((strTemp = sr.ReadLine()) != null)
        //    {
        //        if (strTemp.Contains("<��:�ı���") == true && strTemp.Contains("xml:space=\"preserve\"") == false)
        //        {
        //            strTemp = strTemp.Replace("<��:�ı���", "<��:�ı��� xml:space=\"preserve\"");

        //        }
        //        sw.WriteLine(strTemp);
        //    }
        //    sr.Close();
        //    sw.Close();

        //    XmlDocument xmlDoc = new XmlDocument();
        //    xmlDoc.Load(tempfile);

        //    XmlNameTable table = xmlDoc.NameTable;
        //    namespaceManager = new XmlNamespaceManager(table);
        //    namespaceManager.AddNamespace("��", "http://schemas.uof.org/cn/2003/uof-wordproc");
        //    namespaceManager.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
        //    namespaceManager.AddNamespace("ͼ", "http://schemas.uof.org/cn/2003/graph");

        //    XmlNode root = xmlDoc.DocumentElement;
        //    RestructurePS(xmlDoc, root);
        //    RestructureHyperlinkAndRevision(xmlDoc, root);
        //    RestructureField(xmlDoc, root);
        //    RestructureRPr(xmlDoc, root);

        //    XmlTextWriter resultWriter = new XmlTextWriter(preOutputFileName, Encoding.UTF8);
        //    xmlDoc.Save(resultWriter);
        //    resultWriter.Close();
        //    OutputFilename = preOutputFileName;

        //    return true;
        //}

        private void RestructurePS(XmlDocument xmlDoc, XmlNode root)
        {

            XmlNode body = root.SelectSingleNode("uof:���ִ���/��:����", namespaceManager);
            XmlNodeList Children = body.ChildNodes;
            XmlNode tempNode = null;
            XmlElement tempPChild = xmlDoc.CreateElement("��", "����", "http://schemas.uof.org/cn/2003/uof-wordproc");

            int i = 0;
            while (i < Children.Count)
            {
                XmlNode node = Children[i];
                if (node.GetType().ToString() != "System.Xml.XmlElement")
                {
                    body.RemoveChild(Children[i]);
                }
                else
                {
                    i++;
                }
            }

            try
            {
                foreach (XmlNode child in Children)
                {
                    if (child.GetType().ToString() == "System.Xml.XmlElement")
                    {

                        switch (child.Name.ToString())
                        {
                            case "��:�ֽ�":
                                if (child.PreviousSibling != null && child.NextSibling != null)
                                {
                                    if (child.PreviousSibling.Name == "��:�ֽ�")
                                    {
                                        tempPChild.RemoveAll();
                                        XmlElement pPrChild = xmlDoc.CreateElement("��", "��������", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                        pPrChild.AppendChild(child.PreviousSibling);
                                        tempPChild.AppendChild(pPrChild);
                                        body.InsertBefore(tempPChild.Clone(), child);
                                    }
                                    else
                                    {
                                        body.InsertBefore(tempPChild.Clone(), child);
                                    }

                                }
                                tempNode = child.Clone();


                                break;

                            case "��:����":
                                if (child.NextSibling == null)
                                {
                                    if (child.PreviousSibling.Name.ToString() == "��:�ֽ�")
                                    {
                                        body.RemoveChild(child.PreviousSibling);
                                    }
                                    body.AppendChild(tempNode);

                                    break;
                                }
                                if (child.PreviousSibling.Name.ToString() == "��:�ֽ�")
                                {
                                    body.RemoveChild(child.PreviousSibling);
                                }
                                if (child.NextSibling.Name.ToString() == "��:�ֽ�")
                                {

                                    tempPChild.RemoveAll();
                                    XmlElement pPrChild = xmlDoc.CreateElement("��", "��������", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                    pPrChild.AppendChild(tempNode);
                                    tempPChild.AppendChild(pPrChild);

                                }
                                break;

                            case "��:���ֱ�":
                                if (child.PreviousSibling.Name.ToString() == "��:�ֽ�")
                                {
                                    body.RemoveChild(child.PreviousSibling);
                                }
                                if (child.NextSibling == null)
                                {
                                    if (child.PreviousSibling.Name.ToString() == "��:�ֽ�")
                                    {
                                        body.RemoveChild(child.PreviousSibling);
                                    }
                                    body.AppendChild(tempNode);
                                    break;
                                }
                                if (child.NextSibling.Name.ToString() == "��:�ֽ�")
                                {
                                    tempPChild.RemoveAll();
                                    XmlElement pPrChild = xmlDoc.CreateElement("��", "��������", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                    pPrChild.AppendChild(tempNode);
                                    tempPChild.AppendChild(pPrChild);

                                }

                                break;
                            default:
                                if (child.PreviousSibling.Name.ToString() == "��:�ֽ�")
                                {
                                    body.RemoveChild(child.PreviousSibling);
                                }
                                if (child.NextSibling == null)
                                {
                                    if (child.PreviousSibling.Name.ToString() == "��:�ֽ�")
                                    {
                                        body.RemoveChild(child.PreviousSibling);
                                    }
                                    body.AppendChild(tempNode);
                                    break;
                                }
                                if (child.NextSibling.Name.ToString() == "��:�ֽ�")
                                {
                                    tempPChild.RemoveAll();
                                    XmlElement pPrChild = xmlDoc.CreateElement("��", "��������", "http://schemas.uof.org/cn/2003/uof-wordproc");
                                    pPrChild.AppendChild(tempNode);
                                    tempPChild.AppendChild(pPrChild);

                                }

                                break;
                        }
                    }
                }
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                logger.Error(e.StackTrace);
            }
        }

        private void RestructureHyperlinkAndRevision(XmlDocument xmlDoc, XmlNode root)
        {
            try
            {
                XmlNode body = root.SelectSingleNode("uof:���ִ���", namespaceManager);
                XmlNodeList hyperSetList = xmlDoc.SelectNodes("//uof:���Ӽ�/uof:��������[@uof:Ŀ��]", namespaceManager);
                for (int i = 0; i < hyperSetList.Count; i++)
                {
                    XmlElement hyperSet = (XmlElement)hyperSetList[i];
                    string target = hyperSet.Attributes["uof:Ŀ��"].Value.ToString();
                    hyperSet.Attributes["uof:Ŀ��"].Value = UrlEncodeForWord(target);
                }

                XmlNodeList hyperList = xmlDoc.SelectNodes("//��:����ʼ[@��:����='hyperlink']", namespaceManager);
                for (int i = 0; i < hyperList.Count; i++)
                {
                    XmlElement hyper = (XmlElement)hyperList[i];
                    //yx,change,2010.3.2
                    if (hyper.SelectNodes("ancestor::��:����/��:��/��:�������[@��:��ʶ������='" + hyper.Attributes["��:��ʶ��"].Value + "']", namespaceManager).Count > 0)
                    {
                        dealElement(xmlDoc, hyper);
                    }
                }

                XmlNodeList RevList = xmlDoc.SelectNodes("//��:��/��:�޶���ʼ", namespaceManager);
                for (int i = 0; i < RevList.Count; i++)
                {
                    XmlElement rev = (XmlElement)RevList[i];
                    dealElement(xmlDoc, rev);
                }

                XmlNodeList RevListFromPara = xmlDoc.SelectNodes("//��:����/��:�޶���ʼ", namespaceManager);
                for (int i = 0; i < RevListFromPara.Count; i++)
                {
                    XmlElement revStart = (XmlElement)RevListFromPara[i];
                    XmlElement revStartParent = (XmlElement)revStart.ParentNode;
                    int m = 0;
                    while (m < revStartParent.ChildNodes.Count)
                    {
                        XmlNode node = revStartParent.ChildNodes[m];
                        if (node.GetType().ToString() != "System.Xml.XmlElement")
                        {
                            revStartParent.RemoveChild(revStartParent.ChildNodes[m]);
                        }
                        else
                        {
                            m++;
                        }
                    }
                    if (revStart.NextSibling == null)
                    {
                        string revid = revStart.Attributes["��:��ʶ��"].Value;
                        string revType = revStart.Attributes["��:����"].Value;
                        XmlElement revEnd = (XmlElement)body.SelectSingleNode("//��:�޶�����[@��:��ʼ��ʶ����='" + revid + "']", namespaceManager);
                        XmlElement revEndParent = (XmlElement)revEnd.ParentNode;
                        if (revType == "format")
                        {
                            revEndParent.InsertBefore(revStart, revEndParent.FirstChild);

                        }
                        else if (revType == "delete" || revType == "insert")
                        {
                            XmlElement pPr = (XmlElement)revEndParent.SelectSingleNode("��:��������[not(preceding-sibling::��:�޶���ʼ/@��:��ʶ�� = following-sibling::��:�޶�����/@��:��ʼ��ʶ����)]", namespaceManager);
                            if (pPr != null)
                            {
                                revEndParent.InsertAfter(revStart, pPr);
                            }
                            else
                            {
                                revEndParent.InsertBefore(revStart, revEndParent.FirstChild);
                            }
                        }

                    }
                }
            }
            catch (Exception e)
            {
                logger.Info(e.Message);
                logger.Info(e.StackTrace);
            }
        }

        private void dealElement(XmlDocument xDoc, XmlElement regionStart)
        {
            XmlElement xmlParent = (XmlElement)regionStart.ParentNode;
            if (xmlParent == null)
            { return; }

            XmlElement xmlGrandpa = (XmlElement)xmlParent.ParentNode;
            XmlElement newElement = xDoc.CreateElement("��", "��", "http://schemas.uof.org/cn/2003/uof-wordproc");
            string id = regionStart.Attributes["��:��ʶ��"].Value;

            XmlElement regionEnd = null;
            if (regionStart.Name == "��:�޶���ʼ")
            {
                regionEnd = (XmlElement)xmlParent.SelectSingleNode("��:�޶�����[@��:��ʼ��ʶ����='" + id + "']", namespaceManager);
            }
            else if (regionStart.Name == "��:����ʼ")
            {
                regionEnd = (XmlElement)xmlParent.SelectSingleNode("��:�������[@��:��ʶ������='" + id + "']", namespaceManager);
            }
            if (regionEnd != null)
            {
                return;
            }

            XmlElement regionStartPr = (XmlElement)xmlParent.SelectSingleNode("��:������[not(preceding-sibling::��:�޶���ʼ/@��:��ʶ�� = following-sibling::��:�޶�����/@��:��ʼ��ʶ����)]", namespaceManager);//�������Ӿ�����

            newElement.AppendChild(regionStart.Clone());
            while (regionStart.NextSibling != null)
            {
                newElement.AppendChild(regionStart.NextSibling);
            }
            xmlParent.RemoveChild(regionStart);

            XmlElement eleUncle = (XmlElement)xmlParent.NextSibling;
            string selectStr = "";
            if (regionStart.Name == "��:�޶���ʼ")
            {
                selectStr = "��:�޶�����[@��:��ʼ��ʶ����='" + id + "']";
            }
            else if (regionStart.Name == "��:����ʼ")
            {
                selectStr = "��:�������[@��:��ʶ������='" + id + "']";
            }

            while (eleUncle != null && eleUncle.SelectSingleNode(selectStr, namespaceManager) == null)
            {
                while (eleUncle.ChildNodes.Count > 0)
                {
                    if (eleUncle.ChildNodes[0].GetType().ToString() == "System.Xml.XmlElement")
                    {
                        XmlElement tempa = (XmlElement)eleUncle.ChildNodes[0];
                        if (tempa.Name == "��:������")
                        {
                            regionStartPr = (XmlElement)tempa.Clone();

                            eleUncle.RemoveChild(eleUncle.ChildNodes[0]);
                        }

                    }
                    else
                    {
                        eleUncle.RemoveChild(eleUncle.ChildNodes[0]);
                    }
                    newElement.AppendChild(eleUncle.ChildNodes[0]);
                }
                eleUncle = (XmlElement)eleUncle.NextSibling;
                if (eleUncle != null)
                {
                    xmlGrandpa.RemoveChild(eleUncle.PreviousSibling);
                }
                else
                {
                    xmlGrandpa.RemoveChild(xmlGrandpa.LastChild);
                }
            }

            if (eleUncle != null)
            {
                if (eleUncle.GetType().ToString() == "System.Xml.XmlElement")
                {
                    while (eleUncle.ChildNodes.Count > 0)
                    {

                        XmlNode nodeBrother = eleUncle.ChildNodes[0];
                        if (nodeBrother.Name == "��:������")
                        {
                            XmlElement temp1 = (XmlElement)nodeBrother.Clone();
                            eleUncle.RemoveChild(nodeBrother);
                            eleUncle.AppendChild(temp1);
                        }
                        else
                        {
                            newElement.AppendChild(nodeBrother);
                        }

                        if (nodeBrother.Name == "��:�������" && nodeBrother.Attributes["��:��ʶ������"].Value == id)
                        {
                            newElement.AppendChild(nodeBrother);
                            break;
                        }
                        else if (nodeBrother.Name == "��:�޶�����" && nodeBrother.Attributes["��:��ʼ��ʶ����"].Value == id)
                        {
                            newElement.AppendChild(nodeBrother);
                            break;
                        }
                    }
                }
                else
                {
                    xmlGrandpa.RemoveChild(eleUncle);
                }
                if (eleUncle.HasChildNodes == false)
                {
                    xmlGrandpa.RemoveChild(eleUncle);
                }
            }
            else
            {
                XmlElement endRegion = (XmlElement)xmlParent.SelectSingleNode("following::��:�������[@��:��ʶ������='" + id + "']", namespaceManager);
                string selectStr2 = "";
                if (regionStart.Name == "��:�޶���ʼ")
                {
                    selectStr2 = "following::��:�޶�����[@��:��ʼ��ʶ����='" + id + "']";
                    endRegion = (XmlElement)xmlParent.SelectSingleNode(selectStr2, namespaceManager);
                }
                else if (regionStart.Name == "��:����ʼ")
                {
                    selectStr2 = "following::��:�������[@��:��ʶ������='" + id + "']";
                    endRegion = (XmlElement)xmlParent.SelectSingleNode(selectStr2, namespaceManager);
                }
                if (endRegion != null)
                {
                    newElement.AppendChild(endRegion);
                }

            }

            if (regionStartPr != null && regionStart.Name == "��:����ʼ")
            {
                XmlElement tmp = (XmlElement)regionStartPr.Clone();
                newElement.InsertBefore(tmp, newElement.ChildNodes[0]);
            }
            xmlGrandpa.InsertAfter(newElement, xmlParent);
            if (xmlParent.HasChildNodes == false)
            {
                xmlGrandpa.RemoveChild(xmlParent);
            }
        }

        private void RestructureField(XmlDocument xmlDoc, XmlNode root)
        {
            try
            {
                XmlNode body = root.SelectSingleNode("uof:���ִ���", namespaceManager);
                XmlNodeList fieldList = xmlDoc.SelectNodes("//��:��ʼ", namespaceManager);
                for (int i = 0; i < fieldList.Count; i++)
                {
                    XmlElement field = (XmlElement)fieldList[i];
                    dealFieldElement(xmlDoc, field);
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                logger.Error(ex.StackTrace);
            }
        }


        private void dealFieldElement(XmlDocument xmlDoc, XmlElement field)
        {
            XmlElement fieldCode = (XmlElement)field.NextSibling;
            while (fieldCode != null)
            {
                if (fieldCode.Name != "��:�����")
                {
                    fieldCode = (XmlElement)fieldCode.NextSibling;
                    field.AppendChild(fieldCode.PreviousSibling);
                }
                else
                {
                    break;
                }

            }

        }

        private void RestructureRPr(XmlDocument xmlDoc, XmlNode root)
        {
            try
            {
                XmlNodeList rPrInpPrList = root.SelectNodes("//��:��������/��:������", namespaceManager);

                foreach (XmlNode rPrInpPr in rPrInpPrList)
                {
                    if (rPrInpPr != null)
                    {
                        XmlNodeList rPrList = rPrInpPr.SelectNodes("./*", namespaceManager);
                        for (int i = 0; i < rPrList.Count; i++)
                        {
                            XmlNode rPrNode = (XmlNode)rPrList[i];
                            XmlElement rPr = (XmlElement)rPrList[i];
                            XmlNodeList runs = rPrInpPr.SelectNodes("ancestor::��:����/��:��", namespaceManager);
                            foreach (XmlNode run in runs)
                            {
                                XmlElement rPrInRun = (XmlElement)run.SelectSingleNode("��:������", namespaceManager);
                                if (rPrInRun == null)
                                {
                                    run.InsertBefore(rPrInpPr.Clone(), run.FirstChild);
                                }
                                else
                                {
                                    XmlNodeList rPrListInRun = rPrInRun.SelectNodes("./*", namespaceManager);
                                    int j;
                                    for (j = 0; j < rPrListInRun.Count; j++)
                                    {
                                        XmlElement rPrElementInRun = (XmlElement)rPrListInRun[j];
                                        if (rPr.Name == rPrElementInRun.Name)
                                            break;
                                    }

                                    if (j == rPrListInRun.Count)
                                    {

                                        rPrInRun.AppendChild(rPrNode.Clone());
                                    }
                                }
                            }

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                logger.Error(ex.StackTrace);
            }
        }

        private void addRPrToRun(XmlDocument xmlDoc, XmlNode root, XmlElement rPr)
        {

        }

        private static Regex preserval = new Regex(@"[!$&'()*+,-.:;=@_~/]|\w");

        private string UrlEncodeForWord(string originalUrl)
        {
            StringBuilder buffer = new StringBuilder();
            int tmp = 0;
            char c = '\0';
            for (int i = 0; i < originalUrl.Length; i++)
            {
                c = originalUrl[i];
                tmp = (int)c;

                if (tmp > 127)
                {
                    buffer.Append(c);
                    continue;
                }

                if (preserval.IsMatch(c.ToString()))
                {
                    buffer.Append(c);
                    continue;
                }

                buffer.Append("%" + DecToHex(tmp / 16) + DecToHex(tmp % 16));
            }

            return buffer.ToString();
        }

        private char DecToHex(int dec)
        {
            if (dec > 15 || dec < 0)
                return '\0';
            else if (dec > 9)
                return (char)('A' + dec - 10);
            else
                return (char)('0' + dec);
        }
    }

    public class UOFToOoxTableProcessing : TableProcessing
    {

        private string wordNamespace = @"http://schemas.uof.org/cn/2009/wordproc";

        public UOFToOoxTableProcessing(string fileName)
            : base(fileName)
        {

        }

        /// <summary>
        /// ����һ����Ԫ��
        /// </summary>
        public override void Processing(string fileName)
        {

            XmlNodeList tableList = this.GetTableList(@"//��:���ֱ�_416C");

            //ѭ�����ڵ�
            for (int i = 0; i < tableList.Count; i++)
            {
                XmlNodeList rowList = this.GetRowList(tableList[i], @"��:��_41CD");

                DoubleLinkedList<XmlNode>[] cellMatrix = new DoubleLinkedList<XmlNode>[rowList.Count];
                for (int j = 0; j < cellMatrix.Length; j++) cellMatrix[j] = new DoubleLinkedList<XmlNode>();

                //����ÿһ����
                this.RowProcessing(tableList[i], cellMatrix);
                //TestMethod(i, cellMatrix);
            }

            XmlSource.Save(fileName);
        }

        private void TestMethod(int tableNum, DoubleLinkedList<XmlNode>[] cellMatrix)
        {
            using (StreamWriter fs = new StreamWriter("target" + tableNum.ToString() + ".xml"))
            {
                fs.WriteLine("<��:���>");
                foreach (DoubleLinkedList<XmlNode> list in cellMatrix)
                {
                    fs.WriteLine("<��:��>");
                    foreach (DoubleLinkedList<XmlNode>.DoubleLinkedNode node in list.ToArray())
                    {
                        fs.WriteLine("<" + node.Value.Name + ">");


                        if (this.GetSingleNode(node.Value, @".//��:�ı���_415B") != null)
                        {
                            fs.WriteLine(this.GetSingleNode(node.Value, @".//��:�ı���_415B").InnerText);
                        }
                        else
                        {
                            fs.WriteLine(node.Value.InnerXml);
                        }
                        fs.WriteLine("</" + node.Value.Name + ">");
                    }
                    fs.WriteLine("</��:��>");
                }
                fs.WriteLine("</��:���>");
            }
        }

        //����ÿһ����   
        private void RowProcessing(XmlNode tableNode, DoubleLinkedList<XmlNode>[] cellMatrix)
        {
            XmlNodeList rowList = rowList = this.GetRowList(tableNode, @"��:��_41CD");

            //����ÿһ����Ԫ��
            for (int i = 0; i < rowList.Count; i++)
            {
                this.CellProcessing(rowList[i], i, cellMatrix);
            }

            //���д���            
            CellVerticalSpanProcessing(cellMatrix);

            //�������е�ֵ����ԭ�����Ϣ
            for (int i = 0; i < rowList.Count; i++)
            {
                XmlNodeList cellList = this.GetCellList(rowList[i], @"��:��Ԫ��_41BE");
                List<DoubleLinkedList<XmlNode>.DoubleLinkedNode> newCellList = cellMatrix[i].ToArray();

                int j = 0;
                while (j < cellList.Count && j < newCellList.Count)
                {
                    rowList[i].ReplaceChild(newCellList[j].Value, cellList[j]);
                    j++;
                }
                cellList = this.GetCellList(rowList[i], @"��:��Ԫ��_41BE");
                while (j < newCellList.Count)
                {
                    if (cellList.Count == 0)
                    {
                        rowList[i].AppendChild(newCellList[j].Value);
                    }
                    else
                    {
                        rowList[i].InsertAfter(newCellList[j].Value, cellList[j - 1]);
                    }
                    j++;
                    cellList = this.GetCellList(rowList[i], @"��:��Ԫ��_41BE");
                }

            }
        }

        //����һ����Ԫ��: 1,���ݿ��и������ѵ�Ԫ��. 2,����Ԫ����Ϣ���Ƶ�������
        private void CellProcessing(XmlNode rowNode, int rowIndex, DoubleLinkedList<XmlNode>[] cellMatrix)
        {

            XmlNodeList cellList = this.GetCellList(rowNode, @"��:��Ԫ��_41BE");

            //������У����ݿ��и������ѵ�Ԫ��
            foreach (XmlNode cellNode in cellList)
            {
                DivideCell(rowNode, cellNode);
            }
            //ˢ��
            cellList = this.GetCellList(rowNode, @"��:��Ԫ��_41BE");
            //����Ԫ����Ϣ���Ƶ�������
            CopyCell(cellList, cellMatrix, rowIndex);
        }

        //���д������п��еĵ�Ԫ���������
        private void CellVerticalSpanProcessing(DoubleLinkedList<XmlNode>[] cellMatrix)
        {
            for (int rowIndex = 0; rowIndex < cellMatrix.Length; rowIndex++)
            {
                List<DoubleLinkedList<XmlNode>.DoubleLinkedNode> cellList = cellMatrix[rowIndex].ToArray();

                //������У����ݿ��и������ѵ�Ԫ��
                for (int i = 0; i < cellList.Count; i++)
                {
                    XmlNode cellNode = cellList[i].Value;
                    XmlNode vMerge = this.GetSingleNode(cellNode, @".//��:����_41A6");
                    if (vMerge == null) continue;

                    int vMergeCount = vMerge.InnerText == null ? 0 : Convert.ToInt32(vMerge.InnerText);
                    vMerge.ParentNode.InsertAfter(this.CreatNewNode("��:���кϲ���ʼ", wordNamespace), vMerge);

                    for (int currentRowIndex = rowIndex + 1; currentRowIndex < vMergeCount + rowIndex; currentRowIndex++)
                    {
                        if (currentRowIndex >= cellMatrix.Length) break;

                        XmlNode newCell = this.CreateRowSpanNode();

                        //if (currentRowIndex == rowIndex + vMergeCount - 1)
                        //{
                        //    newCell.AppendChild(this.CreatNewNode("��:���кϲ�����", wordNamespace));
                        //}

                        cellMatrix[currentRowIndex].Insert(i + 1, newCell);
                    }
                }
            }
        }

        //������У����ݿ��и������ѵ�Ԫ��
        private void DivideCell(XmlNode rowNode, XmlNode cellNode)
        {
            XmlNode vMerge = this.GetSingleNode(cellNode, @".//��:����_41A6");
            XmlNode hMerge = this.GetSingleNode(cellNode, @".//��:����_41A7");

            if (hMerge == null) return;

            int hMergeCount = hMerge.InnerText == null ? 0 : Convert.ToInt32(hMerge.InnerText);

            for (int i = 1; i < hMergeCount; i++)
            {

                XmlNode newCell = this.CreatNewNode("��:��Ԫ��_41BE", wordNamespace);//�½���Ԫ��
                newCell.AppendChild(this.CreatNewNode("��:�½��ĵ�Ԫ��", wordNamespace));
                newCell.AppendChild(this.CreatNewNode("��:����ռλ��Ԫ��", wordNamespace));

                if (vMerge != null)//����п�����Ϣ�����Ƶ��½���Ԫ����
                {
                    XmlNode cellAttribute = newCell.AppendChild(this.CreatNewNode("��:��Ԫ������_41B7", wordNamespace));
                    cellAttribute.AppendChild(vMerge.Clone());
                    newCell.AppendChild(cellAttribute);
                }

                //���ѵ�Ԫ��
                rowNode.InsertAfter(newCell, cellNode);
            }

        }

        //����Ԫ����Ϣ���Ƶ�������
        private void CopyCell(XmlNodeList cellList, DoubleLinkedList<XmlNode>[] cellMatrix, int rowIndex)
        {
            foreach (XmlNode cell in cellList)
            {
                cellMatrix[rowIndex].Append(cell);
            }
        }

        //�½�����ռλ��Ԫ��
        private XmlNode CreateRowSpanNode()
        {
            XmlNode newCell = this.CreatNewNode("��:��Ԫ��_41BE", wordNamespace);//�½���Ԫ��
            newCell.AppendChild(this.CreatNewNode("��:�½��ĵ�Ԫ��", wordNamespace));
            newCell.AppendChild(this.CreatNewNode("��:����ռλ��Ԫ��", wordNamespace));

            return newCell;
        }

        protected override XmlNamespaceManager GetXmlNamespaceManager(XmlNameTable nameTable)
        {
            XmlNamespaceManager nm = new XmlNamespaceManager(nameTable);
            nm.AddNamespace("��", "http://schemas.uof.org/cn/2009/wordproc");
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2009/uof");
            return nm;
        }

    }

    public abstract class TableProcessing
    {

        protected XmlNamespaceManager XmlSourceNamespaceManager
        {
            get;
            set;
        }

        public XmlDocument XmlSource
        {
            get;
            set;
        }

        protected abstract XmlNamespaceManager GetXmlNamespaceManager(XmlNameTable nameTable);

        protected XmlNodeList GetCellList(XmlNode rowNode, string cellMarkName)
        {
            return rowNode.SelectNodes(cellMarkName, XmlSourceNamespaceManager);
        }

        protected XmlNodeList GetRowList(XmlNode tableNode, string rowMarkName)
        {
            return tableNode.SelectNodes(rowMarkName, XmlSourceNamespaceManager);
        }

        protected XmlNodeList GetTableList(string tableMarkName)
        {
            return XmlSource.SelectNodes(tableMarkName, XmlSourceNamespaceManager);
        }

        public abstract void Processing(string fileName);

        public TableProcessing(string fileName)
        {
            XmlSource = new XmlDocument();
            XmlSource.Load(fileName);
            XmlSourceNamespaceManager = this.GetXmlNamespaceManager(XmlSource.NameTable);
        }

        protected XmlNode GetSingleNode(XmlNode parentNode, string nodeName)
        {
            return parentNode.SelectSingleNode(nodeName, XmlSourceNamespaceManager);
        }

        protected XmlNode CreatNewNode(string nodeName, string nameSpace)
        {
            return this.XmlSource.CreateNode(XmlNodeType.Element, nodeName, nameSpace);
        }

        protected XmlAttribute CreateNewAttribute(string nodeName)
        {
            return this.XmlSource.CreateAttribute(nodeName);
        }
    }


    public class DoubleLinkedList<T>
    {
        /// <summary>
        /// ˫������ڵ�
        /// </summary>
        public class DoubleLinkedNode
        {
            public T Value
            {
                set;
                get;
            }

            public DoubleLinkedNode Next
            {
                set;
                get;
            }

            public DoubleLinkedNode Previous
            {
                set;
                get;
            }

            public DoubleLinkedNode(T value, DoubleLinkedNode next, DoubleLinkedNode previous)
            {
                this.Value = value;
                this.Next = next;
                this.Previous = previous;
            }

            public DoubleLinkedNode()
                : this(default(T), null, null)
            {

            }
        }

        public DoubleLinkedNode Head
        {
            set;
            get;
        }

        public DoubleLinkedNode Last
        {
            set;
            get;
        }

        public int Count
        {
            set;
            get;
        }

        public DoubleLinkedList()
        {
            this.Head = new DoubleLinkedNode();
            this.Last = this.Head;
            this.Count = 0;
        }

        public void Insert(int position, T value)
        {
            DoubleLinkedNode newNode = new DoubleLinkedNode(value, null, null);

            if (position < 1)
            {
                position = 1;
            }

            if (Count == 0)
            {
                this.Head.Next = newNode;
                newNode.Previous = this.Head;
                this.Last = newNode;
            }
            else if (position > Count)
            {
                Last.Next = newNode;
                newNode.Previous = Last;
                Last = newNode;
            }
            else
            {
                DoubleLinkedNode p = this.Head;
                int currentIndex = 0;
                while (currentIndex != position)
                {
                    p = p.Next;
                    currentIndex++;
                }
                newNode.Next = p;
                newNode.Previous = p.Previous;
                newNode.Previous.Next = newNode;
                p.Previous = newNode;
            }

            Count++;
        }


        public void Append(T value)
        {
            this.Insert(this.Count + 1, value);
        }

        public List<DoubleLinkedNode> ToArray()
        {
            List<DoubleLinkedNode> list = new List<DoubleLinkedNode>();
            DoubleLinkedNode p = this.Head;
            while (p.Next != null)
            {
                p = p.Next;
                list.Add(p);
            }

            return list;
        }

        public List<DoubleLinkedNode> Traversal()
        {
            List<DoubleLinkedNode> list = new List<DoubleLinkedNode>();
            DoubleLinkedNode p = this.Head;
            while (p.Next != null)
            {
                p = p.Next;
                Console.Write(p.Value + " ,");
                list.Add(p);
            }

            return list;
        }

        public DoubleLinkedNode this[int index]
        {
            get
            {
                if (index < 0 || index >= this.Count)
                {
                    return null;
                }

                DoubleLinkedNode p = this.Head.Next;
                int i = 0;
                while (i != index && p != null)
                {
                    p = p.Next;
                    i++;
                }

                return p;
            }
        }

        public void Delete(int index)
        {

        }

        public void Delete(DoubleLinkedNode node)
        {

        }

        public static void Test()
        {
            DoubleLinkedList<int> list = new DoubleLinkedList<int>();
            list.Insert(1, 1);
            list.Insert(0, 2);
            list.Insert(3, 3);
            list.Insert(2, 4);
            for (int i = 0; i < 5; i++) list.Append(i + 5);
            Console.WriteLine(list.Count);
            list.Traversal();
            //Console.WriteLine(list.Head.Next.Value);

            Console.WriteLine(list[0].Value);
        }
    }

}
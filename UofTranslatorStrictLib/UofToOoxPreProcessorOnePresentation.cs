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
    /// The first step pre process in uof->OOX of Presentation
    /// </summary>
    /// <author>huashuran&linyaohu</author>
    class UofToOoxPreProcessorOnePresentation : AbstractProcessor
    {

        private static ILog logger = LogManager.GetLogger(typeof(UofToOoxPreProcessorOneWord).FullName);

        public override bool transform()
        {
            // pretreat to the input .uof file of Presentation
            //preMethod(inputFile);

            FileStream fs = null;
            //XmlUrlResolver resourceResolver = null;
            XmlReader xr = null;

            string extractPath = Path.GetDirectoryName(outputFile) + Path.AltDirectorySeparatorChar;
            string prefix = this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + TranslatorConstants.UOFToOOX_POWERPOINT_LOCATION;
            string uof2ooxpre=extractPath+"uof2ooxpre.xml"; //  调用此部分出异常 所以改为下面语句
            string picture_xml = "data";
            // string uof2ooxpre = extractPath + "tmpDoc1.xml";

            try
            {
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
                xr = XmlReader.Create(extractPath + "uof2ooxpre.xslt");

                XPathDocument doc = new XPathDocument(extractPath + @"content.xml");
                XslCompiledTransform transFrom = new XslCompiledTransform();
                XsltSettings setting = new XsltSettings(true, false);
                XmlUrlResolver xur = new XmlUrlResolver();
                transFrom.Load(xr, setting, xur);
                XPathNavigator nav = ((IXPathNavigable)doc).CreateNavigator();
               // fs = new FileStream(outputFile, FileMode.Create);
                fs = new FileStream(uof2ooxpre, FileMode.Create);
                transFrom.Transform(nav, null, fs);
                fs.Close();

               preMethod(uof2ooxpre);

                //ole对象lyy
               if (Directory.Exists(extractPath + "drawings") && Directory.Exists(extractPath + "embeddings"))
               {
                    string tmpOle = extractPath + Path.AltDirectorySeparatorChar + "tmpOle.xml";
                   XmlDocument xdoc = new XmlDocument();
                   xdoc.Load(outputFile);
                   XmlNameTable nt = xdoc.NameTable;
                   XmlNamespaceManager nm = new XmlNamespaceManager(nt);
                   nm.AddNamespace("w", TranslatorConstants.XMLNS_W);
                   nm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);

                   xdoc = OlePretreatment(xdoc, "uof:UOF", extractPath, nm);
                   xdoc.Save(tmpOle);
                   // OutputFilename = tmpPic;
                   outputFile = tmpOle;
               }

               if (Directory.Exists(extractPath + picture_xml))
               {
                   string tmpPic = extractPath + Path.AltDirectorySeparatorChar + "tmpPic.xml";
                   XmlDocument xdoc = new XmlDocument();
                   xdoc.Load(outputFile);
                   XmlNameTable nt = xdoc.NameTable;
                   XmlNamespaceManager nm = new XmlNamespaceManager(nt);
                   nm.AddNamespace("w", TranslatorConstants.XMLNS_W);
                   nm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);

                   xdoc = PicPretreatment(xdoc, "uof:UOF", extractPath + picture_xml, nm);
                   xdoc.Save(tmpPic);
                   // OutputFilename = tmpPic;
                   outputFile = tmpPic;

               }
                //公式
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
                return true;


            }
            catch (Exception ex)
            {
                logger.Error("Fail in Uof2.0 to OOX pretreatment1: " + ex.Message);
                logger.Error(ex.StackTrace);
                throw new Exception("Fail in Uof2.0 to OOX pretreatment1 of Presentation");
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }

        }


        #region the implementation of preMethod()

        static XmlNodeList paraStyles;

        private void preMethod(string FileLoc)
        {
            XmlDocument xmlDoc;
            XmlNameTable table;
            XmlNamespaceManager nm = null;

            //string tempdoc = Path.GetTempPath().ToString() + "tempdoc.xml";
            XmlTextWriter resultWriter;
          //  XmlTextReader txtReader = new XmlTextReader(FileLoc);
            xmlDoc = new XmlDocument();
            xmlDoc.Load(FileLoc);
            // txtReader.Close();


            table = xmlDoc.NameTable;
            nm = new XmlNamespaceManager(table);
            nm.AddNamespace("", "http://uof2ooxPre");
            nm.AddNamespace("字", TranslatorConstants.XMLNS_UOFWORDPROC);
            nm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);
            nm.AddNamespace("图", TranslatorConstants.XMLNS_UOFGRAPH);
            nm.AddNamespace("表", TranslatorConstants.XMLNS_UOFSPREADSHEET);
            nm.AddNamespace("演",TranslatorConstants.XMLNS_UOFPRESENTATION);
            nm.AddNamespace("a", TranslatorConstants.XMLNS_A);
            nm.AddNamespace("app", TranslatorConstants.XMLNS_APP);
            nm.AddNamespace("cp", TranslatorConstants.XMLNS_CP);
            nm.AddNamespace("dc", TranslatorConstants.XMLNS_DC);
            nm.AddNamespace("dcmitype", TranslatorConstants.XMLNS_DCMITYPE);
            nm.AddNamespace("dcterms", TranslatorConstants.XMLNS_DCTERMS);
            nm.AddNamespace("fo", TranslatorConstants.XMLNS_FO);
            nm.AddNamespace("p", TranslatorConstants.XMLNS_P);
            nm.AddNamespace("r",TranslatorConstants.XMLNS_R);
            nm.AddNamespace("rel", TranslatorConstants.XMLNS_REL);
            nm.AddNamespace("cus", TranslatorConstants.XMLNS_CUS);
            nm.AddNamespace("vt", TranslatorConstants.XMLNS_VT);
            nm.AddNamespace("式样", TranslatorConstants.XMLNS_UOFSTYLES);
            nm.AddNamespace("规则", TranslatorConstants.XMLNS_UOFRULES);
            nm.AddNamespace("图形", TranslatorConstants.XMLNS_UOFGRAPHICS);
            string path = "uof:UOF";
            XmlNodeList eleList = xmlDoc.SelectNodes(path, nm);
            XmlElement uof = (XmlElement)eleList[0];


            //2011-4-20 罗文甜 把母版标题标识符OBJ00001修改为OBJ90001。修改母版标题动画不识别问题
            XmlNodeList shapeLists = xmlDoc.SelectNodes("//图:图形_8062", nm);
            XmlNodeList shapeMaos = xmlDoc.SelectNodes("//uof:锚点_C644", nm);
            XmlNodeList animaShapes = xmlDoc.SelectNodes("//演:序列_6B1B", nm);
            XmlNodeList touchShapes = xmlDoc.SelectNodes("//演:定时_6B2E", nm);
            if (shapeLists != null)
            {
                foreach (XmlNode shapeList in shapeLists)
                {
                    if (shapeList.Attributes.GetNamedItem("标识符_804B").Value == "Obj00001")
                    {
                        shapeList.Attributes.GetNamedItem("标识符_804B").Value = "Obj90001";
                    }
                }
            }
            if (shapeMaos != null)
            {
                foreach (XmlNode shapeMao in shapeMaos)
                {
                    if (shapeMao.Attributes.GetNamedItem("图形引用_C62E").Value == "Obj00001")
                        shapeMao.Attributes.GetNamedItem("图形引用_C62E").Value = "Obj90001";
                }
            }
            if (animaShapes != null)
            {
                foreach (XmlNode animaShape in animaShapes)
                {
                    if (((XmlElement)animaShape).HasAttribute("对象引用_6C28") && animaShape.Attributes.GetNamedItem("对象引用_6C28").Value == "Obj00001")
                        animaShape.Attributes.GetNamedItem("对象引用_6C28").Value = "Obj90001";
                }
            }
            if (touchShapes != null)
            {
                foreach (XmlNode touchShape in touchShapes)
                {
                    if (((XmlElement)touchShape).HasAttribute("触发对象引用_6B34") && touchShape.Attributes.GetNamedItem("触发对象引用_6B34").Value == "Obj00001")
                        touchShape.Attributes.GetNamedItem("触发对象引用_6B34").Value = "Obj90001";
                }
            }
            //end by 罗文甜 2011-4-20

            //09.11.26 马有旭 添加 根据扩展区内容在所有母版锚点中加入页眉、页脚、时间日期内容。
            XmlNodeList masters = xmlDoc.SelectNodes("//uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D", nm);


            //----将基式样引用中的段落属性提取出来加至后继段落式样中------暂不处理
            //基式样引用的段落式样，看相应标识符的被引用式样的有无段落式样引用
            //把每个段落式样中的基式样信息加到段落式样中
            //paraStyles = xmlDoc.SelectNodes("//式样:段落式样_9912|//uof:段落式样", nm);
            paraStyles = xmlDoc.SelectNodes("//式样:段落式样_9912|//式样:段落式样_9905", nm); // 没有找到uof：段落式样相应的元素 
            for (int s = 0; s < paraStyles.Count; s++)
            {
                addProp((XmlElement)paraStyles[s]);
            }


            //为区域开始和区域结束间的 句 添加超级链接
            XmlNodeList textBody = xmlDoc.SelectNodes("uof:UOF/uof:对象集/图形:图形集_7C00/图:图形_8062/图:文本_803C", nm);
            for (int t = 0; t < textBody.Count; t++)
            {
                XmlNodeList run = textBody[t].SelectNodes("图:内容_8043/字:段落_416B/字:句_419D", nm);
                //XmlNodeList blockStart = textBody[t].SelectNodes("字:段落/字:句/字:区域开始", nm);
                //XmlNodeList blockEnd = textBody[t].SelectNodes("字:段落/字:句/字:区域结束", nm);
                //for (int b = 0; b < blockStart.Count; b++)
                //{
                for (int r = 0; r < run.Count; r++)
                {
                    if (run[r].SelectNodes("字:区域开始_4165[@类型_413B='hyperlink']", nm).Count != 0)
                    {
                        if (run[r].SelectNodes("字:区域结束_4167", nm).Count == 0)
                        {
                            run[r + 1].AppendChild(run[r].SelectNodes("字:区域开始_4165", nm)[0].Clone());
                        }
                    }
                }

            }

           
            XmlNodeList ls = xmlDoc.SelectNodes("/uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652", nm);
            if (ls != null)
            {
                for (int i = 0; i < ls.Count; ++i)
                {
                    StringBuilder sb = new StringBuilder();
                    sb.Append("/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F[@页面版式引用_6B27='");
                    sb.Append(ls[i].Attributes["标识符_6B0D"].Value).Append("']");  // 改为页面版式的标识符 李娟 2012.02.14
                    XmlNodeList lslides = xmlDoc.SelectNodes(sb.ToString(), nm);
                    if (lslides.Count > 0)
                    {
                        sb = new StringBuilder();
                        sb.Append("/uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@标识符_6BE8='");
                        sb.Append(lslides[0].Attributes["母版引用_6B26"].Value);
                        sb.Append("']");
                        XmlNode master = xmlDoc.SelectSingleNode(sb.ToString(), nm);
                        XmlNodeList hfanchors = master.SelectNodes("uof:锚点_C644[uof:占位符_C626/@类型_C627='header' or uof:占位符_C626/@类型_C627='date' or uof:占位符_C626/@类型_C627='footer' or uof:占位符_C626/@类型_C627='number']", nm);
                        if (hfanchors != null)
                        {
                            foreach (XmlNode hfanchor in hfanchors)
                            {                           
                                ls[i].AppendChild(hfanchor.Clone());
                            }
                        }
                    }
                }
            }
            XmlNodeList slides = xmlDoc.SelectNodes("uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F", nm);

            //layoutOriNum用于记录初始状态页面版式集中layout的数目，完成layout复制后可将原始layout删除
            int layoutOriNum = xmlDoc.SelectNodes("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652", nm).Count;
            if (layoutOriNum != 0)
            {
                //通过幻灯片中的属性为版式添加子元素，表示所有对其引用的母版的信息
                for (int i = 0; i < slides.Count; i++)
                {
                    XmlElement slideRfLayout = xmlDoc.CreateElement("slide");
                    slideRfLayout.InnerText = slides[i].Attributes["标识符_6B0A"].Value.ToString();

                    if (slides[i].Attributes["母版引用_6B26"] != null && slides[i].Attributes["页面版式引用_6B27"] != null)
                    {
                        string masterRf = (string)slides[i].Attributes["母版引用_6B26"].Value;
                        string layoutRf = (string)slides[i].Attributes["页面版式引用_6B27"].Value;

                        //在layout中添加元素，用于记录跟该版式相关的所有的slide及master的信息
                        if (xmlDoc.SelectSingleNode("//规则:页面版式_B652[@标识符_6B0D='" + layoutRf + "']/master[@masterRef='" + masterRf + "']", nm) == null)
                        {
                            XmlElement masterRfLayout = xmlDoc.CreateElement("master");
                            masterRfLayout.SetAttribute("masterRef", masterRf.ToString());

                            masterRfLayout.AppendChild((XmlNode)slideRfLayout);

                            xmlDoc.SelectSingleNode("//规则:页面版式_B652[@标识符_6B0D='" + layoutRf + "']", nm).AppendChild((XmlNode)masterRfLayout);

                        }
                        else
                        {
                            xmlDoc.SelectSingleNode("//规则:页面版式_B652[@标识符_6B0D='" + layoutRf + "']/master[@masterRef='" + masterRf + "']", nm).AppendChild((XmlNode)slideRfLayout);
                        }

                    }
                    if (slides[i].Attributes["母版引用_6B26"] != null && slides[i].Attributes["页面版式引用_6B27"] == null)
                    {
                        string masterRf = (string)slides[i].Attributes["母版引用_6B26"].Value;

                        //通过layout的标识符，判断是否存在这样的空白layout，隶属于该母版之下。若无，创建之，并在相应的母版中加入版式引用。
                        if (xmlDoc.SelectSingleNode("//规则:页面版式_B652[@标识符_6B0D='lyt" + masterRf + "']", nm) == null)
                        {
                            XmlElement blankLayout = xmlDoc.CreateElement("规则", "页面版式_B652", TranslatorConstants.XMLNS_UOFPRESENTATION);
                            blankLayout.SetAttribute("标识符_6B0D", TranslatorConstants.XMLNS_UOFPRESENTATION, "lyt" + masterRf);

                            XmlElement masterRfLayout = xmlDoc.CreateElement("master");
                            masterRfLayout.SetAttribute("masterRef", masterRf.ToString());

                            masterRfLayout.AppendChild((XmlNode)slideRfLayout);

                            blankLayout.AppendChild((XmlNode)masterRfLayout);

                            XmlElement layoutParent = xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651", nm) as XmlElement;
                            // xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/演:页面版式集", nm).AppendChild((XmlNode)blankLayout);
                            if (xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651", nm) == null)
                            {
                                XmlElement lyoutSet = xmlDoc.CreateElement("规则", "页面版式集_B651", TranslatorConstants.XMLNS_UOFPRESENTATION);
                                //XmlAttribute lyoutSetAttr = xmlDoc.CreateAttribute("uof", "locID", "http://schemas.uof.org/cn/2003/uof");
                                //lyoutSetAttr.Value = "p0017";
                                //lyoutSet.Attributes.SetNamedItem((XmlNode)lyoutSetAttr);                 //uof2.0 无
                                xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D", nm).PrependChild(lyoutSet);
                            }
                            xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651", nm).AppendChild((XmlNode)blankLayout);

                            XmlElement mst = xmlDoc.SelectSingleNode("//演:母版_6C0D[@标识符_6BE8='" + masterRf + "']", nm) as XmlElement;

                            //相应的母版中加入版式引用
                            XmlElement lytRef = xmlDoc.CreateElement("演", "页面版式引用", TranslatorConstants.XMLNS_UOFPRESENTATION);
                            //XmlElement lytRef = xmlDoc.CreateElement("页面版式引用_6BEC", TranslatorConstants.XMLNS_UOFPRESENTATION);
                            lytRef.InnerText = "lyt" + masterRf;
                            mst.AppendChild((XmlNode)lytRef);
                        }
                        XmlElement sld = slides[i] as XmlElement;

                        //幻灯片只有一个版式引用
                        sld.SetAttribute("页面版式引用_6B27", TranslatorConstants.XMLNS_UOFPRESENTATION, "lyt" + masterRf);
                    }

                }

                XmlNodeList layouts = xmlDoc.SelectNodes("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652", nm);
                masters = xmlDoc.SelectNodes("uof:UOF/uof:演示文稿/演:主体/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D", nm);

                //一个版式有几个母版引用就对该版式复制几次
                for (int j = 0; j < layoutOriNum; j++)
                {
                    int lytMstCount = layouts[j].SelectNodes("master").Count;

                    for (int m = 0; m < lytMstCount; m++)
                    {
                        XmlElement layoutClone = (XmlElement)layouts[j].Clone();
                        string oldID = layoutClone.GetAttribute("标识符_6B0D");
                        layoutClone.SetAttribute("标识符_6B0D", m + "CL" + oldID);
                        xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651", nm).AppendChild((XmlNode)layoutClone);

                        int lytMstSldCount = layouts[j].SelectNodes("master")[0].ChildNodes.Count;
                        for (int n = 0; n < lytMstSldCount; n++)
                        {
                            XmlElement sld = xmlDoc.SelectSingleNode("//演:幻灯片_6C0F[@标识符_6B0A='" + layouts[j].SelectNodes("master")[0].ChildNodes[0].InnerText.ToString() + "']", nm) as XmlElement;
                            //幻灯片只有一个版式引用
                            sld.SetAttribute("页面版式引用_6B27", m + "CL" + oldID);
                            layouts[j].SelectNodes("master")[0].RemoveChild(layouts[j].SelectNodes("master")[0].ChildNodes[0]);
                        }
                        XmlElement mst = xmlDoc.SelectSingleNode("//演:母版_6C0D[@标识符_6BE8='" + layouts[j].SelectNodes("master")[0].Attributes[0].Value.ToString() + "']", nm) as XmlElement;
                        //母版可以有多个版式引用
                        XmlElement lytRef = xmlDoc.CreateElement("演", "页面版式引用", TranslatorConstants.XMLNS_UOFPRESENTATION);
                        lytRef.InnerText = m + "CL" + oldID;
                        mst.AppendChild((XmlNode)lytRef);

                        layouts[j].RemoveChild(layouts[j].SelectNodes("master")[0]);
                    }
                }

                //删除原始layouts
                for (int i = 0; i < layoutOriNum; i++)
                {
                    XmlElement layoutParent = xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651", nm) as XmlElement;
                    xmlDoc.SelectSingleNode("uof:UOF/uof:演示文稿/演:公用处理规则/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651", nm).RemoveChild(layoutParent.FirstChild);
                }
            }

                //modify by linyaohu
                XmlNodeList anchors = xmlDoc.SelectNodes("//uof:锚点_C644", nm);
                string assembleList;
                string[] splitAssembleList;
                LinkedList<string> allInSplitedAssembleList;
                allInSplitedAssembleList = new LinkedList<string>();
                LinkedList<string>.Enumerator enumer = allInSplitedAssembleList.GetEnumerator();
                int nodeAttributeNum;
                bool flag;//标识链接状态
                int currentNodePositionInLinkedList;
                for (int i = 0; i < anchors.Count; i++)
                {
                    string selectNode = "uof:UOF/uof:对象集/图形:图形集_7C00/图:图形_8062[@标识符_804B='" + anchors[i].Attributes["图形引用_C62E"].Value + "']";
                    if (xmlDoc.SelectSingleNode(selectNode, nm) != null)
                    {
                        anchors[i].AppendChild(xmlDoc.SelectSingleNode(selectNode, nm).Clone());
                        allInSplitedAssembleList.Clear();
                        enumer = allInSplitedAssembleList.GetEnumerator();
                        flag = false;
                        currentNodePositionInLinkedList = 0;
                        nodeAttributeNum = 0;
                        while (true)
                        {
                            if (flag)
                            {
                                while (enumer.MoveNext())
                                {
                                    currentNodePositionInLinkedList++;

                                    if (enumer.Current.ToString() != "")
                                    {
                                        string objName=enumer.Current.ToString();
                                        string newObjName = objName;
                                        if (objName.Equals("Obj00001"))
                                        {
                                            newObjName = "Obj90001";
                                        }
                                       

                                       // selectNode = "uof:UOF/uof:对象集/图形:图形集_7C00/图:图形_8062[@标识符_804B='" + enumer.Current.ToString() + "']";
                                        selectNode = "uof:UOF/uof:对象集/图形:图形集_7C00/图:图形_8062[@标识符_804B='" + newObjName + "']";
                                        
                                        anchors[i].AppendChild(xmlDoc.SelectSingleNode(selectNode, nm).Clone());
                                        for (nodeAttributeNum = 0; nodeAttributeNum < xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Count; nodeAttributeNum++)
                                        {
                                            if (xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Item(nodeAttributeNum).Name.ToString() == "组合列表_8064")
                                            {
                                                assembleList = xmlDoc.SelectSingleNode(selectNode, nm).Attributes.GetNamedItem("组合列表_8064").Value.ToString();
                                                splitAssembleList = assembleList.Split(new char[1] { ' ' });
                                                foreach (string s in splitAssembleList)
                                                {
                                                    allInSplitedAssembleList.AddLast(s);
                                                }
                                                //链表已更改，重新初始化枚举器
                                                enumer = allInSplitedAssembleList.GetEnumerator();
                                                //把枚举器定位到已遍历结点
                                                for (int current = 0; current < currentNodePositionInLinkedList; current++)
                                                {
                                                    enumer.MoveNext();
                                                }
                                            }
                                        }
                                    }
                                }
                                //已遍历完链表中所有节点，跳出循环
                                flag = false;
                                break;
                            }
                            else if (!flag)
                            {//初始链表
                                for (nodeAttributeNum = 0; nodeAttributeNum < xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Count; nodeAttributeNum++)
                                {
                                    if (xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Item(nodeAttributeNum).Name.ToString() == "组合列表_8064")
                                    {
                                        assembleList = xmlDoc.SelectSingleNode(selectNode, nm).Attributes.GetNamedItem("组合列表_8064").Value.ToString();
                                        splitAssembleList = assembleList.Split(new char[1] { ' ' });
                                        foreach (string s in splitAssembleList)
                                        {
                                            allInSplitedAssembleList.AddLast(s);
                                        }
                                        flag = true;
                                        enumer = allInSplitedAssembleList.GetEnumerator();
                                    }
                                }
                            }
                            if (!flag)
                                break;
                        }


                    }
                }

                // xmlDoc.SelectSingleNode("uof:UOF/uof:对象集", nm).RemoveAll();
                //modify by linyaohu
                XmlNode objectSet = xmlDoc.SelectSingleNode("uof:UOF/uof:对象集", nm);
                XmlNodeList geoList = objectSet.SelectNodes("图形:图形集_7C00", nm);
                foreach (XmlNode geoNode in geoList)
                {
                    xmlDoc.SelectSingleNode("uof:UOF/uof:对象集", nm).RemoveChild(geoNode);
                }

                //added by hsr, 090222
                //在<uof:图形/>下写入段落式样集<lststyle/>
                XmlNodeList shapes = xmlDoc.SelectNodes("//图:图形_8062", nm);
                for (int i = 0; i < shapes.Count; i++)
                {
                    XmlNodeList pStyles = shapes[i].SelectNodes("图:文本_803C/图:内容_8043/字:段落_416B/字:段落属性_419B[@式样引用_419C]", nm);
                    XmlElement lstStyle = xmlDoc.CreateElement("演", "lstStyle", TranslatorConstants.XMLNS_UOFPRESENTATION);

                    for (int j = 0; j < pStyles.Count; j++)
                    {
                        XmlNode styleRef = (xmlDoc.SelectSingleNode("/uof:UOF/uof:式样集/式样:式样集_990B/式样:段落式样集_9911/式样:段落式样_9912[@标识符_4100='" + pStyles[j].Attributes["式样引用_419C"].Value + "'] | /uof:UOF/uof:式样集/式样:式样集_990B/式样:文本式样集_9913/式样:文本式样_9914/式样:段落式样_9905[@标识符_4100='" + pStyles[j].Attributes["式样引用_419C"].Value + "']", nm));
                        if (styleRef != null)
                        {
                            //2011-3-18罗文甜，修改bug，式样信息丢失。这个问题找了好久。。
                            XmlNode styleReltem = styleRef.Clone();
                            lstStyle.AppendChild(styleReltem);
                        }
                        // lstStyle.AppendChild(xmlDoc.SelectSingleNode("uof:UOF/uof:式样集/uof:段落式样[@字:标识符='" + pStyles[j].Attributes["字:式样引用"].Value + "']|uof:UOF/uof:演示文稿/演:公用处理规则/演:文本式样集/演:文本式样/演:段落式样[@字:标识符='" + pStyles[j].Attributes["字:式样引用"].Value + "']", nm).Clone());
                        //lstStyle.AppendChild(xmlDoc.SelectSingleNode("//uof:段落式样[@字:标识符='" + pStyles[j].Attributes["字:式样引用"].Value + "']|//演:段落式样[@字:标识符='" + pStyles[j].Attributes["字:式样引用"].Value + "']", nm).Clone());
                    }

                    shapes[i].AppendChild((XmlNode)lstStyle);

                }

                //2011-1-5 罗文甜 自由曲线的路径预处理
                // 修改路径 李娟 2012.02.14
                XmlNodeList shapeKeyCoordinates = xmlDoc.SelectNodes("//图:路径_801C", nm);
                if (shapeKeyCoordinates != null)
                {
                    foreach (XmlNode shapeKeyCoordinate in shapeKeyCoordinates)
                    {
                        string keyPath;
                        //keyPath = shapeKeyCoordinate.Attributes.GetNamedItem("图:路径值_8069").Value;

                        keyPath = shapeKeyCoordinate.InnerText;
                        string[] PathStr = keyPath.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                        XmlElement tempNode, childNode, childNode1, childNode2;
                        for (int i = 0; i < PathStr.Length; i++)
                        {
                            if (PathStr[i] == "M")
                            {
                                tempNode = xmlDoc.CreateElement("M");
                                childNode = xmlDoc.CreateElement("pt");

                                childNode.SetAttribute("x", PathStr[i + 1]);
                                childNode.SetAttribute("y", PathStr[i + 2]);
                                tempNode.AppendChild(childNode);
                                shapeKeyCoordinate.AppendChild(tempNode);
                            }
                            if (PathStr[i] == "C")
                            {
                                tempNode = xmlDoc.CreateElement("C");
                                childNode = xmlDoc.CreateElement("pt");

                                childNode.SetAttribute("x", PathStr[i + 1]);
                                childNode.SetAttribute("y", PathStr[i + 2]);
                                childNode1 = xmlDoc.CreateElement("pt");
                                childNode1.SetAttribute("x", PathStr[i + 3]);
                                childNode1.SetAttribute("y", PathStr[i + 4]);
                                childNode2 = xmlDoc.CreateElement("pt");
                                childNode2.SetAttribute("x", PathStr[i + 5]);
                                childNode2.SetAttribute("y", PathStr[i + 6]);
                                tempNode.AppendChild(childNode);
                                tempNode.AppendChild(childNode1);
                                tempNode.AppendChild(childNode2);
                                shapeKeyCoordinate.AppendChild(tempNode);
                            }
                            if (PathStr[i] == "L")
                            {
                                tempNode = xmlDoc.CreateElement("L");
                                childNode = xmlDoc.CreateElement("pt");
                                childNode.SetAttribute("x", PathStr[i + 1]);
                                childNode.SetAttribute("y", PathStr[i + 2]);
                                tempNode.AppendChild(childNode);
                                shapeKeyCoordinate.AppendChild(tempNode);
                            }
                        }
                    }
                }
                //2011-01-12罗文甜，阴影预处理
                XmlNodeList shapeShades = xmlDoc.SelectNodes("//图:阴影_8051/uof:偏移量_C61B", nm);
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
                                double angValueTe = Math.Atan2(Math.Abs(x), Math.Abs(y)) * 180 / 3.1415;
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
                                XmlElement ang = xmlDoc.CreateElement("dir");
                                ang.InnerText = Convert.ToString(angValue);
                                shapeShade.AppendChild(ang);
                                XmlElement distance = xmlDoc.CreateElement("dist");
                                distance.InnerText = Convert.ToString(distanceValue);
                                shapeShade.AppendChild(distance);
                            }
                        }

                    }
                }
            

            //马有旭 2010.04.24 动画预处理
            slides = xmlDoc.SelectNodes("//演:幻灯片_6C0F|//演:母版_6C0D", nm);
            foreach (XmlNode slide in slides)
            {
                XmlNodeList seqs = slide.SelectNodes("演:动画_6B1A/演:序列_6B1B[演:效果_6B40/*]", nm);
                XmlNode newAnim = xmlDoc.CreateElement("演:动画_6B1A", nm.LookupNamespace("演"));
                XmlNode anim = slide.SelectSingleNode("演:动画_6B1A", nm);
                XmlNode outterSeq = null;
                XmlNode preNode = null;
                int id = 4;
                long allDelay = 0;
                long thisDelay = 0;
                //bool with = false;
                for (int i = 0; i < seqs.Count; ++i)
                {
                    String animevent = seqs[i].SelectSingleNode("演:定时_6B2E", nm).Attributes["事件_6B2F"].Value;
                    String speed = seqs[i].SelectSingleNode("演:定时_6B2E", nm).Attributes["延时_6B30"].Value;  // uof2.0 定时无属性速度 改为延迟 有待验证 李娟
                    seqs[i].Attributes.Append(xmlDoc.CreateAttribute("id"));
                    seqs[i].Attributes["id"].Value = id.ToString();
                    if (animevent.Equals("on-click"))
                    {
                        //with = false;
                        allDelay = tranUofSpeed(speed);
                        thisDelay = allDelay;
                        preNode = seqs[i];
                        if (outterSeq != null)
                        {
                            newAnim.InsertBefore(outterSeq, newAnim.LastChild);
                        }
                        outterSeq = xmlDoc.CreateElement("演:序列_6B1B", nm.LookupNamespace("演"));
                        outterSeq.InsertAfter(seqs[i], outterSeq.LastChild);
                    }
                    else if (animevent.Equals("after-previous"))
                    {
                        allDelay += thisDelay;
                        thisDelay = tranUofSpeed(speed);
                        seqs[i].Attributes.Append(xmlDoc.CreateAttribute("delay"));
                        seqs[i].Attributes["delay"].Value = allDelay.ToString();
                        preNode = seqs[i];
                        if (outterSeq == null)
                        {
                            outterSeq = xmlDoc.CreateElement("演:序列_6B1B", nm.LookupNamespace("演"));
                        }
                        //else if (with == true)
                        //{
                        //    newAnim.InsertBefore(outterSeq, newAnim.LastChild);
                        //    outterSeq = xmlDoc.CreateElement("演:序列_6B1B", nm.LookupNamespace("演"));
                        //}
                        //with = false;
                        outterSeq.InsertAfter(seqs[i], outterSeq.LastChild);
                    }
                    else
                    {
                        //with = true;
                        if (tranUofSpeed(speed) > thisDelay)
                            thisDelay = tranUofSpeed(speed);
                        if (preNode != null)
                        {
                            preNode.InsertAfter(seqs[i], preNode.LastChild);
                        }
                        else
                        {
                            if (outterSeq == null)
                            {
                                outterSeq = xmlDoc.CreateElement("演:序列_6B1B", nm.LookupNamespace("演"));
                                outterSeq.InsertAfter(seqs[i], outterSeq.LastChild);
                            }
                        }
                        preNode = seqs[i];
                    }
                    ++id;
                    newAnim.InsertAfter(outterSeq, newAnim.LastChild);
                }
                if (seqs.Count > 0)
                    slide.ReplaceChild(newAnim, anim);
            }
            //2010.04.24 end
            //txtReader.Close();
            resultWriter = new XmlTextWriter(outputFile, System.Text.Encoding.UTF8);

            xmlDoc.Save(resultWriter);
            resultWriter.Close();

        }


        private static void FindAssembleGraph(string selectNode, XmlNamespaceManager nm)
        {
            // TODO:find assemble graph
        }

        private void addProp(XmlElement ele)
        {
            string hID = gethID(ele);//获得基式样信息
            XmlElement referencedEle = null;
            if (hID != "EOF")//如果基式样存在
            {
                try
                {
                    for (int r = 0; r < paraStyles.Count; r++)
                    {
                        string idid = paraStyles[r].Attributes["标识符_4100"].Value.ToString();   //改为段落式样的 标识符
                        if (hID == idid)
                        {

                            referencedEle = (XmlElement)paraStyles[r];
                            addProp(referencedEle);
                            compareEqualNode(ele, referencedEle);//比较基式样与当前段落式样
                            ele.RemoveAttribute("基式样引用_4104");
                            break;
                        }
                    }
                }
                catch (Exception exp)
                {
                    throw exp;
                }
            }
        }

        private void compareEqualNode(XmlElement curNode, XmlElement refNode)
        {
            //if (curNode.Name != "a:p" && curNode.Name != "a:lstStyle")
            if (true)
            {
                // try
                // {
                //比较当前节点与被引用节点是否相同
                //0.比较元素的属性是否一致，将原有属性加入新元素中
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
                    //该属性节点为原来的属性节点，拿过来

                    if (q == curAttrList.Count)
                    {
                        //msdn 说 属性不能被添加两次，所以要 clone ， 或者用 CloneNode(true)也可以。还有要进行类型转换，加上（XA）

                        XmlAttribute refAttr = (XmlAttribute)refAttrList[p].Clone();
                        curAttrList.Append(refAttr);
                    }
                }

                //当前节点无子元素，直接由上层继承
                if (curNode.ChildNodes.Count == 0 && refNode.ChildNodes.Count != 0)
                {
                    for (int m = 0; m < refNode.ChildNodes.Count; m++)
                    {
                        curNode.AppendChild((XmlNode)refNode.ChildNodes[m].Clone());
                    }
                }
                if (curNode.ChildNodes.Count == 0 && refNode.ChildNodes.Count == 0)
                {
                    return;
                }
                if (curNode.ChildNodes.Count != 0 && refNode.ChildNodes.Count != 0)
                {
                    if (curNode.ChildNodes.Count == 1 && refNode.ChildNodes.Count == 1 && curNode.ChildNodes[0].GetType().ToString() == "System.Xml.XmlText" && refNode.ChildNodes[0].GetType().ToString() == "System.Xml.XmlText")
                    {
                        return;
                    }
                    else
                    {

                        for (int m = 0; m < refNode.ChildNodes.Count; m++)
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

                                curNode.AppendChild((XmlNode)refNode.ChildNodes[m].Clone());
                            }
                        }
                    }
                }
            }

        }
        //返回当前段落式样的基式样引用的内容。。
        private string gethID(XmlElement currentEle)
        {
            string referenceId = "";
            if (currentEle.Attributes["基式样引用_4104"] != null)
            {
                referenceId = currentEle.Attributes["基式样引用_4104"].Value.ToString();
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
            if (currentEle.Attributes["标识符_4100"] != null)
            {
                Id = currentEle.Attributes["标识符_4100"].Value.ToString();
                return Id;
            }
            else
            {
                return "EOF";
            }
        }

        //转换UOF动画速度为OOXML动画速度
        private long tranUofSpeed(String uofSpeed)
        {
            switch (uofSpeed)
            {
                case "very-slow": return 5000;
                case "slow": return 3000;
                case "medium": return 2000;
                case "fase": return 1000;
                case "very-fast": return 500;
                default: return 500;
            }
        }

        #endregion

    }
}
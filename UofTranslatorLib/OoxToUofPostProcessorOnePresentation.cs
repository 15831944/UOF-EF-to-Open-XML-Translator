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
using System.Reflection;
using System.Resources;
using System.Collections;
using System.Globalization;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// The first step post process in OOX->UOF of Presentation
    /// </summary>
    /// <author>linyaohu</author>
    class OoxToUofPostProcessorOnePresentation : AbstractProcessor
    {

        private static ILog logger = LogManager.GetLogger(typeof(OoxToUofPostProcessorOnePresentation).FullName);
        private event EventHandler progressMessageIntercepted;

        private event EventHandler feedbackMessageIntercepted;

        //protected LinkedList<IProcessor> uof2ooxPreProcessors;

       // protected LinkedList<IProcessor> oox2uofPreProcessors;

       // protected LinkedList<IProcessor> oox2uofPostProcessors;

        private Hashtable messageTypes = new Hashtable();

        private ResourceManager resources;

        private string promptMsg;

        public override bool transform()
        {

            //XmlUrlResolver resourceResolver = null;

            //try
            //{

            //    resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
            //        this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + "Powerpoint.oox2uof");
            //    MainTransform(TranslatorConstants.OOXToUOF_XSL, resourceResolver, originalFile, grpTmp, outputFile);

            //}
            //catch (Exception ex)
            //{
            //    logger.Warn(ex.Message);
            //}


            string numberRefTmp = Path.GetDirectoryName(originalFile) + "\\" + "numberRefTmp.xml";
          //  NumberRef(inputFile, numberRefTmp);
         //   deleteLayoutAnchors(inputFile);
            bool result = true;
          //  XmlTextWriter resultWriter = null;
          

            #region change the IDs in the UOF file to fix a bug
            //更改输出文件的ID值,以使从UOF转化为OPENXML时，Layout的id能正常转换
            //added by SYBASE 2009.07.08
            XmlDocument doc = new XmlDocument();
            //doc.Load(@"C:\Users\v-xipia\Desktop\test.xml");
          //  doc.Load(outputFile);
            doc.Load(inputFile);

            XmlTextWriter tw = null;

            // Create an XmlNamespaceManager to resolve the default namespace.
            XmlNamespaceManager nmgr = new XmlNamespaceManager(doc.NameTable);
            //nsmgr.AddNamespace("bk", "urn:newbooks-schema");
            nmgr.AddNamespace("字", "http://schemas.uof.org/cn/2009/wordproc");
            nmgr.AddNamespace("uof", "http://schemas.uof.org/cn/2009/uof");
            nmgr.AddNamespace("图", "http://schemas.uof.org/cn/2009/graph");
            nmgr.AddNamespace("表", "http://schemas.uof.org/cn/2009/spreadsheet");
            nmgr.AddNamespace("演", "http://schemas.uof.org/cn/2009/presentation");
            nmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            nmgr.AddNamespace("app", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties");
            nmgr.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            nmgr.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
            nmgr.AddNamespace("dcmitype", "http://purl.org/dc/dcmitype/");
            nmgr.AddNamespace("dcterms", "http://purl.org/dc/terms/");
            nmgr.AddNamespace("fo", "http://www.w3.org/1999/XSL/Format");
            nmgr.AddNamespace("p", "http://schemas.openxmlformats.org/presentationml/2006/main");
            nmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            nmgr.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");
            nmgr.AddNamespace("元","http://schemas.uof.org/cn/2009/metadata");
            nmgr.AddNamespace("对象","http://schemas.uof.org/cn/2009/objects");
            nmgr.AddNamespace("式样","http://schemas.uof.org/cn/2009/styles");
            nmgr.AddNamespace("pzip", "urn:u2o:xmlns:post-processings:special");
            nmgr.AddNamespace("图形", TranslatorConstants.XMLNS_UOFGRAPHICS);
            nmgr.AddNamespace("规则", TranslatorConstants.XMLNS_UOFRULES);
            //// Select and display all book titles.
            //XmlNodeList nodeList;
            XmlElement rootOutput = doc.DocumentElement;

            //2010-12-15 罗文甜，增加段落式样的标识符替换
            XmlNodeList paraStyleList = rootOutput.SelectNodes("//式样:段落式样_9912|//式样:段落式样_9905", nmgr);
            string currentParaID = "";
            string tempParaID = "";
            int inum = 1;
            foreach (XmlNode paraStyle in paraStyleList)
            {
                currentParaID = paraStyle.Attributes.GetNamedItem("标识符_4100").Value;
                if (inum < 10)
                    // tempParaID = "tyleAttrId" + "0000" + inum;
                    tempParaID = "tyleAttrId" + "0000" + inum;
                else if (inum < 100)
                    tempParaID = "tyleAttrId" + "000" + inum;
                else if (inum < 1000)
                    tempParaID = "tyleAttrId" + "00" + inum;
                else if (inum < 10000)
                    tempParaID = "tyleAttrId" + "0" + inum;
                else if (inum < 100000)
                    tempParaID = "tyleAttrId" + inum;
                else
                    throw new Exception("Too many paragraphs id");
                paraStyle.Attributes.GetNamedItem("标识符_4100").Value = tempParaID;
                foreach (XmlNode paraStyleTwo in paraStyleList)
                {
                    if (((XmlElement)paraStyleTwo).HasAttribute("基式样引用_4104"))
                    {
                        if (paraStyleTwo.Attributes.GetNamedItem("基式样引用_4104").Value == currentParaID)
                        {
                            paraStyleTwo.Attributes.GetNamedItem("基式样引用_4104").Value = tempParaID;
                        }
                    }
                }

                XmlNodeList paraAttrs = rootOutput.SelectNodes("//字:段落属性_419B", nmgr);
                foreach (XmlNode paraAttr in paraAttrs)
                {
                    if (((XmlElement)paraAttr).HasAttribute("式样引用_419C") && paraAttr.Attributes.GetNamedItem("式样引用_419C").Value == currentParaID)
                    {
                        paraAttr.Attributes.GetNamedItem("式样引用_419C").Value = tempParaID;
                    }
                }
                
                inum++;
            }

            XmlNodeList animiationSequenceList = doc.SelectNodes("//演:序列_6B1B", nmgr);
            //2010-12-08 罗文甜：增加连接线始端和终端的图形引用
            XmlNodeList curveShapList = doc.SelectNodes("//图:连接线规则_8027", nmgr);

            XmlNodeList anchorNodeList;
            anchorNodeList = rootOutput.SelectNodes("//uof:锚点_C644", nmgr);

            //当前的GraphicsID
            string currentGraphicsID="";
            //替换掉的GraphicsID
            //OBJ00003
            string tempGraphicsID="";
            //tempGraphicsID = "OBJ";
            //对当前处理的图像编号计数
            int i = 1;

            foreach (XmlNode anchorNode in anchorNodeList)
            {
                //Console.WriteLine(anchorNode.Name);
                //存储当前GraphicsID,并替换掉
                currentGraphicsID = anchorNode.Attributes.GetNamedItem("图形引用_C62E").Value;
                //Console.WriteLine(currentGraphicsID + "\r\n");

                if (i < 10)
                    tempGraphicsID = "Obj" + "000" + i;
                else if (i < 100)
                    tempGraphicsID = "Obj" + "00" + i;
                else if (i < 1000)
                    tempGraphicsID = "Obj" + "0" + i;
                else if (i < 10000)
                    tempGraphicsID = "Obj" + i;
                else
                    throw new Exception("Too many Graphics id");


                anchorNode.Attributes.GetNamedItem("图形引用_C62E").Value = tempGraphicsID;
                //Console.WriteLine(tempGraphicsID);


                XmlNodeList graphicsNodeList;
                //XmlElement root = doc.DocumentElement;
                graphicsNodeList = rootOutput.SelectNodes("//图:图形_8062", nmgr);


                foreach (XmlNode graphicsNode in graphicsNodeList)
                {

                    if (graphicsNode.Attributes.GetNamedItem("标识符_804B").Value == currentGraphicsID)
                    {

                        //Console.WriteLine(graphicsNode.Name);
                        //Console.WriteLine(graphicsNode.Attributes.GetNamedItem("图:标识符").Value + "\r\n");
                        graphicsNode.Attributes.GetNamedItem("标识符_804B").Value = tempGraphicsID;
                        //Console.WriteLine(currentGraphicsID);
                        //Console.WriteLine();
                        
                        // 动画引用
                        foreach (XmlNode animiationSequence in animiationSequenceList)
                        {
                            if (animiationSequence.Attributes.GetNamedItem("对象引用_6C28").Value == currentGraphicsID)
                            {
                                animiationSequence.Attributes.GetNamedItem("对象引用_6C28").Value = tempGraphicsID;
                            }
                            //2010-11-15 罗文甜：增加触发器的图形引用
                            XmlNode tgtObject = animiationSequence.SelectSingleNode("演:定时_6B2E", nmgr);
                            XmlElement tgtObjectEle = (XmlElement)tgtObject;

                            if (tgtObjectEle.HasAttribute("触发对象引用_6B34"))
                            {
                                if (tgtObjectEle.Attributes.GetNamedItem("触发对象引用_6B34").Value == currentGraphicsID)
                                {
                                    tgtObjectEle.Attributes.GetNamedItem("触发对象引用_6B34").Value = tempGraphicsID;
                                   
                                }
                            }
                           
                        }
                         //2010-12-08 罗文甜：增加连接线始端和终端的图形引用                      
                        foreach (XmlNode curveShap in curveShapList)
                        {
                            if (curveShap.Attributes.GetNamedItem("连接线引用_8028").Value == currentGraphicsID)
                            {
                                curveShap.Attributes.GetNamedItem("连接线引用_8028").Value = tempGraphicsID;
                            }
                            if (((XmlElement)curveShap).HasAttribute("始端对象引用_8029") && curveShap.Attributes.GetNamedItem("始端对象引用_8029").Value == currentGraphicsID)
                            {
                                curveShap.Attributes.GetNamedItem("始端对象引用_8029").Value = tempGraphicsID;
                            }
                            if (((XmlElement)curveShap).HasAttribute("终端对象引用_802A") && curveShap.Attributes.GetNamedItem("终端对象引用_802A").Value == currentGraphicsID)
                            {
                                curveShap.Attributes.GetNamedItem("终端对象引用_802A").Value = tempGraphicsID;
                            }
                        }

                      }


                    //Console.WriteLine();
                }

                i++;
                //Console.WriteLine(nodeList1.Count);
            }


            XmlNodeList groupShapes = doc.SelectNodes("//pzip:archive//图形:图形集_7C00/图:图形_8062[@组合列表_8064]", nmgr);
            XmlNodeList shapes = doc.SelectNodes("//pzip:archive//图形:图形集_7C00/图:图形_8062", nmgr);
            foreach (XmlNode shape in shapes)
            {
                currentGraphicsID = shape.Attributes.GetNamedItem("标识符_804B").Value;

                // only change group shapes
                if(!currentGraphicsID.Contains("Obj"))
                {
                    if (i < 10)
                        tempGraphicsID = "Obj" + "000" + i;
                    else if (i < 100)
                        tempGraphicsID = "Obj" + "00" + i;
                    else if (i < 1000)
                        tempGraphicsID = "Obj" + "0" + i;
                    else if (i < 10000)
                        tempGraphicsID = "Obj" + i;
                    else
                        throw new Exception("Too many Graphics id");

                    shape.Attributes.GetNamedItem("标识符_804B").Value = tempGraphicsID;

                    string newGroupIDString = string.Empty;

                    foreach (XmlNode groupShape in groupShapes)
                    {
                        string groupID = groupShape.Attributes.GetNamedItem("组合列表_8064").Value;
                        string[] graphicIDs = groupID.Split(new char[1] { ' ' });

                        for (int j = 0; j < graphicIDs.Length; j++)
                        {
                            if (graphicIDs[j] == currentGraphicsID)
                            {
                                graphicIDs[j] = tempGraphicsID;
                            }
                        }

                        for (int k = 0; k < graphicIDs.Length; k++)
                        {
                            newGroupIDString += graphicIDs[k] + " ";
                        }
                        groupShape.Attributes.GetNamedItem("组合列表_8064").Value = newGroupIDString.Substring(0, newGroupIDString.Length - 1);
                        newGroupIDString = string.Empty;
                        /*
                        if (groupID.Contains(currentGraphicsID))
                        {
                            groupShape.Attributes.GetNamedItem("图:组合列表").Value = groupID.Replace(currentGraphicsID, tempGraphicsID);
                        }
                        
                        */


                    }

                    i++;
                }
            }
            


            int lastID=0;

            XmlNodeList slideLayouts = doc.SelectNodes("/pzip:archive/pzip:entry[@pzip:target='rules.xml']/规则:公用处理规则_B665/规则:演示文稿_B66D/规则:页面版式集_B651/规则:页面版式_B652", nmgr);
            XmlNodeList refSlideLayouts = doc.SelectNodes("/pzip:archive/pzip:entry[@pzip:target='content.xml']/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F", nmgr);
          //  XmlNodeList refSlideLayouts = doc.SelectNodes("uof:UOF/uof:演示文稿//@演:页面版式引用", nmgr);
            lastID = AddjustID(doc, slideLayouts, refSlideLayouts, "标识符_6B0D", "页面版式引用_6B27", "ID", 1);


            //<!--注销这部分内容李娟 2012 03，26·····································-->
            //String uofns = nmgr.LookupNamespace("uof");
            //XmlNode docroot = doc.SelectSingleNode("uof:UOF", nmgr);
            //XmlNode kzq = docroot.SelectSingleNode("/uof:UOF/uof:扩展区",nmgr);
            
            //if (kzq == null)
            //{
            //    kzq = doc.CreateElement("uof:扩展区",uofns);
            //    docroot.AppendChild(kzq);
            //}
            //XmlElement kz = doc.CreateElement("uof:扩展", uofns);
            //kz.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0051");

            //2010-11-19罗文甜：增加软件名称和软件版本
            //XmlElement rjmc = doc.CreateElement("uof:软件名称", uofns);
            //rjmc.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0052");
            //rjmc.InnerText = "EIOffice";
            //XmlElement rjbb = doc.CreateElement("uof:软件版本", uofns);
            //rjbb.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0053");
            //rjbb.InnerText = "v1.33";
            //XmlElement kznr = doc.CreateElement("uof:扩展内容", uofns);
            //kznr.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0054");
            
            //XmlElement path = doc.CreateElement("uof:路径", uofns);
            //XmlElement nr = doc.CreateElement("uof:内容", uofns);
            //path.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0065");
            //nr.SetAttribute("locID", "http://schemas.uof.org/cn/2003/uof", "u0056");
            //path.InnerText = "演示文稿";
            XmlNodeList masters = doc.SelectNodes("//演:母版_6C0D", nmgr);
            foreach (XmlNode master in masters)
            {
                XmlNodeList list = master.SelectNodes("uof:锚点_C644[uof:占位符_C626/@类型_C627='date' or uof:占位符_C626/@类型_C627='number' or uof:占位符_C626/@类型_C627='header' or uof:占位符_C626/@类型_C627='footer']", nmgr);
                foreach (XmlNode node in list)
                {
                    XmlElement jsx = doc.CreateElement("uof:句属性", "http://schemas.uof.org/cn/2009/uof");
                    XmlElement ymyjlx = doc.CreateElement("uof:页眉页脚类型", "http://schemas.uof.org/cn/2009/uof");
                    XmlNode para = doc.SelectSingleNode("//图:图形_8062[@标识符_804B='" + node.Attributes["图形引用_C62E"].Value + "']/图:文本_803C/图:内容_8043/字:段落_416B", nmgr);
                    String type = String.Empty;
                    if (para != null)
                    {
                        String refid = para.Attributes["标识符_4169"].Value;
                        XmlNode childNode = node.SelectSingleNode("uof:占位符_C626", nmgr);
                        switch (childNode.Attributes["类型_C627"].Value)
                        {
                            case "date":
                                type = "datetime";
                                break;
                            case "number":
                                type = "slidenumber";
                                break;
                            case "header":
                                type = "header";
                                break;
                            case "footer":
                                type = "footer";
                                break;
                        }
                    }
                    //注销这块 李娟 2012.03.26····················································
                    //jsx.SetAttribute("locID", uofns, "w0027");
                    //jsx.SetAttribute("attrList", uofns, "引用 序号");
                    //String num;
                    //if (para.Attributes["序号"] == null)
                    //    num = "1";
                    //else
                    //    num = para.Attributes["序号"].Value;
                    //jsx.SetAttribute("序号", uofns, num);
                    //jsx.SetAttribute("引用", uofns, refid);
                    //ymyjlx.SetAttribute("locID", uofns, "w0031");
                    //ymyjlx.SetAttribute("attrList", uofns, "类型");
                    //ymyjlx.SetAttribute("类型", uofns, type);
                    //jsx.AppendChild(ymyjlx);
                    //nr.AppendChild(jsx);
                }
            }
            //注销这块 李娟2012.03.26
            //kznr.AppendChild(path);
            //kznr.AppendChild(nr);
            ////2010-11-19罗文甜
            //kz.AppendChild(rjmc);
            //kz.AppendChild(rjbb);

            //kz.AppendChild(kznr);
            //kzq.AppendChild(kz);
            //09.11.19 马有旭 添加↑
            //10.03.09 马有旭 添加 修改幻灯片段落ID 原：ID12E3C 改为：shp0graphc0
            //2010-11-24-罗文甜-修改
            XmlNodeList anchors = doc.SelectNodes("//pzip:archive/pzip:entry[@pzip:target='content.xml']/演:演示文稿文档_6C10/演:幻灯片集_6C0E/演:幻灯片_6C0F/uof:锚点_C644 | //pzip:archive/pzip:entry[@pzip:target='content.xml']/演:演示文稿文档_6C10/演:母版集_6C0C/演:母版_6C0D[@类型_6BEA='slide']/uof:锚点_C644[uof:占位符_C626/@类型_C627!='date' and uof:占位符_C626/@类型_C627!='footer' and uof:占位符_C626/@类型_C627!='number']", nmgr);
            //XmlNodeList anchors = doc.SelectNodes("/uof:UOF/uof:演示文稿/演:主体//uof:锚点", nmgr);
            int graphicCount = 0;
            int paraCount = 0;
            foreach (XmlNode anchor in anchors)
            {
                XmlNode graphic = doc.SelectSingleNode("//pzip:archive/pzip:entry[@pzip:target='graphics.xml']/图形:图形集_7C00/图:图形_8062[@标识符_804B='" + anchor.Attributes["图形引用_C62E"].Value + "']", nmgr);
                if (graphic != null)
                {
                    ++graphicCount;
                    paraCount = 0;
                    String newId = "shp" + graphicCount + "graphc";
                    XmlNodeList paragraphs = graphic.SelectNodes("图:文本_803C/图:内容_8043/字:段落_416B", nmgr);
                    foreach (XmlNode para in paragraphs)
                    {
                        XmlAttribute attr = para.Attributes["标识符_4169"];
                        if (attr != null)
                        {
                            String oldId = attr.Value;
                            attr.Value = newId + paraCount;
                            XmlNodeList refs = doc.SelectNodes("//*[@段落引用_6C27='" + oldId + "']", nmgr);
                            foreach(XmlNode refNode in refs)
                            {
                                refNode.Attributes["段落引用_6C27"].Value = newId + paraCount;
                            }
                            ++paraCount;
                        }
                    }
                }
            }
            //注销这部分 李娟 ········································· 2012.03.36
            //删除@序号
            //XmlNodeList paraList = doc.SelectNodes("//字:段落[@序号]",nmgr);
            //foreach (XmlNode para in paraList)
            //{
            //    para.Attributes.Remove(para.Attributes["序号"]);
            //}
            try
            {
                string tmpAdjustID = Path.GetDirectoryName(inputFile) + Path.AltDirectorySeparatorChar + "tmpAdjustID.xml";
                tw = new XmlTextWriter(tmpAdjustID, Encoding.UTF8);
                doc.Save(tw);

                tw.Close();

                XPathDocument xslDoc;
                XmlReaderSettings xrs = new XmlReaderSettings();
                XmlReader source = null;
                XmlWriter writer = null;
                OoxZipResolver zipResolver = null;
                XmlUrlResolver resourceResolver = null;
                try
                {
                    //xrs.ProhibitDtd = true;

                    resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
                    this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + "Powerpoint.oox2uof");

                    xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream("post-processing.xslt"));
                    xrs.XmlResolver = resourceResolver;
                    source = XmlReader.Create(tmpAdjustID);

                    XslCompiledTransform xslt = new XslCompiledTransform();
                    XsltSettings settings = new XsltSettings(true, false);
                    xslt.Load(xslDoc, settings, resourceResolver);

                    //if (!originalFile.Equals(string.Empty))
                    //{
                    //    zipResolver = new OoxZipResolver(originalFile, resourceResolver);
                    //}
                    XsltArgumentList parameters = new XsltArgumentList();
                    parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);

                    // zip format
                    parameters.AddParam("outputFile", "", outputFile);
                    // writer = new OoxZipWriter(inputFile);
                    writer = new UofZipWriter(outputFile);

                    if (zipResolver != null)
                    {
                        xslt.Transform(source, parameters, writer, zipResolver);
                    }
                    else
                    {
                        xslt.Transform(source, parameters, writer);
                    }
                }
                finally
                {
                    if (writer != null)
                        writer.Close();
                    if (source != null)
                        source.Close();
                }
            }
            catch
            {
                throw new Exception("Fail in ajust ids");
            }
            finally
            {
                if (tw != null)
                {
                    tw.Close();
                }
            }

            #endregion

            return result;
        }

        private  void NumberRef(string inputFileName, string outputFileName)
        {
            XmlTextWriter tw = null;
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFileName);
            XmlNameTable xnameTable = xdoc.NameTable;
            XmlNamespaceManager nm = new XmlNamespaceManager(xnameTable);
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2009/uof");
            nm.AddNamespace("字", "http://schemas.uof.org/cn/2009/uofwordproc");
            nm.AddNamespace("图", "http://schemas.uof.org/cn/2009/graph");
            XmlNodeList graphics = xdoc.SelectNodes("//图:图形_8062", nm);
            XmlNodeList paragraphs = null;
            string autoNumberMsg;
            string autoNumberMsgNext;
            XmlNode autoNumber = null;
            XmlNode autoNumberNext = null;
            if (graphics != null)
            {
                foreach (XmlNode graphic in graphics)
                {
                    paragraphs = graphic.SelectNodes(".//字:段落_416B", nm);
                    if (paragraphs.Count > 1)
                    {
                        for (int i = 0; i < paragraphs.Count - 1; i++)
                        {
                            autoNumber = paragraphs[i].SelectSingleNode("字:段落属性_419B/字:自动编号信息_4186", nm);
                            autoNumberNext = paragraphs[i + 1].SelectSingleNode("字:段落属性_419B/字:自动编号信息_4186", nm);
                            if (autoNumber != null && autoNumberNext != null)
                            {
                                autoNumberMsg = paragraphs[i].SelectSingleNode("字:段落属性_419B/字:自动编号信息_4186", nm).Attributes.GetNamedItem("编号引用_4187").Value;
                                autoNumberMsgNext = paragraphs[i + 1].SelectSingleNode("字:段落属性_419B/字:自动编号信息_4186", nm).Attributes.GetNamedItem("编号引用_4187").Value;
                                autoNumber = xdoc.SelectSingleNode("//..//字:自动编号信息_4186[@标识符_4100='" + autoNumberMsg + "']", nm);
                                autoNumberNext = xdoc.SelectSingleNode("//..//字:自动编号信息_4186[@标识符_4100='" + autoNumberMsgNext + "']", nm);
                                if (autoNumberMsg != autoNumberMsgNext)
                                {
                                    if (autoNumber.InnerXml == autoNumberNext.InnerXml)
                                    {
                                        paragraphs[i + 1].SelectSingleNode("字:段落属性_419B/字:自动编号信息_4186", nm).Attributes.GetNamedItem("编号引用_4187").Value =
                                            paragraphs[i].SelectSingleNode("字:段落属性_419B/字:自动编号信息_4186", nm).Attributes.GetNamedItem("编号引用_4187").Value;
                                        autoNumber = null;
                                        autoNumberNext = null;
                                    }
                                }
                            }

                        }
                    }
                }
            }
            tw = new XmlTextWriter(outputFileName, Encoding.UTF8);
            try
            {
                xdoc.Save(tw);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
               // logger.Error(ex.StackTrace);
               // throw new Exception("Fail in number reference treatment");
            }
            finally
            {
                if (tw != null)
                    tw.Close();
            }
        }

        // <summary>
        // addust id
       // </summary>
       // <param name="xdoc">source document</param>
       // <param name="tw">result file</param>
       // <param name="nodeList">source node list</param>
       // <param name="refNodeList">related node list</param>
       // <param name="attributeNameItem">the name item of attribute you want to adjust</param>
       // <param name="prefixID">the prefix of id</param>
       // <param name="startID">start number of ID</param>
       // <returns>total number</returns>
        private int AddjustID(XmlDocument xdoc,
            XmlNodeList nodeList, XmlNodeList refNodeList,
            string attributeNameItem,
            string refAttributeNameItem,
            string prefixID,int startID)
        {
           
            // current id
            string currentID = "";
            // temp id
            string tempID="";

            foreach (XmlNode node in nodeList)
            {
                // get current id
                currentID = node.Attributes.GetNamedItem(attributeNameItem).Value;

                // new format of id 
                if (startID < 10)
                    tempID = prefixID + "000" + startID;
                else if (startID < 100)
                    tempID = prefixID + "00" + startID;
                else if (startID < 1000)
                    tempID = prefixID + "0" + startID;
                else 
                    tempID = prefixID + startID;

                // set id
                node.Attributes.GetNamedItem(attributeNameItem).Value = tempID;

                // change all related ids
                foreach (XmlNode refNode in refNodeList)
                {
                    if (refNode.Attributes.GetNamedItem(refAttributeNameItem).Value == currentID)
                        refNode.Attributes.GetNamedItem(refAttributeNameItem).Value = tempID;
                }

                startID++;
            }
            return startID;

        }

        //2010.03.19 马有旭 添加 删除版式中类型为页眉页脚 日期时间 幻灯片编号 的锚点
        private void deleteLayoutAnchors(String inputfile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(inputFile);
            bool edited = false;
            XmlNamespaceManager nm = new XmlNamespaceManager(doc.NameTable);
            nm.AddNamespace("演", "http://schemas.uof.org/cn/2009/presentation");
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2009/uof");
            XmlNodeList layouts = doc.SelectNodes("//规则:页面版式集_B651/规则:页面版式_B652", nm);
            foreach (XmlNode layout in layouts)
            {
                XmlNodeList anchors = layout.SelectNodes("uof:锚点_C644[uof:占位符_C626/@类型_C627='date' or uof:占位符_C626/@类型_C627='header' or uof:占位符_C626/@类型_C627='footer' or uof:占位符_C626/@类型_C627='number']", nm);
                foreach (XmlNode anchor in anchors)
                {
                    edited = true;
                    layout.RemoveChild(anchor);
                }
            }
            if (edited)
                doc.Save(inputFile);
        }


        protected void MessageCallBack(object sender, XsltMessageEncounteredEventArgs e)
        {
            if (e.Message.StartsWith("progress:"))
            {
                if (progressMessageIntercepted != null)
                {
                    progressMessageIntercepted(this, null);
                }
            }
            else if (e.Message.StartsWith("feedback:"))
            {
               // logger.Warn(e.Message);
                string valuableMsg = ParseFeedbackMessage(e.Message);
                if (valuableMsg.Equals(String.Empty))
                    return;
                if (feedbackMessageIntercepted != null)
                {
                    feedbackMessageIntercepted(this, new UofEventArgs(valuableMsg));
                }
            }
        }

        private string ParseFeedbackMessage(string message)
        {
            string[] subMsg = message.Split(':');
            if (subMsg.Length != 4)
                return String.Empty;
            if (messageTypes.ContainsKey(subMsg[2]))
                return String.Empty;

            messageTypes.Add(subMsg[2], subMsg[2]);

            if (resources == null)
            {
                if (CultureInfo.CurrentCulture.Name.Equals("zh-CN"))
                {
                    resources = new ResourceManager(
                        typeof(UOFTranslator).Namespace + ".resources." + TranslatorConstants.MESSAGE_ZH_CHS_RESX,
                        Assembly.GetExecutingAssembly());
                    promptMsg = " 可能在转换过程中被丢失!";
                }
                else
                {
                    resources = new ResourceManager(
                        typeof(UOFTranslator).Namespace + ".resources." + TranslatorConstants.MESSAGE_EN_RESX,
                        Assembly.GetExecutingAssembly());
                    promptMsg = " may lost during conversion!";
                }
            }

            string result = resources.GetString(subMsg[2]);
            if (result != null)
                return result + promptMsg;
            return string.Empty;
        }

    }
}
using System;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using System.Xml.Xsl;
using System.Reflection;
using System.Text;
using Act.UofTranslator.UofZipUtils;
using System.Windows.Forms;
using System.Configuration;
using log4net;
using log4net.Config;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// Spreadsheet translator
    /// </summary>
    /// <author>linyaohu</author>
    public class SpreadsheetTranslator : UOFTranslator
    {
        private bool isOox2UofPackage = true;

        private static ILog logger = LogManager.GetLogger(typeof(PresentationTranslator).FullName);

        public SpreadsheetTranslator()
        {
            InitalPreProcessors();
            XmlConfigurator.Configure(LogManager.GetRepository(),
                new FileInfo(UOFTranslator.ASSEMBLY_PATH + @"conf\log4net.config"));
            LoadConfigs(ASSEMBLY_PATH + @"conf\config.xml");
        }

        protected override void InitalPreProcessors()
        {
            base.InitalPreProcessors();
            uof2ooxPreProcessors.AddLast(new UofToOoxPreProcessorOneSpreadSheet());
            // uof2ooxPreProcessors.AddLast(new UofToOoxPreProcessorTwoSpreadSheet());
            oox2uofPreProcessors.AddLast(new OoxToUofPreProcessorOneSpreadSheet());
            // oox2uofPostProcessors.AddLast(new OoxToUofPostProcessorOneSpreadsheet());
        }

        private void LoadConfigs(string filename)
        {
            XmlTextReader reader = null;
            bool hasConfinItem = false;
            try
            {
                reader = new XmlTextReader(filename);
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            {
                                if (reader.Name.Equals("oox2uof_package"))
                                {
                                    this.isOox2UofPackage = Convert.ToBoolean(reader.ReadString());
                                    hasConfinItem = true;
                                }
                            }
                            break;
                        default:
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                logger.Error("Fail to read config file: " + e.Message);
                logger.Error(e.StackTrace);
            }
            finally
            {
                reader.Close();
            }
            if (!hasConfinItem)
            {
                logger.Error("Can not find 'oox2uof_package' config item! Use default config!");
            }
            logger.Info("Load config file succcess!");
        }

        protected override void DoUofToOoxMainTransform(string inputFile, string outputFile, string resourceDir)
        {
            XmlUrlResolver resourceResolver;
            XPathDocument xslDoc;
            XmlReaderSettings xrs = new XmlReaderSettings();
            XmlReader source = null;
            XmlWriter writer = null;
            string mainOutput = Path.GetDirectoryName(inputFile) + Path.AltDirectorySeparatorChar + "mainOutput.xml";
            string equAfterMain = Path.GetDirectoryName(inputFile) + Path.AltDirectorySeparatorChar + "equAfterMain.xml";

            try
            {
                xrs.ProhibitDtd = true;
                string xslLocation = TranslatorConstants.UOFToOOX_XSL;
                if (outputFile == null)
                {
                    xslLocation = TranslatorConstants.UOFToOOX_COMPUTE_SIZE_XSL;
                }

                if (resourceDir == null)
                {
                    resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
                        this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + "Excel.uof2oox");
                    xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(xslLocation));
                    xrs.XmlResolver = resourceResolver;
                    source = XmlReader.Create(inputFile);
                }
                else
                {
                    resourceResolver = new XmlUrlResolver();
                    xslDoc = new XPathDocument(resourceDir + "/" + xslLocation);
                    source = XmlReader.Create(resourceDir + "/" + TranslatorConstants.SOURCE_XML, xrs);
                }
                try
                {
                    XslCompiledTransform xslt = new XslCompiledTransform();
                    XsltSettings settings = new XsltSettings(true, false);
                    //Assembly ass = Assembly.Load("excel_uof2oox");
                    //Type t = ass.GetType("uof2oox");
                    // xslt.Load(t);
                    xslt.Load(xslDoc, settings, resourceResolver);
                    XsltArgumentList parameters = new XsltArgumentList();
                    parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                    //if (outputFile != null)
                    //{
                    //    parameters.AddParam("outputFile", "", outputFile);
                    //    writer = new OoxZipWriter(inputFile);
                    //}
                    //else
                    //{
                    //    writer = new XmlTextWriter(new StringWriter());
                    //}

                    //xslt.Transform(source, parameters, writer);
                    XmlTextWriter fs = new XmlTextWriter(mainOutput, Encoding.UTF8);
                    xslt.Transform(source, parameters, fs);
                    fs.Close();

                    SetDrawingTwoCellAnchorValue(mainOutput, equAfterMain);

                    // 增加对公式的互操作支持
                    // EquInter(mainOutput, equAfterMain);

                    xslLocation = TranslatorConstants.UOFToOOX_POSTTREAT_STEP1_XSL;

                    if (resourceDir == null)
                    {
                        resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
                            this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + "Excel.uof2oox");
                        xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(xslLocation));
                        xrs.XmlResolver = resourceResolver;
                        source = XmlReader.Create(equAfterMain);
                    }
                    XslCompiledTransform xslt2 = new XslCompiledTransform();
                    XsltSettings settings2 = new XsltSettings(true, false);
                    xslt.Load(xslDoc, settings2, resourceResolver);

                    XsltArgumentList parameters2 = new XsltArgumentList();
                    parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                    if (outputFile != null)
                    {
                        parameters2.AddParam("outputFile", "", outputFile);
                        writer = new OoxZipWriter(equAfterMain);
                    }
                    else
                    {
                        writer = new XmlTextWriter(new StringWriter());
                    }

                    xslt.Transform(source, parameters2, writer);
                }
                catch (Exception ex)
                {
                    throw ex;
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

        protected override void DoOoxToUofMainTransform(string originalFile, string inputFile, string outputFile, string resourceDir)
        {

            // string tempFilePath = Path.GetDirectoryName(inputFile).ToString();
            //// string mainSheet = tempFilePath + "\\" + "oox2uof.xsl";
            // string mainSheet = tempFilePath + "\\" + "oox2uof.xsl";
            // string grpTmp = tempFilePath + "\\" + "grTmp.xml";
            // string mainTranslatorFile=tempFilePath+"\\"+"tempdoc.xml";
            // string mainOutputFile = tempFilePath + "\\" + "mainOutputFile.xml";
            // string mainOutputFile2 = tempFilePath + "\\" + "mainOutputFile2.xml";

            // XmlDocument grpdoc = new XmlDocument();
            // grpdoc.Load(inputFile);
            // XmlNamespaceManager nmgr = new XmlNamespaceManager(grpdoc.NameTable);
            // nmgr.AddNamespace("ws","http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            // nmgr.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            // nmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            // FileStream fs = new FileStream(grpTmp, FileMode.Create);
            // try
            // {
            //     GrpShPre(grpdoc, nmgr, "xdr", fs);
            // }
            // catch (Exception ex)
            // {
            //     throw ex;

            // }
            // finally
            // {
            //     if (fs != null)
            //         fs.Close();
            // }

            // XmlDocument xmlDoc;
            // XPathDocument doc = new XPathDocument(inputFile);;
            // XslCompiledTransform transFrom = new XslCompiledTransform();
            // XsltSettings setting = new XsltSettings(true, false);
            // XmlUrlResolver xur = new XmlUrlResolver();
            // XPathNavigator nav = ((IXPathNavigable)doc).CreateNavigator();
            // fs= new FileStream(mainTranslatorFile, FileMode.Create);
            // XmlTextReader tr = null;
            // XmlTextWriter tw = null;
            // try
            // {
            //     nav = ((IXPathNavigable)doc).CreateNavigator();
            //     transFrom.Load(mainSheet, setting, xur);
            //     Assembly ass = Assembly.Load("excel_oox2uof");
            //     Type t = ass.GetType("oox2uof");
            //     //transFrom.Load(t);
            //     transFrom.Transform(nav, null, fs);
            //     fs.Close();

            //     tr = new XmlTextReader(mainTranslatorFile);
            //     tw = new XmlTextWriter(mainOutputFile, Encoding.UTF8);
            //     try
            //     {
            //         xmlDoc = new XmlDocument();
            //         xmlDoc.Load(tr);
            //         XmlNamespaceManager nm = new XmlNamespaceManager(xmlDoc.NameTable);
            //         nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            //         nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            //         XmlNodeList data;
            //         data = xmlDoc.SelectNodes("//uof:UOF//..//表:数据有效性", nm);
            //         foreach (XmlNode node in data)
            //         {

            //             XmlElement quar = (XmlElement)node.FirstChild;
            //             string s1 = node.FirstChild.InnerText.ToString();
            //             string s2 = s1.Replace('%', '\'');
            //             quar.InnerText = s2;
            //             node.ReplaceChild(quar, node.FirstChild);

            //         }
            //         data = xmlDoc.SelectNodes("//uof:UOF//..//表:条件格式化", nm);
            //         foreach (XmlNode node in data)
            //         {
            //             if (node.SelectSingleNode("表:区域", nm) != null)
            //             {
            //                 string s1 = "";
            //                 string s2 = "";
            //                 XmlNodeList sheetArea = node.SelectNodes("表:区域", nm);
            //                 foreach (XmlNode areaNode in sheetArea)
            //                 {
            //                     XmlElement conditionFormat = (XmlElement)areaNode;
            //                     s1 = areaNode.InnerText.ToString();
            //                     s2 = s1.Replace('%', '\'');
            //                     conditionFormat.InnerText = s2;
            //                     node.ReplaceChild(conditionFormat, areaNode);
            //                 }
            //             }
            //         }
            //         data = xmlDoc.SelectNodes("//uof:UOF//..//表:数据源", nm);
            //         foreach (XmlNode node in data)
            //         {
            //             if (node.SelectSingleNode("表:系列", nm) != null)
            //             {
            //                 string s1 = "";
            //                 string s2 = "";
            //                 XmlNodeList sheetseries = node.SelectNodes("表:系列", nm);
            //                 foreach (XmlNode seriesNode in sheetseries)
            //                 {
            //                     XmlElement dataSource = (XmlElement)seriesNode;
            //                     if (seriesNode.Attributes.GetNamedItem("表:系列值") != null)
            //                     {
            //                         s1 = seriesNode.Attributes.GetNamedItem("表:系列值").Value.ToString();
            //                         s2 = s1.Replace('%', '\'');
            //                         if (s2.Contains("\'\'"))
            //                         {
            //                             s2 = s2.Replace("\'\'", "\'");
            //                         }
            //                         dataSource.Attributes.GetNamedItem("表:系列值").Value = s2;
            //                     }

            //                     if (seriesNode.Attributes.GetNamedItem("表:系列名") != null)
            //                     {
            //                         s1 = seriesNode.Attributes.GetNamedItem("表:系列名").Value.ToString();
            //                         s2 = s1.Replace('%', '\'');
            //                         if (s2.Contains("\'\'"))
            //                         {
            //                             s2 = s2.Replace("\'\'", "\'");
            //                         }
            //                         dataSource.Attributes.GetNamedItem("表:系列名").Value = s2;
            //                     }

            //                     if (seriesNode.Attributes.GetNamedItem("表:分类名") != null)
            //                     {
            //                         s1 = seriesNode.Attributes.GetNamedItem("表:分类名").Value.ToString();
            //                         s2 = s1.Replace('%', '\'');
            //                         if (s2.Contains("\'\'"))
            //                         {
            //                             s2 = s2.Replace("\'\'", "\'");
            //                         }
            //                         dataSource.Attributes.GetNamedItem("表:分类名").Value = s2;
            //                     }
            //                     //node.ReplaceChild(dataSource, node.SelectSingleNode("表:系列", nm));
            //                     node.ReplaceChild(dataSource, seriesNode);
            //                 }
            //             }

            //         }
            //         xmlDoc.Save(tw);
            //         tw.Close();
            //         tr.Close();

            //     }
            //     catch (Exception ex)
            //     {
            //         throw ex;
            //     }
            //     finally
            //     {
            //         if (tr != null)
            //             tr.Close();
            //         if (tw != null)
            //             tw.Close();
            //     }
            //     XmlDocument xdoc = new XmlDocument();
            //     xdoc.Load(mainOutputFile);
            //     XmlNamespaceManager xmlnm = new XmlNamespaceManager(xdoc.NameTable);
            //     xmlnm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            //     XmlNodeList sheetRoots = xdoc.SelectNodes("//表:工作表", xmlnm);
            //     XmlNodeList filters = null;
            //     try
            //     {
            //         foreach (XmlNode sheetRoot in sheetRoots)
            //         {
            //             filters = sheetRoot.SelectNodes(".//表:范围", xmlnm);
            //             if (filters != null)
            //             {
            //                 string sheetName = sheetRoot.Attributes.GetNamedItem("表:名称").Value.ToString();
            //                 foreach (XmlNode filter in filters)
            //                 {
            //                     string rangeVal = filter.InnerText.ToString();
            //                     string[] addVals = rangeVal.Split(new char[1] { ':' });
            //                     string newRangVal = "\'" + sheetName + "\'" + "!$" + addVals[0].Substring(0, 1) + '$' + addVals[0].Substring(1, addVals[0].Length - 1) + ":$" + addVals[1].Substring(0, 1) + '$' + addVals[1].Substring(1, addVals[1].Length - 1);
            //                     filter.InnerText = newRangVal;
            //                     newRangVal = "";
            //                 }
            //             }
            //             filters = null;
            //         }
            //         tw = new XmlTextWriter(mainOutputFile2, Encoding.UTF8);
            //         xdoc.Save(tw);
            //         tw.Close();
            //     }
            //     catch (Exception ex)
            //     {
            //         logger.Error("error in processing the filter", ex);
            //     }
            //     finally
            //     {
            //         if (tw != null)
            //             tw.Close();
            //     }

            // 图片预处理
            XmlDocument xmlDoc = new XmlDocument();
            string wordPrePath = Path.GetDirectoryName(inputFile) + Path.DirectorySeparatorChar;
            string picture_xml = "XLSX\\xl\\media";
            string tmpPic = wordPrePath + "tmpPic.xml";
            string custom_xml = "XLSX\\customXml";
            string customPath = wordPrePath + custom_xml;
            string tmpCustom = wordPrePath + "custom.xml";

            if (Directory.Exists(wordPrePath + picture_xml))
            {
                XmlNamespaceManager nms = new XmlNamespaceManager(xmlDoc.NameTable);
                nms.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
                nms.AddNamespace("w", TranslatorConstants.XMLNS_W);
                xmlDoc.Load(inputFile);
                xmlDoc = PicPretreatment(xmlDoc, "ws:spreadsheets", wordPrePath + picture_xml, nms);
                XmlTextWriter picWriter = new XmlTextWriter(tmpPic, Encoding.UTF8);
                xmlDoc.Save(picWriter);
                picWriter.Close();
                inputFile = tmpPic;
            }

            // manage the custom xml part
            if (Directory.Exists(customPath))
            {
                xmlDoc.Load(inputFile);
                XmlNamespaceManager nms = new XmlNamespaceManager(xmlDoc.NameTable);

                nms.AddNamespace("w", TranslatorConstants.XMLNS_W);
                nms.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
                xmlDoc = CustomXMPretreatment(xmlDoc, "ws:spreadsheets", customPath, nms);
                XmlTextWriter customWriter = new XmlTextWriter(tmpCustom, Encoding.UTF8);
                xmlDoc.Save(customWriter);
                customWriter.Close();
                inputFile = tmpCustom;
            }


            // add by linyaohu new function the smartArt 2013-03-05
            XmlDocument doc = new XmlDocument();
            doc.Load(inputFile);
            string tmpSmartArtFile = wordPrePath + "tmpSmartArt.xml";
            XmlNamespaceManager nm = new XmlNamespaceManager(doc.NameTable);
            nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
            nm.AddNamespace("xdr", TranslatorConstants.XMLNS_XDR);
            nm.AddNamespace("r", TranslatorConstants.XMLNS_PR);
            nm.AddNamespace("dsp", TranslatorConstants.XMLNS_DSP);

            XmlNodeList workSheets = doc.SelectNodes("ws:spreadsheets/ws:spreadsheet", nm);
            for (int i = 1; i < workSheets.Count + 1; i++)
            {
                XmlNode smartArtRel = doc.SelectSingleNode("ws:spreadsheets/ws:spreadsheet[" + i + "]/ws:Drawings/xdr:wsDr//r:Relationships", nm);
                if (smartArtRel != null)
                {
                    XmlNodeList smartArtRelDrawings = smartArtRel.SelectNodes("r:Relationship[@Type='http://schemas.microsoft.com/office/2007/relationships/diagramDrawing']", nm);
                    foreach (XmlNode smartArtRelDr in smartArtRelDrawings)
                    {
                        string target = smartArtRelDr.Attributes["Target"].InnerText;
                        string targetFile = wordPrePath + "XLSX\\XL\\" + target.Substring(3);

                        XmlDocument smartArtDrawing = new XmlDocument();
                        smartArtDrawing.Load(targetFile);

                        // add the drawing file
                        XmlNode drawingFile = doc.CreateElement("dsp", "drawing", TranslatorConstants.XMLNS_DSP);
                        XmlAttribute drawingFileName = doc.CreateAttribute("id");

                        // target="../diagrams/xxx.xml", get the xxx.xml
                        drawingFileName.Value = target.Substring(12);
                        drawingFile.Attributes.Append(drawingFileName);
                        drawingFile.InnerXml = smartArtDrawing.SelectSingleNode("dsp:drawing", nm).InnerXml;
                        smartArtRel.ParentNode.AppendChild(drawingFile);
                    }
                }
            }
            XmlTextWriter tmpSmartArtWriter = new XmlTextWriter(tmpSmartArtFile, Encoding.UTF8);
            doc.Save(tmpSmartArtWriter);
            tmpSmartArtWriter.Close();
            inputFile = tmpSmartArtFile;


            // main transform
            XmlUrlResolver resourceResolver = null;

            try
            {

                resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
                    this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + "Excel.oox2uof");
                MainTransform(TranslatorConstants.OOXToUOF_XSL, resourceResolver, originalFile, inputFile, outputFile);

                //XPathDocument xslDoc;
                //    XmlReaderSettings xrs = new XmlReaderSettings();
                //    XmlReader source = null;
                //    XmlWriter writer = null;
                //    OoxZipResolver zipResolver = null;
                //    string zipXMLFileName = "input.xml";

                //    try
                //    {
                //        //xrs.ProhibitDtd = true;

                //        // xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(directionXSL));
                //        xrs.XmlResolver = resourceResolver;
                //        string sr = ZipXMLFile(inputFile);
                //        ZipReader archive = ZipFactory.OpenArchive(sr);
                //        source = XmlReader.Create(archive.GetEntry(zipXMLFileName));

                //        XslCompiledTransform transFrom = new XslCompiledTransform();
                //        XsltSettings settings = new XsltSettings(true, false);


                //        Assembly ass = Assembly.Load("excel_oox2uof");//调用ppt_oox2uof.dll程序集
                //        Type t = ass.GetType("oox2uof");//name=oox2uof.xslt
                //        transFrom.Load(t);

                //        if (!originalFile.Equals(string.Empty))
                //        {
                //            zipResolver = new OoxZipResolver(originalFile, resourceResolver);
                //        }
                //        XsltArgumentList parameters = new XsltArgumentList();
                //        parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);

                //        // zip format
                //        parameters.AddParam("outputFile", "", outputFile);
                //        // writer = new OoxZipWriter(inputFile);
                //        writer = new UofZipWriter(outputFile);

                //        if (zipResolver != null)
                //        {

                //            transFrom.Transform(source, parameters, writer);
                //        }
                //        else
                //        {

                //            transFrom.Transform(source, parameters, writer);
                //        }
                //    }
                //    finally
                //    {
                //        if (writer != null)
                //            writer.Close();
                //        if (source != null)
                //            source.Close();
                //    }
            }
            catch (Exception ex)
            {
                logger.Warn(ex.Message);
            }

        }

      

        /// <summary>
        ///  获取自定义列的宽度
        /// </summary>
        /// <param name="sheetPath">sheet*.xml的位置</param>
        /// <returns>所有列的实际宽度</returns>
        private static double[] GetSheetColWidth(XmlNode corrSheetNode, XmlNamespaceManager nm)
        {


            // 默认对应到UOF宽度为54
            string defaultColWidth = "54";

            // 列对应换算公式
            double formula = 54 / 8.38;
            double defaultWidth = Convert.ToDouble(defaultColWidth);
            if (corrSheetNode.SelectSingleNode(".//ws:sheetFormatPr", nm).Attributes["defaultColWidth"] != null)
            {
                defaultColWidth = corrSheetNode.SelectSingleNode(".//ws:sheetFormatPr", nm).Attributes["defaultColWidth"].Value;
                defaultWidth = Convert.ToDouble((defaultColWidth)) * formula;
            }

            int maxCol = 65536;
            double[] colWidth = new double[maxCol];
            for (int i = 1; i < maxCol; i++)
            {
                colWidth[i] = defaultWidth;
            }

            XmlNodeList cols = corrSheetNode.SelectNodes(".//ws:cols/ws:col", nm);
            foreach (XmlNode col in cols)
            {
                string location = col.Attributes["max"].Value;
                string setWidth = col.Attributes["width"].Value;
                double setWidthDouble = Convert.ToDouble(setWidth) * formula;
                colWidth[Convert.ToInt32(location)] = setWidthDouble;
            }
            return colWidth;
        }

        /// <summary>
        ///  获取自定义列的宽度
        /// </summary>
        /// <param name="sheetPath">sheet*.xml的位置</param>
        /// <returns>所有列的实际宽度</returns>
        private static double[] GetSheetRowHeight(XmlNode corrSheetNode, XmlNamespaceManager nm)
        {


            // 默认对应到UOF行高13.5
            string defaultRowHeight = "13.5";

            double defaultHeight = Convert.ToDouble(defaultRowHeight);
            //2014-4-15 update by Qihy
            if (corrSheetNode.SelectSingleNode("//ws:sheetFormatPr", nm) != null && corrSheetNode.SelectSingleNode("//ws:sheetFormatPr", nm).Attributes["defaultRowHeight"] != null)
            {
                defaultRowHeight = corrSheetNode.SelectSingleNode("//ws:sheetFormatPr", nm).Attributes["defaultRowHeight"].Value;
                defaultHeight = Convert.ToDouble(defaultRowHeight);
            }

            int maxRow = 65536;
            double[] rowHeight = new double[maxRow];
            for (int i = 1; i < maxRow; i++)
            {
                rowHeight[i] = defaultHeight;
            }

            XmlNodeList rows = corrSheetNode.SelectNodes("ws:worksheet/ws:sheetData/ws:row", nm);
            foreach (XmlNode row in rows)
            {
                if (row.Attributes["ht"] != null)
                {
                    string location = row.Attributes["r"].Value;
                    string setHeight = row.Attributes["ht"].Value;
                    double setHeightDouble = Convert.ToDouble(setHeight);
                    rowHeight[Convert.ToInt32(location)] = setHeightDouble;
                }
            }
            return rowHeight;
        }



        private static XmlNode SetTwoCellAnchorOff(XmlNode twoCellAnchorChild, double[] colWidth, double[] rowHeight, XmlNamespaceManager nm)
        {
            // XmlNode from = twoCellAnchorFrom.SelectSingleNode("xdr:from", nm);
            //string oriColValue = twoCellAnchorChild.SelectSingleNode("xdr:col", nm).InnerText;
            string oriColOffValue = twoCellAnchorChild.SelectSingleNode("xdr:colOff", nm).InnerText;
            string oriRowOffValue = twoCellAnchorChild.SelectSingleNode("xdr:rowOff", nm).InnerText;

            //int col = Convert.ToInt32(oriColValue);
            double colOff = Convert.ToDouble(oriColOffValue);
            double colValue = 0;
            int adjustColValue = 1;
            double colOffValue = 0.0;

            //int[] colWidth = GetSheetColWidth(corrSheet);
            for (int i = 1; i < colWidth.Length; i++)
            {
                colValue += colWidth[i];
                if (colValue > colOff)
                {
                    adjustColValue = i - 1;
                    colOffValue = colOff - (colValue - colWidth[i]);
                    break;
                }
            }

            double rowOff = Convert.ToDouble(oriRowOffValue);
            double rowValue = 0;
            int adjustRowValue = 1;
            double rowOffValue = 0.0;

            //int[] colWidth = GetSheetColWidth(corrSheet);
            for (int j = 1; j < rowHeight.Length; j++)
            {
                rowValue += rowHeight[j];
                if (rowValue > rowOff)
                {
                    adjustRowValue = j - 1;
                    rowOffValue = rowOff - (rowValue - rowHeight[j]);
                    break;
                }
            }

            // 绝对值=列*54+偏移*28.3/360000
            twoCellAnchorChild.SelectSingleNode("xdr:col", nm).InnerText = Convert.ToString(adjustColValue);
            twoCellAnchorChild.SelectSingleNode("xdr:colOff", nm).InnerText = Convert.ToString(Convert.ToInt32(colOffValue * 360000 / 28.3));
            twoCellAnchorChild.SelectSingleNode("xdr:row", nm).InnerText = Convert.ToString(adjustRowValue);
            twoCellAnchorChild.SelectSingleNode("xdr:rowOff", nm).InnerText = Convert.ToString(Convert.ToInt32(rowOffValue * 360000 / 28.3));


            return twoCellAnchorChild;
        }

        private static XmlNode SetTwoCellAnchorOff2(XmlNode twoCellAnchorChild, double[] colWidth, double[] rowHeight, XmlNamespaceManager nm)
        {
            // XmlNode from = twoCellAnchorFrom.SelectSingleNode("xdr:from", nm);
            //string oriColValue = twoCellAnchorChild.SelectSingleNode("xdr:col", nm).InnerText;
            string oriColOffValue = twoCellAnchorChild.SelectSingleNode("xdr:colOff", nm).InnerText;
            string oriRowOffValue = twoCellAnchorChild.SelectSingleNode("xdr:rowOff", nm).InnerText;

            //int col = Convert.ToInt32(oriColValue);
            double colOff = Convert.ToDouble(oriColOffValue);
            double colValue = 0;
            int adjustColValue = 1;
            double colOffValue = 0.0;

            //int[] colWidth = GetSheetColWidth(corrSheet);
            for (int i = 1; i < colWidth.Length; i++)
            {
                colValue += colWidth[i];
                if (colValue > colOff)
                {
                    adjustColValue = i - 1;
                    colOffValue = colOff - (colValue - colWidth[i]);
                    break;
                }
            }

            double rowOff = Convert.ToDouble(oriRowOffValue);
            double rowValue = 0;
            int adjustRowValue = 1;
            double rowOffValue = 0.0;

            //int[] colWidth = GetSheetColWidth(corrSheet);
            for (int j = 1; j < rowHeight.Length; j++)
            {
                rowValue += rowHeight[j];
                if (rowValue > rowOff)
                {
                    adjustRowValue = j - 1;
                    rowOffValue = rowOff - (rowValue - rowHeight[j]);
                    break;
                }
            }

            // 绝对值=列*54+偏移*28.3/360000
            twoCellAnchorChild.SelectSingleNode("xdr:col", nm).InnerText = Convert.ToString(adjustColValue);
            twoCellAnchorChild.SelectSingleNode("xdr:colOff", nm).InnerText = Convert.ToString(Convert.ToInt32(colOffValue * 360000 / 28.3));
            twoCellAnchorChild.SelectSingleNode("xdr:row", nm).InnerText = Convert.ToString(adjustRowValue + 1);
            twoCellAnchorChild.SelectSingleNode("xdr:rowOff", nm).InnerText = Convert.ToString(Convert.ToInt32(rowOffValue * 360000 / 28.3));


            return twoCellAnchorChild;
        }

        /// <summary>
        /// 设置drawing*.xml文件中所有from和to的实际绝对坐标值（非相对值）
        /// </summary>
        /// <param name="xlsxLocation">XLSX文件夹位置</param>
        private static void SetDrawingTwoCellAnchorValue(string mainOutputFile, string equAfterMain)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(mainOutputFile);

            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("pzip", TranslatorConstants.XMLNS_PZIP);
            nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
            nm.AddNamespace("xdr", TranslatorConstants.XMLNS_XDR);

            XmlNodeList entryNodes = xdoc.SelectNodes("//pzip:entry[@pzip:target]", nm);
            LinkedList<XmlNode> drawingNodes = new LinkedList<XmlNode>();
            drawingNodes.Clear();
            foreach (XmlNode entryNode in entryNodes)
            {
                string target = entryNode.Attributes["pzip:target"].Value;
                //if (target.Contains("xl/drawings/drawing"))
                if (target.Contains("xl/drawings/drawing") && !target.Contains("xl/drawings/drawingchart"))
                {
                    drawingNodes.AddLast(entryNode);
                }
            }

            LinkedList<XmlNode>.Enumerator enumer = drawingNodes.GetEnumerator();
            while (enumer.MoveNext())
            {
                XmlNode current = enumer.Current;
                string target = current.Attributes["pzip:target"].Value;

                // 19: "xl/drawings/drawing*.xml"从*.xml处开始匹配
                string corrSheet = "xl/worksheets/sheet" + target.Substring(19);
                XmlNode corrSheetNode = xdoc.SelectSingleNode("//pzip:entry[@pzip:target='" + corrSheet + "']", nm);
                //2014-4-15 update by Qihy
                double[] colWidth = null;
                if (corrSheetNode != null)
                {
                    colWidth = GetSheetColWidth(corrSheetNode, nm);
                }
                double[] rowHeight = null;
                if (corrSheetNode != null)
                {
                    rowHeight = GetSheetRowHeight(corrSheetNode, nm);
                }

                XmlNodeList twoCellAnchors = current.SelectNodes(".//xdr:twoCellAnchor", nm);


                foreach (XmlNode twoCellAnchor in twoCellAnchors)
                {

                    SetTwoCellAnchorOff(twoCellAnchor.SelectSingleNode("xdr:from", nm), colWidth, rowHeight, nm);
                    SetTwoCellAnchorOff2(twoCellAnchor.SelectSingleNode("xdr:to", nm), colWidth, rowHeight, nm);
                }

                XmlNodeList oneCellAnchors = current.SelectNodes(".//xdr:oneCellAnchor", nm);
                foreach (XmlNode oneCellAnchor in oneCellAnchors)
                {
                    SetTwoCellAnchorOff(oneCellAnchor.SelectSingleNode("xdr:from", nm), colWidth, rowHeight, nm);
                }

                //xdoc.Save(mainOutputFile);

            }

            XmlTextWriter tw = new XmlTextWriter(equAfterMain, Encoding.UTF8);
            xdoc.Save(tw);
            tw.Close();
        }

    }
}


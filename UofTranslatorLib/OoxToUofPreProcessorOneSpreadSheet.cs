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
using System.Collections;
using System.Data;
using System.Text.RegularExpressions;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// pretreatment of Spreadsheet from OOX->UOF
    /// </summary>
    /// <author>Linyaohu</authord>
    class OoxToUofPreProcessorOneSpreadSheet : AbstractProcessor
    {
        private static ILog logger = LogManager.GetLogger(typeof(OOXToUOFPreProcessorOnePresentation).FullName);

        public override bool transform()
        {
            string extractPath = Path.GetDirectoryName(outputFile) + "\\";
            string extractXlsxPath = extractPath + "XLSX" + "\\";
            if (!Directory.Exists(extractPath))
                Directory.CreateDirectory(extractPath);
            if (!Directory.Exists(extractXlsxPath))
                Directory.CreateDirectory(extractXlsxPath);
            ZipReader archive = ZipFactory.OpenArchive(inputFile);
            archive.ExtractOfficeDocument(inputFile, extractXlsxPath);
            string prefix = this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + ".Excel.oox2uof";
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

            SetDrawingTwoCellAnchorValue(extractXlsxPath);

            string preOutPutFileName = extractPath + "sheetMResult.xml";
            string normalToRCFileName = extractPath + "normalToRCResult.xml";
            //FileStream fs = new FileStream(preOutPutFileName, FileMode.Create);
            XmlTextWriter fs = new XmlTextWriter(preOutPutFileName, Encoding.UTF8);

            AdjustCommentStructure(extractPath + @"XLSX\xl\");

            //MoveChartRelFiles(extractPath + @"XLSX\xl\charts\_rels\", extractPath + @"XLSX\xl\charts\");
            addChartFileName(extractPath + @"XLSX\xl\charts\");
            MoveChartRelFiles(extractPath + @"XLSX\xl\charts\_rels\", extractPath + @"XLSX\xl\charts\");
            addChartToDrawing(extractPath + @"XLSX\xl\drawings\");
            addRelsToSheet(extractPath + @"XLSX\xl\worksheets\");
            addRelsToDrawing(extractPath + @"XLSX\xl\drawings\");
            addHyperlinks(extractPath + @"XLSX\xl\worksheets");
            addformular(extractPath + @"XLSX\xl\worksheets");
            AddxiaoGroup(extractPath + @"XLSX\xl\worksheets");

            //AddGroup(extractPath + @"XLSX\xl\worksheets");
            cellFormatTrans(extractPath + @"XLSX\xl\styles.xml");
            try
            {
                File.Copy(extractPath + @"XLSX\xl\workbook.xml", extractPath + "workbook.xml", true);
                XmlReader xr = XmlReader.Create(extractPath + @"\workbook.xml");
                // XPathDocument doc = new XPathDocument(extractPath + @"\workboko.xml");
                XslCompiledTransform transFrom = new XslCompiledTransform();
                XsltSettings setting = new XsltSettings(true, false);
                XmlUrlResolver xur = new XmlUrlResolver();
                transFrom.Load(extractPath + @"\pretreatment.xsl", setting, xur);
                //XPathNavigator nav = ((IXPathNavigable)doc).CreateNavigator();
                transFrom.Transform(xr, null, fs); ;
                fs.Close();
            }
            catch (Exception e)
            {
                logger.Error("Fail in OoxToUofPreProcessorOneSpreadSheet: " + e.Message);
                logger.Error(e.StackTrace);
                isSuccess = false;
                throw new Exception("Fail in OoxToUofPreProcessorOneSpreadSheet");
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }
            string dataValidationPre = extractPath + "dataValdPre.xml";
            DataValidationPretreatment(preOutPutFileName, dataValidationPre);

            NormalToRC(dataValidationPre, normalToRCFileName);

            //add talbe styles
            AddTablesStyle(extractPath, normalToRCFileName, outputFile);

            // RC 引用处理
            return isSuccess;

        }

        #region fix row&Height strench bug
        /// <summary>
        ///  获取自定义列的宽度
        /// </summary>
        /// <param name="sheetPath">sheet*.xml的位置</param>
        /// <returns>所有列的实际宽度</returns>
        private static double[] GetSheetColWidth(string sheetPath)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(sheetPath);

            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);

            // 默认对应到UOF宽度为54
            string defaultColWidth = "54";

            // 列对应换算公式
            double formula = 54 / 8.38;
            double defaultWidth = Convert.ToDouble(defaultColWidth);
            //2014-4-15 update by Qihy
            if (xdoc.SelectSingleNode("//ws:sheetFormatPr", nm) != null && xdoc.SelectSingleNode("//ws:sheetFormatPr", nm).Attributes["defaultColWidth"] != null)
            {
                defaultColWidth = xdoc.SelectSingleNode("//ws:sheetFormatPr", nm).Attributes["defaultColWidth"].Value;
                defaultWidth = Convert.ToDouble((defaultColWidth)) * formula;
            }

            int maxCol = 65536;
            double[] colWidth = new double[maxCol];
            for (int i = 1; i < maxCol; i++)
            {
                colWidth[i] = defaultWidth;
            }

            XmlNodeList cols = xdoc.SelectNodes("//ws:cols/ws:col", nm);
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
        private static double[] GetSheetRowHeight(string sheetPath)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(sheetPath);

            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);

            // 默认对应到UOF行高13.5
            string defaultRowHeight = "13.5";

            double defaultHeight = Convert.ToDouble(defaultRowHeight);
            //2014-4-15 update by Qihy
            if (xdoc.SelectSingleNode("//ws:sheetFormatPr", nm) != null && xdoc.SelectSingleNode("//ws:sheetFormatPr", nm).Attributes["defaultRowHeight"] != null)
            {
                defaultRowHeight = xdoc.SelectSingleNode("//ws:sheetFormatPr", nm).Attributes["defaultRowHeight"].Value;
                defaultHeight = Convert.ToDouble(defaultRowHeight);
            }

            int maxRow = 65536;
            double[] rowHeight = new double[maxRow];
            for (int i = 1; i < maxRow; i++)
            {
                rowHeight[i] = defaultHeight;
            }

            XmlNodeList rows = xdoc.SelectNodes("//ws:row", nm);
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

        private static XmlNode SetTwoCellAnchorOff(XmlNode twoCellAnchorChild, string corrSheet, double[] colWidth, double[] rowHeight, XmlNamespaceManager nm)
        {
            // XmlNode from = twoCellAnchorFrom.SelectSingleNode("xdr:from", nm);
            string oriColValue = twoCellAnchorChild.SelectSingleNode("xdr:col", nm).InnerText;
            string oriColOffValue = twoCellAnchorChild.SelectSingleNode("xdr:colOff", nm).InnerText;
            int col = Convert.ToInt32(oriColValue);
            double colValue = 0;
            //int[] colWidth = GetSheetColWidth(corrSheet);
            for (int i = 1; i <= col; i++)
            {
                colValue += colWidth[i];
            }


            // row
            string oriRowValue = twoCellAnchorChild.SelectSingleNode("xdr:row", nm).InnerText;
            string oriRowOffValue = twoCellAnchorChild.SelectSingleNode("xdr:rowOff", nm).InnerText;
            int row = Convert.ToInt32(oriRowValue);
            double rowValue = 0;
            //int[] colWidth = GetSheetColWidth(corrSheet);
            for (int j = 1; j <= row; j++)
            {
                rowValue += rowHeight[j];
            }


            // 绝对值=列*54+偏移*28.3/360000
            twoCellAnchorChild.SelectSingleNode("xdr:colOff", nm).InnerText = Convert.ToString((colValue + Convert.ToInt32(oriColOffValue) * 28.3 / 360000));
            twoCellAnchorChild.SelectSingleNode("xdr:rowOff", nm).InnerText = Convert.ToString((rowValue + Convert.ToInt32(oriRowOffValue) * 28.3 / 360000));

            return twoCellAnchorChild;
        }

        /// <summary>
        /// 设置drawing*.xml文件中所有from和to的实际绝对坐标值（非相对值）
        /// </summary>
        /// <param name="xlsxLocation">XLSX文件夹位置</param>
        private static void SetDrawingTwoCellAnchorValue(string xlsxLocation)
        {
            string drawingLocation = xlsxLocation + Path.AltDirectorySeparatorChar + "xl" + Path.AltDirectorySeparatorChar + "drawings" + Path.AltDirectorySeparatorChar;
            //string sheetLocation = xlsxLocation + Path.AltDirectorySeparatorChar + "xl" + Path.AltDirectorySeparatorChar + "worksheets"+Path.AltDirectorySeparatorChar;
            if (Directory.Exists(drawingLocation))
            {
                string[] drawingFiles = Directory.GetFiles(drawingLocation);
                int i = 0;
                foreach (string drawingFile in drawingFiles)
                {
                    FileInfo fi = new FileInfo(drawingFile);
                    i++;

                    // 不处理vmlDrawing*.xml的情况
                    //2014-4-15 update by Qihy,修复bug3148 (Transition&Strict)OOX->UOF->OOX 第一轮转换失败
                    //if (!fi.Name.Contains("vml") && !fi.Name.Contains("drawingchartsheet"))
                    if (!fi.Name.Contains("vml"))
                    {
                        // 7: drawing1.xml 只需获取1.xml
                        string drawingSubString = fi.Name.Substring(7);

                        string sheetLocation = GetsheetPosition(xlsxLocation, i);
                        // string corrSheet = sheetLocation + "sheet" + drawingSubString;
                        string corrSheet = xlsxLocation + Path.AltDirectorySeparatorChar + "xl" + Path.AltDirectorySeparatorChar + sheetLocation;

                        double[] colWidth = GetSheetColWidth(corrSheet);
                        double[] rowHeight = GetSheetRowHeight(corrSheet);

                        XmlDocument xdoc = new XmlDocument();
                        xdoc.Load(drawingFile);
                        XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
                        nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
                        nm.AddNamespace("xdr", TranslatorConstants.XMLNS_XDR);

                        XmlNodeList twoCellAnchors = xdoc.SelectNodes("//xdr:twoCellAnchor", nm);
                        foreach (XmlNode twoCellAnchor in twoCellAnchors)
                        {
                            SetTwoCellAnchorOff(twoCellAnchor.SelectSingleNode("xdr:from", nm), corrSheet, colWidth, rowHeight, nm);
                            SetTwoCellAnchorOff(twoCellAnchor.SelectSingleNode("xdr:to", nm), corrSheet, colWidth, rowHeight, nm);
                        }
                        XmlNodeList oneCellAnchors = xdoc.SelectNodes("//xdr:oneCellAnchor", nm);
                        foreach (XmlNode oneCellAnchor in oneCellAnchors)
                        {
                            SetTwoCellAnchorOff(oneCellAnchor.SelectSingleNode("xdr:from", nm), corrSheet, colWidth, rowHeight, nm);
                            // SetTwoCellAnchorOff(twoCellAnchor.SelectSingleNode("xdr:to", nm), corrSheet, colWidth, nm);
                        }
                        xdoc.Save(drawingFile);
                    }
                }

            }

        }

        private static String GetsheetPosition(string xlsxLocation, int i)
        {
            String sheetLocation = "";
            string chartsheetLocation = xlsxLocation + Path.AltDirectorySeparatorChar + "xl" + Path.AltDirectorySeparatorChar + "chartsheets" + Path.AltDirectorySeparatorChar;
            string workbookLocation = xlsxLocation + Path.AltDirectorySeparatorChar + "xl" + Path.AltDirectorySeparatorChar + "workbook.xml";
            string workbookRLocation = xlsxLocation + Path.AltDirectorySeparatorChar + "xl" + Path.AltDirectorySeparatorChar + "_rels" + Path.AltDirectorySeparatorChar + "workbook.xml.rels";
            XmlDocument xdoc1 = new XmlDocument();
            XmlDocument xdoc2 = new XmlDocument();
            string workbookFile = workbookLocation;
            string workbookRFile = workbookRLocation;

            xdoc1.Load(workbookFile);
            xdoc2.Load(workbookRFile);
            XmlNamespaceManager nm1 = new XmlNamespaceManager(xdoc1.NameTable);
            XmlNamespaceManager nm2 = new XmlNamespaceManager(xdoc2.NameTable);
            nm1.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
            nm1.AddNamespace("r", TranslatorConstants.XMLNS_R);
            nm2.AddNamespace("rel", TranslatorConstants.XMLNS_REL);
            XmlNodeList sheets = xdoc1.SelectNodes("//ws:sheet", nm1);
            XmlNodeList relationships = xdoc2.SelectNodes("//rel:Relationship", nm2);
            String wbId = "rId1";
            int j = 0;
            foreach (XmlNode sheet in sheets)
            {
                j++;
                if (j == i)
                {
                    wbId = sheet.Attributes["r:id"].Value;
                    break;
                }

            }
            foreach (XmlNode relationship in relationships)
            {
                String wbrId = relationship.Attributes["Id"].Value;
                if (wbrId.Equals(wbId))
                {
                    sheetLocation = relationship.Attributes["Target"].Value;
                }

            }
            return sheetLocation;
        }

        #endregion

        #region improve benmark

        private void AdjustCommentStructure(string xlPath)
        {
            string wsRelsPath = xlPath + Path.AltDirectorySeparatorChar + "worksheets/_rels/";
            string commentsPath = xlPath;
            string wsPath = xlPath + Path.AltDirectorySeparatorChar + "worksheets/";

            if (Directory.Exists(wsRelsPath))
            {
                string[] sheetsRels = Directory.GetFiles(wsRelsPath);
                if (sheetsRels.Length > 0)
                {
                    foreach (string sheetRel in sheetsRels)
                    {
                        string[] commentDrawingRel = GetRelComment(sheetRel);
                        if (commentDrawingRel[0] != null)
                        {
                            MoveCommentContent(xlPath, sheetRel, commentDrawingRel[0], commentDrawingRel[1]);
                        }
                    }
                }
            }

        }

        /// <summary>
        ///  Get the relation comment file from sheet* relation file
        /// </summary>
        /// <param name="sheetRel"></param>
        /// <returns></returns>
        private string[] GetRelComment(string sheetRel)
        {
            string[] commentDrawingRel = new string[2];
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(sheetRel);
            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("pr", TranslatorConstants.XMLNS_REL);

            XmlNode commentTargetNode = xdoc.SelectSingleNode("pr:Relationships/pr:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments']", nm);

            string target = string.Empty;
            if (commentTargetNode != null)
            {
                target = commentTargetNode.Attributes["Target"].Value;

                // target=../comment*.xml
                commentDrawingRel[0] = target.Substring(3);
                XmlNode vmlDrawingNode = xdoc.SelectSingleNode("pr:Relationships/pr:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing']", nm);

                if (vmlDrawingNode != null)
                {
                    string vmlTarget = vmlDrawingNode.Attributes["Target"].Value;

                    // vmlTarget=../drawings/vmlDrawing1.vml
                    commentDrawingRel[1] = vmlTarget.Substring(3);
                }
            }
                return commentDrawingRel;
        }

        private void MoveCommentContent(string xlPath, string sheetRel, string commentFile,string drawingFile)
        {
            FileInfo fi = new FileInfo(sheetRel);
            string sheetFileName = fi.Name;
            string sheetFile = sheetFileName.Substring(0, sheetFileName.IndexOf("."))+".xml";
            string sheetFilePath = xlPath + Path.AltDirectorySeparatorChar + "worksheets/" + sheetFile;
            string commentFilePath = xlPath + Path.AltDirectorySeparatorChar + commentFile;
            string vmlDrawingFilePath = xlPath + Path.AltDirectorySeparatorChar + drawingFile;

            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(commentFilePath);
            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);

            XmlDocument xSheetDoc = new XmlDocument();
            xSheetDoc.Load(sheetFilePath);
            XmlNamespaceManager nms = new XmlNamespaceManager(xSheetDoc.NameTable);
            nms.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
            nms.AddNamespace("v", TranslatorConstants.XMLNS_V);

            XmlDocument vDoc = new XmlDocument();
            if (File.Exists(vmlDrawingFilePath))
            {
                vDoc.Load(vmlDrawingFilePath);
            }
            XmlNamespaceManager nmv = new XmlNamespaceManager(vDoc.NameTable);
            nmv.AddNamespace("v", TranslatorConstants.XMLNS_V);



            // 5: 'sheet*.xml中的‘sheet’长度'
            string phNo = "comments" + sheetFile.Substring(5, sheetFile.IndexOf(".") - 5);

            // i:第几个comment元素
            int i = 1;

            XmlNodeList commentList = xdoc.SelectNodes("ws:comments/ws:commentList/ws:comment", nm);
            foreach (XmlNode comment in commentList)
            {
                XmlNode cmTmpNode = xSheetDoc.ImportNode(comment, true);

                // 增加 图形引用属性＝'comments*.*, 是否显示visibility
                XmlAttribute phNoRef = xSheetDoc.CreateAttribute("phNoRef");
                phNoRef.Value = phNo + "." + i;
                cmTmpNode.Attributes.Append(phNoRef);

                if (vDoc != null)
                {
                    XmlNode shape = vDoc.SelectSingleNode("//v:shape[" + i + "]", nmv);
                    if (shape != null)
                    {
                        string style = shape.Attributes["style"].Value;

                        XmlAttribute vsblty = xSheetDoc.CreateAttribute("visibility");
                        vsblty.Value = "hidden";
                        if (style.Contains("visibility"))
                        {
                            cmTmpNode.Attributes.Append(vsblty);
                        }
                    }
                }

                i++;

                string commentRef = comment.Attributes["ref"].Value;
                string commentRow = Regex.Replace(commentRef, "[A-Z]", "", RegexOptions.IgnoreCase);

                XmlNode worksheetRow = xSheetDoc.SelectSingleNode("ws:worksheet/ws:sheetData/ws:row[@r='" + commentRow + "']", nms);
                XmlNode worksheetRowCol = worksheetRow.SelectSingleNode("ws:c[@r='" + commentRef + "']", nms);
                if (worksheetRowCol != null)
                {
                    worksheetRowCol.AppendChild(cmTmpNode);
                }
                else
                {
                    XmlElement col = xSheetDoc.CreateElement("ws", "c", TranslatorConstants.XMLNS_WS);
                    XmlAttribute colR = xSheetDoc.CreateAttribute("r");
                    colR.Value = commentRef;
                    col.Attributes.Append(colR);

                    col.AppendChild(cmTmpNode);
                    worksheetRow.AppendChild(col);
                }
            }

            xSheetDoc.Save(sheetFilePath);
        }

        #endregion

        /// <summary>
        ///  add the XL/tables/table*.xml to pretreatment1
        /// </summary>
        /// <param name="inpute"></param>
        /// <param name="output"></param>
        private static void AddTablesStyle(string tmpPath, string inputFile, string outputFile)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFile);

            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);

            string tablePath = tmpPath + Path.AltDirectorySeparatorChar + "XLSX" + Path.AltDirectorySeparatorChar + "xl" + Path.AltDirectorySeparatorChar + "tables";
            if (Directory.Exists(tablePath))
            {
                String[] tableXMLFiles = Directory.GetFiles(tablePath);
                foreach (string tableXMLFile in tableXMLFiles)
                {
                    FileInfo fi = new FileInfo(tableXMLFile);
                    XmlElement tableNode = xdoc.CreateElement("ws", "tables", TranslatorConstants.XMLNS_WS);
                    XmlAttribute tableName = xdoc.CreateAttribute("tableName");
                    tableName.Value = fi.Name;
                    tableNode.Attributes.Append(tableName);


                    XmlDocument xDocTables = new XmlDocument();
                    xDocTables.Load(tableXMLFile);
                    XmlNode tableTmpNode = xdoc.ImportNode(xDocTables.LastChild, true);
                    tableNode.AppendChild(tableTmpNode);

                    xdoc.SelectSingleNode("ws:spreadsheets", nm).AppendChild(tableNode);

                }
                xdoc.Save(outputFile);
            }
            else
            {
                File.Copy(inputFile, outputFile);
            }
        }

        private static void cellFormatTrans(String path)
        {

            XmlDocument style = new XmlDocument();
            style.Load(path);
            XmlNameTable nTable = style.NameTable;
            XmlNamespaceManager nm = new XmlNamespaceManager(nTable);
            nm.AddNamespace("ws", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            XmlNodeList r = style.SelectNodes("//ws:numFmts/ws:numFmt", nm);
            if (r != null)
            {
                string type = "";
                string formatCode = "";
                foreach (XmlNode n in r)
                {
                    formatCode = n.Attributes[1].Value.ToString();
                    GetProperNumFmts(ref type, ref formatCode);

                    XmlAttribute formatType = style.CreateAttribute("formatType");
                    formatType.Value = type;
                    n.Attributes[1].Value = formatCode;
                    n.Attributes.Append(formatType);
                }
            }

            style.Save(path);

        }

        private static void GetProperNumFmts(ref string type, ref string value)
        {
            string newValue = "";
            string label = "";
            if (value.Contains("¥"))
                label = "¥";
            else if (value.Contains("$"))
                label = "$";
            else
                label = "P";

            value = value.Replace(label, "O");

            if (value.Equals(@"0.00_ ;[Red]\-0.00\ "))
            {
                newValue = "0.00_ ;[Red]-0.00_ ";
                type = "number";
            }
            else if (value.Equals(@"#,##0.00_ ;[Red]\-#,##0.00\ "))//row3
            {
                newValue = "#,##0.00_ ;[Red]-#,##0.00_ ";
                type = "number";
            }
            else if (value.Equals(@"0.00_ "))
            {
                newValue = "0.00_ ";
                type = "number";
            }
            else if (value.Equals("0.00;[Red]0.00"))
            {
                newValue = "0.00;[Red]0.00";
                type = "number";
            }
            else if (value.Equals(@"0.00_);\(0.00\)"))
            {
                newValue = "0.00_);(0.00)";
                type = "number";
            }
            else if (value.Equals(@"0.00_);[Red]\(0.00\)"))
            {
                newValue = "0.00_);[Red](0.00)";
                type = "number";
            }
            else if (value.Equals("0.000_ "))
            {
                newValue = "0.000_ ";
                type = "number";
            }
            else if (value.Equals("[$O-804]#,##0.00;[$O-804]\\-#,##0.00"))//row4
            {
                newValue = "[$￥-411]#,##0.00;-[$￥-411]#,##0.00";
                type = "currency";
            }
            else if (value.Equals("#,##0.000[O₮-450];\\-#,##0.000[O₮-450]"))//row5
            {
                newValue = "￥#,##0.000;￥-#,##0.000";
                type = "currency";
            }
            //else if (value.Equals(""))//row18
            //{
            //    newValue = "#,##0.00";
            //    type = "currency";
            //}
            else if (value.Equals("\"O\"#,##0.00;[Red]\"O\"\\-#,##0.00"))
            {
                newValue = "O#,##0.00;[Red]O-#,##0.00";
                type = "currency";
            }
            else if (value.Equals("\"O\"#,##0.00;\"O\"\\-#,##0.00"))
            {
                newValue = "O#,##0.00;O-#,##0.00";
                type = "currency";
            }
            else if (value.Equals("\"O\"#,##0.00;[Red]\"O\"#,##0.00"))
            {
                newValue = "O#,##0.00;[Red]O#,##0.00";
                type = "currency";
            }
            else if (value.Equals("\"O\"#,##0.00_);[Red]\\(\"O\"#,##0.00\\)"))
            {
                newValue = "O#,##0.00_);[Red](O#,##0.00)";
                type = "currency";
            }
            else if (value.Equals("\"O\"#,##0.00_);\\(\"O\"#,##0.00\\)"))
            {
                newValue = "O#,##0.00_);(O#,##0.00)";
                type = "currency";
            }
            else if (value.Equals("_ \"O\"* #,##0_ ;_ \"O\"* \\-#,##0_ ;_ \"O\"* \"-\"_ ;_ @_ "))
            {
                newValue = "[$¥-804]#,##0;-[$¥-804]#,##0";
                type = "currency";
            }
            else if (value.Equals("[OO-1009]#,##0"))
            {
                newValue = "[$$-1009]#,##0;-[$$-1009]#,##0";
                type = "currency";
            }
            // 20130325 add by xuzhenwei BUG_2754:功能测试 部分单元格格式不正确 start
            else if (value.Equals("_-[$$-409]* #,##0.00_ ;_-[$$-409]* \\-#,##0.00\\ ;_-[$$-409]* \"-\"??_ ;_-@_ "))//row6
            {
                newValue = "_-[$￥-411]* #,##0.000_-;-[$￥-411]* #,##0.000_-;_-[$￥-411]* \"-\"???_-;_-@_-";
                type = "accounting";
            }
            //end
            else if (value.Equals("_ [$O-804]* #,##0.000_ ;_ [$O-804]* \\-#,##0.000_ ;_ [$O-804]* \"-\"???_ ;_ @_ "))//row6
            {
                newValue = "_-[$￥-411]* #,##0.000_-;-[$￥-411]* #,##0.000_-;_-[$￥-411]* \"-\"???_-;_-@_-";
                type = "accounting";
            }
            else if (value.Equals("_-[OO-409]* #,##0.00_ ;_-[OO-409]* \\-#,##0.00\\ ;_-[OO-409]* \"-\"??_ ;_-@_ "))//row6
            {
                newValue = "_-[$$-409]* #,##0.00_ ;_-[$$-409]* -#,##0.00 ;_-[$$-409]* \"-\"??_ ;_-@_ ";
                type = "accounting";
            }
            else if (value.Equals("_ [O€-2]\\ * #,##0.00_ ;_ [O€-2]\\ * \\-#,##0.00_ ;_ [O€-2]\\ * \"-\"??_ ;_ @_ "))//D2
            {
                newValue = "_ [$€-2] * #,##0.00_ ;_ [$€-2] * -#,##0.00_ ;_ [$€-2] * \"-\"??_ ;_ @_ ";
                type = "accounting";
            }
            else if (value.Equals("_ [$O-804]* #,##0.00_ ;_ [$O-804]* \\-#,##0.00_ ;_ [$O-804]* \"-\"??_ ;_ @_ "))//C2会计 日语 小数2位，numberFormat.xlsx  
            {
                newValue = "_-[$￥-411]* #,##0.00_-;-[$￥-411]* #,##0.00_-;_-[$￥-411]* \"-\"??_-;_-@_-";
                type = "accounting";
            }
            // 20130325 update by xuzhenwei BUG_2754:功能测试 部分单元格格式不正确 and BUG_2730:集成OO-UOF 单元格格式自定义转换后为常规 start
            else if (value.Equals("\"￥\"#,##0.00;\"￥\"\\-#,##0.00"))//自定义类型的会计格式
            {
                newValue = "\"￥\"#,##0.00;\"￥\"\\-#,##0.00";
                type = "accounting";
            }
            else if (value.Equals("_-[$O-411]* #,##0.00_-;\\-[$O-411]* #,##0.00_-;_-[$O-411]* \"-\"??_-;_-@_-"))//会计 日本
            {
                newValue = "_-[$¥-411]* #,##0.00_-;-[$¥-411]* #,##0.00_-;_-[$¥-411]* \"-\"??_-;_-@_-";
                type = "accounting";
            }
            else if (value.Equals("_-[O£-809]* #,##0.00_-;\\-[O£-809]* #,##0.00_-;_-[O£-809]* \"-\"??_-;_-@_-"))//E2
            {
                newValue = "_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* \"-\"??_-;_-@_- ";
                type = "accounting";
            }
            else if (value.Equals("_ [Ofr.-100C]\\ * #,##0.00_ ;_ [Ofr.-100C]\\ * \\-#,##0.00_ ;_ [Ofr.-100C]\\ * \"-\"??_ ;_ @_ "))//E2
            {
                newValue = "_ [$SFr.-100C] * #,##0.00_ ;_ [$SFr.-100C] * -#,##0.00_ ;_ [$SFr.-100C] * \"-\"??_ ;_ @_ ";
                type = "accounting";
            }
            // end
            else if (value.Contains("_ \"O\"* #,##0.00_ ;_ \"O\"* \\-#,##0.00_ ;_ \"O\"* \"-\"??_ ;_ @_ "))
            {
                newValue = "_ O* #,##0.00_ ;_ O* -#,##0.00_ ;_ O* \"-\"??_ ;_ @_ ";
                type = "accounting";
            }
            else if (value.Contains("_-\\O* #,##0.00_ ;_-\\O* \\-#,##0.00\\ ;_-\\O* \"-\"??_ ;_-@_ "))
            {
                newValue = "_-O* #,##0.00_ ;_-O* -#,##0.00 ;_-O* \"-\"??_ ;_-@_ ";
                type = "accounting";
            }
            else if (value.Equals("[O-409]yyyy/m/d\\ h:mm\\ AM/PM;@"))//row7
            {
                newValue = "yyyy-m-d h:mm AM/PM;@";
                type = "date";
            }
            else if (value.Equals("[DBNum1][O-804]yyyy\"年\"m\"月\"d\"日\";@"))//row8
            {
                newValue = "[DBNum1]yyyy\"年\"m\"月\"d\"日\";@";
                type = "date";
            }
            else if (value.Equals("[O-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy"))//row9
            {
                newValue = "yyyy\"年\"m\"月\"d\"日\";@";
                type = "date";
            }
            else if (value.Equals("[O-404]aaa;@"))//row10
            {
                newValue = "aaa;@";
                type = "date";
            }
            else if (value.Equals("[O-409]dd\\-mmm\\-yy;@"))//row11
            {
                newValue = "dd-mmm-yy;@";
                type = "date";
            }
            else if (value.Equals("[O-409]mmmm\\-yy;@"))//row11
            {
                newValue = "mmmm-yy;@";
                type = "date";
            }
            else if (value.Equals("yyyy\\-m\\-d;@"))
            {
                newValue = "yyyy-m-d;@";
                type = "date";
            }
            else if (value.Contains("[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\";@"))
            {
                newValue = "[DBNum1]yyyy\"年\"m\"月\"d\"日\";@";
                type = "date";
            }
            else if (value.Contains("[DBNum1][$-804]yyyy\"年\"m\"月\";@"))
            {
                newValue = "[DBNum1]yyyy\"年\"m\"月\";@";
                type = "date";
            }

            else if (value.Contains("[DBNum1][$-804]m\"月\"d\"日\";@"))
            {
                newValue = "[DBNum1]m\"月\"d\"日\";@";
                type = "date";
            }
            else if (value.Contains("yyyy\"年\"m\"月\"d\"日\";@"))
            {
                newValue = "yyyy\"年\"m\"月\"d\"日\";@";
                type = "date";
            }
            else if (value.Contains("yyyy\"年\"m\"月\";@"))
            {
                newValue = "yyyy\"年\"m\"月\";@";
                type = "date";
            }
            else if (value.Contains("m\"月\"d\"日\";@"))
            {
                newValue = "m\"月\"d\"日\";@";
                type = "date";
            }


            else if (value.Contains("[$-804]aaaa;@") || value.Contains("[O-804]aaaa;@"))
            {
                newValue = "aaaa;@";
                type = "date";
            }

            else if (value.Contains("[$-804]aaa;@"))
            {
                newValue = "aaa;@";
                type = "date";
            }
            else if (value.Contains("yyyy/m/d;@"))
            {
                newValue = "yyyy-m-d;@";
                type = "date";
            }
            else if (value.Contains("[O-409]yyyy/m/d\\ h:mm\\ AM/PM;@"))
            {
                newValue = "yyyy-m-d h:mm AM/PM;@";
                type = "date";
            }
            else if (value.Contains("yyyy/m/d\\ h:mm;@"))
            {
                newValue = "yyyy-m-d h:mm;@";
                type = "date";
            }


            else if (value.Contains("yy/m/d;@"))
            {
                newValue = "m-d;@";
                type = "date";
            }

            else if (value.Contains("mm/dd/yy;@"))
            {
                newValue = "mm-dd-yy;@";
                type = "date";
            }
            else if (value.Contains("[$-409]d\\-mmm\\-yy;@"))
            {
                newValue = "d-mmm;@";
                type = "date";
            }
            else if (value.Contains("[$-409]d\\-mmm\\-yy;@"))
            {
                newValue = "d-mmm-yy;@";
                type = "date";
            }
            else if (value.Contains(@"[$-409]dd\-mmm\-yy;@"))
            {
                newValue = "dd-mmm-yy;@";
                type = "date";
            }

            else if (value.Contains(@"[$-409]mmm\-yy;@"))
            {
                newValue = "mmm-yy;@";
                type = "date";
            }
            else if (value.Contains(@"[$-409]mmmm\-yy;@"))
            {
                newValue = "mmmm-yy;@";
                type = "date";
            }
            else if (value.Contains("[$-409]mmmmm;@"))
            {
                newValue = @"mmmmm;@";
                type = "date";
            }
            else if (value.Contains(@"[$-409]mmmmm\-yy;@"))
            {
                newValue = "mmmmm-yy;@";
                type = "date";
            }

           //2014-3-24 add by Qihy bug_3186 互操作OO-UOF-OO 日期格式14-Mar转换不正确的问题(此处Equals("[$-409]d\\-mmm;@"")或者Equals(@"[$-409]d\-mmm;@"")都不匹配) start
            else if (value.Contains("-mmm;@"))//日期类型的单元格格式
            {
                newValue = "d-mmm;@";
                type = "date";
            }
            // end
            else if (value.Equals("0.0"))//H2
            {
                newValue = "0.0";
                type = "custom";
            }
            // 20130514 update by xuzhenwei BUG_2920 第三轮回归 oo-uof 功能测试 单元格数字格式不正确 start
            /*else if (value.Contains("0_ "))//添加
            {
                newValue = "0_ ";
                type = "custom";
            }*/
            else if (value.Equals("_ * #,##0_ ;_ * \\-#,##0_ ;_ * \"-\"??_ ;_ @_ "))
            {
                newValue = "_ * #,##0_ ;_ * -#,##0_ ;_ * \"-\"_ ;_ @_ ";
                type = "custom";
            }
            // end
            // 20130226 add by xuzhenwei BUG_2669:单元格内容显示格式化 start
            else if (value.Equals("h\"时\"mm\"分\";@"))//L2
            {
                newValue = "h\"时\"mm\"分\";@";
                type = "time";
            }
            else if (value.Equals("h\"时\"mm\"分\"ss\"秒\";@"))//N2
            {
                newValue = "h\"时\"mm\"分\"ss\"秒\";@";
                type = "time";
            }
            else if (value.Equals("[O-409]h:mm:ss\\ AM/PM;@") || value.Equals("[O-F400]h:mm:ss\\ AM/PM"))
            {
                newValue = "h:mm:ss AM/PM;@";
                type = "time";
            }
            else if (value.Equals("[$-409]h:mm\\ AM/PM;@"))
            {
                newValue = "h:mm AM/PM;@";
                type = "time";
            }
            else if (value.Equals("[DBNum1][O-804]上午/下午h\"时\"mm\"分\";@"))
            {
                newValue = "上午/下午h\"时\"mm\"分\";@";
                type = "time";
            }
            else if (value.Equals("上午/下午h\"时\"mm\"分\";@"))
            {
                newValue = "上午/下午h\"时\"mm\"分\";@";
                type = "time";
            }
            else if (value.Equals("[DBNum1][O-804]h\"时\"mm\"分\";@"))
            {
                newValue = "h\"时\"mm\"分\";@";
                type = "time";
            }
            else if (value.Equals("h:mm:ss;@"))
            {
                newValue = "h:mm:ss;@";
                type = "time";
            }
            else if (value.Equals("h:mm;@"))
            {
                newValue = "h:mm;@";
                type = "time";
            }
            //else if (value.Equals(""))//row12
            //{
            //    newValue = "0.00%";
            //    type = "percentage";
            //}
            else if (value.Contains("0.00%"))
            {
                newValue = "0.00%";
                type = "percentage";
            }
            else if (value.Contains("0.000%"))
            {
                newValue = "0.000%";
                type = "percentage";
            }
            else if (value.Contains("0.0000%"))
            {
                newValue = "0.0000%";
                type = "percentage";
            }
            // 20130226 xuzhenwei end 
            //else if (value.Equals(""))//row13
            //{
            //    newValue = "# ?/?";
            //    type = "fraction";
            //}
            else if (value.Equals("#\\ ??/100"))//row14
            {
                newValue = "# ??/100";
                type = "fraction";
            }
            else if (value.Equals("#\\ ?/8"))//row15
            {
                newValue = "# ?/8";
                type = "fraction";
            }
            else if (value.Equals("#\\ ???/???"))//row16
            {
                newValue = "# ???/???";
                type = "fraction";
            }
            //else if (value.Equals(""))//row17
            //{
            //    newValue = "0.00E+00";
            //    type = "scientific";
            //}
            else if (value.Equals("0.000E+00"))//row17
            {
                newValue = "0.000E+00";
                type = "scientific";
            }
            else if (value.Equals("[DBNum2][O-804]General"))//row19
            {
                newValue = "[DBNum2]General";
                type = "specialization";
            }
            else if (value.Equals("[DBNum1][O-804]General"))//row20
            {
                newValue = "[DBNum1]General";
                type = "specialization";
            }
            else if (value.Equals("000000"))//row21
            {
                newValue = "000000";
                type = "specialization";
            }
            //20130326 add by xuzhenwei bug_2729 集成OO-UOF 字体加粗效果丢失；A8-B35单元格数字格式不正确 start
            else if (value.Equals("ddd"))//用户自定义类型的单元格格式
            {
                newValue = "ddd";
                type = "custom";
            }
            else if (value.Equals("d\" 日\""))//用户自定义类型的单元格格式
            {
                newValue = "d\" 日\"";
                type = "custom";
            }
            // end

           //2014-3-20 add by Qihy bug_3116 互操作OO-UOF-OO 某年；某月不正确以及货币颜色不显示的问题 start
            else if (value.Equals("0\\ \"年\""))//用户自定义类型的单元格格式
            {
                newValue = "0\\ \"年\"";
                type = "custom";
            }
            else if (value.Equals("0\\ \"月\""))//用户自定义类型的单元格格式
            {
                newValue = "0\\ \"月\"";
                type = "custom";
            }
            else if (value.Equals("\"￥\"#,##0;[Red]\"￥\"\\-#,##0"))//用户自定义类型的单元格格式
            {
                newValue = "￥#,##0;[Red]￥-#,##0";
                type = "custom";
            }
            else if (value.Equals("\"￥\"#,##0.00;[Red]\"￥\"\\-#,##0.00"))//用户自定义类型的单元格格式
            {
                newValue = "￥#,##0.00;[Red]￥-#,##0.00";
                type = "custom";
            }
            // end

             //2014-3-24 add by Qihy  start
            else if (value.Equals("0_);\\(0\\)"))//用户自定义类型的单元格格式
            {
                newValue = "0_);(0)";
                type = "number";
            }
            // end
            //2014-4-21 add by Qihy bug_3209
            else
            {
                //newValue = "general";
                //type = "general";
                newValue = value;
                type = "custom";
            }

            // end

            newValue = newValue.Replace("O", label);
            value = newValue;
            //return newValue;
        }

        private static void addChartToDrawing(String path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
                return;
            foreach (FileInfo drawing in dir.GetFiles())
            {
                if (drawing.Extension.ToLower().Equals(".xml"))
                {
                    FileInfo relsFile = new FileInfo(drawing.Directory.FullName + "/_rels/" + drawing.Name + ".rels");
                    if (relsFile.Exists)
                    {
                        XmlDocument rels = new XmlDocument();
                        rels.Load(relsFile.FullName);
                        XmlNameTable nTable = rels.NameTable;
                        XmlNamespaceManager nm = new XmlNamespaceManager(nTable);
                        nm.AddNamespace("ws", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                        nm.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");
                        XmlNodeList r = rels.SelectNodes("//r:Relationship[@Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart']/@Target", nm);
                        foreach (XmlNode n in r)
                        {
                            XmlDocument chart = new XmlDocument();
                            chart.Load(drawing.Directory.FullName + @"\" + n.Value);
                            //removeChartElement(chart, nm);
                            // removecatAxElement(chart, nm);
                            // removevalAxElement(chart, nm);
                            XmlDocument drawDoc = new XmlDocument();
                            drawDoc.Load(drawing.FullName);
                            XmlNode nn = drawDoc.ImportNode(chart.DocumentElement, true);
                            drawDoc.DocumentElement.AppendChild(nn);
                            drawDoc.Save(drawing.FullName);
                        }
                    }
                }
            }
        }
        private static void addRelsToSheet(String path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
                return;
            foreach (FileInfo sheet in dir.GetFiles())
            {
                if (sheet.Extension.ToLower().Equals(".xml"))
                {
                    FileInfo relsFile = new FileInfo(sheet.Directory.FullName + "/_rels/" + sheet.Name + ".rels");
                    if (relsFile.Exists)
                    {
                        XmlDocument sheetDoc = new XmlDocument();
                        XmlDocument relsDoc = new XmlDocument();
                        sheetDoc.Load(sheet.FullName);
                        relsDoc.Load(relsFile.FullName);
                        XmlNode node = sheetDoc.ImportNode(relsDoc.DocumentElement, true);
                        sheetDoc.DocumentElement.AppendChild(node);
                        sheetDoc.Save(sheet.FullName);
                    }
                }
            }
        }
        private static void addRelsToDrawing(String path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
                return;
            foreach (FileInfo sheet in dir.GetFiles())
            {
                if (sheet.Extension.ToLower().Equals(".xml"))
                {
                    FileInfo relsFile = new FileInfo(sheet.Directory.FullName + "/_rels/" + sheet.Name + ".rels");
                    if (relsFile.Exists)
                    {
                        XmlDocument sheetDoc = new XmlDocument();
                        XmlDocument relsDoc = new XmlDocument();
                        sheetDoc.Load(sheet.FullName);
                        relsDoc.Load(relsFile.FullName);
                        XmlNode node = sheetDoc.ImportNode(relsDoc.DocumentElement, true);
                        sheetDoc.DocumentElement.AppendChild(node);
                        sheetDoc.Save(sheet.FullName);
                    }
                }
            }
        }
        private static void addChartFileName(String path)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (!dir.Exists)
                return;

            foreach (FileInfo chart in dir.GetFiles())
            {
                if (chart.Extension.ToLower().Equals(".xml"))
                {
                    XmlDocument doc = new XmlDocument();
                    doc.Load(chart.FullName);
                    XmlAttribute attr = doc.CreateAttribute("filename");
                    attr.Value = chart.Name;
                    doc.DocumentElement.Attributes.Append(attr);
                    FileInfo relsFile = new FileInfo(chart.DirectoryName + @"\_rels\" + chart.Name + ".rels");
                    if (relsFile.Exists)
                    {
                        XmlDocument relsDoc = new XmlDocument();
                        relsDoc.Load(relsFile.FullName);
                        XmlNode node = doc.ImportNode(relsDoc.DocumentElement, true);
                        doc.DocumentElement.AppendChild(node);
                    }
                    doc.Save(chart.FullName);

                }
            }
        }

        private void MoveChartRelFiles(string inputDir, string outputDir)
        {
            // if the direcotry exist
            if (Directory.Exists(inputDir) && Directory.Exists(outputDir))
            {
                string[] chartRelFiles = Directory.GetFiles(inputDir, "*");
                foreach (string chartRelFile in chartRelFiles)
                {
                    // 匹配文件名
                    string fileNameWE = chartRelFile.Substring(chartRelFile.LastIndexOf('\\') + 1, chartRelFile.LastIndexOf('.') - chartRelFile.LastIndexOf('\\') - 1);
                    string fileName = outputDir + fileNameWE;

                    // charts 中存在该文件
                    if (File.Exists(fileName))
                    {
                        XmlDocument chartFileDoc = new XmlDocument();
                        XmlDocument chartFileRelDoc = new XmlDocument();
                        XmlTextWriter tw = null;
                        try
                        {
                            chartFileDoc.Load(fileName);
                            chartFileRelDoc.Load(chartRelFile);
                            string tmpFile = outputDir + "tmp.xml";

                            XmlNameTable nTable = chartFileDoc.NameTable;
                            XmlNamespaceManager nm = new XmlNamespaceManager(nTable);
                            nm.AddNamespace("ws", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                            nm.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

                            XmlNameTable nTableRel = chartFileRelDoc.NameTable;
                            XmlNamespaceManager nmRel = new XmlNamespaceManager(nTableRel);
                            nmRel.AddNamespace("pr", "http://schemas.openxmlformats.org/package/2006/relationships");

                            XmlNode pictNode = chartFileDoc.CreateNode(XmlNodeType.Element, "ws", "picture", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                            // 去掉声明;
                            XmlNode relsNode = chartFileRelDoc.SelectSingleNode("pr:Relationships", nmRel);
                            pictNode.InnerXml = relsNode.OuterXml;
                            chartFileDoc.SelectSingleNode("c:chartSpace", nm).AppendChild(pictNode);
                            tw = new XmlTextWriter(tmpFile, Encoding.UTF8);
                            chartFileDoc.Save(tw);
                            tw.Close();
                            File.Delete(fileName);
                            File.Move(tmpFile, fileName);
                        }
                        catch (Exception ex)
                        {
                            logger.Error(ex.Message);
                            logger.Error(ex.StackTrace);
                        }
                        finally
                        {
                            if (tw != null)
                            {
                                tw.Close();
                            }
                        }
                    }
                }
            }
        }

        private void DataValidationPretreatment(string input, string output)
        {
            XmlTextWriter tw = new XmlTextWriter(output, Encoding.UTF8);
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(input);
            XmlNamespaceManager nm = new XmlNamespaceManager(xmlDoc.NameTable);
            nm.AddNamespace("ws", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            XmlNodeList spreadSheets = xmlDoc.SelectNodes("ws:spreadsheets/ws:spreadsheet", nm);
            foreach (XmlNode spreadSheet in spreadSheets)
            {
                XmlNode dataValidations = spreadSheet.SelectSingleNode("ws:worksheet/ws:dataValidations", nm);
                if (dataValidations != null)
                {
                    XmlNodeList originalDataValidation = dataValidations.SelectNodes("ws:dataValidation", nm);
                    foreach (XmlNode node in originalDataValidation)
                    {
                        if (node.Attributes.GetNamedItem("sqref") != null)
                        {
                            string s1 = node.Attributes.GetNamedItem("sqref").Value.ToString();
                            string[] s2 = s1.Split(new char[1] { ' ' });
                            if (s2.Length > 1)
                            {
                                dataValidations.RemoveChild(node);
                                foreach (string s in s2)
                                {
                                    XmlNode xn = xmlDoc.CreateNode(XmlNodeType.Element, "ws:dataValidation", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                                    xn = node.CloneNode(true);
                                    xn.Attributes.GetNamedItem("sqref").Value = s;
                                    dataValidations.AppendChild(xn);
                                }
                            }
                        }
                    }
                    XmlNode att = xmlDoc.CreateNode(XmlNodeType.Attribute, "count", "");
                    dataValidations.Attributes.SetNamedItem(att);
                    dataValidations.Attributes.GetNamedItem("count").Value = dataValidations.ChildNodes.Count.ToString();
                }
            }

            xmlDoc.Save(tw);
            tw.Close();
        }

        private static void removeChartElement(XmlDocument chart, XmlNamespaceManager nmgr)
        {
            nmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            XmlNodeList chartElements = chart.SelectNodes("//c:plotArea/*", nmgr);
            bool isFirst = true;
            foreach (XmlNode n in chartElements)
            {
                if (n.Name.EndsWith("Chart"))
                {
                    if (isFirst)
                        isFirst = false;
                    else
                    {
                        n.ParentNode.RemoveChild(n);
                    }
                }
            }
        }


        private static void removecatAxElement(XmlDocument chart, XmlNamespaceManager nmgr)
        {
            nmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            XmlNodeList catAxElements = chart.SelectNodes("//c:plotArea/c:catAx", nmgr);
            bool isFirst = true;
            foreach (XmlNode n in catAxElements)
            {


                if (isFirst)
                    isFirst = false;
                else
                {
                    n.ParentNode.RemoveChild(n);
                }

            }
        }

        private static void removevalAxElement(XmlDocument chart, XmlNamespaceManager nmgr)
        {
            nmgr.AddNamespace("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            XmlNodeList valAxElements = chart.SelectNodes("//c:plotArea/c:valAx", nmgr);
            bool isFirst = true;
            foreach (XmlNode n in valAxElements)
            {


                if (isFirst)
                    isFirst = false;
                else
                {
                    n.ParentNode.RemoveChild(n);
                }

            }
        }

        private static void addformular(String path)
        {
            String[] sheetFileNames = Directory.GetFiles(path, "*.xml");
            foreach (String sheetFileName in sheetFileNames)
            {
                XmlDocument sheetDoc = new XmlDocument();
                sheetDoc.Load(sheetFileName);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(sheetDoc.NameTable);
                nsmgr.AddNamespace("ws", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                XmlNodeList colset = sheetDoc.SelectNodes("//ws:row/ws:c", nsmgr);
                foreach (XmlNode cols in colset)
                {
                    if (cols != null)
                    {
                        XmlNodeList collist = cols.SelectNodes("ws:f[contains(@ref,':')]", nsmgr);
                        foreach (XmlNode col in collist)
                        {
                            String attrValue = col.Attributes["ref"].Value;
                            String before = attrValue.Substring(0, attrValue.IndexOf(':'));
                            String after = attrValue.Substring(attrValue.IndexOf(':') + 1, attrValue.Length - attrValue.IndexOf(':') - 1);
                            Char startLetter = before.Substring(0, 1).ToCharArray()[0];
                            Int32 startNumber;
                            if (before.Length == 2)
                            {
                                startNumber = Int32.Parse(before.Substring(1));
                            }
                            else
                            {
                                startNumber = Int32.Parse(before.Substring(2));
                            }
                            Char endLetter = after.Substring(0, 1).ToCharArray()[0];
                            Int32 endNumber;
                            if (after.Length == 2)
                            {
                                endNumber = Int32.Parse(after.Substring(1));
                            }
                            else
                            {
                                endNumber = Int32.Parse(after.Substring(2));
                            }
                            cols.RemoveChild(col);
                            for (char c = startLetter; c <= endLetter; ++c)
                            {
                                for (int i = startNumber; i <= endNumber; ++i)
                                {
                                    String newValue = "";
                                    newValue += c;
                                    newValue += i;
                                    XmlNode newcol = col.Clone();
                                    newcol.Attributes["ref"].Value = newValue;
                                    cols.AppendChild(newcol);
                                }
                            }
                        }
                        //sheetDoc.Save(sheetFileName);
                    }
                }

                sheetDoc.Save(sheetFileName);
            }
        }

        private static void addHyperlinks(String path)
        {
            String[] sheetFileNames = Directory.GetFiles(path, "*.xml");
            foreach (String sheetFileName in sheetFileNames)
            {
                XmlDocument sheetDoc = new XmlDocument();
                sheetDoc.Load(sheetFileName);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(sheetDoc.NameTable);
                nsmgr.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                XmlNode hyperlinks = sheetDoc.SelectSingleNode("//s:hyperlinks", nsmgr);
                if (hyperlinks != null)
                {
                    XmlNodeList hyperlinklist = hyperlinks.SelectNodes("s:hyperlink[contains(@ref,':')]", nsmgr);
                    foreach (XmlNode hyperlink in hyperlinklist)
                    {
                        String attrValue = hyperlink.Attributes["ref"].Value;
                        String before = attrValue.Substring(0, attrValue.IndexOf(':'));
                        String after = attrValue.Substring(attrValue.IndexOf(':') + 1, attrValue.Length - attrValue.IndexOf(':') - 1);
                        Char startLetter = before.Substring(0, 1).ToCharArray()[0];
                        Int32 startNumber = Int32.Parse(before.Substring(1));
                        Char endLetter = after.Substring(0, 1).ToCharArray()[0];
                        Int32 endNumber = Int32.Parse(after.Substring(1));
                        hyperlinks.RemoveChild(hyperlink);
                        for (char c = startLetter; c <= endLetter; ++c)
                        {
                            for (int i = startNumber; i <= endNumber; ++i)
                            {
                                String newValue = "";
                                newValue += c;
                                newValue += i;
                                XmlNode newHyperlink = hyperlink.Clone();
                                newHyperlink.Attributes["ref"].Value = newValue;
                                hyperlinks.AppendChild(newHyperlink);
                            }
                        }
                    }
                    sheetDoc.Save(sheetFileName);
                }
            }
        }
        private static void AddxiaoGroup(string path)
        {
            String[] sheetFileNames = Directory.GetFiles(path, "*.xml");
            foreach (String sheetFileName in sheetFileNames)
            {
                XmlDocument sheetDoc = new XmlDocument();
                sheetDoc.Load(sheetFileName);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(sheetDoc.NameTable);
                nsmgr.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                nsmgr.AddNamespace("ws", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                XmlNodeList rows = sheetDoc.SelectNodes(".//ws:row", nsmgr);
                XmlNodeList cols = sheetDoc.SelectNodes(".//ws:col", nsmgr);
                //处理列
                Hashtable hashCol = new Hashtable();
                fillHashTable(cols, hashCol);
                //处理行
                Hashtable hashRow = new Hashtable();
                fillHashTable(rows, hashRow);
                List<XmlNode> groupList = new List<XmlNode>();
                //处理列
                if (hashCol.Count > 0)
                    for (int alevel = 1; alevel <= hashCol.Count; alevel++)
                    {
                        List<XmlNode> colList = (List<XmlNode>)hashCol[alevel];
                        if (colList.Count > 0)
                        {
                            List<int> colArrayList = new List<int>();

                            foreach (XmlNode node in colList)
                            {
                                XmlAttribute attMin = node.Attributes["min"];
                                int min = 0, max = 0;
                                if (attMin != null)
                                {
                                    int tmpR = int.Parse(attMin.InnerText.Trim());
                                    min = tmpR;
                                }
                                XmlAttribute attMax = node.Attributes["max"];
                                if (attMax != null)
                                {
                                    int tmp = int.Parse(attMax.InnerText.Trim());
                                    max = tmp;
                                }
                                if (min != 0 && max != 0)
                                {
                                    for (int xiao = min; xiao <= max; xiao++)
                                    {
                                        colArrayList.Add(xiao);
                                    }
                                }
                            }
                            int[] colArray = colArrayList.ToArray();
                            int start = colArray[0];
                            int end = 0;
                            if (colArray.Length == 1)
                            {
                                groupList.Add(getOneGroup(sheetDoc, colArray[0], colArray[0], GroupType.Col));
                            }
                            else
                            {
                                for (int i = 1; i < colArray.Length; i++)
                                {
                                    if (colArray[i] - colArray[i - 1] != 1)
                                    { //刚好多一行
                                        end = colArray[i - 1];
                                        groupList.Add(getOneGroup(sheetDoc, start, end, GroupType.Col));
                                        start = colArray[i];
                                    }
                                    if (i == colArray.Length - 1)
                                    {
                                        end = colArray[i];
                                        groupList.Add(getOneGroup(sheetDoc, start, end, GroupType.Col));
                                    }
                                }
                            }
                        }
                    }
                //处理行
                if (hashRow.Count > 0)
                    for (int level = 1; level <= hashRow.Count; level++)
                    {
                        List<XmlNode> rowList = (List<XmlNode>)hashRow[level];
                        if (rowList.Count > 0)
                        {
                            int[] rowArray = new int[rowList.Count];
                            int index = 0;
                            foreach (XmlNode node in rowList)
                            {
                                XmlAttribute attR = node.Attributes["r"];
                                if (attR != null)
                                {
                                    int tmpR = int.Parse(attR.InnerText.Trim());
                                    rowArray[index++] = tmpR;
                                }
                            }
                            int start = rowArray[0];
                            int end = 0;
                            if (rowArray.Length == 1)
                            {
                                groupList.Add(getOneGroup(sheetDoc, rowArray[0], rowArray[0], GroupType.Row));
                            }
                            else
                            {
                                for (int i = 1; i < rowArray.Length; i++)
                                {
                                    if (rowArray[i] - rowArray[i - 1] != 1)
                                    { //刚好多一行
                                        end = rowArray[i - 1];
                                        groupList.Add(getOneGroup(sheetDoc, start, end, GroupType.Row));
                                        start = rowArray[i];
                                    }
                                    if (i == rowArray.Length - 1)
                                    {
                                        end = rowArray[i];
                                        groupList.Add(getOneGroup(sheetDoc, start, end, GroupType.Row));
                                    }
                                }
                            }
                        }
                    }

                XmlNode groupSet = sheetDoc.CreateElement("GroupSet");
                foreach (XmlNode group in groupList)
                {
                    groupSet.AppendChild(group);
                }
                XmlNode worksheet = sheetDoc.SelectSingleNode(".//ws:worksheet", nsmgr);
                if (worksheet != null)
                    worksheet.AppendChild(groupSet);
                sheetDoc.Save(sheetFileName);
            }
        }
        enum GroupType { Col, Row }
        private static XmlNode getOneGroup(XmlDocument sheetDoc, int start, int end, GroupType type)
        {
            XmlNode nodeGroup = sheetDoc.CreateElement("Group");
            XmlAttribute attStart = sheetDoc.CreateAttribute("Start");
            attStart.InnerText = start.ToString();
            nodeGroup.Attributes.Append(attStart);
            XmlAttribute attEnd = sheetDoc.CreateAttribute("End");
            attEnd.InnerText = end.ToString();
            nodeGroup.Attributes.Append(attEnd);
            XmlAttribute attPosition = sheetDoc.CreateAttribute("Position");
            if (type == GroupType.Col)
                attPosition.InnerText = "Col";
            else if (type == GroupType.Row)
                attPosition.InnerText = "Row";
            nodeGroup.Attributes.Append(attPosition);
            return nodeGroup;
        }
        /// <summary>
        /// 填充一个hash表
        /// </summary>
        /// <param name="cols"></param>
        /// <param name="hashCol"></param>
        private static void fillHashTable(XmlNodeList cols, Hashtable hashCol)
        {
            foreach (XmlNode col in cols)
            {
                XmlAttribute attOutlineLevel = col.Attributes["outlineLevel"];
                if (attOutlineLevel != null)
                {
                    int level = int.Parse(attOutlineLevel.InnerText.Trim());
                    for (int i = level; i > 0; i--)
                    {
                        if (hashCol.ContainsKey(i))
                        {
                            List<XmlNode> colList = (List<XmlNode>)hashCol[i];
                            colList.Add(col);

                        }
                        else
                        {
                            List<XmlNode> colList = new List<XmlNode>();
                            colList.Add(col);
                            hashCol.Add(i, colList);
                        }
                    }
                }
            }
        }
        private static void AddGroup(string path)
        {
            //string temp = "";
            String[] sheetFileNames = Directory.GetFiles(path, "*.xml");
            foreach (String sheetFileName in sheetFileNames)
            {
                XmlDocument sheetDoc = new XmlDocument();
                sheetDoc.Load(sheetFileName);
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(sheetDoc.NameTable);
                nsmgr.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                nsmgr.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                nsmgr.AddNamespace("ws", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                XmlNodeList rows = sheetDoc.SelectNodes(".//ws:row", nsmgr);
                XmlNodeList cols = sheetDoc.SelectNodes(".//ws:col", nsmgr);


                //DataSet ds_row = GetGroupList(rows);
                //DataSet ds_col = GetGroupList(cols);
                DataSet ds_row = new DataSet("GroupSetDs");
                DataTable dt_row = new DataTable("GroupSetDt");
                DataColumn c1_row = new DataColumn("Start");
                DataColumn c2_row = new DataColumn("End");
                //DataColumn c3 = new DataColumn("Seq");
                dt_row.Columns.Add(c1_row);
                dt_row.Columns.Add(c2_row);
                //dt.Columns.Add(c3);
                ds_row.Tables.Add(dt_row);
                GetGroupList(rows, ref ds_row);

                DataSet ds_col = new DataSet("GroupSetDs");
                DataTable dt_col = new DataTable("GroupSetDt");
                DataColumn c1_col = new DataColumn("Start");
                DataColumn c2_col = new DataColumn("End");
                //DataColumn c3 = new DataColumn("Seq");
                dt_col.Columns.Add(c1_col);
                dt_col.Columns.Add(c2_col);
                //dt.Columns.Add(c3);
                ds_col.Tables.Add(dt_col);
                GetGroupList(cols, ref ds_col);

                XmlElement groupSet = sheetDoc.CreateElement("GroupSet");
                for (int i = 0; i < ds_row.Tables["GroupSetDt"].Rows.Count; i++)
                {
                    XmlElement subGroup = sheetDoc.CreateElement("Group");
                    XmlAttribute Start = sheetDoc.CreateAttribute("Start");
                    Start.Value = ds_row.Tables["GroupSetDt"].Rows[i]["Start"].ToString();

                    XmlAttribute End = sheetDoc.CreateAttribute("End");
                    End.Value = ds_row.Tables["GroupSetDt"].Rows[i]["End"].ToString();

                    XmlAttribute Pos = sheetDoc.CreateAttribute("Position");
                    Pos.Value = "Row";
                    subGroup.Attributes.Append(Start);
                    subGroup.Attributes.Append(End);
                    subGroup.Attributes.Append(Pos);
                    groupSet.AppendChild(subGroup);
                }


                for (int i = 0; i < ds_col.Tables["GroupSetDt"].Rows.Count; i++)
                {
                    XmlElement subGroup = sheetDoc.CreateElement("Group");
                    XmlAttribute Start = sheetDoc.CreateAttribute("Start");
                    Start.Value = ds_col.Tables["GroupSetDt"].Rows[i]["Start"].ToString();

                    XmlAttribute End = sheetDoc.CreateAttribute("End");
                    End.Value = ds_col.Tables["GroupSetDt"].Rows[i]["End"].ToString();

                    XmlAttribute Pos = sheetDoc.CreateAttribute("Position");
                    Pos.Value = "Col";
                    subGroup.Attributes.Append(Start);
                    subGroup.Attributes.Append(End);
                    subGroup.Attributes.Append(Pos);
                    groupSet.AppendChild(subGroup);
                }
                /*   * */
                sheetDoc.DocumentElement.AppendChild(groupSet);

                sheetDoc.Save(sheetFileName);

            }

            //return temp;
        }

        public static void GetGroupList(XmlNodeList list, ref DataSet ds)
        {


            // ArrayList startList = new ArrayList();
            //ArrayList endList = new ArrayList();
            ArrayList existList = new ArrayList();
            ArrayList valueList = new ArrayList();

            Boolean key = false;
            string type = "";

            foreach (XmlNode row in list)
            {
                if (row.Name == "row")
                    type = "row";
                else
                    type = "col";

                key = false;
                for (int i = 0; i < row.Attributes.Count; i++)
                {
                    if (row.Attributes[i].Name == "outlineLevel")
                        key = true;
                }
                if (key)
                {
                    valueList.Add(row.Attributes["outlineLevel"].Value.ToString());
                }
                else
                {
                    valueList.Add(0);
                }
                existList.Add(0);
            }
            for (int i = 0; i < valueList.Count; i++)
            {
                if (int.Parse(valueList[i].ToString()) > 0)
                {
                    if (int.Parse(existList[i].ToString()) == 0)
                    {
                        for (int j = i + 1; j < valueList.Count; j++)
                        {
                            if (int.Parse(valueList[j].ToString()) < int.Parse(valueList[i].ToString()))
                            {
                                //if (int.Parse(existList[j].ToString()) == 0)
                                //{
                                //temp += Path.GetFileName(sheetFileName) + ":Start At " + i + " End At " + (int)(j - 1) + "\n";
                                //string rowNo = ;
                                //startList.Add(list[i].Attributes["r"].Value.ToString());
                                //endList.Add(list[j - 1].Attributes["r"].Value.ToString());

                                DataRow dr = ds.Tables[0].NewRow();
                                if (type == "row")
                                {
                                    dr["Start"] = list[i].Attributes["r"].Value.ToString();
                                    dr["End"] = list[j - 1].Attributes["r"].Value.ToString();
                                }
                                else
                                {
                                    dr["Start"] = list[i].Attributes["min"].Value.ToString();
                                    dr["End"] = list[j - 1].Attributes["min"].Value.ToString();
                                }
                                ds.Tables[0].Rows.Add(dr);
                                existList[i] = 1;
                                break;
                                //}
                            }
                            if (int.Parse(valueList[i].ToString()) == int.Parse(valueList[j].ToString()))
                            {
                                if (int.Parse(existList[j].ToString()) == 0)
                                {
                                    existList[j] = 1;
                                }
                            }
                        }


                        if (i > 0)
                        {
                            if (int.Parse(valueList[i].ToString()) - int.Parse(valueList[i - 1].ToString()) > 1)
                            {
                                for (int k = i + 1; k < valueList.Count; k++)
                                {
                                    if (int.Parse(valueList[k].ToString()) < int.Parse(valueList[i].ToString()) - 1)
                                    {
                                        //if (int.Parse(existList[j].ToString()) == 0)
                                        //{
                                        //temp += Path.GetFileName(sheetFileName) + ":Start At " + i + " End At " + (int)(k - 1) + "\n";
                                        //startList.Add(list[i].Attributes["r"].Value.ToString());
                                        //endList.Add(list[k - 1].Attributes["r"].Value.ToString());
                                        DataRow dr = ds.Tables["GroupSetDt"].NewRow();
                                        if (type == "row")
                                        {
                                            dr["Start"] = list[i].Attributes["r"].Value.ToString();
                                            dr["End"] = list[k - 1].Attributes["r"].Value.ToString();
                                        }
                                        else
                                        {
                                            dr["Start"] = list[i].Attributes["min"].Value.ToString();
                                            dr["End"] = list[k - 1].Attributes["min"].Value.ToString();
                                        }
                                        ds.Tables["GroupSetDt"].Rows.Add(dr);
                                        existList[i] = 1;
                                        break;
                                        //}
                                    }
                                    if (int.Parse(valueList[i].ToString()) - 1 == int.Parse(valueList[k].ToString()))
                                    {
                                        if (int.Parse(existList[k].ToString()) == 0)
                                        {
                                            existList[k] = 1;
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
            // return ds;
        }


        #region RC引用处理

        private static void NormalToRC(string inputFile, string outputFile)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFile);
            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("ws", TranslatorConstants.XMLNS_WS);
            XmlNode refMode = xdoc.SelectSingleNode("ws:spreadsheets/ws:workbook/ws:calcPr", nm).Attributes["refMode"];

            // @refMode="R1C1"
            if (refMode != null)
            {
                XmlNodeList worksheets = xdoc.SelectNodes("ws:spreadsheets/ws:spreadsheet/ws:worksheet", nm);
                foreach (XmlNode worksheet in worksheets)
                {
                    XmlNodeList rows = worksheet.SelectNodes("ws:sheetData/ws:row", nm);
                    foreach (XmlNode row in rows)
                    {
                        XmlNode cf = row.SelectSingleNode("ws:c[ws:f]", nm);
                        if (cf != null)
                        {
                            string rV = cf.Attributes["r"].Value;
                            string formula = cf.SelectSingleNode("ws:f", nm).InnerText;
                            cf.SelectSingleNode("ws:f", nm).InnerText = CalNormalToRC(formula, rV);
                        }
                    }

                }

            }
            XmlTextWriter tw = new XmlTextWriter(outputFile, Encoding.UTF8);
            try
            {
                xdoc.Save(tw);
            }
            catch
            {
                throw;
            }
            finally
            {
                tw.Close();
            }

        }


        /// <summary>
        ///  计算每一个函数的RC转换
        /// </summary>
        /// <param name="formula">函数公式</param>
        /// <param name="oriR">结果所在位置，如O6</param>
        /// <returns>返回更改后的公式</returns>
        private static string CalNormalToRC(string formula, string oriR)
        {

            Regex range = new Regex("([A-Z]+[0-9]+[:][A-Z]+[0-9]+)");
            //string formula = "DAVERAGE(A5:F22,\"语文\",B26:B27)";

            MatchCollection mc = range.Matches(formula);
            if (mc.Count > 0)
            {
                foreach (Match m in mc)
                {
                    string[] rc1 = m.Value.Split(new char[] { ':' });
                    string r1 = CalRC(rc1[0], oriR);
                    string r2 = CalRC(rc1[1], oriR);
                    formula = formula.Replace(m.Value, r1 + ":" + r2);

                }
            }

            return formula;
        }

        /// <summary>
        /// 把A1:B2中的A1换成相对于currentCell的RC表示
        /// </summary>
        /// <param name="range">当前的范围</param>
        /// <param name="currentCell">当前单元格地址</param>
        /// <returns>RC表达</returns>
        public static string CalRC(string range, string currentCell)
        {
            string result = "R";
            string cellCol = GetTheLetter(currentCell);
            string cellRow = GetTheNum(currentCell);
            string rangeCol = GetTheLetter(range);
            string rangeRow = GetTheNum(range);


            int resultRow = Convert.ToInt32(rangeRow) - Convert.ToInt32(cellRow);
            if (resultRow != 0)
            {
                result += "[" + resultRow.ToString() + "]";
            }

            double col = ConvertLetterToNum(rangeCol) - ConvertLetterToNum(cellCol);
            int resultCol = Convert.ToInt32(col);
            if (resultCol != 0)
            {
                result += "C[" + resultCol.ToString() + "]";
            }
            else
            {
                result += "C";
            }


            return result;
        }


        /// <summary>
        /// get the letter
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string GetTheLetter(string range)
        {
            return Regex.Replace(range, "[0-9]", "", RegexOptions.IgnoreCase);
        }

        /// <summary>
        /// get the number
        /// </summary>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string GetTheNum(string range)
        {
            return Regex.Replace(range, "[a-z]", "", RegexOptions.IgnoreCase); ;
        }
        /// <summary>
        ///  计算字母转换成的数字
        /// </summary>
        /// <param name="letter"></param>
        /// <returns></returns>
        public static double ConvertLetterToNum(string letter)
        {
            double num = 0;
            int maxLen = letter.Length - 1;
            int curr = 0;
            for (int i = maxLen; i >= 0; i--)
            {
                char lt = letter[curr++];
                int alph = ConvertAlphToNum(lt);

                // 26：A-Z的值
                num += Math.Pow(26, i) * alph;

            }

            return num;
        }

        /// <summary>
        /// 计算字母的数字大小，A对应的是1，B对应的是2
        /// </summary>
        /// <param name="alph">字母</param>
        /// <returns>数字</returns>
        private static int ConvertAlphToNum(char alph)
        {
            //64：表示数字大小
            return Convert.ToInt32(alph) - 64;
        }

        #endregion

    }
}


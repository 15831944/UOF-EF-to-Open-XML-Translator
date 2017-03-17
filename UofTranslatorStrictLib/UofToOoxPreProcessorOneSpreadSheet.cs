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
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Collections;
namespace Act.UofTranslator.UofTranslatorStrictLib
{
    /// <summary>
    /// The first step pretreatment of spreadsheet from UOF->OOX
    /// </summary>
    /// <author>linyaohu</author>
    class UofToOoxPreProcessorOneSpreadSheet : AbstractProcessor
    {
        private static ILog logger = LogManager.GetLogger(typeof(UofToOoxPreProcessorOneSpreadSheet).FullName);

        public override bool transform()
        {
            //if (!inputFile.Equals(string.Empty))
            //{
            //    preMethod(inputFile);
            //    return true;
            //}
            //return false;
            XmlTextWriter fs = null;
            XmlUrlResolver resourceResolver = null;
            XmlReader xr = null;

            string extractPath = Path.GetDirectoryName(outputFile) + Path.AltDirectorySeparatorChar;
            string prefix = this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + TranslatorConstants.UOFToOOX_EXCEL_LOCATION;
            string uof2ooxpre = extractPath + "uof2ooxpre.xml";
            string picture_xml = "data";

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
                        int Length = 204800;
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

                resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(), prefix);
                //xr = XmlReader.Create(((ResourceResolver)resourceResolver).GetInnerStream("uof2oopre.xslt"));
                xr = XmlReader.Create(extractPath + "pretreatment.xsl");
                //处理分组集
                this.dealGroupLevel(extractPath + @"content.xml");
                XPathDocument doc = new XPathDocument(extractPath + @"content.xml");
                XslCompiledTransform transFrom = new XslCompiledTransform();
                XsltSettings setting = new XsltSettings(true, false);
                XmlUrlResolver xur = new XmlUrlResolver();
                transFrom.Load(xr, setting, xur);
                XPathNavigator nav = ((IXPathNavigable)doc).CreateNavigator();
                fs = new XmlTextWriter(outputFile, Encoding.UTF8);
                // fs = new FileStream(uof2ooxpre, FileMode.Create);
                transFrom.Transform(nav, null, fs);
                fs.Close();

                //preMethod(uof2ooxpre);

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
                // 公式
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

                string rcToNormal = extractPath + "RCPre.xml";

                RCToNormal(outputFile, rcToNormal);
                outputFile = rcToNormal;
                return true;


            }
            catch (Exception ex)
            {
                logger.Error("Fail in Uof2.0 to OOX pretreatment1: " + ex.Message);
                logger.Error(ex.StackTrace);
                //return false;
                throw new Exception("Fail in Uof2.0 to OOX pretreatment1 of spreadsheet");
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }
        /// <summary>
        /// 增加uof到ooxml中分组集相关处理
        /// </summary>
        /// <param name="content">content.xml路径</param>
        private void dealGroupLevel(string content)
        {
            if (File.Exists(content))
            {//文件存在 
                XmlDocument doc = new XmlDocument();
                doc.Load(content);
                //增加命名空间
                XmlNamespaceManager nm = new XmlNamespaceManager(doc.NameTable);
                //nm.AddNamespace("uof", "http://schemas.uof.org/cn/2009/uof");
                nm.AddNamespace("表", TranslatorConstants.XMLNS_UOFSPREADSHEET);
                XmlNodeList nodeListContentSet = doc.SelectNodes("//表:工作表内容_E80E", nm);
                foreach (XmlNode excelContent in nodeListContentSet)
                {
                    XmlNode nodeGroup = excelContent.SelectSingleNode("表:分组集_E7F6", nm);
                    if (nodeGroup != null)
                    {//存在分组集 
                        this.dealColorRow(doc, nodeGroup, RowOrColType.Col, excelContent, nm);
                        this.dealColorRow(doc, nodeGroup, RowOrColType.Row, excelContent, nm);
                        this.tailDeal(nodeGroup, nm);
                    }
                }
                doc.Save(content);
            }
        }
        enum RowOrColType { Row, Col }
        private void dealColorRow(XmlDocument doc, XmlNode nodeGroup, RowOrColType type, XmlNode excelContent, XmlNamespaceManager nm)
        {
            XmlNodeList list = null;


            if (type == RowOrColType.Row)
            {
                list = nodeGroup.SelectNodes("表:行_E842", nm);
            }
            else if (type == RowOrColType.Col)
            {
                list = nodeGroup.SelectNodes("表:列_E841", nm);
            }
            if (list.Count > 0)
            {//存在
                List<int[]> listScope = new List<int[]>();
                int minRowNo = 100000;
                int maxRowNo = 0;
                foreach (XmlNode col in list)
                {
                    int[] scope = new int[2];
                    //20120122 gaoyuwei bug2648 分级不正确 start
                    int start = int.Parse(col.Attributes["起始_E73A"].InnerText);
                    int end = int.Parse(col.Attributes["终止_E73B"].InnerText);
                    //end
                    if (start < minRowNo) minRowNo = start;
                    if (end > maxRowNo) maxRowNo = end;
                    scope[0] = start;
                    scope[1] = end;

                    listScope.Add(scope);
                }
                XmlNode arrayRow = null;

                for (int i = minRowNo; i <= maxRowNo; i++)
                {
                    if (type == RowOrColType.Row)
                        arrayRow = excelContent.SelectSingleNode("表:行_E7F1[@行号_E7F3='" + i + "']", nm);
                    else if (type == RowOrColType.Col)
                        arrayRow = excelContent.SelectSingleNode("表:列_E7EC[@列号_E7ED='" + i + "']", nm);

                    if (arrayRow == null)
                    {
                        XmlNode addNode = null;
                        if (type == RowOrColType.Col)
                        {
                            addNode = doc.CreateElement("表", "列_E7EC", TranslatorConstants.XMLNS_UOFSPREADSHEET);
                            XmlAttribute attNo = doc.CreateAttribute("列号_E7ED");
                            attNo.InnerText = i.ToString();
                            addNode.Attributes.Append(attNo);//添加列号
                            XmlAttribute attWidth = doc.CreateAttribute("列宽_E7EE");
                            attWidth.InnerText = "54.0";
                            addNode.Attributes.Append(attWidth);//添加列宽
                            XmlAttribute attStyle = doc.CreateAttribute("式样引用_E7BD");
                            attStyle.InnerText = "CELLSTYLE_2";
                            addNode.Attributes.Append(attStyle);//添加式样引用
                            XmlAttribute attAuto = doc.CreateAttribute("是否自适应列宽_E7F0");
                            attAuto.InnerText = "true";
                            addNode.Attributes.Append(attAuto);//是否自适应列宽_E7F0

                        }
                        else if (type == RowOrColType.Row)
                        {
                            addNode = doc.CreateElement("表", "行_E7F1", TranslatorConstants.XMLNS_UOFSPREADSHEET);
                            XmlAttribute attNo = doc.CreateAttribute("行号_E7F3");
                            attNo.InnerText = i.ToString();
                            addNode.Attributes.Append(attNo);//添加行号
                            XmlAttribute attHeight = doc.CreateAttribute("行高_E7F4");
                            attHeight.InnerText = "22.5";
                            addNode.Attributes.Append(attHeight);
                        }
                        excelContent.AppendChild(addNode);
                        arrayRow = addNode;
                    }
                    // 20130415 update by linyaohu BUG_2827:互操作oo-uof（编辑）-oo 分组是否隐藏转换不正确 用来处理多余的空列输出 start
                    else
                    {
                        if (arrayRow.Attributes["跨度_E7EF"] != null)
                        {
                            string spaceValue = arrayRow.Attributes["跨度_E7EF"].Value;

                            int space = Convert.ToInt32(spaceValue);

                            i += space;
                        }
                    }
                    // end

                    XmlAttribute att = doc.CreateAttribute("outlineLevel");
                    att.InnerText = this.backLevel(i, listScope).ToString();
                    arrayRow.Attributes.Append(att);


                    //arrayRow = doc.CreateElement("row");

                    //XmlAttribute attR = doc.CreateAttribute("r");
                    //attR.InnerText = (i + 1).ToString();
                    //arrayRow.Attributes.Append(attR);

                    //XmlAttribute attHeight = doc.CreateAttribute("ht");
                    //XmlNode nodeCol = excelContent.SelectSingleNode("表:行_E7F1[@行号_E7F3='" + i + "']", nm);
                    //double height = 14.25;
                    //if (nodeCol != null)
                    //{
                    //    XmlAttribute att = nodeCol.Attributes["行高_E7F4"];
                    //    if (att != null)
                    //    {
                    //        string s = att.InnerText;
                    //        height = double.Parse(s);
                    //    }
                    //}
                    //attHeight.InnerText = height.ToString();//行高
                    //arrayRow.Attributes.Append(attHeight);
                    //XmlAttribute attCus = doc.CreateAttribute("customHeight");
                    //attCus.InnerText = "1";
                    //arrayRow.Attributes.Append(attCus);

                    //XmlAttribute attLevel = doc.CreateAttribute("outlineLevel");
                    //attLevel.InnerText = this.backLevel(i, listScope).ToString();
                    //arrayRow.Attributes.Append(attLevel);

                    //XmlAttribute attDescent = doc.CreateAttribute("x14ac", "dyDescent", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");
                    //attDescent.InnerText = "0.15";
                    //arrayRow.Attributes.Append(attDescent);

                    //myRows.AppendChild(arrayRow);

                }
            }
        }
        /// <summary>
        /// 返回分组等级
        /// </summary>
        /// <param name="colNo"></param>
        /// <param name="listScope"></param>
        /// <returns></returns>
        private int backLevel(int colNo, List<int[]> listScope)
        {
            int ret = 0;
            foreach (int[] array in listScope)
            {
                if (colNo <= array[1] && colNo >= array[0])
                    ret++;
            }
            return ret;
        }
        /// <summary>
        /// 重新整理含有分组集的工作表列和行
        /// </summary>
        /// <param name="nodeGroup"></param>
        /// <param name="nm"></param>
        private void tailDeal(XmlNode nodeGroup, XmlNamespaceManager nm)
        {
            XmlNode excelContent = nodeGroup.ParentNode;
            XmlNodeList listCol = excelContent.SelectNodes("表:列_E7EC", nm);
            foreach (XmlNode node in this.Resort(listCol, "列号_E7ED"))
            {
                excelContent.RemoveChild(node);
                excelContent.AppendChild(node);
            }
            XmlNodeList listRow = excelContent.SelectNodes("表:行_E7F1", nm);
            foreach (XmlNode node in this.Resort(listRow, "行号_E7F3"))
            {
                excelContent.RemoveChild(node);
                excelContent.AppendChild(node);
            }
            excelContent.RemoveChild(nodeGroup);
            excelContent.AppendChild(nodeGroup);
        }
        private XmlNode[] Resort(XmlNodeList list, string type)
        {
            XmlNode[] retArray = new XmlNode[list.Count];
            int index = 0;
            foreach (XmlNode node in list)
            {
                retArray[index++] = node;
            }
            for (int i = 0; i < list.Count; i++)
            {
                for (int j = i + 1; j < list.Count; j++)
                {
                    XmlNode tmp = null;
                    XmlAttribute attI = retArray[i].Attributes[type];
                    XmlAttribute attJ = retArray[j].Attributes[type];
                    if (int.Parse(attI.InnerText.Trim()) > int.Parse(attJ.InnerText.Trim()))
                    {
                        tmp = retArray[i];
                        retArray[i] = retArray[j];
                        retArray[j] = tmp;
                    }
                }
            }
            return retArray;

        }
        private void preMethod(string FileLoc)
        {
            string transPath = Path.GetDirectoryName(outputFile);
            //string tmpFileChangeAttibute = transPath + "\\" + "tmp1.xml";
            string tmpFileConvertRowNumber = transPath + "\\tmp_rowNumber.xml";
            string tmpFile = transPath + "\\" + "tmpFile.xml";
            string tmpFileMoveGrpShapes = transPath + "\\" + "tmp_moveGShapes.xml";
            string tmpFileNumFmts = transPath + "\\" + "tmp_numFmt.xml";
            string temp = transPath + "\\" + "tmp.xml";
            string tempK = transPath + "\\" + "tmpK.xml";
            string tempG = transPath + "\\" + "tmpG.xml";
            //string tmpFileAddConditionalOrder = transPath + "\\tmp3.xml";

            //ChangeAnchorAttributeName(FileLoc, tmpFileChangeAttibute);

            // convert row number
            ConvertRowIntTo26(FileLoc, tmpFileConvertRowNumber);

            FileStream fs = new FileStream(tmpFile, FileMode.Create);
            FileStream fs2 = null;
            try
            {
                NumberFormatsTransfer(tmpFileConvertRowNumber, tmpFileNumFmts);
                // move all shapes of group-shapes
                //PreGrpShaps(tmpFileChangeAttibute, tmpFileMoveGrpShapes);
                PreGrpShaps(tmpFileNumFmts, tmpFileMoveGrpShapes);

                GetKeyPoints(tmpFileMoveGrpShapes, tempK);
                AddGroup(tempK, tempG);
                AddPosition(tempG, temp);
                XmlUrlResolver resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(), this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + ".Excel.uof2oox");
                XmlReader xr;
                xr = XmlReader.Create(((ResourceResolver)resourceResolver).GetInnerStream("pretreatment.xsl"));

                // step 1
                XPathDocument doc = new XPathDocument(temp);
                // XPathDocument doc = new XPathDocument(FileLoc);
                XslCompiledTransform transFrom = new XslCompiledTransform();
                XsltSettings setting = new XsltSettings(true, false);
                XmlUrlResolver xur = new XmlUrlResolver();
                //transFrom.Load(xr, setting, xur);
                Assembly ass = Assembly.Load("excel_uof2oox");
                Type t = ass.GetType("pretreatment");
                transFrom.Load(t);
                XPathNavigator nav = ((IXPathNavigable)doc).CreateNavigator();
                transFrom.Transform(nav, null, fs);
                fs.Close();

                XmlDocument xdoc = new XmlDocument();
                xdoc.Load(tmpFile);

                fs2 = new FileStream(outputFile, FileMode.Create);
                xdoc.Save(fs2);

            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                logger.Error(ex.StackTrace);
                throw new Exception("Fail in UofToOoxPreProcessorOneSpreasSheet");
            }
            finally
            {
                if (fs != null)
                    fs.Close();
                if (fs2 != null)
                {
                    fs2.Close();
                }
            }
        }

        private void AddPosition(string inputFileName, string destFileName)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFileName);

            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);
            nm.AddNamespace("表", TranslatorConstants.XMLNS_UOFSPREADSHEET);
            XmlNodeList sheetList = xdoc.SelectNodes("//表:工作表", nm);

            foreach (XmlNode sheet in sheetList)
            {
                XmlNodeList chartNodesList = sheet.SelectNodes("表:工作表内容/表:图表", nm);
                XmlNodeList anchorNodesList = sheet.SelectNodes("表:工作表内容/uof:锚点", nm);
                XmlNodeList colList = sheet.SelectNodes("表:工作表内容/表:列", nm);
                XmlNodeList rowList = sheet.SelectNodes("表:工作表内容/表:行", nm);
                if (chartNodesList != null)
                {
                    foreach (XmlNode chart in chartNodesList)
                    {
                        double xCor = double.Parse(chart.Attributes["表:x坐标"].Value.ToString());
                        double yCor = double.Parse(chart.Attributes["表:y坐标"].Value.ToString());
                        double tempData = 0.0;
                        tempData = xCor / 72.0;
                        int maxCol = (int)tempData;
                        tempData = yCor / 19.0;
                        int maxRow = (int)tempData;
                        double sum = 0.0;
                        double sumPre = 0.0;
                        bool has;
                        int _fromRow = 0;
                        int _fromCol = 0;
                        int _fromRowOff = 0;
                        int _fromColOff = 0;

                        int _toRow = 0;
                        int _toCol = 0;
                        int _toRowOff = 0;
                        int _toColOff = 0;

                        double width = double.Parse(chart.Attributes["表:宽度"].Value.ToString());
                        double height = double.Parse(chart.Attributes["表:高度"].Value.ToString());

                        tempData = (width + xCor) / 72.0;
                        int _maxCol = (int)tempData;
                        tempData = (height + yCor) / 19.0;
                        int _maxRow = (int)tempData;

                        //form col
                        for (int i = 1; i < maxCol + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (colList != null)
                            {
                                foreach (XmlNode col in colList)
                                {
                                    if (int.Parse(col.Attributes["表:列号"].Value.ToString()) == i)
                                    {
                                        has = true;
                                        temp = double.Parse(col.Attributes["表:列宽"].Value.ToString()) * 4 / 3;
                                        sum += temp;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 72.0;
                                sum += 72.0;
                            }
                            if (sum > xCor)
                            {
                                _fromCol = i - 1;
                                _fromColOff = Convert.ToInt32((xCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute fromCol = xdoc.CreateAttribute("fromCol");
                                fromCol.Value = _fromCol.ToString();
                                XmlAttribute fromColOff = xdoc.CreateAttribute("fromColOff");
                                fromColOff.Value = _fromColOff.ToString();

                                chart.Attributes.Append(fromCol);
                                chart.Attributes.Append(fromColOff);
                                break;
                            }
                        }
                        if (sum <= xCor)
                        {
                            _fromCol = maxCol;
                            _fromColOff = Convert.ToInt32((xCor - sum) * 360000 / 28.3);
                            XmlAttribute fromCol = xdoc.CreateAttribute("fromCol");
                            fromCol.Value = _fromCol.ToString();
                            XmlAttribute fromColOff = xdoc.CreateAttribute("fromColOff");
                            fromColOff.Value = _fromColOff.ToString();

                            chart.Attributes.Append(fromCol);
                            chart.Attributes.Append(fromColOff);
                        }

                        //from row
                        sum = 0;
                        sumPre = 0;
                        for (int i = 1; i < maxRow + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (rowList != null)
                            {
                                foreach (XmlNode row in rowList)
                                {
                                    if (int.Parse(row.Attributes["表:行号"].Value.ToString()) == i)
                                    {
                                        if (row.Attributes.GetNamedItem("表:行高") != null)
                                        {
                                            temp = double.Parse(row.Attributes["表:行高"].Value.ToString()) * 19 / 12;
                                        }
                                        else
                                            temp = 14.25 * 19 / 12;
                                        sum += temp;
                                        has = true;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 19;
                                sum += 19;
                            }
                            if (sum >= yCor)
                            {
                                //fromYOff = sum - temp;
                                _fromRow = i - 1;
                                _fromRowOff = Convert.ToInt32((yCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute fromRow = xdoc.CreateAttribute("fromRow");
                                fromRow.Value = _fromRow.ToString();
                                XmlAttribute fromRowOff = xdoc.CreateAttribute("fromRowOff");
                                fromRowOff.Value = _fromRowOff.ToString();

                                chart.Attributes.Append(fromRow);
                                chart.Attributes.Append(fromRowOff);
                                //sum = sumPre;
                                break;
                            }
                            sumPre = sum;
                        }
                        if (sum <= yCor)
                        {
                            _fromRow = maxRow;
                            _fromRowOff = Convert.ToInt32((yCor - sum) * 360000 / 28.3);
                            XmlAttribute fromRow = xdoc.CreateAttribute("fromRow");
                            fromRow.Value = _fromRow.ToString();
                            XmlAttribute fromRowOff = xdoc.CreateAttribute("fromRowOff");
                            fromRowOff.Value = _fromRowOff.ToString();

                            chart.Attributes.Append(fromRow);
                            chart.Attributes.Append(fromRowOff);
                        }

                        //to col
                        sum = 0;
                        for (int i = 1; i < _maxCol + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (colList != null)
                            {
                                foreach (XmlNode col in colList)
                                {
                                    if (int.Parse(col.Attributes["表:列号"].Value.ToString()) == i)
                                    {
                                        has = true;
                                        temp = double.Parse(col.Attributes["表:列宽"].Value.ToString()) * 4 / 3;
                                        sum += temp;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 72.0;
                                sum += 72.0;
                            }
                            if (sum > width + xCor)
                            {
                                _toCol = i - 1;
                                _toColOff = Convert.ToInt32((width + xCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute toCol = xdoc.CreateAttribute("toCol");
                                toCol.Value = _toCol.ToString();
                                XmlAttribute toColOff = xdoc.CreateAttribute("toColOff");
                                toColOff.Value = _toColOff.ToString();

                                chart.Attributes.Append(toCol);
                                chart.Attributes.Append(toColOff);
                                break;
                            }
                        }
                        if (sum <= width + xCor)
                        {
                            _toCol = _maxCol;
                            _toColOff = Convert.ToInt32((width + xCor - sum) * 360000 / 28.3);
                            XmlAttribute toCol = xdoc.CreateAttribute("toCol");
                            toCol.Value = _toCol.ToString();
                            XmlAttribute toColOff = xdoc.CreateAttribute("toColOff");
                            toColOff.Value = _toColOff.ToString();

                            chart.Attributes.Append(toCol);
                            chart.Attributes.Append(toColOff);
                        }

                        //to row
                        sum = 0;
                        for (int i = 1; i < _maxRow + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (rowList != null)
                            {
                                foreach (XmlNode row in rowList)
                                {
                                    if (int.Parse(row.Attributes["表:行号"].Value.ToString()) == i)
                                    {
                                        if (row.Attributes.GetNamedItem("表:行高") != null)
                                        {
                                            temp = double.Parse(row.Attributes["表:行高"].Value.ToString()) * 19 / 12;
                                        }
                                        else
                                            temp = 14.25 * 19 / 12;
                                        sum += temp;
                                        has = true;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 19.0;
                                sum += 19.0;
                            }
                            if (sum >= height + yCor)
                            {
                                //toYOff = sum - temp;
                                _toRow = i - 1;
                                _toRowOff = Convert.ToInt32((height + yCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute toRow = xdoc.CreateAttribute("toRow");
                                toRow.Value = _toRow.ToString();
                                XmlAttribute toRowOff = xdoc.CreateAttribute("toRowOff");
                                toRowOff.Value = _toRowOff.ToString();

                                chart.Attributes.Append(toRow);
                                chart.Attributes.Append(toRowOff);
                                //sum = sumPre;
                                break;
                            }
                            sumPre = sum;
                        }
                        if (sum <= height + yCor)
                        {
                            _toRow = _maxRow;
                            _toRowOff = Convert.ToInt32((height + yCor - sum) * 360000 / 28.3);
                            XmlAttribute toRow = xdoc.CreateAttribute("toRow");
                            toRow.Value = _toRow.ToString();
                            XmlAttribute toRowOff = xdoc.CreateAttribute("toRowOff");
                            toRowOff.Value = _toRowOff.ToString();

                            chart.Attributes.Append(toRow);
                            chart.Attributes.Append(toRowOff);
                        }


                    }
                }

                //Anchor Position
                if (anchorNodesList != null)
                {
                    foreach (XmlNode anchor in anchorNodesList)
                    {
                        double xCor = double.Parse(anchor.Attributes["uof:x坐标"].Value.ToString());
                        double yCor = double.Parse(anchor.Attributes["uof:y坐标"].Value.ToString());
                        double tempData = 0.0;
                        tempData = xCor / 54.0;
                        int maxCol = (int)tempData;
                        tempData = yCor / 14.25;
                        int maxRow = (int)tempData;
                        double sum = 0.0;
                        double sumPre = 0.0;
                        bool has;
                        int _fromRow = 0;
                        int _fromCol = 0;
                        int _fromRowOff = 0;
                        int _fromColOff = 0;

                        int _toRow = 0;
                        int _toCol = 0;
                        int _toRowOff = 0;
                        int _toColOff = 0;

                        double width = double.Parse(anchor.Attributes["uof:宽度"].Value.ToString());
                        double height = double.Parse(anchor.Attributes["uof:高度"].Value.ToString());

                        tempData = (width + xCor) / 54.0;
                        int _maxCol = (int)tempData;
                        tempData = (height + yCor) / 14.25;
                        int _maxRow = (int)tempData;

                        //form col
                        for (int i = 1; i < maxCol + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (colList != null)
                            {
                                foreach (XmlNode col in colList)
                                {
                                    if (int.Parse(col.Attributes["表:列号"].Value.ToString()) == i)
                                    {
                                        has = true;
                                        temp = double.Parse(col.Attributes["表:列宽"].Value.ToString());
                                        sum += temp;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 72.0;
                                sum += 72.0;
                            }
                            if (sum > xCor)
                            {
                                _fromCol = i - 1;
                                _fromColOff = Convert.ToInt32((xCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute fromCol = xdoc.CreateAttribute("fromCol");
                                fromCol.Value = _fromCol.ToString();
                                XmlAttribute fromColOff = xdoc.CreateAttribute("fromColOff");
                                fromColOff.Value = _fromColOff.ToString();

                                anchor.Attributes.Append(fromCol);
                                anchor.Attributes.Append(fromColOff);
                                break;
                            }
                        }
                        if (sum <= xCor)
                        {
                            _fromCol = maxCol;
                            _fromColOff = Convert.ToInt32((xCor - sum) * 360000 / 28.3);
                            XmlAttribute fromCol = xdoc.CreateAttribute("fromCol");
                            fromCol.Value = _fromCol.ToString();
                            XmlAttribute fromColOff = xdoc.CreateAttribute("fromColOff");
                            fromColOff.Value = _fromColOff.ToString();

                            anchor.Attributes.Append(fromCol);
                            anchor.Attributes.Append(fromColOff);
                        }

                        //from row
                        sum = 0;
                        sumPre = 0;
                        for (int i = 1; i < maxRow + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (rowList != null)
                            {
                                foreach (XmlNode row in rowList)
                                {
                                    if (int.Parse(row.Attributes["表:行号"].Value.ToString()) == i)
                                    {
                                        if (row.Attributes.GetNamedItem("表:行高") != null)
                                        {
                                            temp = double.Parse(row.Attributes["表:行高"].Value.ToString());
                                        }
                                        else
                                            temp = 14.25;
                                        sum += temp;
                                        has = true;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 19;
                                sum += 19;
                            }
                            if (sum >= yCor)
                            {
                                //fromYOff = sum - temp;
                                _fromRow = i - 1;
                                _fromRowOff = Convert.ToInt32((yCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute fromRow = xdoc.CreateAttribute("fromRow");
                                fromRow.Value = _fromRow.ToString();
                                XmlAttribute fromRowOff = xdoc.CreateAttribute("fromRowOff");
                                fromRowOff.Value = _fromRowOff.ToString();

                                anchor.Attributes.Append(fromRow);
                                anchor.Attributes.Append(fromRowOff);
                                //sum = sumPre;
                                break;
                            }
                            sumPre = sum;
                        }
                        if (sum <= yCor)
                        {
                            _fromRow = maxRow;
                            _fromRowOff = Convert.ToInt32((yCor - sum) * 360000 / 28.3);
                            XmlAttribute fromRow = xdoc.CreateAttribute("fromRow");
                            fromRow.Value = _fromRow.ToString();
                            XmlAttribute fromRowOff = xdoc.CreateAttribute("fromRowOff");
                            fromRowOff.Value = _fromRowOff.ToString();

                            anchor.Attributes.Append(fromRow);
                            anchor.Attributes.Append(fromRowOff);
                        }

                        //to col
                        sum = 0;
                        for (int i = 1; i < _maxCol + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (colList != null)
                            {
                                foreach (XmlNode col in colList)
                                {
                                    if (int.Parse(col.Attributes["表:列号"].Value.ToString()) == i)
                                    {
                                        has = true;
                                        temp = double.Parse(col.Attributes["表:列宽"].Value.ToString());
                                        sum += temp;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 72.0;
                                sum += 72.0;
                            }
                            if (sum > width + xCor)
                            {
                                _toCol = i - 1;
                                _toColOff = Convert.ToInt32((width + xCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute toCol = xdoc.CreateAttribute("toCol");
                                toCol.Value = _toCol.ToString();
                                XmlAttribute toColOff = xdoc.CreateAttribute("toColOff");
                                toColOff.Value = _toColOff.ToString();

                                anchor.Attributes.Append(toCol);
                                anchor.Attributes.Append(toColOff);
                                break;
                            }
                        }
                        if (sum <= width + xCor)
                        {
                            _toCol = _maxCol;
                            _toColOff = Convert.ToInt32((width + xCor - sum) * 360000 / 28.3);
                            XmlAttribute toCol = xdoc.CreateAttribute("toCol");
                            toCol.Value = _toCol.ToString();
                            XmlAttribute toColOff = xdoc.CreateAttribute("toColOff");
                            toColOff.Value = _toColOff.ToString();

                            anchor.Attributes.Append(toCol);
                            anchor.Attributes.Append(toColOff);
                        }

                        //to row
                        sum = 0;
                        for (int i = 1; i < _maxRow + 1; i++)
                        {
                            double temp = 0.0;
                            has = false;
                            if (rowList != null)
                            {
                                foreach (XmlNode row in rowList)
                                {
                                    if (int.Parse(row.Attributes["表:行号"].Value.ToString()) == i)
                                    {
                                        if (row.Attributes.GetNamedItem("表:行高") != null)
                                        {
                                            temp = double.Parse(row.Attributes["表:行高"].Value.ToString());
                                        }
                                        else
                                            temp = 14.25;
                                        sum += temp;
                                        has = true;
                                        break;
                                    }
                                }
                            }
                            if (!has)
                            {
                                temp = 19.0;
                                sum += 19.0;
                            }
                            if (sum >= height + yCor)
                            {
                                //toYOff = sum - temp;
                                _toRow = i - 1;
                                _toRowOff = Convert.ToInt32((height + yCor - sum + temp) * 360000 / 28.3);
                                XmlAttribute toRow = xdoc.CreateAttribute("toRow");
                                toRow.Value = _toRow.ToString();
                                XmlAttribute toRowOff = xdoc.CreateAttribute("toRowOff");
                                toRowOff.Value = _toRowOff.ToString();

                                anchor.Attributes.Append(toRow);
                                anchor.Attributes.Append(toRowOff);
                                //sum = sumPre;
                                break;
                            }
                            sumPre = sum;
                        }
                        if (sum <= height + yCor)
                        {
                            _toRow = _maxRow;
                            _toRowOff = Convert.ToInt32((height + yCor - sum) * 360000 / 28.3);
                            XmlAttribute toRow = xdoc.CreateAttribute("toRow");
                            toRow.Value = _toRow.ToString();
                            XmlAttribute toRowOff = xdoc.CreateAttribute("toRowOff");
                            toRowOff.Value = _toRowOff.ToString();

                            anchor.Attributes.Append(toRow);
                            anchor.Attributes.Append(toRowOff);
                        }


                    }
                }

            }

            FileStream fs = null;


            try
            {
                fs = new FileStream(destFileName, FileMode.Create);
                xdoc.Save(fs);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw new Exception("Fail in adding condition order");
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }

            }
            /**/
        }
        private void GetKeyPoints(string fileName, string destFileName)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(fileName);

            XmlNameTable table = xmlDoc.NameTable;
            XmlNamespaceManager nm = new XmlNamespaceManager(table);
            nm.AddNamespace("", "http://uof2ooxPre");
            nm.AddNamespace("字", "http://schemas.uof.org/cn/2003/uof-wordproc");
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("图", "http://schemas.uof.org/cn/2003/graph");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            nm.AddNamespace("演", "http://schemas.uof.org/cn/2003/uof-slideshow");
            nm.AddNamespace("a", "http://purl.oclc.org/ooxml/drawingml/main");
            nm.AddNamespace("app", "http://purl.oclc.org/ooxml/officeDocument/extendedProperties");
            nm.AddNamespace("cp", "http://schemas.openxmlformats.org/package/2006/metadata/core-properties");
            nm.AddNamespace("dc", "http://purl.org/dc/elements/1.1/");
            nm.AddNamespace("dcmitype", "http://purl.org/dc/dcmitype/");
            nm.AddNamespace("dcterms", "http://purl.org/dc/terms/");
            nm.AddNamespace("fo", "http://www.w3.org/1999/XSL/Format");
            nm.AddNamespace("p", "http://purl.oclc.org/ooxml/presentationml/main");
            nm.AddNamespace("r", "http://purl.oclc.org/ooxml/officeDocument/relationships");
            nm.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");
            nm.AddNamespace("cus", "http://purl.oclc.org/ooxml/officeDocument/customProperties");
            nm.AddNamespace("vt", "http://purl.oclc.org/ooxml/officeDocument/docPropsVTypes");

            XmlNodeList shapeKeyCoordinates = xmlDoc.SelectNodes("//图:关键点坐标", nm);
            if (shapeKeyCoordinates != null)
            {
                foreach (XmlNode shapeKeyCoordinate in shapeKeyCoordinates)
                {
                    string keyPath;
                    keyPath = shapeKeyCoordinate.Attributes.GetNamedItem("图:路径").Value;
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
            xmlDoc.Save(destFileName);
        }
        private void NumberFormatsTransfer(string fileName, string destFileName)
        {
            XmlDocument document = new XmlDocument();
            document.Load(fileName);
            XmlNameTable nameTable = document.NameTable;
            XmlNamespaceManager nm = new XmlNamespaceManager(nameTable);
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            FileStream fs = new FileStream(destFileName, FileMode.Create);
            XmlNodeList nodeList = document.SelectNodes("/uof:UOF/uof:式样集/uof:单元格式样", nm);

            if (nodeList != null)
            {
                foreach (XmlNode node in nodeList)
                {
                    XmlNode NumFmt = node.SelectSingleNode("./表:数字格式", nm);
                    if (NumFmt != null)
                    {
                        string type = NumFmt.Attributes["表:分类名称"].Value.ToString();
                        string value = NumFmt.Attributes["表:格式码"].Value.ToString();
                        if (type != "general")
                        {
                            string newValue = GetProperNumFmts(type, value);
                            NumFmt.Attributes["表:格式码"].Value = newValue;
                        }
                    }

                }
            }
            /* * */
            document.Save(fs);
            fs.Close();
        }

        private string GetProperNumFmts(string type, string value)
        {
            string newValue = "";
            string label = "";
            switch (type)
            {
                case "number":
                    if (value.Equals("0.00_ ;[Red]-0.00_ "))
                    {
                        newValue = @"0.00_ ;[Red]\-0.00\ ";
                    }
                    else if (value.Equals("0.00_ "))
                    {
                        newValue = @"0.00_ ";
                    }
                    else if (value.Equals("0.00;[Red]0.00"))
                    {
                        newValue = "0.00;[Red]0.00";
                    }
                    else if (value.Equals("0.00_);(0.00)"))
                    {
                        newValue = @"0.00_);\(0.00\)";
                    }
                    else if (value.Equals("0.00_);[Red](0.00)"))
                    {
                        newValue = @"0.00_);[Red]\(0.00\)";
                    }
                    else
                        newValue = "General";
                    break;

                case "currency":

                    if (value.Contains("￥"))
                        label = "￥";
                    else if (value.Contains("$"))
                        label = "$";
                    else
                        label = "￥";

                    value = value.Replace(label, "O");
                    if (value.Equals("O#,##0.00;[Red]O-#,##0.00"))
                    {
                        newValue = "\"O\"#,##0.00;[Red]\"O\"\\-#,##0.00";
                    }
                    else if (value.Equals("O#,##0.00;O-#,##0.00"))
                    {
                        newValue = "\"O\"#,##0.00;\"O\"\\-#,##0.00";
                    }
                    else if (value.Equals("O#,##0.00;[Red]O#,##0.00"))
                    {
                        newValue = "\"O\"#,##0.00;[Red]\"O\"#,##0.00";
                    }
                    else if (value.Equals("O#,##0.00_);[Red](O#,##0.00)"))
                    {
                        newValue = "\"O\"#,##0.00_);[Red]\\(\"O\"#,##0.00\\)";
                    }
                    else if (value.Equals("O#,##0.00_);(O#,##0.00)"))
                    {
                        newValue = "\"O\"#,##0.00_);\\(\"O\"#,##0.00\\)";
                    }
                    else
                        newValue = "\"O\"#,##0.00;\"O\"\\-#,##0.00";

                    newValue = newValue.Replace("O", label);
                    break;

                case "accounting":
                    if (value.Contains("￥"))
                        label = "￥";
                    else if (value.Contains("$"))
                        label = "$";
                    else
                        label = "￥";
                    value = value.Replace(label, "O");
                    if (value.Contains("_ O* #,##0.00_ ;_ O* -#,##0.00_ ;_ O* \"-\"??_ ;_ @_ "))
                    {
                        newValue = "_ \"O\"* #,##0.00_ ;_ \"O\"* \\-#,##0.00_ ;_ \"O\"* \"-\"??_ ;_ @_ ";
                    }
                    else if (value.Contains("_-O* #,##0.00_ ;_-O* -#,##0.00 ;_-O* \"-\"??_ ;_-@_ "))
                    {
                        newValue = "_-\\O* #,##0.00_ ;_-\\O* \\-#,##0.00\\ ;_-\\O* \"-\"??_ ;_-@_ ";
                    }
                    else
                        newValue = "_-\\O* #,##0.00_ ;_-\\O* \\-#,##0.00\\ ;_-\\O* \"-\"??_ ;_-@_ ";

                    newValue = newValue.Replace("O", label);
                    break;

                case "date":
                    if (value.Contains("[DBNum1]yyyy\"年\"m\"月\"d\"日\";@"))
                    {
                        newValue = "[DBNum1][$-804]yyyy\"年\"m\"月\"d\"日\";@";
                    }
                    else if (value.Contains("[DBNum1]yyyy\"年\"m\"月\";@"))
                    {
                        newValue = "[DBNum1][$-804]yyyy\"年\"m\"月\";@";
                    }

                    else if (value.Contains("[DBNum1]m\"月\"d\"日\";@"))
                    {
                        newValue = "[DBNum1][$-804]m\"月\"d\"日\";@";
                    }
                    else if (value.Contains("yyyy\"年\"m\"月\"d\"日\";@"))
                    {
                        newValue = "yyyy\"年\"m\"月\"d\"日\";@";
                    }
                    else if (value.Contains("yyyy\"年\"m\"月\";@"))
                    {
                        newValue = "yyyy\"年\"m\"月\";@";
                    }
                    else if (value.Contains("m\"月\"d\"日\";@"))
                    {
                        newValue = "m\"月\"d\"日\";@";
                    }


                    else if (value.Contains("aaaa;@"))
                    {
                        newValue = "[$-804]aaaa;@";
                    }

                    else if (value.Contains("aaa;@"))
                    {
                        newValue = "[$-804]aaa;@";
                    }
                    else if (value.Contains("yyyy-m-d;@"))
                    {
                        newValue = "yyyy/m/d;@";
                    }
                    else if (value.Contains("yyyy-m-d h:mm AM/PM;@"))
                    {
                        newValue = "[$-409]yyyy/m/d\\ h:mm\\ AM/PM;@";
                    }
                    else if (value.Contains("yyyy-m-d h:mm;@"))
                    {
                        newValue = "yyyy/m/d\\ h:mm;@";
                    }


                    else if (value.Contains("m-d;@"))
                    {
                        newValue = "yy/m/d;@";
                    }

                    else if (value.Contains("mm-dd-yy;@"))
                    {
                        newValue = "mm/dd/yy;@";
                    }
                    else if (value.Contains("d-mmm;@"))
                    {
                        newValue = "[$-409]d\\-mmm\\-yy;@";
                    }
                    else if (value.Contains("d-mmm-yy;@"))
                    {
                        newValue = "[$-409]d\\-mmm\\-yy;@";
                    }
                    else if (value.Contains("dd-mmm-yy;@"))
                    {
                        newValue = @"[$-409]dd\-mmm\-yy;@";
                    }



                    else if (value.Contains("mmm-yy;@"))
                    {
                        newValue = @"[$-409]mmm\-yy;@";
                    }
                    else if (value.Contains("mmmm-yy;@"))
                    {
                        newValue = @"[$-409]mmmm\-yy;@";
                    }
                    else if (value.Contains("mmmmm;@"))
                    {
                        newValue = @"[$-409]mmmmm;@";
                    }
                    else if (value.Contains("mmmmm-yy;@"))
                    {
                        newValue = @"[$-409]mmmmm\-yy;@";
                    }


                    else
                        newValue = "_-\\O* #,##0.00_ ;_-\\O* \\-#,##0.00\\ ;_-\\O* \"-\"??_ ;_-@_ ";
                    break;

                case "percentage":
                    if (value.Contains("0.00%"))
                        newValue = "0.00%";
                    break;

                default:

                    newValue = "General";
                    break;

            }


            /*
             * if (value.Equals("O#,##0.00;[Red]O-#,##0.00"))
                    {
                        newValue = "&quot;O&quot;#,##0.00;[Red]&quot;O&quot;\\-#,##0.00";
                    }
                    else if (value.Equals("O#,##0.00;O-#,##0.00"))
                    {
                        newValue = "&quot;O&quot;#,##0.00;&quot;O&quot;\\-#,##0.00";
                    }
                    else if (value.Equals("O#,##0.00;[Red]O#,##0.00"))
                    {
                        newValue = "&quot;O&quot;#,##0.00;[Red]&quot;O&quot;#,##0.00";
                    }
                    else if (value.Equals("O#,##0.00_);[Red](O#,##0.00)"))
                    {
                        newValue = "&quot;O&quot;#,##0.00_);[Red]\\(&quot;O&quot;#,##0.00\\)";
                    }
                    else if (value.Equals("O#,##0.00_);(O#,##0.00)"))
                    {
                        newValue = "&quot;O&quot;#,##0.00_);\\(&quot;O&quot;#,##0.00\\)";
                    }
             * */


            /*

              if (type == "number")
              {
                  char[] seperator = new char[1];
                  seperator[0] = ';';
                  string[] value_split = new string[2];
                  value_split = value.Split(seperator[0]);

                  if (value_split.Length == 2)
                  {
                      value_split[1] = value_split[1].Replace("-", @"\-");
                      value_split[1] = value_split[1].Replace("_", @"\");
                      value_split[1] = value_split[1].Replace("(", @"\(");
                      value_split[1] = value_split[1].Replace(")", @"\)");
                      newValue = value_split[0] + ";" + value_split[1];
                  }
              }
            
             if (type == "currency")
              {
                  char[] seperator = new char[1];
                  seperator[0] = ';';
                  string[] value_split = new string[2];
                  value_split = value.Split(seperator[0]);

                  if (value_split.Length == 2)
                  {
                      value_split[1] = value_split[1].Replace("-", @"\-");
                      value_split[1] = value_split[1].Replace("_", @"\");
                      value_split[1] = value_split[1].Replace("(", @"\(");
                      value_split[1] = value_split[1].Replace(")", @"\)");
                      newValue = value_split[0] + ";" + value_split[1];

                      newValue = newValue.Replace("￥", @"&quot;￥&quot;");
                      newValue = newValue.Replace("US$", @"&quot;US$&quot;");
                  }
              }
               * */
            return newValue;
        }

        #region util to fix bugs or do something for pretreatment

        private void ChangeAnchorAttributeName(string fileName, string destFileName)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(fileName);
            XmlNameTable nameTable = xdoc.NameTable;
            XmlNamespaceManager nm = new XmlNamespaceManager(nameTable);
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            FileStream fsChangeAttr = new FileStream(destFileName, FileMode.Create);

            XmlNodeList oldAnchors = null;
            XmlNode workContent = null;
            int sheetsNum = xdoc.SelectNodes("/uof:UOF/uof:电子表格/表:主体/表:工作表", nm).Count;
            string workContentString = "";
            string anchorString = "";
            for (int i = 1; i < sheetsNum + 1; i++)
            {
                workContent = xdoc.SelectSingleNode("//uof:UOF/uof:电子表格/表:主体/表:工作表[" + i + "]" + "/表:工作表内容", nm);
                oldAnchors = xdoc.SelectNodes("//uof:UOF/uof:电子表格/表:主体/表:工作表[" + i + "]" + "/表:工作表内容/uof:锚点", nm);

                if (workContent != null && oldAnchors != null)
                {
                    foreach (XmlNode anchor in oldAnchors)
                    {
                        anchorString += anchor.OuterXml;
                        workContent.RemoveChild(anchor);
                    }

                    workContentString = workContent.InnerXml;
                    if (anchorString.Contains("表:x坐标"))
                    {
                        anchorString = anchorString.Replace("表:x坐标", "uof:x坐标");
                    }
                    if (anchorString.Contains("表:y坐标"))
                    {
                        anchorString = anchorString.Replace("表:y坐标", "uof:y坐标");
                    }
                    if (anchorString.Contains("表:宽度"))
                    {
                        anchorString = anchorString.Replace("表:宽度", "uof:宽度");
                    }
                    if (anchorString.Contains("表:高度"))
                    {
                        anchorString = anchorString.Replace("表:高度", "uof:高度");
                    }

                    workContentString += anchorString;
                    workContent.InnerXml = workContentString;

                }
                workContentString = "";
                anchorString = "";
            }
            try
            {
                xdoc.Save(fsChangeAttr);
                fsChangeAttr.Close();
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                logger.Error(ex.StackTrace);
            }
            finally
            {
                if (fsChangeAttr != null)
                    fsChangeAttr.Close();
            }
        }


        private void PreGrpShaps(string inputFileName, string destFileName)
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputFileName);
            XmlTextWriter tw = null;

            XmlNameTable nTable = xmlDoc.NameTable;
            XmlNamespaceManager nm = new XmlNamespaceManager(nTable);
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            nm.AddNamespace("图", "http://schemas.uof.org/cn/2003/graph");
            XmlNodeList sheets = xmlDoc.SelectNodes("//表:工作表", nm);
            XmlNodeList anchors = null;
            string assembleList;
            string[] splitAssembleList;
            LinkedList<string> allInSplitedAssembleList;
            allInSplitedAssembleList = new LinkedList<string>();
            LinkedList<string>.Enumerator enumer = allInSplitedAssembleList.GetEnumerator();
            int nodeAttributeNum;
            bool flag = false;//标识链接状态
            int currentNodePositionInLinkedList;

            foreach (XmlNode sheet in sheets)
            {
                anchors = sheet.SelectNodes(".//uof:锚点[@uof:图形引用]", nm);
                for (int i = 0; i < anchors.Count; i++)
                {
                    string selectNode = "uof:UOF/uof:对象集/图:图形[@图:标识符='" + anchors[i].Attributes["uof:图形引用"].Value + "']";
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
                                selectNode = "uof:UOF/uof:对象集/图:图形[@图:标识符='" + enumer.Current.ToString() + "']";

                                anchors[i].AppendChild(xmlDoc.SelectSingleNode(selectNode, nm).Clone());
                                for (nodeAttributeNum = 0; nodeAttributeNum < xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Count; nodeAttributeNum++)
                                {
                                    if (xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Item(nodeAttributeNum).Name.ToString() == "图:组合列表")
                                    {
                                        assembleList = xmlDoc.SelectSingleNode(selectNode, nm).Attributes.GetNamedItem("图:组合列表").Value.ToString().TrimEnd(' ');
                                        splitAssembleList = assembleList.Split(new char[1] { ' ' });
                                        foreach (string s in splitAssembleList)
                                        {
                                            allInSplitedAssembleList.AddLast(s);
                                        }
                                        enumer = allInSplitedAssembleList.GetEnumerator();
                                        for (int current = 0; current < currentNodePositionInLinkedList; current++)
                                        {
                                            enumer.MoveNext();
                                        }
                                    }
                                }
                            }
                            flag = false;
                            break;
                        }
                        else if (!flag)
                        {
                            for (nodeAttributeNum = 0; nodeAttributeNum < xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Count; nodeAttributeNum++)
                            {
                                if (xmlDoc.SelectSingleNode(selectNode, nm).Attributes.Item(nodeAttributeNum).Name.ToString() == "图:组合列表")
                                {
                                    assembleList = xmlDoc.SelectSingleNode(selectNode, nm).Attributes.GetNamedItem("图:组合列表").Value.ToString().TrimEnd(' ');
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
                        {
                            break;
                        }
                    }
                }

            }
            try
            {
                tw = new XmlTextWriter(destFileName, Encoding.UTF8);
                xmlDoc.Save(tw);
                tw.Close();
            }
            catch (Exception ex)
            {
                logger.Error("fail in moving the group-shapes state");
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

        private void AddOrderOnCF(string inputFileName, string destFileName)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFileName);

            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");

            XmlNodeList conditionalFormatSet = xdoc.SelectNodes(".//表:条件格式化", nm);
            if (conditionalFormatSet != null)
            {
                int i = 0;
                foreach (XmlNode conditionalFormat in conditionalFormatSet)
                {
                    XmlNodeList conditions = conditionalFormat.SelectNodes("表:条件", nm);
                    foreach (XmlNode condition in conditions)
                    {
                        // add attribue of order
                        XmlAttribute order = xdoc.CreateAttribute("order");
                        order.Value = i.ToString();
                        i++;

                        condition.Attributes.Append(order);
                    }
                }
            }

            FileStream fs = null;
            try
            {
                fs = new FileStream(destFileName, FileMode.Create);
                xdoc.Save(fs);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw new Exception("Fail in adding condition order");
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }

            }

        }


        /// <summary>
        /// 把INT数转成字母，A-Z，之后是AA-AZ...
        /// </summary>
        /// <param name="inputInt">输入一个整数</param>
        /// <returns>字母串</returns>
        private string IntTo26(int inputInt)
        {
            int i = 0;
            if (inputInt > 0)
            {
                i = inputInt;
            }
            else
            {
                return "";
            }

            char[] alphabet = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
            Stack<char> result = new Stack<char>();

            while (i > 0)
            {
                // 整除时，进‘Z’前一位退1
                if ((i % 26) == 0)
                {
                    result.Push('Z');
                    i -= 26;
                    i = i / 26;
                }
                else
                {
                    result.Push(alphabet[i % 26 - 1]);
                    i = i / 26;
                }
            }

            string string26 = "";
            while (result.Count > 0)
            {
                string26 += result.Pop().ToString();
            }

            return string26;
        }


        /// <summary>
        /// 把字母转换成数字，AA对应的是27
        /// </summary>
        /// <param name="inputAlphabet"></param>
        /// <returns></returns>
        private int Convert26ToInt(string inputAlphabet)
        {
            char[] input = inputAlphabet.Trim().ToUpper().ToCharArray();
            int result = 0;

            for (int i = 0; i < input.Length; i++)
            {
                int alphaVal = Convert.ToInt32(input[i]) - 64;
                result += alphaVal * (int)Math.Pow(26, input.Length - i - 1);
            }

            return result;
        }


        private void ConvertRowIntTo26(string inputFileName, string outputFileName)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFileName);

            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");

            XmlNodeList sheets = xdoc.SelectNodes("//表:工作表", nm);
            foreach (XmlNode sheet in sheets)
            {
                XmlNodeList rows = sheet.SelectNodes("表:工作表内容/表:行", nm);
                if (rows != null)
                {
                    foreach (XmlNode row in rows)
                    {
                        XmlNodeList cells = row.SelectNodes("表:单元格", nm);
                        if (cells != null)
                        {
                            foreach (XmlNode cell in cells)
                            {
                                int rowNum = Convert.ToInt32(cell.Attributes.GetNamedItem("表:列号").Value.ToString());
                                //cell.Attributes.GetNamedItem("表:列号").Value = IntTo26(rowNum);
                                XmlAttribute row26 = xdoc.CreateAttribute("表:列号值", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
                                row26.Value = IntTo26(rowNum);
                                cell.Attributes.Append(row26);
                            }
                        }
                    }
                }
            }

            FileStream fs = null;
            try
            {
                fs = new FileStream(outputFileName, FileMode.Create);
                xdoc.Save(fs);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw new Exception("Fail in convert row numbe of PreProcessorOne");
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }
            }
        }


        private void AddGroup(string inputFileName, string destFileName)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFileName);



            //FileStream fs2 = new FileStream(@"c:\test.txt", FileMode.Create);
            // StreamWriter sw = new StreamWriter(fs2);


            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            nm.AddNamespace("表", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
            XmlNodeList sheetList = xdoc.SelectNodes("//表:工作表", nm);

            foreach (XmlNode sheet in sheetList)
            {

                XmlNode sheetCon = sheet.SelectSingleNode("表:工作表内容", nm);
                XmlNodeList rowsList = sheet.SelectNodes("表:工作表内容/表:行", nm);
                XmlNodeList colsList = sheet.SelectNodes("表:工作表内容/表:列", nm);
                int rowCount = rowsList.Count;
                int max = 0;
                XmlNode rowGroup = sheet.SelectSingleNode("表:工作表内容/表:分组集", nm);

                if (rowGroup != null)
                {
                    //int rowCount = 0;

                    try
                    {
                        XmlNodeList groupList = rowGroup.SelectNodes("表:行", nm);
                        foreach (XmlNode group in groupList)
                        {
                            int rowNo = int.Parse(group.Attributes["表:终止"].Value.ToString());
                            max = max > rowNo ? max : rowNo;
                        }
                        int groupCount = groupList.Count;
                        ArrayList rows = new ArrayList(max);
                        for (int i = 0; i < max; i++)
                            rows.Add(0);
                        int start = 0;
                        int end = 0;

                        foreach (XmlNode group in groupList)
                        {
                            string a = group.Attributes["表:起始"].Value.ToString();
                            string b = group.Attributes["表:终止"].Value.ToString();
                            start = int.Parse(a);
                            end = int.Parse(b);


                            for (int temp = start - 1; temp < end; temp++)
                                rows[temp] = int.Parse(rows[temp].ToString()) + 1;
                        }



                        for (int k = 0; k < max; k++)
                        {
                            //sw.WriteLine(rows[k].ToString());
                            XmlAttribute levelAttr = xdoc.CreateAttribute("level");
                            levelAttr.Value = rows[k].ToString();

                            //foreach (XmlNode row in rowsList)
                            //{
                            //if (rowsList[k].Attributes["表:行号"].Value == k.ToString())
                            //levelAttr.AppendChild(rowsList[k]);
                            rowsList[k].Attributes.Append(levelAttr);
                            // }
                        }

                        rows.Clear();
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex.Message);
                    }


                    try
                    {
                        max = 0;
                        XmlNodeList groupList = rowGroup.SelectNodes("表:列", nm);
                        foreach (XmlNode group in groupList)
                        {
                            int colNo = int.Parse(group.Attributes["表:终止"].Value.ToString());
                            max = max > colNo ? max : colNo;
                        }
                        int groupCount = groupList.Count;
                        ArrayList cols = new ArrayList(max);
                        for (int i = 0; i < max; i++)
                            cols.Add(0);
                        int start = 0;
                        int end = 0;

                        foreach (XmlNode group in groupList)
                        {
                            string a = group.Attributes["表:起始"].Value.ToString();
                            string b = group.Attributes["表:终止"].Value.ToString();
                            start = int.Parse(a);
                            end = int.Parse(b);


                            for (int temp = start - 1; temp < end; temp++)
                                cols[temp] = int.Parse(cols[temp].ToString()) + 1;
                        }



                        bool has = false;
                        for (int k = 0; k < max; k++)
                        {
                            has = false;
                            //sw.WriteLine(rows[k].ToString());
                            XmlAttribute levelAttrCols = xdoc.CreateAttribute("level");
                            levelAttrCols.Value = cols[k].ToString();

                            //foreach (XmlNode row in rowsList)
                            //{
                            //if (rowsList[k].Attributes["表:行号"].Value == k.ToString())
                            //levelAttr.AppendChild(rowsList[k]);
                            // rowsList[k].Attributes.Append(levelAttr);

                            // }

                            foreach (XmlNode col in colsList)
                            {

                                if (int.Parse(col.Attributes["表:列号"].Value.ToString()) == k + 1)
                                {
                                    has = true;

                                    col.Attributes.Append(levelAttrCols);

                                }
                            }
                            if (!has)
                            {

                                XmlElement colEle = xdoc.CreateElement("表:列", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
                                XmlAttribute attr_locId = xdoc.CreateAttribute("uof:locID", "http://schemas.uof.org/cn/2003/uof");
                                attr_locId.Value = "s0048";
                                XmlAttribute attr_attrList = xdoc.CreateAttribute("uof:attrList", "http://schemas.uof.org/cn/2003/uof");
                                attr_attrList.Value = "列号 隐藏 列宽 式样引用 跨度 是否自适应列宽";
                                XmlAttribute attr_colNum = xdoc.CreateAttribute("表:列号", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
                                attr_colNum.Value = (k + 1).ToString();
                                XmlAttribute attr_colWidth = xdoc.CreateAttribute("表:列宽", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
                                attr_colWidth.Value = "54.0";
                                XmlAttribute attr_ref = xdoc.CreateAttribute("表:式样引用", "http://schemas.uof.org/cn/2003/uof-spreadsheet");
                                attr_ref.Value = "CELLSTYLE_0";

                                colEle.Attributes.Append(attr_locId);
                                colEle.Attributes.Append(attr_attrList);
                                colEle.Attributes.Append(attr_colNum);
                                colEle.Attributes.Append(attr_colWidth);
                                colEle.Attributes.Append(attr_ref);
                                colEle.Attributes.Append(levelAttrCols);

                                sheetCon.AppendChild(colEle);


                            }
                        }

                        cols.Clear();
                    }
                    catch (Exception ex)
                    {
                        logger.Warn(ex.Message);
                    }
                }
            }


            //sw.Close();
            //fs2.Close();


            FileStream fs = null;


            try
            {
                fs = new FileStream(destFileName, FileMode.Create);
                xdoc.Save(fs);
            }
            catch (Exception ex)
            {
                logger.Error(ex.Message);
                throw new Exception("Fail in adding condition order");
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                }

            }

        }


        #endregion


        #region RC转换成Normal形式


        private static void RCToNormal(string inputFile, string outputFile)
        {
            XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFile);
            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("uof", TranslatorConstants.XMLNS_UOF);
            nm.AddNamespace("规则", TranslatorConstants.XMLNS_UOFRULES);
            nm.AddNamespace("表", TranslatorConstants.XMLNS_UOFSPREADSHEET);

            XmlTextWriter tw = new XmlTextWriter(outputFile, Encoding.UTF8);
            string refMode = xdoc.SelectSingleNode("uof:UOF/规则:公用处理规则_B665/规则:电子表格_B66C/规则:是否RC引用_B634", nm).InnerText;

            // @refMode="R1C1"
            if (refMode == "true")
            {
                XmlNodeList worksheets = xdoc.SelectNodes("uof:UOF/表:电子表格文档_E826/表:工作表集/表:单工作表", nm);
                foreach (XmlNode worksheet in worksheets)
                {
                    XmlNodeList rows = worksheet.SelectNodes("表:工作表_E825/表:工作表内容_E80E/表:行_E7F1", nm);
                    foreach (XmlNode row in rows)
                    {
                        XmlNode cf = row.SelectSingleNode("表:单元格_E7F2[表:数据_E7B3/表:公式_E7B5]", nm);
                        if (cf != null)
                        {
                            string cV = cf.Attributes["列号_E7BC"].Value;
                            string rV = row.Attributes["行号_E7F3"].Value;
                            string formula = cf.SelectSingleNode("表:数据_E7B3/表:公式_E7B5", nm).InnerText;
                            cf.SelectSingleNode("表:数据_E7B3/表:公式_E7B5", nm).InnerText = CalNormalToRC(formula, Convert.ToInt32(rV), Convert.ToInt32(cV));
                        }
                    }

                }

            }
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
        private static string CalNormalToRC(string formula, int row, int col)
        {

            Regex range = new Regex(@"^[R]");
            //string formula = "DAVERAGE(A5:F22,\"语文\",B26:B27)";
            string[] ranges = formula.Split(new char[] { ',' });

            if (ranges.Length == 1)
            {
                //string range1 = ranges[0].Substring(ranges[0].IndexOf("(") + 1, ranges[0].Length - ranges[0].IndexOf("(") - 2);
                if (ranges[0].Contains(":"))
                {
                    string range1 = ranges[0].Substring(ranges[0].IndexOf("(") + 1, ranges[0].Length - ranges[0].IndexOf("(") - 2);
                    string[] rc1 = range1.Split(new char[] { ':' });

                    if (range.IsMatch(range1))
                    {
                        string r1 = RevCalRC(rc1[0], row, col);
                        string r2 = RevCalRC(rc1[1], row, col);
                        formula = formula.Replace(range1, r1 + ":" + r2);
                    }
                }
                
            }
            else if (ranges.Length > 2)
            {
                string range1 = ranges[0].Substring(ranges[0].IndexOf("(") + 1);
                string[] rc1 = range1.Split(new char[] { ':' });

                if (range.IsMatch(range1))
                {
                    string r1 = RevCalRC(rc1[0], row, col);
                    string r2 = RevCalRC(rc1[1], row, col);
                    formula = formula.Replace(range1, r1 + ":" + r2);
                }



                for (int i = 1; i < ranges.Length - 1; i++)
                {
                    if (range.IsMatch(ranges[i]))
                    {
                        string[] rc = ranges[i].Split(new char[] { ':' });
                        string r1 = RevCalRC(rc[0], row, col);
                        string r2 = RevCalRC(rc[1], row, col);
                        formula = formula.Replace(range1, r1 + ":" + r2);
                    }
                }

                string rangeLast = ranges[ranges.Length - 1].Substring(0, ranges[ranges.Length - 1].Length - 1);
                string[] rcl = rangeLast.Split(new char[] { ':' });
                string rl1 = RevCalRC(rcl[0], row, col);
                string rl2 = RevCalRC(rcl[1], row, col);
                formula = formula.Replace(rangeLast, rl1 + ":" + rl2);
            }
            return formula;
        }


        /// <summary>
        /// 把RC表示转换成A1：B1的形式
        /// </summary>
        /// <param name="rc">RC字符串</param>
        /// <param name="row">行号</param>
        /// <param name="col">列号</param>
        /// <returns>正常形式</returns>
        public static string RevCalRC(string rc, int row, int col)
        {
            int r = 0;
            int c = 0;
            string[] rcStr = rc.Split(new char[] { 'C' });

            // R[]的形式，而不是R
            string rStr = rcStr[0];
            if (rStr.Length > 1)
            {
                string tmp = rStr.Substring(rStr.IndexOf('[') + 1, rStr.IndexOf(']') - rStr.IndexOf('[') - 1);
                r = Convert.ToInt32(tmp) + row;
            }
            else
            {

                r = row;
            }

            string cStr = rcStr[1];
            if (cStr.Length > 1)
            {
                string tmp = cStr.Substring(cStr.IndexOf('[') + 1, cStr.IndexOf(']') - cStr.IndexOf('[') - 1);
                c = Convert.ToInt32(tmp) + col;
            }
            else
            {
                c = col;
            }

            return ConvertNumToLetter(c) + r;
        }


        /// <summary>
        ///  把数字转换成对应的字母
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        private static char ConvertNumToAlph(int num)
        {
            //64：表示数字大小
            return Convert.ToChar(num + 64);
        }


        /// <summary>
        ///  把数字转换成对应的字母
        /// </summary>
        /// <param name="num"></param>
        /// <returns></returns>
        public static string ConvertNumToLetter(int num)
        {
            string result = string.Empty;

            double l = Math.Log(num, 26);
            int length = Convert.ToInt32(Math.Floor(l));

            for (int i = length; i >= 0; i--)
            {
                int dgt = 0;
                for (int k = 1; k < 26; k++)
                {
                    double tmp = Math.Pow(26, i) * k;
                    if (tmp > num)
                    {
                        dgt = k - 1;
                        break;
                    }
                }
                result += ConvertNumToAlph(dgt);
                double co = Math.Pow(26, i) * dgt;
                num -= Convert.ToInt32(co);
            }

            return result;
        }




        #endregion
    }
}
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
using Act.UofTranslator.UofZipUtils;
using System.IO;
using System.Xml.XPath;
using System.Xml.Xsl;
using log4net;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// The first step pre process
    /// </summary>
    /// <author>fangchunyan</author>
    class OoxToUofPreProcessorOneWord : AbstractProcessor
    {
        private static ILog logger = LogManager.GetLogger(typeof(OoxToUofPreProcessorOneWord).FullName);

        public override bool transform()
        {
            //string Document_xml = "word/document.xml";//word�µ�doucument.xml�ĵ�
            string picture_xml="word\\media";
           // string tmpFile = Path.GetTempPath() + "tmpDoc.xml";
           // string tmpFile2 = Path.GetTempPath() + "sndtmpDoc.xml";
           // Guid gPath = Guid.NewGuid();
           // string wordPrePath = Path.GetTempPath() + gPath.ToString() + "\\";
         //   if (!Directory.Exists(wordPrePath))
           //     Directory.CreateDirectory(wordPrePath);
            string wordPrePath = Path.GetDirectoryName(outputFile)+"\\";//����ĵ�·��
            string mediaPath = wordPrePath + picture_xml;

            string tblVtcl=wordPrePath+"tblVertical.xml";
            string tmpFile = wordPrePath + "tmpDoc.xml";
            string tmpFile2 = wordPrePath + "sndtmpDoc.xml";//Ԥ�����е��м��ĵ�
            string tmpFile3 = wordPrePath + "thrtmpDoc.xml";//Ԥ�����е��м��ĵ�
            string preOutputFileName = wordPrePath + "tempdoc.xml";//Ԥ����������м��ĵ�
           
            XmlReader source = null;
            XmlWriter writer = null;
            ZipReader archive = ZipFactory.OpenArchive(inputFile);
            archive.ExtractOfficeDocument(inputFile, wordPrePath);
            //archive.ExtractUOFDocument(inputFile, wordPrePath);
            bool isSuccess = true;
            try
            {
                OoxToUofTableProcessing tableProcessing2 = new OoxToUofTableProcessing(wordPrePath+"word"+Path.AltDirectorySeparatorChar+"document.xml");
                tableProcessing2.Processing(tblVtcl);

                //�ҵ�Ԥ�����һ����ʽ����pretreatmentStep1.xsl
                XPathDocument xpdoc = UOFTranslator.GetXPathDoc(TranslatorConstants.OOXToUOF_PRETREAT_STEP1_XSL, TranslatorConstants.OOXToUOF_WORD_LOCATION);
                XslCompiledTransform xslt = new XslCompiledTransform();
                xslt.Load(xpdoc);

               // source = XmlReader.Create(archive.GetEntry(Document_xml));//��ѹ���õ�document.xml�ĵ�
                source = XmlReader.Create(tblVtcl);
                writer = new XmlTextWriter(tmpFile, Encoding.UTF8);//tmpDoc.xml          
                xslt.Transform(source, writer);//document.xml--����Ԥ����.xsl--tmpDoc.xml
                if (writer != null) writer.Close();
                if (source != null) source.Close();

                //�ڶ���Ԥ����
                xpdoc = UOFTranslator.GetXPathDoc(TranslatorConstants.OOXToUOF_PRETREAT_STEP2_XSL, TranslatorConstants.OOXToUOF_WORD_LOCATION);
                xslt = new XslCompiledTransform();
                xslt.Load(xpdoc);
                writer = new XmlTextWriter(tmpFile2, Encoding.UTF8);//sndtemDoc.xml
                xslt.Transform(tmpFile, writer);//tmpDoc--����Ԥ����2.xsl--sndtemDoc.xml
                if (writer != null) writer.Close();

                //������Ԥ����
                xpdoc = UOFTranslator.GetXPathDoc(TranslatorConstants.OOXToUOF_PRETREAT_STEP3_XSL, TranslatorConstants.OOXToUOF_WORD_LOCATION);
                xslt = new XslCompiledTransform();
                xslt.Load(xpdoc);
                // xslt.Transform(tmpFile2, outputFile);
                xslt.Transform(tmpFile2, tmpFile3);//����Ԥ����3--tempdoc.xml
                
                //2011/6/17zhaobj����ӰԤ����
                XmlNamespaceManager nms = null;
                XmlDocument xmlDoc;
                XmlTextWriter resultWriter = null;
                xmlDoc = new XmlDocument();
                xmlDoc.Load(tmpFile3);

                nms = new XmlNamespaceManager(xmlDoc.NameTable);
                nms.AddNamespace("a",TranslatorConstants.XMLNS_A);
                nms.AddNamespace("w", TranslatorConstants.XMLNS_W);

                XmlNodeList shapeShadeList = xmlDoc.SelectNodes("//a:outerShdw", nms);                   
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
                
              
                resultWriter = new XmlTextWriter(preOutputFileName, System.Text.Encoding.UTF8);
                xmlDoc.Save(resultWriter);
                resultWriter.Close();
                 
                /*   
                 
                 
                 
                 */

                OutputFilename = preOutputFileName;

                //ͼƬԤ����
                //if (Directory.Exists(wordPrePath + picture_xml))
                //{
                //    xmlDoc.Load(preOutputFileName);
                //    DirectoryInfo mediaInfo = new DirectoryInfo(wordPrePath + picture_xml);
                //    FileInfo[] medias = mediaInfo.GetFiles();
                //    XmlNode root = xmlDoc.SelectSingleNode("w:document", nms);
                //    XmlElement mediaNode=xmlDoc.CreateElement("w","media",TranslatorConstants.XMLNS_W);
                //    foreach (FileInfo media in medias)
                //    {
                //        XmlElement mediaFileNode = xmlDoc.CreateElement("u2opic", "picture", "urn:u2opic:xmlns:post-processings:special");
                //        mediaFileNode.SetAttribute("target", "urn:u2opic:xmlns:post-processings:special", media.FullName);
                //        mediaNode.AppendChild(mediaFileNode);
                //    }
                //    root.AppendChild(mediaNode);
                //    xmlDoc.Save(tmpPic);
                //    OutputFilename = tmpPic;
                //}

                
            }
            catch (Exception e)
            {
                logger.Error("Fail in OoxToUofPreProcessorOneWord: " + e.Message);
                logger.Error(e.StackTrace);
                isSuccess = false;
            }
            finally
            {
                if (writer != null) 
                    writer.Close();
                if (source != null) 
                    source.Close();
                if (archive != null)
                    archive.Close();
                if (File.Exists(tmpFile2))
                    File.Delete((tmpFile2));
            }
            return isSuccess;
        }

               
    }
    
   
    
    public class OoxToUofTableProcessing : TableProcessing
    {

        private readonly string wordNamespace = @"http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        public OoxToUofTableProcessing(string fileName)
            : base(fileName)
        {

        }

        public override void Processing(string outputFile)
        {
            XmlNodeList tableList = this.GetTableList(@"//w:tbl");

            //ѭ�����ڵ�
            for (int i = 0; i < tableList.Count; i++)
            {
                XmlNodeList rowList = this.GetRowList(tableList[i], @"w:tr");

                DoubleLinkedList<XmlNode>[] cellMatrix = new DoubleLinkedList<XmlNode>[rowList.Count];
                for (int j = 0; j < cellMatrix.Length; j++) cellMatrix[j] = new DoubleLinkedList<XmlNode>();

                //����ÿһ����
                this.RowProcessing(tableList[i], cellMatrix);
            }

            XmlSource.Save(outputFile);
        }

        //����ÿһ����   
        private void RowProcessing(XmlNode tableNode, DoubleLinkedList<XmlNode>[] cellMatrix)
        {
            XmlNodeList rowList = this.GetRowList(tableNode, @"w:tr");

            //����ÿһ����Ԫ��
            for (int i = 0; i < rowList.Count; i++)
            {
                this.CellProcessing(rowList[i], i, cellMatrix);
            }

            //���д���            
            CellVerticalSpanProcessing(cellMatrix);

            //ɾ���½��ĵ�Ԫ��
            this.DeleteNewCreatedCell(cellMatrix);
        }


        //����һ����Ԫ��: 1,���ݿ��и������ѵ�Ԫ��. 2,����Ԫ����Ϣ���Ƶ�������
        private void CellProcessing(XmlNode rowNode, int rowIndex, DoubleLinkedList<XmlNode>[] cellMatrix)
        {
            XmlNodeList cellList = this.GetCellList(rowNode, @"w:tc");

            //������У����ݿ��и������ѵ�Ԫ��
            foreach (XmlNode cellNode in cellList)
            {
                DivideCell(rowNode, cellNode);
            }

            //ˢ��
            cellList = this.GetCellList(rowNode, @"w:tc");
            //����Ԫ����Ϣ���Ƶ�������
            CopyCell(cellList, cellMatrix, rowIndex);
        }

        //����Ԫ����Ϣ���Ƶ�������
        private void CopyCell(XmlNodeList cellList, DoubleLinkedList<XmlNode>[] cellMatrix, int rowIndex)
        {
            foreach (XmlNode cell in cellList)
            {
                cellMatrix[rowIndex].Append(cell);
            }
        }

        //������У����ݿ��и������ѵ�Ԫ��
        private void DivideCell(XmlNode rowNode, XmlNode cellNode)
        {
            XmlNode hMerge = this.GetSingleNode(cellNode, @"w:tcPr/w:gridSpan");

            if (hMerge == null) return;

            int hMergeCount = hMerge.Attributes["w:val"] == null ? 0 : Convert.ToInt32(hMerge.Attributes["w:val"].Value);

            for (int i = 1; i < hMergeCount; i++)
            {
                XmlNode newCell = cellNode.CloneNode(true);

                newCell.AppendChild(this.CreatNewNode("w:newCell", wordNamespace));

                //���ѵ�Ԫ��
                rowNode.InsertAfter(newCell, cellNode);
            }
        }

        //���д������ݿ��и����ҵ���ʼ������Ŀ���λ��
        private void CellVerticalSpanProcessing(DoubleLinkedList<XmlNode>[] cellMatrix)
        {
            for (int rowIndex = 0; rowIndex < cellMatrix.Length; rowIndex++)
            {
                List<DoubleLinkedList<XmlNode>.DoubleLinkedNode> cellList = cellMatrix[rowIndex].ToArray();

                //������У����ݿ��и����ҵ���ʼ������Ŀ���λ��
                for (int columnIndex = 0; columnIndex < cellList.Count; columnIndex++)
                {
                    XmlNode cellNode = cellList[columnIndex].Value;
                    XmlNode vMerge = this.GetSingleNode(cellNode, @"w:tcPr/w:vMerge[@w:val='restart']");
                    if (vMerge == null) continue;

                    int currentRowIndex = rowIndex + 1;
                    //������
                    int rowSpan = 1;
                    while (currentRowIndex < cellMatrix.Length)//�ӵ�ǰ������Ѱ����ͬ�еĽ��
                    {
                        XmlNode node = cellMatrix[currentRowIndex].ToArray()[columnIndex].Value;
                        XmlNode verticalMerge = this.GetSingleNode(node, @"w:tcPr/w:vMerge");

                        //�µĺϲ�
                        if (verticalMerge == null || (verticalMerge.Attributes[@"w:val"] != null && verticalMerge.Attributes[@"w:val"].Value == @"restart"))
                        {
                            break;
                        }
                        rowSpan++;
                        currentRowIndex++;
                    }

                    //��Ҫ����һ�Ͻڵ��м�����ʼ���еĿ�ʼ�У������У���������
                    XmlAttribute rowSpanCount = this.CreateNewAttribute("w:rowSpanCount");
                    rowSpanCount.Value = rowSpan.ToString();
                    vMerge.Attributes.Append(rowSpanCount);
                }
            }
        }

        //ɾ���½��ĵ�Ԫ��
        private void DeleteNewCreatedCell(DoubleLinkedList<XmlNode>[] cellMatrix)
        {
            for (int rowIndex = 0; rowIndex < cellMatrix.Length; rowIndex++)
            {
                List<DoubleLinkedList<XmlNode>.DoubleLinkedNode> cellList = cellMatrix[rowIndex].ToArray();
                //�Ӻ���ǰ��ɾ��
                for (int cellIndex = cellList.Count - 1; cellIndex >= 0; cellIndex--)
                {
                    XmlNode cell = cellList[cellIndex].Value;
                    XmlNode newCell = this.GetSingleNode(cell, @".//w:newCell");
                    if (newCell != null)
                    {
                        cell.ParentNode.RemoveChild(cell);
                    }

                }
            }
        }

        #region override methods

        protected override XmlNamespaceManager GetXmlNamespaceManager(XmlNameTable nameTable)
        {
            XmlNamespaceManager nm = new XmlNamespaceManager(nameTable);
            nm.AddNamespace("w", wordNamespace);
            return nm;
        }

        #endregion
    }
}

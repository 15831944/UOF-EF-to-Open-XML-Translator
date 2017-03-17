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
using System.Xml.XPath;
using log4net;
using System.Xml.Xsl;
using System.Reflection;
using System.IO;
using System.Collections;
using System.Resources;
using System.Globalization;
using System.Xml;
using Act.UofTranslator.UofZipUtils;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// This is a abstract class base Translator, it provides common functions.
    /// </summary>
    /// <author>linwei</author>
    /// <modifier>linyaohu</modifier>
    public abstract class UOFTranslator : IUOFTranslator
    {

        private event EventHandler progressMessageIntercepted;

        private event EventHandler feedbackMessageIntercepted;

        protected LinkedList<IProcessor> uof2ooxPreProcessors;

        protected LinkedList<IProcessor> oox2uofPreProcessors;

        protected LinkedList<IProcessor> oox2uofPostProcessors;

        private static ILog logger = LogManager.GetLogger(typeof(UOFTranslator).FullName);

        private Hashtable messageTypes = new Hashtable();

        private ResourceManager resources;

        private string promptMsg;

        protected string mainOutput;

        public static string ASSEMBLY_PATH = Assembly.GetExecutingAssembly().Location.Substring(0,
            Assembly.GetExecutingAssembly().Location.LastIndexOf(Path.DirectorySeparatorChar) + 1);

        protected Guid guid;

        protected virtual void InitalPreProcessors()
        {
            //AccessControl(System.IO.Path.GetDirectoryName(this.GetType().Assembly.Location) + @"\conf\config.xml");
            uof2ooxPreProcessors = new LinkedList<IProcessor>();
            oox2uofPreProcessors = new LinkedList<IProcessor>();

            oox2uofPostProcessors = new LinkedList<IProcessor>();

        }

        protected abstract void DoUofToOoxMainTransform(string inputFile, string outputFile, string resourceDir);

        protected abstract void DoOoxToUofMainTransform(string originalFile, string inputFile, string outputFile, string resourceDir);

        private void DoOoxToUofTransform(string inputFile, string outputFile, string resourceDir)
        {

            guid = Guid.NewGuid();//新的文件夹
            if (!IsOox(inputFile))
            {
                throw new NotAnOoxDocumentException(inputFile + " 不是OOX文件!");
            }

            LinkedList<IProcessor>.Enumerator enumer = oox2uofPreProcessors.GetEnumerator();//双链表：枚举出预处理次数
            string input = inputFile;
            string outputPath = Path.GetTempPath().ToString() + guid.ToString();
            string output = outputPath + Path.DirectorySeparatorChar + "tmpDoc";
            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }
            int i = 0;
            while (enumer.MoveNext())//预处理
            {
                i++;
                enumer.Current.OriginalFilename = inputFile;
                enumer.Current.InputFilename = input;
                enumer.Current.OutputFilename = output + i + ".xml";
                enumer.Current.transform();//预处理转换
                input = enumer.Current.OutputFilename;//预处理得到的文档
            }



            DoOoxToUofMainTransform(inputFile, enumer.Current.OutputFilename, outputFile, resourceDir);//主转换
            string mainOutputOriginalFile = enumer.Current.OutputFilename;//主转换得到的文档
            enumer = oox2uofPostProcessors.GetEnumerator();//后处理次数

            while (enumer.MoveNext())//后处理
            {
                enumer.Current.OriginalFilename = mainOutputOriginalFile;
                enumer.Current.InputFilename = mainOutput;
                enumer.Current.OutputFilename = outputFile;
                enumer.Current.transform();//后处理转换
            }
            // Directory.Delete(outputPath, true);
        }

        private void DoUofToOoxTransform(string inputFile, string outputFile, string resourceDir)
        {
            guid = Guid.NewGuid();
            if (!IsUof(inputFile))
            {
                throw new NotAnUofDocumentException(inputFile + " 不是UOF文件");
            }

            LinkedList<IProcessor>.Enumerator enumer = uof2ooxPreProcessors.GetEnumerator();
            string outputPath = Path.GetTempPath().ToString() + guid.ToString();
            string output = outputPath + Path.DirectorySeparatorChar + "tmpDoc";
            if (!Directory.Exists(outputPath))
                Directory.CreateDirectory(outputPath);
            //  string input = EIPicturePre(inputFile, outputPath);

            ZipReader archive = ZipFactory.OpenArchive(inputFile);
            archive.ExtractUOFDocument(inputFile, outputPath);

            string input = inputFile;

            int i = 0;
            while (enumer.MoveNext())
            {
                i++;
                enumer.Current.OriginalFilename = inputFile;
                enumer.Current.InputFilename = input;
                enumer.Current.OutputFilename = output + i + ".xml";
                enumer.Current.transform();
                input = enumer.Current.OutputFilename;
            }

            DoUofToOoxMainTransform(enumer.Current.OutputFilename, outputFile, resourceDir);
            //  Directory.Delete(outputPath, true);
        }

        #region util functions

        protected bool IsUof(string fileName)
        {
            // TODO: implement
            return true;
        }

        protected bool IsOox(string fileName)
        {
            // TODO: implement
            return true;
        }

        public static XPathDocument GetXPathDoc(string xslname, string location)
        {
            return new XPathDocument(Assembly.GetExecutingAssembly().GetManifestResourceStream(typeof(UOFTranslator).Namespace
                + "." + TranslatorConstants.RESOURCE_LOCATION + "." + location + "." + xslname));
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
                logger.Warn(e.Message);
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

        /// <summary>
        ///  main transform which needs the orginal File
        /// </summary>
        /// <param name="directionXSL">transform direction</param>
        /// <param name="resourceResolver">xsl location</param>
        /// <param name="originalFile">original File</param>
        /// <param name="inputFile">File after pretreatment</param>
        /// <param name="outputFile">output file</param>
        protected void MainTransform(string directionXSL, XmlUrlResolver resourceResolver, string originalFile, string inputFile, string outputFile)
        {

            XPathDocument xslDoc;
            XmlReaderSettings xrs = new XmlReaderSettings();
            XmlReader source = null;
            XmlWriter writer = null;
            OoxZipResolver zipResolver = null;
            string zipXMLFileName = "input.xml";

            try
            {
                //xrs.ProhibitDtd = true;

                xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(directionXSL));
                xrs.XmlResolver = resourceResolver;
                string sr = ZipXMLFile(inputFile);
                ZipReader archive = ZipFactory.OpenArchive(sr);
                source = XmlReader.Create(archive.GetEntry(zipXMLFileName));

                XslCompiledTransform xslt = new XslCompiledTransform();
                XsltSettings settings = new XsltSettings(true, false);
                xslt.Load(xslDoc, settings, resourceResolver);

                if (!originalFile.Equals(string.Empty))
                {
                    zipResolver = new OoxZipResolver(originalFile, resourceResolver);
                }
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

        /// <summary>
        /// zip the big xml file
        /// </summary>
        /// <param name="inputFile">input xml file</param>
        /// <returns>zip file</returns>
        protected static string ZipXMLFile(string inputFile)
        {
            string dir = Path.GetDirectoryName(inputFile);
            string zipXMLFile = dir + Path.AltDirectorySeparatorChar + "toMainSheet.zip";
            ZlibZipWriter zw = new ZlibZipWriter(zipXMLFile);
            zw.AddEntry("input.xml");

            FileInfo fi = new FileInfo(inputFile);
            StreamReader sr = new StreamReader(inputFile);

            while (!sr.EndOfStream)
            {
                byte[] tmpbuffer = Encoding.UTF8.GetBytes(sr.ReadLine());
                zw.Write(tmpbuffer, 0, tmpbuffer.Length);
            }
            zw.Flush();
            zw.Close();
            return zipXMLFile;
        }
        /// <summary>
        ///  pretreatment of picture
        /// </summary>
        /// <param name="xmlDoc">input file stream</param>
        /// <param name="fireNodeName">first node</param>
        /// <param name="picPath">picture location</param>
        /// <param name="nms">xml namespace manager</param>
        /// <returns>result stream</returns>
        protected XmlDocument PicPretreatment(XmlDocument xmlDoc, string fireNodeName, string picPath, XmlNamespaceManager nms)
        {
            DirectoryInfo mediaInfo = new DirectoryInfo(picPath);
            FileInfo[] medias = mediaInfo.GetFiles();
            XmlNode root = xmlDoc.SelectSingleNode(fireNodeName, nms);
            XmlElement mediaNode = xmlDoc.CreateElement("w", "media", TranslatorConstants.XMLNS_W);
            foreach (FileInfo media in medias)
            {
                XmlElement mediaFileNode = xmlDoc.CreateElement("u2opic", "picture", "urn:u2opic:xmlns:post-processings:special");
                mediaFileNode.SetAttribute("target", "urn:u2opic:xmlns:post-processings:special", media.FullName);
                mediaNode.AppendChild(mediaFileNode);
            }
            root.AppendChild(mediaNode);
            return xmlDoc;
        }
        //ole对象lyy
        protected XmlDocument OlePretreatment(XmlDocument xmlDoc, string fireNodeName, string olePath, XmlNamespaceManager nms)
        {
            DirectoryInfo mediaInfo = new DirectoryInfo(olePath + "ppt\\embeddings");
            FileInfo[] medias = mediaInfo.GetFiles();
            XmlNode root = xmlDoc.SelectSingleNode(fireNodeName, nms);
            XmlElement mediaNode = xmlDoc.CreateElement("w", "ole", TranslatorConstants.XMLNS_W);
            foreach (FileInfo media in medias)
            {
                XmlElement mediaFileNode = xmlDoc.CreateElement("u2oole", "embeding", "urn:u2oole:xmlns:post-processings:special");
                mediaFileNode.SetAttribute("target", "urn:u2oole:xmlns:post-processings:special", media.FullName);
                mediaNode.AppendChild(mediaFileNode);
            }
            if (Directory.Exists(olePath + "ppt\\drawings"))
            {
                mediaInfo = new DirectoryInfo(olePath + "ppt\\drawings");
                medias = mediaInfo.GetFiles();
                foreach (FileInfo media in medias)
                {
                    XmlElement mediaFileNode = xmlDoc.CreateElement("u2oole", "drawing", "urn:u2oole:xmlns:post-processings:special");
                    mediaFileNode.SetAttribute("target", "urn:u2oole:xmlns:post-processings:special", media.FullName);
                    mediaNode.AppendChild(mediaFileNode);
                }
            }
            if (Directory.Exists(olePath + "ppt\\drawings\\_rels"))
            {
                mediaInfo = new DirectoryInfo(olePath + "ppt\\drawings\\_rels");
                medias = mediaInfo.GetFiles();
                foreach (FileInfo media in medias)
                {
                    XmlElement mediaFileNode = xmlDoc.CreateElement("u2oole", "drawingrel", "urn:u2oole:xmlns:post-processings:special");
                    mediaFileNode.SetAttribute("target", "urn:u2oole:xmlns:post-processings:special", media.FullName);
                    mediaNode.AppendChild(mediaFileNode);
                }
            }
            root.AppendChild(mediaNode);
            return xmlDoc;
        }
        /// <summary>
        /// pretreatment of custom xml data,OOXM to UOF
        /// </summary>
        /// <param name="xmlDoc">input file stream</param>
        /// <param name="firstNodeName">first node</param>
        /// <param name="nms">name space manager</param>
        /// <returns>result stream</returns>
        protected XmlDocument CustomXMPretreatment(XmlDocument xmlDoc, string firstNodeName, string customXMLPath,XmlNamespaceManager nms)
        {
            DirectoryInfo customDirInfo = new DirectoryInfo(customXMLPath);
            DirectoryInfo customDirRelInfo=new DirectoryInfo(customXMLPath+Path.AltDirectorySeparatorChar+"_rels");
            FileInfo[] items = customDirInfo.GetFiles();
            FileInfo[] rels = customDirRelInfo.GetFiles();

            XmlNode root = xmlDoc.SelectSingleNode(firstNodeName, nms);
            XmlElement customXMLNode = xmlDoc.CreateElement("w", "customXML", TranslatorConstants.XMLNS_W);

            // manage the items
            foreach (FileInfo item in items)
            {
                XmlElement itemFileNode = xmlDoc.CreateElement("w","customItem", TranslatorConstants.XMLNS_W);
                itemFileNode.SetAttribute("fileName", "customXml/"+item.Name);
                XmlDocument itemDoc = new XmlDocument();
                itemDoc.Load(item.FullName);
               
                XmlNode content = xmlDoc.ImportNode(itemDoc.DocumentElement, true);
                itemFileNode.AppendChild(content);
                customXMLNode.AppendChild(itemFileNode);
            }

            // manage the rels
           // XmlElement relsNode = xmlDoc.CreateElement("w", "_rels", TranslatorConstants.XMLNS_W);
            foreach (FileInfo rel in rels)
            {
               // XmlElement relNode = xmlDoc.CreateElement("w", "_rel", TranslatorConstants.XMLNS_W);
                XmlElement relNode = xmlDoc.CreateElement("w", "customItem", TranslatorConstants.XMLNS_W);
                relNode.SetAttribute("fileName","customXml/_rels/"+ rel.Name);
                XmlDocument relDoc = new XmlDocument();
                relDoc.Load(rel.FullName);

                XmlNode content = xmlDoc.ImportNode(relDoc.DocumentElement, true);
                relNode.AppendChild(content);
                //relsNode.AppendChild(relNode);
                customXMLNode.AppendChild(relNode);
            }
            //customXMLNode.AppendChild(relsNode);
            root.AppendChild(customXMLNode);

            return xmlDoc;
        }

        //     /// <summary>
        //     ///  get the entry of xsl file
        //     /// </summary>
        //     /// <param name="xslLocation">xsl location</param>
        //     /// <returns>xml url resolver</returns>
        //     protected XmlUrlResolver XSLResourceResolver(string xslLocation)
        //     {
        //         return new ResourceResolver(Assembly.GetExecutingAssembly(),
        //this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + xslLocation);

        //     }

        #endregion

        #region IUOFTranslator Members

        public void AddProgressMessageListener(EventHandler listener)
        {
            progressMessageIntercepted += listener;
        }

        public void AddFeedbackMessageListener(EventHandler listener)
        {
            feedbackMessageIntercepted += listener;
        }

        public void UofToOox(string inputFile, string outputFile)
        {
            DoUofToOoxTransform(inputFile, outputFile, null);
        }

        public void UofToOoxWithExternalResources(string inputFile, string outputFile, string resourceDir)
        {
            DoUofToOoxTransform(inputFile, outputFile, resourceDir);
        }

        public void UofToOoxComputeSize(string inputFile)
        {
            DoUofToOoxTransform(inputFile, null, null);
        }

        public void OoxToUof(string inputFile, string outputFile)
        {
            DoOoxToUofTransform(inputFile, outputFile, null);
        }

        public void OoxToUofWithExternalResources(string inputFile, string outputFile, string resourceDir)
        {
            DoOoxToUofTransform(inputFile, outputFile, resourceDir);
        }

        public void OoxToUofComputeSize(string inputFile)
        {
            DoOoxToUofTransform(inputFile, null, null);
        }


        #endregion

        /*
        public void AccessControl(string fileName)
        {
            try
            {

                // string configFilePath = System.IO.Path.GetDirectoryName(this.GetType().Assembly.Location) + @"\conf\config.xml";

                System.Security.AccessControl.FileSecurity fileSecurity = System.IO.File.GetAccessControl(fileName);

                fileSecurity.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule("Everyone",

                    System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow));

                System.IO.File.SetAccessControl(fileName, fileSecurity);

            }

            catch (Exception ex)
            {
                throw new Exception("fail in set the file's access control of " + fileName);
            }

        }
    */

        #region changeGrpUtil

        Stack<XmlNode> grpStack = new Stack<XmlNode>();
        double _offX = 0;
        double _offY = 0;
        double _extX = 0;
        double _extY = 0;
        double _chOffX = 0;
        double _chOffY = 0;
        double _chExtX = 0;
        double _chExtY = 0;
        protected void GrpShPre(XmlDocument xdoc, XmlNamespaceManager nm, string prefix, FileStream resultFile)
        {

            // powerpoint
            if (prefix == "p")
            {
                // grpStack.Clear();
                XmlNodeList slds = xdoc.SelectNodes("//p:sld |//p:sldMaster | //p:sldLayout | //p:notesMaster |//p:handoutMaster | //p:notes", nm);
                foreach (XmlNode sld in slds)
                {
                    XmlNodeList grpSps = sld.SelectNodes("p:cSld/p:spTree/p:grpSp", nm);
                    foreach (XmlNode grpSp in grpSps)
                    {
                        grpSpTransform(grpSp, nm, prefix);
                    }
                }
            }

            // spreadsheet
            else if (prefix == "xdr")
            {
                XmlNodeList spreadSheets = xdoc.SelectNodes("//ws:spreadsheet", nm);
                foreach (XmlNode sheet in spreadSheets)
                {
                    XmlNodeList twoCellAnchor = sheet.SelectNodes("ws:Drawings/xdr:wsDr/xdr:twoCellAnchor[xdr:grpSp]", nm);
                    foreach (XmlNode node in twoCellAnchor)
                    {
                        grpSpTransform(node.SelectSingleNode("xdr:grpSp", nm), nm, prefix);
                    }
                }
            }
            xdoc.Save(resultFile);
        }

        private XmlNode grpSpTransform(XmlNode grpSp, XmlNamespaceManager nm, string prefix)
        {

            foreach (XmlNode gs in grpSp.ChildNodes)
            {


                // bool flag = false;
                if (gs.LocalName == "grpSpPr")
                {
                    // 初始栈
                    if (grpStack.Count == 0)
                    {
                        grpStack.Push(gs.SelectSingleNode("a:xfrm", nm));
                    }

                        // 转换之后再入栈
                    else
                    {
                        /*
                       
                        if (flag == true)
                        {
                            grpStack.Pop();
                            flag = false;
                        }
                        */
                        XmlNode lastGrpSpPr = grpStack.Peek();
                        GetValues(lastGrpSpPr, nm);
                        /*
                        double oldX = Convert.ToDouble(gs.SelectSingleNode("a:xfrm/a:off", nm).Attributes["x"].Value);
                        double oldY = Convert.ToDouble(gs.SelectSingleNode("a:xfrm/a:off", nm).Attributes["y"].Value);
                        double oldExtX = Convert.ToDouble(gs.SelectSingleNode("a:xfrm/a:ext", nm).Attributes["cx"].Value);
                        double oldExtY = Convert.ToDouble(gs.SelectSingleNode("a:xfrm/a:ext", nm).Attributes["cy"].Value);

                        gs.SelectSingleNode("a:xfrm/a:off", nm).Attributes["x"].Value = (_offX + (oldX - _chOffX) * _extX / _chExtX).ToString();
                        gs.SelectSingleNode("a:xfrm/a:off", nm).Attributes["y"].Value = (_offY + (oldY - _chOffY) * _extY / _chExtY).ToString();
                        gs.SelectSingleNode("a:xfrm/a:ext", nm).Attributes["cx"].Value = (oldExtX * _extX / _chExtX).ToString();
                        gs.SelectSingleNode("a:xfrm/a:ext", nm).Attributes["cy"].Value = (oldExtY * _extY / _chExtY).ToString();
                         * */


                        grpStack.Push(SetValues(gs, nm, "").SelectSingleNode("a:xfrm", nm));
                    }
                }
                else if (gs.LocalName == "grpSp")
                {
                    grpSpTransform(gs, nm, prefix);
                }
                else if (gs.LocalName == "sp" || gs.LocalName == "pic" || gs.LocalName == "cxnSp")
                {
                    // flag = true;


                    double oldX = Convert.ToDouble(gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:off", nm).Attributes["x"].Value);
                    double oldY = Convert.ToDouble(gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:off", nm).Attributes["y"].Value);
                    double oldExtX = Convert.ToDouble(gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:ext", nm).Attributes["cx"].Value);
                    double oldExtY = Convert.ToDouble(gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:ext", nm).Attributes["cy"].Value);
                    XmlNode node = grpStack.Peek();
                    GetValues(node, nm);

                    gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:off", nm).Attributes["x"].Value = (_offX + (oldX - _chOffX) * _extX / _chExtX).ToString();
                    gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:off", nm).Attributes["y"].Value = (_offY + (oldY - _chOffY) * _extY / _chExtY).ToString();
                    gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:ext", nm).Attributes["cx"].Value = (oldExtX * _extX / _chExtX).ToString();
                    gs.SelectSingleNode(prefix + ":spPr/a:xfrm/a:ext", nm).Attributes["cy"].Value = (oldExtY * _extY / _chExtY).ToString();

                    //SetValues(gs, nm, prefix + ":spPr/");
                }
            }

            // pop the stack
            grpStack.Pop();
            return grpSp;
        }

        private void GetValues(XmlNode node, XmlNamespaceManager nm)
        {
            _offX = Convert.ToDouble(node.SelectSingleNode("a:off", nm).Attributes["x"].Value);
            _offY = Convert.ToDouble(node.SelectSingleNode("a:off", nm).Attributes["y"].Value);
            _extX = Convert.ToDouble(node.SelectSingleNode("a:ext", nm).Attributes["cx"].Value);
            _extY = Convert.ToDouble(node.SelectSingleNode("a:ext", nm).Attributes["cy"].Value);
            _chOffX = Convert.ToDouble(node.SelectSingleNode("a:chOff", nm).Attributes["x"].Value);
            _chOffY = Convert.ToDouble(node.SelectSingleNode("a:chOff", nm).Attributes["y"].Value);
            _chExtX = Convert.ToDouble(node.SelectSingleNode("a:chExt", nm).Attributes["cx"].Value);
            _chExtY = Convert.ToDouble(node.SelectSingleNode("a:chExt", nm).Attributes["cy"].Value);

        }

        private XmlNode SetValues(XmlNode node, XmlNamespaceManager nm, string prefix)
        {
            double oldX = Convert.ToDouble(node.SelectSingleNode(prefix + "a:xfrm/a:off", nm).Attributes["x"].Value);
            double oldY = Convert.ToDouble(node.SelectSingleNode(prefix + "a:xfrm/a:off", nm).Attributes["y"].Value);
            double oldExtX = Convert.ToDouble(node.SelectSingleNode(prefix + "a:xfrm/a:ext", nm).Attributes["cx"].Value);
            double oldExtY = Convert.ToDouble(node.SelectSingleNode(prefix + "a:xfrm/a:ext", nm).Attributes["cy"].Value);

            node.SelectSingleNode(prefix + "a:xfrm/a:off", nm).Attributes["x"].Value = (_offX + (oldX - _chOffX) * _extX / _chExtX).ToString();
            node.SelectSingleNode(prefix + "a:xfrm/a:off", nm).Attributes["y"].Value = (_offY + (oldY - _chOffY) * _extY / _chExtY).ToString();
            node.SelectSingleNode(prefix + "a:xfrm/a:ext", nm).Attributes["cx"].Value = (oldExtX * _extX / _chExtX).ToString();
            node.SelectSingleNode(prefix + "a:xfrm/a:ext", nm).Attributes["cy"].Value = (oldExtY * _extY / _chExtY).ToString();

            return node;
        }

        #endregion

        

        #region Translate chart to Picture

        /// <summary>
        ///  get the embeded chart data
        /// </summary>
        /// <param name="chartTypeNode">chart type node (eg:c:barChart)</param>
        /// <param name="nm">name space</param>
        /// <returns>chart data</returns>
        public static DataTable GetChartData(XmlNode chartTypeNode, XmlNamespaceManager nm)
        {
            DataTable dt = new DataTable();
            XmlNodeList series = chartTypeNode.SelectNodes("//c:ser", nm);

            string[] sersName = GetSeriesName(series, nm);
            for (int i = 0; i < sersName.Length; i++)
            {
                // 散点图需存储X和Y的值
                if (chartTypeNode.Name == "c:scatterChart" || chartTypeNode.Name == "c:bubbleChart")
                {
                    dt.Columns.Add(sersName[i] + "XValue", Type.GetType("System.Double"));
                    dt.Columns.Add(sersName[i] + "YValue", Type.GetType("System.Double"));
                }
                else
                {
                    dt.Columns.Add(sersName[i], Type.GetType("System.Double"));
                }
            }

            // get the categories' name
            //  XmlNodeList cts = series[0].SelectNodes("c:cat/c:strRef/c:strCache/c:pt", nm);
            string[] ctName = GetCategoryName(series[0], nm);

            if (ctName.Length > 0)
            {
                //  get the series' value
                for (int i = 0; i < ctName.Length; i++)
                {
                    DataRow dr = dt.NewRow();
                    for (int j = 0; j < series.Count; j++)
                    {
                        string serName = series[j].SelectSingleNode("c:tx/c:strRef/c:strCache/c:pt/c:v", nm).InnerText;
                        string val = series[j].SelectSingleNode("c:val/c:numRef/c:numCache/c:pt[" + (i + 1) + "]/c:v", nm).InnerText;

                        dr[serName] = val;
                    }
                    dt.Rows.Add(dr);
                }
            }
            else //无分类，比如scatter
            {
                if (chartTypeNode.Name == "c:scatterChart" || chartTypeNode.Name == "c:bubbleChart")
                {
                    string ptCount=string.Empty;                    
                    XmlNode numRefptCount = chartTypeNode.SelectSingleNode("c:ser/c:xVal/c:numRef/c:numCache/c:ptCount", nm);
                    if (numRefptCount != null)
                    {
                        ptCount = numRefptCount.Attributes["val"].Value;
                    }
                    else
                    {
                        ptCount = chartTypeNode.SelectSingleNode("c:ser/c:xVal/c:strRef/c:strCache/c:ptCount", nm).Attributes["val"].Value;
                    }
                    int xCount = Convert.ToInt32(ptCount);

                    int i = 1;
                    for (int k = 0; k < xCount; k++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int j = 0; j < series.Count; j++)
                        {
                            string serName = series[j].SelectSingleNode("c:tx/c:strRef/c:strCache/c:pt/c:v", nm).InnerText;

                            if (chartTypeNode.Name == "c:scatterChart" || chartTypeNode.Name == "c:bubbleChart")
                            {
                                string xVal = string.Empty;
                                string yVal = string.Empty;

                                XmlNode xValueNode = series[j].SelectSingleNode("c:xVal/c:numRef/c:numCache/c:pt[" + (k + 1) + "]/c:v", nm);
                                if (xValueNode != null)
                                {
                                    xVal = xValueNode.InnerText;
                                }
                                else
                                {
                                    xVal = i.ToString();
                                }

                                XmlNode yValueNode = series[j].SelectSingleNode("c:yVal/c:numRef/c:numCache/c:pt[" + (k + 1) + "]/c:v", nm);
                                if (yValueNode != null)
                                {
                                    yVal = yValueNode.InnerText;
                                }
                                else
                                {
                                    yVal = series[j].SelectSingleNode("c:yVal/c:strRef/c:strCache/c:pt[" + (k + 1) + "]/c:v", nm).InnerText;
                                }
                                dr[serName + "XValue"] = xVal;
                                dr[serName + "YValue"] = yVal;
                            }
                            
                        }
                        i++;
                        dt.Rows.Add(dr);
                    }
                }

            }
            //foreach (XmlNode ser in series)
            //{
            //    XmlNodeList cvs = ser.SelectNodes("c:val/c:numRef/c:numCache/c:pt", nm);
            //    for (int j = 0; j < ctName.Length; j++)
            //    {

            //        string serName = ser.SelectSingleNode("c:tx/c:strRef/c:strCache/c:pt/c:v", nm).InnerText;
            //        string val = cvs[j].SelectSingleNode("c:v", nm).InnerText;

            //        DataRow dr = dt.NewRow();
            //        dr[serName] = val;

            //        dt.Rows.Add(dr);

            //    }

            //}

            return dt;
        }

        /// <summary>
        ///  get the series name
        /// </summary>
        /// <param name="series">series node</param>
        /// <param name="nm">name space</param>
        /// <returns>series name</returns>
        public static string[] GetSeriesName(XmlNodeList series, XmlNamespaceManager nm)
        {

            // XmlNodeList series = doc.SelectNodes("//c:ser", nm);
            string[] seriesName = new string[series.Count];
            for (int i = 0; i < series.Count; i++)
            {
                seriesName[i] = series[i].SelectSingleNode("c:tx/c:strRef/c:strCache/c:pt/c:v", nm).InnerText;
            }

            return seriesName;
        }

        /// <summary>
        ///  get the category name
        /// </summary>
        /// <param name="ser">series node</param>
        /// <param name="nm">name space</param>
        /// <returns>category name</returns>
        public static string[] GetCategoryName(XmlNode ser, XmlNamespaceManager nm)
        {

            XmlNodeList cts;
            XmlNode strRef = ser.SelectSingleNode("c:cat/c:strRef", nm);
            if (strRef != null)
            {
                cts = ser.SelectNodes("c:cat/c:strRef/c:strCache/c:pt", nm);
            }
            else
            {
                cts = ser.SelectNodes("c:cat/c:numRef/c:numCache/c:pt", nm);
            }
            string[] ctsName = new string[cts.Count];
            for (int i = 0; i < cts.Count; i++)
            {
                ctsName[i] = cts[i].SelectSingleNode("c:v", nm).InnerText;
            }

            return ctsName;
        }

        /// <summary>
        ///  get the series' name
        /// </summary>
        /// <param name="chartFile">chart xml file</param>
        /// <returns>series' name</returns>
        public static string[] GetSeriesName(string chartFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(chartFile);
            XmlNamespaceManager nm = new XmlNamespaceManager(doc.NameTable);
            nm.AddNamespace("c", TranslatorConstants.XMLNS_C);

            XmlNodeList series = doc.SelectNodes("//c:ser", nm);
            string[] seriesName = new string[series.Count];
            for (int i = 0; i < series.Count; i++)
            {
                seriesName[i] = series[i].SelectSingleNode("c:tx/c:strRef/c:strCache/c:pt/c:v", nm).InnerText;
            }

            return seriesName;
        }

        /// <summary>
        /// get the categories name
        /// </summary>
        /// <param name="chartFile">chart xml file</param>
        /// <returns>categories name</returns>
        public static string[] GetCategoryName(string chartFile)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(chartFile);
            XmlNamespaceManager nm = new XmlNamespaceManager(doc.NameTable);
            nm.AddNamespace("c", TranslatorConstants.XMLNS_C);

            XmlNode ser = doc.SelectSingleNode("//c:ser[1]", nm);
            XmlNodeList cts;
            XmlNode strRef = ser.SelectSingleNode("c:cat/c:strRef", nm);
            if (strRef != null)
            {
                cts = ser.SelectNodes("c:cat/c:strRef/c:strCache/c:pt", nm);
            }
            else
            {
                cts = ser.SelectNodes("c:cat/c:numRef/c:numCache/c:pt", nm);
            }
            string[] ctsName = new string[cts.Count];
            for (int i = 0; i < cts.Count; i++)
            {
                ctsName[i] = cts[i].SelectSingleNode("c:v", nm).InnerText;
            }

            return ctsName;
        }


        /// <summary>
        ///  Check the chart cotains how many chart Types (Combo type)
        /// </summary>
        /// <param name="xdoc">chart file</param>
        /// <param name="nm">name space manager</param>
        /// <returns>chart type nodes</returns>
        public static LinkedList<XmlNode> ChkChartTypeNodes(XmlDocument xdoc, XmlNamespaceManager nm)
        {
            LinkedList<XmlNode> chartTypeNodes = new LinkedList<XmlNode>();
            chartTypeNodes.Clear();
            XmlNodeList plotAreaChildNodes = xdoc.SelectSingleNode("//c:plotArea", nm).ChildNodes;
            foreach (XmlNode plotAreaChildNode in plotAreaChildNodes)
            {
                string nodeName = plotAreaChildNode.Name;
                if (nodeName.Contains("Chart"))
                {
                    chartTypeNodes.AddLast(plotAreaChildNode);
                }
            }

            return chartTypeNodes;
        }

        /// <summary>
        ///  Determine the corresponding chart type according to the origianl chart node
        /// </summary>
        /// <param name="oriChartNode">original chart node</param>
        /// <param name="nm">name space</param>
        /// <returns>framework chart type</returns>
        public static SeriesChartType DtmCorrChartType(XmlNode oriChartNode, XmlNamespaceManager nm)
        {
            string oriChartTypes = oriChartNode.Name;
            switch (oriChartTypes)
            {
                case "c:barChart":
                    {
                        string barDir = oriChartNode.SelectSingleNode("c:barDir", nm).Attributes["val"].Value;
                        string grouping = oriChartNode.SelectSingleNode("c:grouping", nm).Attributes["val"].Value;
                        if (barDir.Equals("col"))
                        {
                            switch (grouping)
                            {
                                case "standard": return SeriesChartType.Column;
                                case "clustered": return SeriesChartType.RangeColumn;
                                case "stacked": return SeriesChartType.StackedColumn;
                                case "percentStacked": return SeriesChartType.StackedColumn100;
                                default: return SeriesChartType.Column;
                            }
                        }
                        else // =bar
                        {
                            switch (grouping)
                            {
                                case "standard": return SeriesChartType.Bar;
                                case "clustered": return SeriesChartType.RangeBar;
                                case "stacked": return SeriesChartType.StackedBar;
                                case "percentStacked": return SeriesChartType.StackedBar100;
                                default: return SeriesChartType.Bar;
                            }
                        }
                    }
                case "c:bar3DChart":
                    {
                        string barDir = oriChartNode.SelectSingleNode("c:barDir", nm).Attributes["val"].Value;
                        string grouping = oriChartNode.SelectSingleNode("c:grouping", nm).Attributes["val"].Value;
                        if (barDir.Equals("col"))
                        {
                            switch (grouping)
                            {
                                case "standard": return SeriesChartType.Column;
                                case "clustered": return SeriesChartType.RangeColumn;
                                case "stacked": return SeriesChartType.StackedColumn;
                                case "percentStacked": return SeriesChartType.StackedColumn100;
                                default: return SeriesChartType.Column;
                            }
                        }
                        else // =bar
                        {
                            switch (grouping)
                            {
                                case "standard": return SeriesChartType.Bar;
                                case "clustered": return SeriesChartType.RangeBar;
                                case "stacked": return SeriesChartType.StackedBar;
                                case "percentStacked": return SeriesChartType.StackedBar100;
                                default: return SeriesChartType.Bar;
                            }
                        }
                    }
                case "c:lineChart":
                    {
                        string grouping = oriChartNode.SelectSingleNode("c:grouping", nm).Attributes["val"].Value;
                        switch (grouping)
                        {
                            // 目前framework暂不支持这几种不同的方式
                            case "standard": return SeriesChartType.Line;
                            case "clustered": return SeriesChartType.Line;
                            case "stacked": return SeriesChartType.Line;
                            case "percentStacked": return SeriesChartType.Line;
                            default: return SeriesChartType.Line;
                        }
                    }
                case "c:line3DChart":
                    {
                        string grouping = oriChartNode.SelectSingleNode("c:grouping", nm).Attributes["val"].Value;
                        switch (grouping)
                        {
                            // 目前framework暂不支持这几种不同的方式
                            case "standard": return SeriesChartType.Line;
                            case "clustered": return SeriesChartType.Line;
                            case "stacked": return SeriesChartType.Line;
                            case "percentStacked": return SeriesChartType.Line;
                            default: return SeriesChartType.Line;
                        }
                    }
                case "c:pieChart": return SeriesChartType.Pie;
                case "c:pie3DChart": return SeriesChartType.Pie;
                case "c:ofPieChart": return SeriesChartType.Pie;
                case "c:doughnutChart": return SeriesChartType.Doughnut;
                case "c:areaChart":
                    {
                        string grouping = oriChartNode.SelectSingleNode("c:grouping", nm).Attributes["val"].Value;
                        switch (grouping)
                        {
                            case "standard": return SeriesChartType.Area;
                            case "clustered": return SeriesChartType.Area;
                            case "stacked": return SeriesChartType.StackedArea;
                            case "percentStacked": return SeriesChartType.StackedArea100;
                            default: return SeriesChartType.Area;
                        }
                    }
                case "c:area3DChart":
                    {
                        string grouping = oriChartNode.SelectSingleNode("c:grouping", nm).Attributes["val"].Value;
                        switch (grouping)
                        {
                            case "standard": return SeriesChartType.Area;
                            case "clustered": return SeriesChartType.Area;
                            case "stacked": return SeriesChartType.StackedArea;
                            case "percentStacked": return SeriesChartType.StackedArea100;
                            default: return SeriesChartType.Area;
                        }
                    }
                case "c:scatterChart":
                    {
                        string scatterStyle = oriChartNode.SelectSingleNode("c:scatterStyle", nm).Attributes["val"].Value;
                        switch (scatterStyle)
                        {
                            // 目前frameWork不支持这几种的细分
                            case "line": return SeriesChartType.Point;
                            case "lineMarker": return SeriesChartType.FastPoint;
                            case "marker": return SeriesChartType.Point;
                            case "none": return SeriesChartType.Point;
                            case "smooth": return SeriesChartType.Spline;
                            case "smoothMarker": return SeriesChartType.Spline;
                            default: return SeriesChartType.Point;
                        }
                    }
                case "c:bubbleChart": return SeriesChartType.Bubble;
                case "c:bubble3D": return SeriesChartType.Bubble;
                case "c:stockChart": return SeriesChartType.Stock;
                case "c:surfaceChart": return SeriesChartType.SplineArea;
                case "c:surface3DChart": return SeriesChartType.SplineArea;
                case "c:radarChart":
                    {
                        string radarStyle = oriChartNode.SelectSingleNode("c:radarStyle", nm).Attributes["val"].Value;
                        switch (radarStyle)
                        {
                            // 目前frameWork不支持这几种的细分
                            case "marker": return SeriesChartType.Radar;
                            case "filled": return SeriesChartType.Radar;
                            case "standard": return SeriesChartType.Radar;
                            default: return SeriesChartType.Radar;
                        }
                    }
                default: return SeriesChartType.Column;
            }
        }

        /// <summary>
        /// get the title's text
        /// </summary>
        /// <param name="paragraphNode">a:p</param>
        /// <param name="nm">name sapce</param>
        /// <returns>title</returns>
        public static string GetTitleText(XmlNode paragraphNode, XmlNamespaceManager nm)
        {
            string title = string.Empty;
            if (paragraphNode != null)
            {
                XmlNodeList rows = paragraphNode.SelectNodes("a:r", nm);
                foreach (XmlNode row in rows)
                {
                    title += row.SelectSingleNode("a:t", nm).InnerText;
                }
            }
            return title;
        }

        public static bool ChkValueShownAsLabel(XmlNode dLbls, XmlNamespaceManager nm)
        {
            bool result = false;
            if (dLbls != null)
            {
                XmlNode showVal = dLbls.SelectSingleNode("c:showVal", nm);
                string val = showVal.Attributes["val"].Value;
                if (val == "1")
                {
                    result = true;
                }
            }
            return result;
        }

        public static void SaveChartAsPic(string chartXMLFile, string outputImageFile, ChartImageFormat imageFormat)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(chartXMLFile);
            XmlNamespaceManager nm = new XmlNamespaceManager(doc.NameTable);
            nm.AddNamespace("c", TranslatorConstants.XMLNS_C);
            nm.AddNamespace("a", TranslatorConstants.XMLNS_A);

            Chart embededChart = new Chart();

            // TODO
            Legend l = new Legend();//初始化一个图例的实例
            l.TableStyle = LegendTableStyle.Wide;
            l.LegendStyle = LegendStyle.Table;
            l.Alignment = System.Drawing.StringAlignment.Near;//设置图表的对齐方式(中间对齐，靠近原点对齐，远离原点对齐)
            l.BackColor = System.Drawing.Color.Black;//设置图例的背景颜色
            //l.DockedToChartArea = embededChart.ChartAreas[0].Name;//设置图例要停靠在哪个区域上
            // l.Docking = System.Windows.Forms.DataVisualization.Charting.Docking.Right;//设置停靠在图表区域的位置(底部、顶部、左侧、右侧)
            //l.Font = new System.Drawing.Font("Trebuchet MS", 8.25F, System.Drawing.FontStyle.Bold);//设置图例的字体属性
            l.IsTextAutoFit = false;//设置图例文本是否可以自动调节大小
            l.LegendStyle = System.Windows.Forms.DataVisualization.Charting.LegendStyle.Column;//设置显示图例项方式(多列一行、一列多行、多列多行)
            l.Name = "Embeded Chart Legend";//设置图例的名称
            l.TitleAlignment = System.Drawing.StringAlignment.Far;
            embededChart.Legends.Add(l.Name);
            embededChart.Legends[0].Docking = Docking.Right;

            LinkedList<DataTable> chartDTs = new LinkedList<DataTable>();
            chartDTs.Clear();
            LinkedList<XmlNode> chartTypeNodes = ChkChartTypeNodes(doc, nm);
            LinkedList<XmlNode>.Enumerator enumer = chartTypeNodes.GetEnumerator();
            while (enumer.MoveNext())
            {
                DataTable chartDT = GetChartData(enumer.Current, nm);
                chartDTs.AddLast(chartDT);
            }

            // bind the data
            LinkedList<DataTable>.Enumerator chartData = chartDTs.GetEnumerator();
            chartData.MoveNext();
            embededChart.DataSource = chartData.Current;

            // TODO 3d and grid
            embededChart.ChartAreas.Add("Series");
            embededChart.ChartAreas["Series"].AxisX.MajorGrid.Enabled = false;
            embededChart.ChartAreas["Series"].AxisX.MinorGrid.Enabled = false;
            //embededChart.ChartAreas["Series"].Area3DStyle.Enable3D = true;

            // TODO: title
            XmlNode chartTitle = doc.SelectSingleNode("//c:chart/c:title/c:tx/c:rich/a:p", nm);
            embededChart.Titles.Add(GetTitleText(chartTitle, nm));
            embededChart.Titles[0].Font = new System.Drawing.Font("Trebuchet MS", 21.6F, System.Drawing.FontStyle.Regular);
            embededChart.Titles[0].Alignment = System.Drawing.ContentAlignment.TopCenter;

            // embededChart.Titles.Add("left title");
            //embededChart.Titles[1].DockedToChartArea = "Series";
            // embededChart.Titles[1].Docking = Docking.Left;
            //  embededChart.Titles[1].Alignment = System.Drawing.ContentAlignment.MiddleLeft;

            // set the series
            string[] seriesName;
            string[] ctsName = GetCategoryName(chartXMLFile);
            LinkedList<XmlNode>.Enumerator chartTypeSeries = chartTypeNodes.GetEnumerator();
            while (chartTypeSeries.MoveNext())
            {
                XmlNodeList currentSeries = chartTypeSeries.Current.SelectNodes("c:ser", nm);
                seriesName = GetSeriesName(currentSeries, nm);
                for (int i = 0; i < seriesName.Length; i++)
                {
                    embededChart.Series.Add(seriesName[i]);
                }
                for (int i = 0; i < seriesName.Length; i++)
                {
                    SeriesChartType chartType = DtmCorrChartType(chartTypeSeries.Current, nm);

                    if (chartType == SeriesChartType.Point || chartType == SeriesChartType.FastPoint || chartType == SeriesChartType.Spline || chartType == SeriesChartType.Bubble)
                    {
                        embededChart.Series[seriesName[i]].XValueMember = seriesName[i] + "XValue";
                        embededChart.Series[seriesName[i]].YValueMembers = seriesName[i] + "YValue";
                    }
                    else
                    {
                        embededChart.Series[seriesName[i]].XValueType = ChartValueType.String;
                        embededChart.Series[seriesName[i]].YValueMembers = seriesName[i];
                    }

                    // 显示数据值 TODO
                    embededChart.Series[seriesName[i]].IsValueShownAsLabel = ChkValueShownAsLabel(doc.SelectSingleNode("//c:dLbls", nm), nm);


                    embededChart.Series[seriesName[i]].ChartType = chartType;

                    // TODO
                    embededChart.Series[seriesName[i]].Legend = "Embeded Chart Legend";
                }

            }

            embededChart.DataBind();


            for (int j = 0; j < ctsName.Length; j++)
            {
                embededChart.Series[0].Points[j].AxisLabel = ctsName[j];
            }

            embededChart.Width = 800;
            embededChart.Height = 600;
            //embededChart.BorderlineWidth = 0;

            embededChart.SaveImage(outputImageFile, imageFormat);

        }

        #endregion

    }
}

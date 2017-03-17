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
using System.Collections;
using System.IO;
using System.Xml;
using System.Text;
using System.Diagnostics;
using Act.UofTranslator.UofZipUtils;
using log4net;
using System.Collections.Generic;

namespace Act.UofTranslator.UofTranslatorStrictLib
{
    /// <summary>
    /// Zip archiving states
    /// </summary>
    /// <author>linwei, sunqingzhi</author>
       /// <summary>
    /// An <c>XmlWriter</c> implementation for serializing the xml stream to a zip archive.
    /// All the necessary information for creating the archive and its entries is picked up 
    /// from the incoming xml stream and must conform to the following specification :
    /// 
    /// TODO : XML schema
    /// 
    /// example :
    /// 
    /// <c>&lt;pzip:archive pzip:target="path"&gt;</c>
    /// 	<c>&lt;pzip:entry pzip:target="relativePath"&gt;</c>
    /// 		<c>&lt;-- xml fragment --&lt;</c>
    /// 	<c>&lt;/pzip:entry&gt;</c>
    /// 	<c>&lt;-- other zip entries --&lt;</c>
    /// <c>&lt;/pzip:archive&gt;</c>
    /// 
    /// </summary>
    class UofZipWriter : XmlWriter
    {
        private const string ZIP_POST_PROCESS_NAMESPACE = "urn:u2o:xmlns:post-processings:special";
        private const string PIC_POST_PROCESS_NAMESPACE = "urn:u2opic:xmlns:post-processings:special";
        private const string PART_ELEMENT = "entry";
        private const string ARCHIVE_ELEMENT = "archive";
        private const string PIC_ELEMENT = "picture";
        //ole对象添加
        private const string OLE_EMBEDDING = "embeding";
        private const string OLE_DRAW = "drawing";
        private const string OLE_DRAWRELS = "drawingrel";
        private const string OLE_POST_PROCESS_NAMESPACE = "urn:u2oole:xmlns:post-processings:special";

        /// <summary>
        /// The zip archive
        /// </summary>
        private ZipWriter zipOutputStream;

        private ProcessingState processingState = ProcessingState.None;

        private Stack elements;
        private Stack attributes;

        /// <summary>
        /// A delegate <c>XmlWriter</c> that actually feeds the zip output stream. 
        /// </summary>
        private XmlWriter delegateWriter = null;

        /// <summary>
        /// The delegate settings
        /// </summary>
        private XmlWriterSettings delegateSettings = null;

        /// <summary>
        /// 源文件
        /// </summary>
        private string sourcefile;

        /// <summary>
        /// Source attribute of the currently processed binary file
        /// </summary>
        private string binarySource;

        /// <summary>
        /// 所有需要处理的二进制文件
        /// </summary>
        private IList<string> binaries;

        /// <summary>
        /// 日志记录器
        /// </summary>
        private static ILog logger = logger = LogManager.GetLogger(typeof(OoxZipWriter).FullName);

        /// <summary>
        /// Constructor
        /// </summary>
        public UofZipWriter(string inputfile)
        {
            elements = new Stack();
            attributes = new Stack();
            binaries = new List<string>();

            delegateSettings = new XmlWriterSettings();
            delegateSettings.OmitXmlDeclaration = false;
            delegateSettings.CloseOutput = false;
            delegateSettings.Encoding = Encoding.UTF8;
            delegateSettings.Indent = false;

            // If we use a new delegate per entry in the archive, 
            // XML conformance will be checked at the document level.
            // It is not possible to check XML conformance at the docuement level
            // with a single delegate that writes all the entries.
            // We must then use ConformanceLevel.Fragment and the xml declaration will be missing.
            delegateSettings.ConformanceLevel = ConformanceLevel.Document;

            sourcefile = inputfile;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose managed resources
                if (delegateWriter != null)
                    delegateWriter.Close();
                if (zipOutputStream != null)
                    zipOutputStream.Dispose();
            }
            base.Dispose(disposing);

        }

        /// <summary>
        /// Delegates <c>WriteStartElement</c> calls when the element's prefix does not 
        /// match a zip command.  
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="localName"></param>
        /// <param name="ns"></param>
        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            logger.Debug("[startElement] prefix=" + prefix + " localName=" + localName + " ns=" + ns);
            elements.Push(new Node(prefix, localName, ns));

            // not a pic processing instruction
            if (!ZIP_POST_PROCESS_NAMESPACE.Equals(ns) && !PIC_POST_PROCESS_NAMESPACE.Equals(ns) && !OLE_POST_PROCESS_NAMESPACE.Equals(ns))
            {
                if (delegateWriter != null)
                {
                    logger.Debug("{WriteStartElement=" + localName + "} delegate");
                    delegateWriter.WriteStartElement(prefix, localName, ns);
                }
                else
                {
                    logger.Debug("[WriteStartElement=" + localName + " } delegate is null");
                }
            }
        }

        /// <summary>
        /// Delegates <c>WriteEndElement</c> calls when the element's prefix does not 
        /// match a zip command. 
        /// Otherwise, close the archive or flush the delegate writer. 
        /// </summary>
        public override void WriteEndElement()
        {
            Node elt = (Node)elements.Pop();

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE) && !elt.Ns.Equals(PIC_POST_PROCESS_NAMESPACE)&& !elt.Ns.Equals(OLE_POST_PROCESS_NAMESPACE))
            {
                logger.Debug("delegate - </" + elt.Name + ">");
                if (delegateWriter != null)
                {
                    delegateWriter.WriteEndElement();
                }
            }
            else
            {
                switch (elt.Name)
                {
                    case ARCHIVE_ELEMENT://archive
                        if (zipOutputStream != null)
                        {
                            logger.Info("[closing archive]");
                            logger.Info("Copying binary files!");
                            CopyBinaryData();
                            zipOutputStream.Close();
                            zipOutputStream = null;
                        }
                        if (processingState == ProcessingState.EntryWaiting)
                        {
                            processingState = ProcessingState.None;
                        }
                        break;
                    case PART_ELEMENT://entry
                        if (delegateWriter != null)
                        {
                            logger.Info("[end part]");
                            delegateWriter.WriteEndDocument();
                            delegateWriter.Flush();
                            delegateWriter.Close();
                            delegateWriter = null;
                        }
                        if (processingState == ProcessingState.EntryStarted)
                        {
                            processingState = ProcessingState.EntryWaiting;
                        }
                        break;
                    case OLE_DRAW:
                    case OLE_DRAWRELS:
                    case OLE_EMBEDDING://ole
                    case PIC_ELEMENT://picture
                        if (binarySource != null)
                        {
                            logger.Info("add picture: target=" + binarySource);
                            if (!binaries.Contains(binarySource))
                            {
                                binaries.Add(binarySource);
                            }
                            binarySource = "";
                        }
                        break;
                }

            }
        }

        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            Node elt = (Node)elements.Peek();
            logger.Debug("[WriteStartAttribute] prefix=" + prefix + " localName" + localName + " ns=" + ns + " element=" + elt.Name);

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE) && !elt.Ns.Equals(PIC_POST_PROCESS_NAMESPACE) && !elt.Ns.Equals(OLE_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteStartAttribute(prefix, localName, ns);
                }
            }
            else
            {
                // we only store attributes of tied to zip processing instructions
                attributes.Push(new Node(prefix, localName, ns));
            }
        }

        public override void WriteEndAttribute()
        {
            Node elt = (Node)elements.Peek();
            logger.Debug("[WriteEndAttribute] element=" + elt.Name);

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE) && !elt.Ns.Equals(PIC_POST_PROCESS_NAMESPACE) && !elt.Ns.Equals(OLE_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteEndAttribute();
                }
            }
            else
            {
                Node attribute = (Node)attributes.Pop();
            }
        }

        // TODO: throw an exception if "target" attribute not set
        public override void WriteString(string text)
        {
            Node elt = (Node)elements.Peek();

            if (!elt.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE) && !elt.Ns.Equals(PIC_POST_PROCESS_NAMESPACE) && !elt.Ns.Equals(OLE_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteString(text);
                }
            }
            else
            {
                // Pick up the target attribute 
                if (attributes.Count > 0)
                {
                    Node attr = (Node)attributes.Peek();
                    if (attr.Ns.Equals(ZIP_POST_PROCESS_NAMESPACE) || attr.Ns.Equals(PIC_POST_PROCESS_NAMESPACE)|| attr.Ns.Equals(OLE_POST_PROCESS_NAMESPACE))
                    {
                        switch (elt.Name)
                        {
                            case ARCHIVE_ELEMENT://archive
                                // Prevent nested archive creation
                                if (processingState == ProcessingState.None && attr.Name.Equals("target"))
                                {
                                    logger.Info("creating archive : " + text);

                                    // 输出包的路径放在式样单中间文档元素<archive Target="包路径名">中了，
                                    // 从中提取即可，当然也可以直接传递过来
                                    zipOutputStream = ZipFactory.CreateArchive(text);
                                    processingState = ProcessingState.EntryWaiting;
                                }
                                break;
                            case PART_ELEMENT://entry
                                // Prevent nested entry creation
                                if (processingState == ProcessingState.EntryWaiting && attr.Name.Equals("target"))
                                {
                                    logger.Info("creating new part : " + text);

                                    //此处的text是entry的target属性的值，也就是文件的路径
                                    zipOutputStream.AddEntry(text);
                                    delegateWriter = XmlWriter.Create(zipOutputStream, delegateSettings);
                                    processingState = ProcessingState.EntryStarted;
                                    delegateWriter.WriteStartDocument();
                                }
                                break;
                            case OLE_DRAW:
                            case OLE_DRAWRELS:
                            case OLE_EMBEDDING://ole
                            case PIC_ELEMENT://picture
                                if (processingState != ProcessingState.None)
                                {
                                    if (attr.Name.Equals("target"))
                                    {
                                        binarySource += text;
                                        logger.Info("picture: target=" + binarySource);
                                    }
                                }
                                break;
                        }
                    }
                }
            }
        }

        #region 重载的其他wirte方法，基本没加其他处理

        public override void WriteFullEndElement()
        {
            this.WriteFullEndElement();
        }

        public override void WriteStartDocument()
        {
            // nothing to do here
        }

        public override void WriteStartDocument(bool b)
        {
            // nothing to do here
        }

        public override void WriteEndDocument()
        {
            // nothing to do here
        }

        public override void WriteDocType(string name, string pubid, string sysid, string subset)
        {
            // nothing to do here
        }

        public override void WriteCData(string s)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteCData(s);
            }
        }

        public override void WriteComment(string s)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteComment(s);
            }
        }

        public override void WriteProcessingInstruction(string name, string text)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteProcessingInstruction(name, text);
            }
        }

        public override void WriteEntityRef(String name)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteEntityRef(name);
            }
        }

        public override void WriteCharEntity(char c)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteCharEntity(c);
            }
        }

        public override void WriteWhitespace(string s)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteWhitespace(s);
            }
        }

        public override void WriteSurrogateCharEntity(char lowChar, char highChar)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteSurrogateCharEntity(lowChar, highChar);
            }
        }

        public override void WriteChars(char[] buffer, int index, int count)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteChars(buffer, index, count);
            }
        }

        public override void WriteRaw(char[] buffer, int index, int count)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteRaw(buffer, index, count);
            }
        }

        public override void WriteRaw(string data)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteRaw(data);
            }
        }

        public override void WriteBase64(byte[] buffer, int index, int count)
        {
            if (delegateWriter != null)
            {
                delegateWriter.WriteBase64(buffer, index, count);
            }
        }

        public override WriteState WriteState
        {
            // nothing smart to do here
            get
            {
                return delegateWriter.WriteState;
            }
        }

        #endregion

        public override void Close()
        {
            // zipStream and delegate are closed elsewhere.... if everything else is fine
            if (delegateWriter != null)
            {
                delegateWriter.Close();
                delegateWriter = null;
            }
            if (zipOutputStream != null)
            {
                zipOutputStream.Close();
                zipOutputStream = null;
            }
        }

        public override void Flush()
        {
            // nothing to do
        }

        public override string LookupPrefix(String ns)
        {
            if (delegateWriter != null)
            {
                return delegateWriter.LookupPrefix(ns);
            }
            else
            {
                return null;
            }
        }

        private const int BUFFER_SIZE = 4096;

        /// <summary>
        /// copy binary data (currently, only picture) in to zip archive
        /// </summary>
        private void CopyBinaryData()
        {
            int size = binaries.Count;
            if (size == 0)
                return;
            if (zipOutputStream == null)
                return;

            string newFilename = "", source = "";
            string pictureLocation = "data";
            BinaryWriter writer = new BinaryWriter(zipOutputStream);
            for (int i = 0; i < size; i++)
            {
                source = binaries[i];

                // judge the uof file type according to the the folder name which define by upper step
                //switch (Directory.GetParent(source).Name.ToString())
                //{
                //    case "wordPic": pictureLocation = "word/media"; break;
                //    case "pptPic": pictureLocation = "ppt/media"; break;
                //    case "excelPic": pictureLocation = "xl/media"; break;
                //    default: break;
                //}
                if(source.Contains("media"))
                {
                    newFilename = pictureLocation + "\\" + source.Substring(source.LastIndexOf("\\") + 1);
                
                }
                else if(source.Contains("embeddings"))//拷贝ole对象
                {
                    newFilename = "embeddings\\" + source.Substring(source.LastIndexOf("\\") + 1);
                
                }
                else
                {
                    newFilename = "drawings\\" + source.Substring(source.LastIndexOf("\\") + 1);
                
                }
                try
                {
                    zipOutputStream.AddEntry(newFilename);
                    byte[] buffer = File.ReadAllBytes(source);
                    writer.Write(buffer);
                }
                catch (Exception e)
                {
                    logger.Error("Fail to copy picture: " + source + " into zip archive! " + e.Message);
                    logger.Error(e.StackTrace);
                }
                finally
                {
                    File.Delete(source);
                }
            }
            writer.Close();
        }

        /// <summary>
        /// Simple representation of elements or attributes nodes
        /// </summary>
        private class Node
        {

            private string name;
            public string Name
            {
                get { return name; }
            }

            private string prefix;
            public string Prefix
            {
                get { return prefix; }
            }

            private string ns;
            public string Ns
            {
                get { return ns; }
            }

            public Node(string prefix, string name, string ns)
            {
                this.prefix = prefix;
                this.name = name;
                this.ns = ns;
            }
        }
    }
}

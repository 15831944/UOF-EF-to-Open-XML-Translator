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
using System.Text;
using System.IO;
using System.Xml;
using System.Diagnostics;
using log4net;

namespace Act.UofTranslator.UofTranslatorStrictLib
{
    /// <summary>
    /// This class is used to handle the post processing for UOF.Currently, 
    /// we just replace the picture element with base64 content.
    /// </summary>
    /// <author>sunqingzhi, linwei</author>
    class UofWriter : XmlWriter
    {

        private const string UOF_POST_PROCESS_NAMESPACE = "urn:o2upic:xmlns:post-processings:special";

        private const string PIC_ELEMENT = "picture";

        private const int BUFFER_SIZE = 1024;

        private XmlWriter delegateWriter = null;

        private XmlWriterSettings delegateSettings = null;

        private XmlResolver resolver;

        private Stack elements;

        private Stack attributes;

        private string binarySource;

        private string outputpath;

        private Stream bufferStream;

        private static ILog logger =  logger = LogManager.GetLogger(typeof(UofWriter).FullName);

        private bool isUofPackaged = true;

        public bool IsUofPackaged
        {
            get
            {
                return this.isUofPackaged;
            }
            set
            {
                if (null == (object)value)
                    throw new ArgumentNullException("The value of IsUofPackaged can't be null");
                this.isUofPackaged = value;
            }
        }

        public UofWriter(XmlResolver res, string outputfile)
        {
            elements = new Stack();
            attributes = new Stack();

            outputpath = outputfile;
            delegateSettings = new XmlWriterSettings();
            delegateSettings.OmitXmlDeclaration = false;
            delegateSettings.CloseOutput = true;
            delegateSettings.Encoding = Encoding.UTF8;
            delegateSettings.Indent = false;
            delegateSettings.ConformanceLevel = ConformanceLevel.Document;
            bufferStream = new BufferedStream(new MemoryStream(BUFFER_SIZE));
            delegateWriter = XmlWriter.Create(bufferStream, delegateSettings);
            resolver = res;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (delegateWriter != null)
                {
                    delegateWriter.Flush();
                    delegateWriter.Close();
                }

            }
            base.Dispose(disposing);

        }

        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            logger.Debug("[startElement] prefix=" + prefix + " localName=" + localName + " ns=" + ns);
            elements.Push(new Node(prefix, localName, ns));
            if (!UOF_POST_PROCESS_NAMESPACE.Equals(ns))
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

        public override void WriteEndElement()
        {
            Node elt = (Node)elements.Pop();

            if (!elt.Ns.Equals(UOF_POST_PROCESS_NAMESPACE))
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
                    case PIC_ELEMENT:
                        if (binarySource != null)
                        {   
                            byte[] img = null;
                            try
                            {
                                img = GetBinaryBytes(binarySource);
                            }
                            catch (Exception e)
                            {
                                logger.Error("Fail to get binary souce: " + binarySource + "\n" + e.Message + "\n" + e.StackTrace);
                            }
                            if (null == img)
                            {
                                break;
                            }

                            if (this.isUofPackaged)
                            {
                                this.delegateWriter.WriteStartElement("uof", "数据", null);
                                this.delegateWriter.WriteAttributeString("uof", "locID", null, "u0037");
                                delegateWriter.WriteBase64(img, 0, img.Length);
                                this.delegateWriter.WriteEndElement();
                            }
                            else
                            {
                                FileStream fileStream = null;
                                string filename = binarySource.Substring(binarySource.LastIndexOf("/") + 1);
                                string fullname = this.outputpath.Substring(0, this.outputpath.LastIndexOf("\\") + 1) + filename;
                                
                                this.delegateWriter.WriteStartElement("uof", "路径", null);
                                this.delegateWriter.WriteAttributeString("uof", "locID", null, "u0038");
                                this.delegateWriter.WriteString(filename);
                                this.delegateWriter.WriteEndElement();

                                try
                                {
                                    fileStream = File.Create(fullname);
                                    fileStream.Write(img, 0, img.Length);
                                    logger.Info("Success to create file: " + fullname);
                                }
                                catch (Exception e)
                                {
                                    logger.Error("Fail to create file: " + fullname + "\n" + e.Message);
                                    logger.Error(e.StackTrace);
                                }
                                finally
                                {
                                    fileStream.Close();
                                }
                            }

                            binarySource = null;
                        }
                        break;
                }
            }
        }


        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            Node elt = (Node)elements.Peek();
            logger.Debug("[WriteStartAttribute] prefix=" + prefix + " localName" + localName + " ns=" + ns + " element=" + elt.Name);

            if (!elt.Ns.Equals(UOF_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteStartAttribute(prefix, localName, ns);
                }
            }
            else
            {
                attributes.Push(new Node(prefix, localName, ns));
            }
        }

        public override void WriteEndAttribute()
        {
            Node elt = (Node)elements.Peek();
            logger.Debug("[WriteEndAttribute] element=" + elt.Name);

            if (!elt.Ns.Equals(UOF_POST_PROCESS_NAMESPACE))
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

        public override void WriteString(string text)
        {
            Node elt = (Node)elements.Peek();

            if (!elt.Ns.Equals(UOF_POST_PROCESS_NAMESPACE))
            {
                if (delegateWriter != null)
                {
                    delegateWriter.WriteString(text);
                }
            }
            else
            {
                if (attributes.Count > 0)
                {
                    Node attr = (Node)attributes.Peek();  
                    switch (elt.Name)
                    {
                        case PIC_ELEMENT:
                            if (attr.Name.Equals("target"))
                            {
                                binarySource += text;
                                logger.Info("picture target=" + binarySource);
                            }
                            break;
                    }  
                }
            }
        }

        #region override other functions

        public override void WriteFullEndElement()
        {
            this.WriteFullEndElement();
        }

        public override void WriteStartDocument()
        {
            delegateWriter.WriteStartDocument();
        }

        public override void WriteStartDocument(bool b)
        {
            // nothing to do here
        }

        public override void WriteEndDocument()
        {
            delegateWriter.WriteEndDocument();
            // nothing to do here,用不着，因为一切由delegatewriter来代替
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

        #endregion

        public override WriteState WriteState
        {
            get
            {
                return delegateWriter.WriteState;
            }
        }


        public override void Close()
        {
            if (delegateWriter != null)
            {
                delegateWriter.Flush();

                FileStream fileStream = null;
                try
                {
                    fileStream = new FileStream(outputpath, FileMode.Create);
                    byte[] tmp = new byte[BUFFER_SIZE];
                    int bytesRead = 0;
                    bufferStream.Seek(0, SeekOrigin.Begin);
                    bytesRead = bufferStream.Read(tmp, 0, 3);
                    if (tmp[0] != 0xEF || tmp[1] != 0xBB || tmp[2] != 0xBF || bytesRead!=3)
                    {
                        bufferStream.Seek(0, SeekOrigin.Begin);
                    }
                    while ((bytesRead = bufferStream.Read(tmp, 0, BUFFER_SIZE)) > 0)
                    {
                        fileStream.Write(tmp, 0, bytesRead);
                    }
                }
                catch (Exception e)
                {
                    logger.Error(e.Message);
                    logger.Error(e.StackTrace);
                }
                finally
                {
                    delegateWriter.Close();
                    delegateWriter = null;
                    if (fileStream != null)
                        fileStream.Close();
                }
            }
        }

        public override void Flush()
        {
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

        private byte[] GetBinaryBytes(string relativeUri)
        {
            Uri absoluteUri = resolver.ResolveUri(null, relativeUri);
            Stream sourceStream = (Stream)resolver.GetEntity(absoluteUri, null, Type.GetType("System.IO.Stream"));
            if (null == sourceStream)
            {
                logger.Error("Fail to get stream for relativeUri: " + relativeUri);
                return null;
            }

            int size = (int)sourceStream.Length;
            BinaryReader binaryReader = new BinaryReader(sourceStream);
            byte[] imgBuffer = null;
            imgBuffer = binaryReader.ReadBytes(size);
            binaryReader.Close();

            logger.Info("GetBinaryBytes for relativeUri: " + relativeUri + " bytes copied: " + size);
            return imgBuffer;
        }

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

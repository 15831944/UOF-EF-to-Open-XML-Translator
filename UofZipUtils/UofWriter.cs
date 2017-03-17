using System;
using System.Collections;
using System.Text;
using System.IO;
using System.Xml;
using System.Diagnostics;

namespace Act.UofConverter.UofZipUtils
{
    /// <summary>
    /// this class is used to handle the post processing for uof,like copying the pictures 
    /// from binary to base64 formatter.
    /// 
    /// 实现方法: 在转换的过程中对于图片所在的对象集，要有图片标签，
    ///           并且标签的信息完备，但内容（即base64代码）为空.  
    /// </summary>
   public class UofWriter:XmlWriter
    {

        private const string UOF_POST_PROCESS_NAMESPACE = "urn:act:xmlns:post-processings:special";
        
        private const string PIC_ELEMENT = "picture";

        /// <summary>
        /// 对于delegatewriter,它存在的目的是保留被重载的xmlwriter的方法，比如如果UofWriter重载了
        /// XmlWriter的WriterStartElement()方法，添加了一些特殊处理，比如对获取的指定元素<图片></图片>
        /// 要在其中添加代码，但原来的功能就不能用了，为即添加特殊处理，还能使用原来的操作，
        /// 就在被重载的方法体中使用delegatewriter.WriteStartElement(),这样就保留了原来的作用.
        /// </summary>
        XmlWriter delegateWriter = null;
       
        /// <summary>
        /// delegate settings
        /// </summary>
        private XmlWriterSettings delegateSettings = null;

        /// <summary>
        /// get the binary images from oox
        /// </summary>
        private XmlResolver resolver;

        private Stack elements;
        private Stack attributes;

        /// <summary>
        /// Source attribute of the currently processed binary file
        /// </summary>
        private string binarySource;
        private string outputpath;
       
        
        /// <summary>
        /// Constructor. the outputfile should contain the extension name--uof,like xxx.uof.
        /// </summary>
        public UofWriter(XmlResolver res,string outputfile)
        {
            elements = new Stack();
            attributes = new Stack();

            outputpath = outputfile;
            delegateSettings = new XmlWriterSettings();
            delegateSettings.OmitXmlDeclaration = false;
            delegateSettings.CloseOutput = false;
            delegateSettings.Encoding = Encoding.Default;
            delegateSettings.Indent = false;
            delegateSettings.ConformanceLevel = ConformanceLevel.Document;
  
            delegateWriter = XmlWriter.Create(outputpath, delegateSettings);
            resolver = res;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Dispose managed resources
                if (delegateWriter != null)
                {
                    // delegateWriter.WriteEndDocument();
                    delegateWriter.Flush();
                    delegateWriter.Close();
                }
               
            }
            base.Dispose(disposing);

        }

        /// <summary>
        /// delegatewriter's WriteStartElement method will be called if no special element matches.
        /// </summary>
        /// <param name="prefix"></param>
        /// <param name="localName"></param>
        /// <param name="ns"></param>
        public override void WriteStartElement(string prefix, string localName, string ns)
        {
            Debug.WriteLine("[startElement] prefix=" + prefix + " localName=" + localName + " ns=" + ns);
            elements.Push(new Node(prefix, localName, ns));
            // not a zip processing instruction
            if (!UOF_POST_PROCESS_NAMESPACE.Equals(ns))
            {
                if (delegateWriter != null)
                {
                    Debug.WriteLine("{WriteStartElement=" + localName + "} delegate");
                    delegateWriter.WriteStartElement(prefix, localName, ns);
                }
                else
                {
                    Debug.WriteLine("[WriteStartElement=" + localName + " } delegate is null");
                }
            }         
        }

        /// <summary>
        ///     
        /// </summary>
        public override void WriteEndElement()
        {
            Node elt = (Node)elements.Pop();

            if (!elt.Ns.Equals(UOF_POST_PROCESS_NAMESPACE))
            {
                Debug.WriteLine("delegate - </" + elt.Name + ">");
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
                        if (binarySource != null )
                        {
                            //call delegatewriter.WriteString(xxx) to write base64
                            CopyBase64RightNow(binarySource);
                            binarySource = null; 
                        }
                        break;
                }
            }
        }


        public override void WriteStartAttribute(string prefix, string localName, string ns)
        {
            Node elt = (Node)elements.Peek();
            Debug.WriteLine("[WriteStartAttribute] prefix=" + prefix + " localName" + localName + " ns=" + ns + " element=" + elt.Name);

            if (!elt.Ns.Equals(UOF_POST_PROCESS_NAMESPACE))
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
            Debug.WriteLine("[WriteEndAttribute] element=" + elt.Name);

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

        // TODO: throw an exception if "target" attribute not set
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
                // Pick up the target attribute 
                if (attributes.Count > 0)
                {
                    Node attr = (Node)attributes.Peek();
                    if (attr.Ns.Equals(UOF_POST_PROCESS_NAMESPACE))
                    {
                        switch (elt.Name)
                        {
                            case PIC_ELEMENT: 
                                if (attr.Name.Equals("source"))
                                {
                                    binarySource += text;
                                    Debug.WriteLine("copy source=" + binarySource);
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
            // nothing smart to do here
            get
            {
                return delegateWriter.WriteState;
            }
        }


        public override void Close()
        {
            // zipStream and delegate are closed elsewhere.... if everything else is fine
            if (delegateWriter != null)
            {
                //delegateWriter.WriteEndDocument();
                delegateWriter.Flush();
                delegateWriter.Close();
                delegateWriter = null;
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

     

        /// <summary>
        /// Transfer a binary file to base64. 
        /// The source archive is handled by the resolver, while the destination archive 
        /// corresponds to our zipOutputStream.  
        /// </summary>
        /// <param name="source">Relative path inside the source archive</param>
        private void CopyBase64RightNow(String source)
        {
            Stream sourceStream = GetStream(source);

            if (sourceStream != null )
            {
                int size = (int)sourceStream.Length;
                byte[] img = new byte[size];
                BinaryReader f = new BinaryReader(sourceStream);
                img = f.ReadBytes(size);
                f.Close();
                delegateWriter.WriteBase64(img,0,size);
                
                Debug.WriteLine("CopyBase64 : " + source + " --> uof, bytes copied = " + size);
            }
        }

        private Stream GetStream(string relativeUri)
        {
            Uri absoluteUri = resolver.ResolveUri(null, relativeUri);
            return (Stream)resolver.GetEntity(absoluteUri, null, Type.GetType("System.IO.Stream"));
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

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Xsl;

namespace Act.UofTranslator.UofTranslatorStrictLib
{
    /// <summary>
    /// This is a abstract class base processor, it provides common functions.
    /// </summary>
    /// <author>linwei</author>
    abstract class AbstractProcessor : IProcessor    
    {

        protected string inputFile = "";

        protected string outputFile = "";

        protected string originalFile = "";

        protected string resourceDir = "";

        #region IProcessor Members

        public abstract bool transform();
        
        public string InputFilename//属性
        {
            get
            {
                return inputFile;
            }
            set
            {
                inputFile = value;
            }
        }

        public string OutputFilename
        {
            get
            {
                return outputFile;
            }
            set
            {
                outputFile = value;
            }
        }

        public string OriginalFilename
        {
            get
            {
                return originalFile;
            }
            set
            {
                originalFile = value;
            }
        }

        public string ResourceDir
        {
            get
            {
                return resourceDir;
            }
            set
            {
                resourceDir = value;
            }
        }
        #endregion


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
                // 20130402 update by xuzhenwei BUG_2727:集成OO-UOF 预定义图形大小不一致；填充，边框转换前后不一致；三维效果丢失；部分文本框对齐方式转换前后不一致。start
                // 先不处理*.wdp格式的图片
                if (media.Extension != ".wdp")
                {
                    XmlElement mediaFileNode = xmlDoc.CreateElement("u2opic", "picture", "urn:u2opic:xmlns:post-processings:special");
                    mediaFileNode.SetAttribute("target", "urn:u2opic:xmlns:post-processings:special", media.FullName);
                    mediaNode.AppendChild(mediaFileNode);
                }
                //end
            }
            root.AppendChild(mediaNode);
            return xmlDoc;
        }
        protected XmlDocument OlePretreatment(XmlDocument xmlDoc, string fireNodeName, string olePath, XmlNamespaceManager nms)
        {
            DirectoryInfo mediaInfo = new DirectoryInfo(olePath + "embeddings");
            FileInfo[] medias = mediaInfo.GetFiles();
            XmlNode root = xmlDoc.SelectSingleNode(fireNodeName, nms);
            XmlElement mediaNode = xmlDoc.CreateElement("w", "ole", TranslatorConstants.XMLNS_W);
            foreach (FileInfo media in medias)
            {
                XmlElement mediaFileNode = xmlDoc.CreateElement("u2oole", "embeding", "urn:u2oole:xmlns:post-processings:special");
                mediaFileNode.SetAttribute("target", "urn:u2oole:xmlns:post-processings:special", media.FullName);
                mediaNode.AppendChild(mediaFileNode);
            }
            if (Directory.Exists(olePath + "drawings"))
            {
                mediaInfo = new DirectoryInfo(olePath + "drawings"); 
                medias = mediaInfo.GetFiles();
                foreach (FileInfo media in medias)
                {
                    if (media.Extension!=".rels")
                    {
                        XmlElement mediaFileNode = xmlDoc.CreateElement("u2oole", "drawingrel", "urn:u2oole:xmlns:post-processings:special");
                        mediaFileNode.SetAttribute("target", "urn:u2oole:xmlns:post-processings:special", media.FullName);
                        mediaNode.AppendChild(mediaFileNode);
                    }
                    else
                    {
                     
                            XmlElement mediaFileNode = xmlDoc.CreateElement("u2oole", "drawing", "urn:u2oole:xmlns:post-processings:special");
                            mediaFileNode.SetAttribute("target", "urn:u2oole:xmlns:post-processings:special", media.FullName);
                            mediaNode.AppendChild(mediaFileNode);
                    }
                }
            }
            root.AppendChild(mediaNode);
            return xmlDoc;
        }
    }
}

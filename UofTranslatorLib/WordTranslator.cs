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
using System.Text.RegularExpressions;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// Core conversion methods
    /// </summary>
    /// <author>linwei</author>
    public class WordTranslator : UOFTranslator
    {


        private bool isOox2UofPackage = true;

        private static ILog logger = LogManager.GetLogger(typeof(WordTranslator).FullName);

        public WordTranslator()
        {

            InitalPreProcessors();
            XmlConfigurator.Configure(LogManager.GetRepository(),
                new FileInfo(UOFTranslator.ASSEMBLY_PATH + @"conf\log4net.config"));//log4net,记录日志
            LoadConfigs(ASSEMBLY_PATH + @"conf\config.xml");
        }

        protected override void InitalPreProcessors()
        {
            base.InitalPreProcessors();
            uof2ooxPreProcessors.AddLast(new UofToOoxPreProcessorOneWord());
            //uof2ooxPreProcessors.AddLast(new UofToOoxPreProcessorTwoWord());
            //uof2ooxPreProcessors.AddLast(new UofToOoxPreProcessorThreeWord());
            oox2uofPreProcessors.AddLast(new OoxToUofPreProcessorOneWord());
            //oox2uofPostProcessors.AddLast(new OoxToUofPostProcessorOneWord());

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
                        this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + TranslatorConstants.UOFToOOX_WORD_LOCATION);
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
                XslCompiledTransform xslt = new XslCompiledTransform();
                XsltSettings settings = new XsltSettings(true, false);
                xslt.Load(xslDoc, settings, resourceResolver);

                XsltArgumentList parameters = new XsltArgumentList();
                parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                if (outputFile != null)
                {
                    parameters.AddParam("outputFile", "", outputFile);
                    writer = new OoxZipWriter(inputFile);
                  //  writer = new UofZipWriter(inputFile);
                }
                else
                {
                    writer = new XmlTextWriter(new StringWriter());
                }
                xslt.Transform(source, parameters, writer);
                mainOutput = outputFile;
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
            XmlUrlResolver resourceResolver = null;
            string picture_xml = "word\\media";
            string custom_xml = "customXml";
            string wordPrePath = Path.GetDirectoryName(inputFile) + "\\";//输出文档路径
            string mediaPath = wordPrePath + picture_xml;
            string customPath = wordPrePath + custom_xml;
            string tmpPic = wordPrePath + "tmpPic.xml";//增加对图片的处理
            string tmpCustom = wordPrePath + "custom.xml";
            XmlDocument xmlDoc = new XmlDocument();
            try
            {

                // chart转换成图片
                string charts_xml = wordPrePath + "word\\charts";
                if (Directory.Exists(charts_xml))
                {
                    string[] charts = Directory.GetFiles(charts_xml, "chart*");
                    if (charts.Length > 0)
                    {
                        if (!Directory.Exists(mediaPath))
                        {
                            Directory.CreateDirectory(mediaPath);
                        }
                    }
                    try
                    {
                        foreach (string chartXMLFile in charts)
                        {
                            SaveChartAsPic(chartXMLFile, mediaPath + Path.AltDirectorySeparatorChar + Path.GetFileName(chartXMLFile) + ".jpg", System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg);
                        }
                    }
                    catch (Exception ex)
                    {
                        logger.Warn("Save chart occured an error!");
                    }
                }

                if (Directory.Exists(mediaPath))
                {
                    xmlDoc.Load(inputFile);
                    XmlNamespaceManager nms = new XmlNamespaceManager(xmlDoc.NameTable);
                
                    nms.AddNamespace("w", TranslatorConstants.XMLNS_W);
                    xmlDoc = PicPretreatment(xmlDoc, "w:document", mediaPath, nms);
                    xmlDoc.Save(tmpPic);
                    inputFile = tmpPic;
                }

                // manage the custom xml part
                if (Directory.Exists(customPath))
                {
                    xmlDoc.Load(inputFile);
                    XmlNamespaceManager nms = new XmlNamespaceManager(xmlDoc.NameTable);

                    nms.AddNamespace("w", TranslatorConstants.XMLNS_W);
                    xmlDoc = CustomXMPretreatment(xmlDoc, "w:document", customPath, nms);
                    xmlDoc.Save(tmpCustom);
                    inputFile = tmpCustom;
                }

                resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
                    this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + TranslatorConstants.OOXToUOF_WORD_LOCATION);//resources.word.oox2uof

                MainTransform(TranslatorConstants.OOXToUOF_XSL, resourceResolver, originalFile, inputFile, outputFile);

            }
            catch (Exception ex)
            {
                logger.Warn(ex.Message);
            }

            //XmlUrlResolver resourceResolver = null;
            //XPathDocument xslDoc = null;
            //XmlReaderSettings xrs = new XmlReaderSettings();
            //XmlReader source = null;
            //XmlWriter writer = null;
            //OoxZipResolver zipResolver = null;

            //try
            //{
            //    xrs.ProhibitDtd = true;//禁用DTD

            //    string xslLocation = TranslatorConstants.OOXToUOF_XSL;//oox2uof.xsl
            //    if (outputFile == null)
            //    {
            //        xslLocation = TranslatorConstants.OOXToUOF_COMPUTE_SIZE_XSL;//oox2uof-compute-size.xsl
            //    }
            //    if (resourceDir == null)//程序中resourceDir为null
            //    {
            //        //Assembly获得当前程序集。this.GetType().Namespace是指UofTranslatorLib。"."代表"/"
            //        resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
            //            this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + TranslatorConstants.OOXToUOF_WORD_LOCATION);//resources.word.oox2uof
            //        xslDoc = new XPathDocument(((ResourceResolver)resourceResolver).GetInnerStream(xslLocation));//从程序集中加载指定的请单资源,oox2uof.xsl

            //        source = XmlReader.Create(inputFile);//预处理的中间文档
            //    }
            //    else
            //    {
            //        resourceResolver = new XmlUrlResolver();
            //        xslDoc = new XPathDocument(resourceDir + "/" + xslLocation);
            //        source = XmlReader.Create(resourceDir + "/" + TranslatorConstants.SOURCE_XML, xrs);//source.xml
            //    }
            //    XslCompiledTransform xslt = new XslCompiledTransform();
            //    XsltSettings settings = new XsltSettings(true, false);//XsltSettings指定执行 XSLT 样式表时要支持的 XSLT 功能
            //    xslt.Load(xslDoc, settings, resourceResolver);//oox2uof.xsl,settings,程序集中的位置
            //    zipResolver = new OoxZipResolver(originalFile, resourceResolver);//把源程序解压
            //    XsltArgumentList parameters = new XsltArgumentList();//包含数目可变的参数（这些参数是 XSLT 参数，或者是扩展对象）。
            //    parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);


            //    if (outputFile != null)
            //    {
            //        parameters.AddParam("outputFile", "", outputFile);
            //      //  writer = new OoxZipWriter(inputFile);
            //    //  writer = new UofWriter(zipResolver, outputFile);
            //      writer = new UofZipWriter();

            //     mainOutput = outputFile;

            //    }
            //    else
            //    {
            //        writer = XmlWriter.Create(Path.GetDirectoryName(inputFile) + "\\" + "temp.uof");
            //    }
            //  //  xslt.Transform(source, parameters, writer);
            //    xslt.Transform(source, parameters, writer, zipResolver);//读预处理的中间文档,参数,写到输出文件,zip
            //}
            //finally
            //{
            //    if (writer != null)
            //        writer.Close();
            //    if (source != null)
            //        source.Close();
            //    if (zipResolver != null)
            //        zipResolver.Dispose();
            //}
        }
    }




}

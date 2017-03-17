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
using System.Resources;


namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// Presetation translator
    /// </summary>
    /// <author>linyaohu</author>
    public class PresentationTranslator : UOFTranslator
    {
        private bool isOox2UofPackage = true;

        private static ILog logger = LogManager.GetLogger(typeof(PresentationTranslator).FullName);

        public PresentationTranslator()
        {
            InitalPreProcessors();
            XmlConfigurator.Configure(LogManager.GetRepository(),
                new FileInfo(UOFTranslator.ASSEMBLY_PATH + @"conf\log4net.config"));
            LoadConfigs(ASSEMBLY_PATH + @"conf\config.xml");
        }


        protected override void InitalPreProcessors()
        {
            base.InitalPreProcessors();
            uof2ooxPreProcessors.AddLast(new UofToOoxPreProcessorOnePresentation());
            // uof2ooxPreProcessors.AddLast(new UofToOoxPreProcessorTwoPresentation());
            oox2uofPreProcessors.AddLast(new OOXToUOFPreProcessorOnePresentation());
            oox2uofPostProcessors.AddLast(new OoxToUofPostProcessorOnePresentation());
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
                        this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + "Powerpoint.uof2oox");
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
                    xslt.Load(xslDoc, settings, resourceResolver);
                    //Assembly ass = Assembly.Load("ppt_uof2oox");
                    //Type t = ass.GetType("uof2oox");
                    //xslt.Load(t);
                    XsltArgumentList parameters = new XsltArgumentList();
                    parameters.XsltMessageEncountered += new XsltMessageEncounteredEventHandler(MessageCallBack);
                    if (outputFile != null)
                    {
                        parameters.AddParam("outputFile", "", outputFile);
                        parameters.AddParam("FileType", "", "Prsentation");
                        writer = new OoxZipWriter(inputFile);
                    }
                    else
                    {
                        writer = new XmlTextWriter(new StringWriter());
                    }
                    xslt.Transform(source, parameters, writer);
                }
                catch (Exception ex)
                {
                    logger.Error("Fail in the main translator of Presentation", ex);
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
        bool isInitialized = false;
        XslCompiledTransform transFrom = null;
        //罗文甜：增加预编译
        private void Initialize()
        {
            //XslCompiledTransform transFrom = new XslCompiledTransform();
            transFrom = new XslCompiledTransform();
            XsltSettings setting = new XsltSettings(true, false);
            XmlUrlResolver xur = new XmlUrlResolver();
            try
            {
                Assembly ass = Assembly.Load("ppt_oox2uof");//调用ppt_oox2uof.dll程序集
                Type t = ass.GetType("oox2uof");//name=oox2uof.xslt
                transFrom.Load(t);
                isInitialized = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        protected override void DoOoxToUofMainTransform(string originalFile, string inputFile, string outputFile, string resourceDir)
        {

            string wordPrePath = Path.GetDirectoryName(inputFile) + Path.DirectorySeparatorChar;
            string grpTmp = Path.GetDirectoryName(inputFile) + "\\" + "grpTmp.xml";//把grpTmp.xml放在temp下的缓存文件夹下
            string drawingRelTmp = Path.GetDirectoryName(inputFile) + "\\" + "drawingRelTmp.xml";            
            string picture_xml = "ppt\\media";
            string mediaPath=wordPrePath+picture_xml;
            string custom_xml = "customXml";
            string customPath = wordPrePath + custom_xml;
            string tmpCustom = wordPrePath + "custom.xml";

                    XmlDocument xdoc = new XmlDocument();
            xdoc.Load(inputFile);//tempdoc1.xml
            XmlNamespaceManager nm = new XmlNamespaceManager(xdoc.NameTable);
            nm.AddNamespace("p", TranslatorConstants.XMLNS_P);
            nm.AddNamespace("a", TranslatorConstants.XMLNS_A);
            nm.AddNamespace("w", TranslatorConstants.XMLNS_W);
            nm.AddNamespace("dsp", TranslatorConstants.XMLNS_DSP);
            FileStream fs = new FileStream(grpTmp, FileMode.Create);

            try
            {
                GrpShPre(xdoc, nm, "p", fs);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (fs != null)
                    fs.Close();
            }

            // chart转换成图片
            string charts_xml = wordPrePath+"ppt\\charts";
            if (Directory.Exists(charts_xml))
            {
                string[] charts = Directory.GetFiles(charts_xml, "chart*");
                if(charts.Length>0)
                {
                    if(!Directory.Exists(mediaPath))
                    {
                        Directory.CreateDirectory(mediaPath);
                    }
                }
                foreach (string chartXMLFile in charts)
                {
                    SaveChartAsPic(chartXMLFile, mediaPath + Path.AltDirectorySeparatorChar + Path.GetFileName(chartXMLFile) + ".jpg", System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Jpeg);
                }
            }
            //ole对象处理lyy
            string tmpPic = wordPrePath + "tmpOle.xml";
            if (Directory.Exists(wordPrePath + "ppt\\embeddings"))
            {
                xdoc.Load(grpTmp);
                xdoc = OlePretreatment(xdoc, "p:presentation", wordPrePath, nm);
                xdoc.Save(tmpPic);
                grpTmp = tmpPic;
            }
            // 图片预处理

            tmpPic = wordPrePath + "tmpPic.xml";
            if (Directory.Exists(mediaPath))
            {
                xdoc.Load(grpTmp);
                xdoc = PicPretreatment(xdoc, "p:presentation", mediaPath, nm);
                xdoc.Save(tmpPic);
                grpTmp = tmpPic;
            }

            // manage the custom xml part
            if (Directory.Exists(customPath))
            {
                xdoc.Load(grpTmp);
                XmlNamespaceManager nms = new XmlNamespaceManager(xdoc.NameTable);

                nms.AddNamespace("w", TranslatorConstants.XMLNS_W);
                nms.AddNamespace("p", TranslatorConstants.XMLNS_P);
                xdoc = CustomXMPretreatment(xdoc, "p:presentation", customPath, nm);
                xdoc.Save(tmpCustom);
                grpTmp = tmpCustom;
            }

            AddDrawingRel(xdoc, nm, drawingRelTmp);

            XmlReaderSettings xrs = new XmlReaderSettings();
            XmlReader source = null;

            string toMainSheet = ZipXMLFile(drawingRelTmp);
            string zipXMLFileName = "input.xml";

           // string sr = ZipXMLFile(inputFile);
            ZipReader archive = ZipFactory.OpenArchive(toMainSheet);
            source = XmlReader.Create(archive.GetEntry(zipXMLFileName));

            //xdoc.Load(drawingRelTmp);

            XslCompiledTransform transFrom = new XslCompiledTransform();
            XsltSettings setting = new XsltSettings(true, false);
            XmlUrlResolver xur = new XmlUrlResolver();
            XPathNavigator nav = ((IXPathNavigable)xdoc).CreateNavigator();//root
            fs = null;
            string mainOutputFile = Path.GetDirectoryName(inputFile) + "\\" + "mainOutputFile.xml";//缓存文件中保存mainOutputFile.xml


            fs = new FileStream(mainOutputFile, FileMode.Create);
            string mainSheet = Path.GetDirectoryName(inputFile).ToString() + "\\" + TranslatorConstants.OOXToUOF_XSL;//从缓存中调用oox2uof.xsl
            transFrom.Load(mainSheet, setting, xur);
            //Assembly ass = Assembly.Load("ppt_oox2uof");//调用ppt_oox2uof.dll程序集
            //Type t = ass.GetType("oox2uof");//name=oox2uof.xslt
            //transFrom.Load(t);
            //transFrom.Transform(nav, null, fs);

            transFrom.Transform(source, null, fs);
            fs.Close();
            mainOutput = mainOutputFile;


            //XmlUrlResolver resourceResolver = null;

            //try
            //{

            //    resourceResolver = new ResourceResolver(Assembly.GetExecutingAssembly(),
            //        this.GetType().Namespace + "." + TranslatorConstants.RESOURCE_LOCATION + "." + "Powerpoint.oox2uof");
           // MainTransform(TranslatorConstants.OOXToUOF_XSL, resourceResolver, originalFile, grpTmp, outputFile);

            ////}
            //catch (Exception ex)
            //{
            //    logger.Warn(ex.Message);
            //}


        }

        /// <summary>
        ///  add the drawing relationships to the files bescause of the smartArt picture fill
        /// </summary>
        /// <param name="xdoc">origianl file</param>
        /// <param name="resultFile">result</param>
        public static void AddDrawingRel(XmlDocument xdoc, XmlNamespaceManager nm, string resultFile)
        {
            string drawingRelPath = Path.GetDirectoryName(resultFile) + Path.AltDirectorySeparatorChar + "ppt\\diagrams\\_rels";

            // smartArt has picture fill
            if (Directory.Exists(drawingRelPath))
            {
                XmlNode root = xdoc.SelectSingleNode("p:presentation", nm);
                string[] drawingRels = Directory.GetFiles(drawingRelPath, "drawing*");
                for (int i = 0; i < drawingRels.Length; i++)
                {
                    XmlNode drawingRel = xdoc.CreateElement("dsp", "drawingRelationship", TranslatorConstants.XMLNS_DSP);
                    XmlAttribute att = xdoc.CreateAttribute("id");
                    att.Value = Path.GetFileName(drawingRels[i]);
                    drawingRel.Attributes.Append(att);

                    XmlDocument relsDoc = new XmlDocument();
                    relsDoc.Load(drawingRels[i]);
                    drawingRel.InnerXml = relsDoc.LastChild.InnerXml;
                    root.AppendChild(drawingRel);
                }
            }

            xdoc.Save(resultFile);

        }


    }
}

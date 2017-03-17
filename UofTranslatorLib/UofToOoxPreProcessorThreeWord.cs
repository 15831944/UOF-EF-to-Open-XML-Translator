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

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// The third step pre process
    /// </summary>
    /// <author>linwei</author>
    class UofToOoxPreProcessorThreeWord : AbstractProcessor
    {
        private static ILog logger = LogManager.GetLogger(typeof(UofToOoxPreProcessorThreeWord).FullName);

        public override bool transform()
        {
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load(inputFile);
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("uof", "http://schemas.uof.org/cn/2003/uof");
            XmlNodeList nodeList = xmlDoc.SelectNodes("//uof:UOF//uof:对象集//uof:其他对象", nsmgr);

            //add by linyaohu
            string pre3OutputFileName = Path.GetDirectoryName(inputFile).ToString() + "\\"+"tmpDoc3.xml";

            if (nodeList.Count != 0)
            {
                //modify by linyaohu
               // string path = inputFile.Substring(inputFile.LastIndexOf("\\") + 1);
               // path = path.Remove(path.IndexOf("."));
               // path = Path.Combine(Path.GetDirectoryName(inputFile), path);
                string path = Path.GetDirectoryName(inputFile) + "\\" + "wordPic";
                try
                {
                    Directory.CreateDirectory(path);
                }
                catch (Exception e)
                {
                    logger.Error("Fail to create tmp directory for UofToOoxPreProcessorStepThree:\n" + e.Message);
                    logger.Error(e.StackTrace);
                    return false;
                }

                string fileType = "";
                string filename = "";
                int i = 0;
                int readBytes = 0;
                byte[] buffer = new byte[1000];
                XmlElement replaceNode;
                XmlNodeReader reader;
                FileStream picFileStream = null;
                BinaryWriter bw = null;

                foreach (XmlNode node in nodeList)
                {
                    i++;
                    reader = new XmlNodeReader(node);
                    reader.Read();
                    fileType = reader.GetAttribute("uof:公共类型");
                    if (fileType == null)
                    {
                        fileType = "jpg";
                    }
                    filename = path + @"\image" + i + "." + fileType;
                    reader.Read();

                    try
                    {
                        picFileStream = new FileStream(filename,
                            FileMode.OpenOrCreate, FileAccess.Write, FileShare.Read);
                        bw = new BinaryWriter(picFileStream);
                        while ((readBytes = reader.ReadElementContentAsBase64(buffer, 0, 1000)) > 0)
                        {
                            bw.Write(buffer, 0, readBytes);
                        }

                    }
                    catch (Exception e)
                    {
                        logger.Error("Fail to read base64 content or write to file: \n" + e.Message);
                        logger.Error(e.StackTrace);
                    }
                    finally
                    {
                        bw.Close();
                        picFileStream.Close();
                    }
                    replaceNode = xmlDoc.CreateElement("u2opic", "picture", "urn:u2opic:xmlns:post-processings:special");
                    replaceNode.SetAttribute("target", "urn:u2opic:xmlns:post-processings:special", filename);
                    node.ReplaceChild(replaceNode, node.FirstChild);
                }
            }

            try
            { 
                //XmlTextWriter resultWriter = new XmlTextWriter(outputFile, Encoding.UTF8);
                XmlTextWriter resultWriter = new XmlTextWriter(pre3OutputFileName, Encoding.UTF8);
                xmlDoc.Save(resultWriter);
                resultWriter.Close();
                OutputFilename = pre3OutputFileName;
            }
            catch (Exception e)
            {
                logger.Error("Fail to save temp uof file for UofToOoxPreProcessorStepThree: \n" + e.Message);
                logger.Error(e.StackTrace);
                return false;
            }

            return true;
        }
    }
}
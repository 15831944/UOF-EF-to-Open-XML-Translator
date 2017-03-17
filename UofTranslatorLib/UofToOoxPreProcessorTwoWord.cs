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
using System.Xml.Xsl;
using System.Xml;
using log4net;
using System.IO;

namespace Act.UofTranslator.UofTranslatorLib
{
    /// <summary>
    /// The second step 
    /// </summary>
    /// <author>fangchunyan</author>
    class UofToOoxPreProcessorTwoWord : AbstractProcessor
    {
        private static ILog logger = LogManager.GetLogger(typeof(UofToOoxPreProcessorTwoWord).FullName);

        public override bool transform()
        {
            XmlReader source = null;
            XmlWriter writer = null;
            bool isSuccess = true;
          //  string pre2OutputFileName = Path.GetDirectoryName(inputFile).ToString() +"\\"+ "tmpDoc2.xml";
            try
            {
                XPathDocument xpdoc = UOFTranslator.GetXPathDoc(TranslatorConstants.UOFToOOX_PRETREAT_STEP2_XSL, TranslatorConstants.UOFToOOX_WORD_LOCATION);
                XslCompiledTransform xslt = new XslCompiledTransform();
                xslt.Load(xpdoc);
                source = XmlReader.Create(inputFile);
                writer = new XmlTextWriter(outputFile, Encoding.UTF8);
               // writer = new XmlTextWriter(pre2OutputFileName, Encoding.UTF8);
              
                xslt.Transform(source, writer);
               // OutputFilename = pre2OutputFileName;
            }
            catch (Exception e)
            {
                logger.Error(e.Message);
                logger.Error(e.StackTrace);
                isSuccess = false;
            }
            finally
            {
                if (writer != null) 
                    writer.Close();
                if (source != null) 
                    source.Close();
            }
            return isSuccess;
        }
    }
}

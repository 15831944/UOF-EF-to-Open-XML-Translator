/*
 * Copyright (c) 2006, Tsinghua University, China
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

namespace UofAddinLib
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:      配置文件管理类
    ///     Author:         邓追
    ///     Create Date:    2007-08-28
    /// </summary>
    public class ConfigManager
    {
        /// <summary>
        ///     Whether the uof file should be packaged
        /// </summary>
        private bool isOox2UofPackage;

        /// <summary>
        ///     Whether the error is ignore when translation
        /// </summary>
        private bool isErrorIgnored;

        /// <summary>
        ///  Whether uof wants to be translated to transitional OOX
        /// </summary>
        private bool isUofToTransitioanlOOX;

        /// <summary>
        ///  Wheter uof wants to be translated to strict OOX
        /// </summary>
        private bool isUofToStrictOOX;

        /// <summary>
        ///     The path of the config file
        /// </summary>
        private string configfile;

        /// <summary>
        ///     Whether the uof file should be packaged
        /// </summary>
        public bool IsOox2UofPackage
        {
            set { isOox2UofPackage = value; }
            get { return isOox2UofPackage; }
        }

        /// <summary>
        ///     Whether the error is ignore when translation
        /// </summary>
        public bool IsErrorIgnored
        {
            set { isErrorIgnored = value; }
            get { return isErrorIgnored; }
        }

        /// <summary>
        ///  Whether uof wants to be translated to transitional OOX
        /// </summary>
        public bool IsUofToTransitioanlOOX
        {
            set { isUofToTransitioanlOOX = value; }
            get { return isUofToTransitioanlOOX; }
        }

        /// <summary>
        ///  Wheter uof wants to be translated to strict OOX
        /// </summary>
        public bool IsUofToStrictOOX
        {
            set { isUofToStrictOOX = value; }
            get { return isUofToStrictOOX; }
        }

        /// <summary>
        ///     The path of the config file
        /// </summary>
        public string ConfigFile
        {
            set { configfile = value; }
            get { return configfile; }
        }

        /// <summary>
        ///     Constructor of ConfigManager
        /// </summary>
        /// <param name="filename">The path of the config file</param>
        ///     Author:         邓追
        ///     Create Date:    2007-05-24
        public ConfigManager(string filename)
        {
            configfile = filename;
            isOox2UofPackage = true;
            isErrorIgnored = false;
            isUofToTransitioanlOOX = true;
            isUofToStrictOOX = false;
        }

        

        /// <summary>
        ///     Load the config file
        /// </summary>
        ///     Author:         邓追
        ///     Create Date:    2007-05-24
        public void LoadConfig()
        {
            XmlTextReader reader = null;
            bool isConfigOneFound = false;
            bool isConfigTwoFound = false;
            bool isConfigThreeFound = false;
            this.IsUofToStrictOOX = false;
            this.IsUofToTransitioanlOOX = false;
            
            string exceptionString = "";
            try
            {
                reader = new XmlTextReader(configfile);
                while (reader.Read())
                {
                    switch (reader.NodeType)
                    {
                        case XmlNodeType.Element:
                            {
                                if (reader.Name.Equals("oox2uof_package"))
                                {
                                    this.isOox2UofPackage = Convert.ToBoolean(reader.ReadString());
                                    isConfigOneFound = true;
                                }
                                else if (reader.Name.Equals("ignore_error"))
                                {
                                    this.isErrorIgnored = Convert.ToBoolean(reader.ReadString());
                                    isConfigTwoFound = true;
                                }
                                else if (reader.Name.Equals("ooxml_version"))
                                {
                                    string ooxmlVersion = reader.ReadString();
                                    if (ooxmlVersion.ToLower().Equals("transitional"))
                                    {
                                        this.isUofToTransitioanlOOX = true;
                                        this.isUofToStrictOOX = false;
                                    }
                                    else
                                    {
                                        this.isUofToTransitioanlOOX = false;
                                        this.isUofToStrictOOX = true;
                                    }
                                    isConfigThreeFound=true;
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
                throw e;
            }
            finally
            {
                reader.Close();
            }
            if (!isConfigOneFound)
            {
                exceptionString += TranslatorRes.ConfigOneNotFound + System.Environment.NewLine;
            }
            if (!isConfigTwoFound)
            {
                exceptionString += TranslatorRes.ConfigTwoNotFound + System.Environment.NewLine;
            }
            if (!isConfigThreeFound)
            {
                exceptionString += TranslatorRes.ConfigThreeNotFound + System.Environment.NewLine;
            }
            if (!(isConfigOneFound && isConfigTwoFound))
            {
                throw new Exception(exceptionString);
            }
        }

        /// <summary>
        ///     Save the config file
        /// </summary>
        ///     Author:         邓追
        ///     Create Date:    2007-05-24
        public void SaveConfig()
        {
            XmlTextWriter writer = null;
            try
            {
                
                writer = new XmlTextWriter(@configfile, Encoding.UTF8);
                writer.Formatting = Formatting.Indented;
                writer.WriteStartDocument();
                writer.WriteComment(" UofConverter的配置文件 ");
                writer.WriteStartElement("Configuration");
                writer.WriteComment(" oox到uof转换是打包还是存到一个目录下");
                writer.WriteComment(" false表示存到一个目录下，true表示打成一个包");
                writer.WriteElementString("oox2uof_package", Convert.ToString(isOox2UofPackage));
                writer.WriteComment(" 是否对用户提示转换过程中的错误");
                writer.WriteComment(" false为提示，true为不提示");
                writer.WriteElementString("ignore_error", Convert.ToString(isErrorIgnored));
                writer.WriteComment("uof 到OOX的版本");
                writer.WriteComment("transitional 为转换成transitional版本");
                writer.WriteComment("strict 为转换成 strict版本");

                if (this.IsUofToTransitioanlOOX)
                {
                    writer.WriteElementString("ooxml_version", "transitional");
                }
                else
                {
                    writer.WriteElementString("ooxml_version", "strict");
                }
                writer.WriteEndElement();
                writer.WriteEndDocument();
                writer.Flush();
            }
            catch (System.IO.IOException ee)
            {

                throw ee;
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                writer.Close();
            }
        }
    }
}

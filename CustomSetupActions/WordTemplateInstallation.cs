using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using Microsoft.Win32;

namespace CustomSetupActions
{
    /// <summary>
    /// Modify word template config file, replaceing the value [TARGETDIR] in config file with path of installing directory
    /// </summary>
    public class WordTemplateInstallation
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="wEventId"></param>
        /// <param name="uFlags"></param>
        /// <param name="dwItem1"></param>
        /// <param name="dwItem2"></param>
        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        [CustomAction]
        public static ActionResult ModifyWordTemplate(Session session)
        {
            session.Log("Begin ModifyWordTemplate");

            Record record = new Record(0);


            try
            {
                string targetDir = session.CustomActionData["InstallFolder"].ToString();

                //replace [TARGETDIR] in xml files
                string installPath = targetDir;
                string uofTemplatePath = installPath + @"uoftemplate\";
                string egovFile = uofTemplatePath + "templates_e_gov.xml";
                string infoUserFile = uofTemplatePath + "templates_info_user.xml";
                string standFile = uofTemplatePath + "templates_standard.xml";

                //string[] templateFiles = Directory.GetDirectories("wordTemplate");
                //if (templateFiles.Length == 0)
                //{
                //    targetDir = "http://sourceforge.net/projects/uof-translator/files/UOF%20Translator%20V5.0/UOF%20Template/";
                //}

                targetDir = uofTemplatePath;

                //Modify installing path
                string findPattern = @"\[TARGETDIR\]";
                DoReplace(egovFile, targetDir, findPattern);
                DoReplace(infoUserFile, targetDir, findPattern);
                DoReplace(standFile, targetDir, findPattern);

                // template install on the local disk or not 
                string[] dots = Directory.GetFiles(uofTemplatePath, "*.dotx");
                //System.Windows.Forms.MessageBox.Show("dots count=" + dots.Length);
                if (dots.Length == 0)
                {
                    //targetDir = "http://sourceforge.net/projects/uof-translator/files/UOF%20Translator%20V5.0/UOF%20Template/";
                    targetDir = "https://sourceforge.net/projects/uof-translator/files/UOF%20Translator%20V5.1/UOF%20Template/";
                   // targetDir = "http://211.71.14.202/uoftemplate/";
                }                

                //System.Windows.Forms.MessageBox.Show("targetDir=" + targetDir);

                //MODIFY RESOURCE LOCATION
                string findPattern2 = @"\[TARGETURL\]";
                DoReplace(egovFile, targetDir, findPattern2);
                DoReplace(infoUserFile, targetDir, findPattern2);
                DoReplace(standFile, targetDir, findPattern2);

                //Modify language id
                int languageID = GetOfficeLanguageID();//default to english
                string languageMarkup = @"\[languageID\]";
                DoReplace(egovFile, languageID.ToString(), languageMarkup);
                DoReplace(infoUserFile, languageID.ToString(), languageMarkup);
                DoReplace(standFile, languageID.ToString(), languageMarkup);
                
                RemoveTmpContent();
            }
            catch (Exception ex)
            {
                record[0] = ex.Message;
                session.Message(InstallMessage.Error, record);
                return ActionResult.Failure;
            }

            return ActionResult.Success;
        }

        //Get SOFTWARE\Microsoft\Office Language Code, 2052 for Chinese and 1033 for English
        static private int GetOfficeLanguageID()
        {
            int languageID = 1033;
            string reg2010 = @"SOFTWARE\Microsoft\Office\14.0\Common\LanguageResources";
            string reg = @"SOFTWARE\Microsoft\Office\15.0\Common\LanguageResources";

            try
            {
                if (Environment.Is64BitOperatingSystem)
                {
                    RegistryKey registryBase = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry64);
                    RegistryKey key = registryBase.OpenSubKey(reg);
                    RegistryKey key2010 = registryBase.OpenSubKey(reg2010);
                    if (key != null && key.GetValue("UILanguage") != null)
                    {
                        languageID = (int)key.GetValue("UILanguage");
                    }
                    else
                    {
                        if (key2010 != null && key2010.GetValue("UILanguage") != null)
                        {
                            languageID = (int)key2010.GetValue("UILanguage");
                        }
                    }
                }
                else
                {
                    RegistryKey key = Registry.CurrentUser.OpenSubKey(reg);
                    RegistryKey key2010 = Registry.CurrentUser.OpenSubKey(reg2010);
                    if (key != null && key.GetValue("UILanguage") != null)
                    {
                        languageID = (int)key.GetValue("UILanguage");
                    }
                    else
                    {
                        if (key2010 != null && key2010.GetValue("UILanguage") != null)
                        {
                            languageID = (int)key2010.GetValue("UILanguage");
                        }
                    }
                }
            }
            catch
            {
            }
            return languageID;
        }

        [CustomAction]
        public static ActionResult RemoveWordTemplateReg(Session session)
        {
            session.Log("Begin removing word template register information");
            Record record = new Record(0);

            try
            {
                RemoveTmpContent();
            }
            catch (Exception ex)
            {
                record[0] = ex.Message;
                session.Message(InstallMessage.Error, record);
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        /// <summary>
        /// remove temp template content
        /// </summary>
        static private void RemoveTmpContent()
        {

            string reg = @"SOFTWARE\Microsoft\Office\15.0\Common\Spotlight\";
            string reg2010 = @"SOFTWARE\Microsoft\Office\14.0\Common\Spotlight\";

            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(reg);
                if (key != null)
                {
                    key.DeleteSubKeyTree("Content");
                }
                key = Registry.LocalMachine.OpenSubKey(reg);
                if (key != null)
                {
                    key.DeleteSubKeyTree("Content");
                }
                key = Registry.CurrentUser.OpenSubKey(reg2010);
                if (key != null)
                {
                    key.DeleteSubKeyTree("Content");
                }
                key = Registry.LocalMachine.OpenSubKey(reg2010);
                if (key != null)
                {
                    key.DeleteSubKeyTree("Content");
                }
            }
            catch
            {
            }
        }

        /// <summary>
        /// Replace text content with specific content
        /// </summary>
        /// <param name="fileFullName">File Name</param>
        /// <param name="replacedBy">target text</param>
        /// <param name="findPattern">be replaced text</param>
        static private void DoReplace(string fileFullName, string replacedBy, string findPattern)
        {
            string result = string.Empty;
            string inputText = string.Empty;

            Regex r = new Regex(findPattern, RegexOptions.IgnoreCase);

            try
            {
                if (File.Exists(fileFullName))
                {
                    using (StreamReader sr = new StreamReader(fileFullName))
                    {
                        inputText = sr.ReadToEnd();
                    }

                    
                    // Compare the regular expression.
                    if (r.IsMatch(inputText))
                    {
                        result = r.Replace(inputText, replacedBy);

                        // Save Change
                        using (StreamWriter sw = new StreamWriter(fileFullName))
                        {
                            sw.Write(result);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }
    }
}

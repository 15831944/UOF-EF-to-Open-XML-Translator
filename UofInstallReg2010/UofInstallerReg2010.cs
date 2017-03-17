using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using Microsoft.Office.Core;
using MSword = Microsoft.Office.Interop.Word;


namespace UofInstallReg2010
{

    /// <summary>
    ///     Module ID:      
    ///     Depiction:          修改模板路径，并注册COM
    ///     Author:             朴鑫 SYBASE v-xipia 汪冲 Jacky Wang 
    ///     Create Date:        2009-01-22
    ///     Modified by:        
    ///     Last Modified Date: 2009-06-04
    /// </summary>
    [RunInstaller(true)]
    public partial class UofInstallerReg2010 : Installer
    {

        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);


        public UofInstallerReg2010()
        {
            InitializeComponent();
        }


        public override void Install(System.Collections.IDictionary stateSaver)
        {
             // System.Windows.Forms.MessageBox.Show("UOFReg 2010");
            base.Install(stateSaver);
           // System.Windows.Forms.MessageBox.Show(stateSaver.ToString());

           //   System.Windows.Forms.MessageBox.Show(this.Context.Parameters.ContainsValue("TARGETDIR").ToString());

            string targetDir = this.Context.Parameters["TARGETDIR"].ToString();

            //replace [TARGETDIR] in xml files
            string installPath = targetDir;
            string UofTemplatePath = installPath + @"uoftemplate\";
            string egovFile = UofTemplatePath + "templates_e_gov.xml";
            string infoUserFile = UofTemplatePath + "templates_info_user.xml";
            string standFile = UofTemplatePath + "templates_standard.xml";

            string replacedBy = targetDir;
            string findPattern = @"\[TARGETDIR\]";

            doReplace(egovFile, replacedBy, findPattern);
            doReplace(infoUserFile, replacedBy, findPattern);
            doReplace(standFile, replacedBy, findPattern);


           // System.Windows.Forms.MessageBox.Show("do some replace");
            // modify the config.xml file's authority, add everyone with full control. To fix the bug that add-in options access to config.xml is denied. 

            try
            {
                string configFilePath = System.IO.Path.GetDirectoryName(this.GetType().Assembly.Location) + @"\conf\config.xml";
                System.Security.AccessControl.FileSecurity fileSecurity = System.IO.File.GetAccessControl(configFilePath);
                fileSecurity.AddAccessRule(new System.Security.AccessControl.FileSystemAccessRule("Everyone",
                    System.Security.AccessControl.FileSystemRights.FullControl, System.Security.AccessControl.AccessControlType.Allow));
                System.IO.File.SetAccessControl(configFilePath, fileSecurity);
            }
            catch (Exception ex)
            {
                throw ex;
            }

         
            int lcid = GetLcid();
   //System.Windows.Forms.MessageBox.Show("lcid="+lcid);

            string languageid = @"\[languageID\]";

            doReplace(egovFile, lcid.ToString(), languageid);
            doReplace(infoUserFile, lcid.ToString(), languageid);
            doReplace(standFile, lcid.ToString(), languageid);

             // System.Windows.Forms.MessageBox.Show("some replace done");
            // add by huizhong , modify the language id in templates XML files for different office version (2052 for Chinese and 1033 for English)
            //RegistryKey machine;
            //RegistryKey rk;
            //machine = Registry.LocalMachine;
            //rk = machine.OpenSubKey("SOFTWARE\\Microsoft\\Office\\Common");
            //Object ver = rk.GetValue("LastAccessInstall");
            //string vernumber = "";
            //if (ver != null)
            //{

            //    vernumber = ((int)ver).ToString() + ".0";
            //}

            //string keypath = "SOFTWARE\\Microsoft\\Office\\" + vernumber + "\\Common\\LanguageResources";
            //rk = machine.OpenSubKey(keypath);
            //Object olanguage = rk.GetValue("SKULanguage");
            //rk.Close();

            //// string languageid = @"\[languageID\]";
            //string lcidver = "";
            //if (olanguage != null)
            //{
            //    lcidver = ((int)olanguage).ToString();

            //}
            //doReplace(egovFile, lcidver, languageid);
            //doReplace(infoUserFile, lcidver, languageid);
            //doReplace(standFile, lcidver, languageid);

            //  System.Windows.Forms.MessageBox.Show("all replace done!");

            //////////////////////////////////////////////////////////////////////////////

            string targetFile = targetDir + "UofTranslator2010.exe";
            //System.Windows.Forms.MessageBox.Show(targetFile);
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(@targetFile);
            //psi.RedirectStandardOutput = true; 
            psi.Arguments = "/regserver";
            psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            psi.UseShellExecute = false;
            System.Diagnostics.Process listFiles;
            listFiles = System.Diagnostics.Process.Start(psi);
            //System.IO.StreamReader myOutput = listFiles.StandardOutput; 
            listFiles.WaitForExit();

            //if (listFiles.HasExited)
            //{
            //    System.Windows.Forms.MessageBox.Show("Action Successful");
            //}

            //refesh the desktop

            SHChangeNotify(0x8000000, 0, IntPtr.Zero, IntPtr.Zero);

        }

        private int GetLcid()
        {
            //  System.Windows.Forms.MessageBox.Show("modify lcid");
            // add by huizhong , modify the language id in templates XML files for different office version (2052 for Chinese and 1033 for English)
            int lcid = 1033;
            try
            {
                Microsoft.Office.Interop.Word.Application thisApplication = new Microsoft.Office.Interop.Word.ApplicationClass();
                LanguageSettings languageSettings = (LanguageSettings)thisApplication.LanguageSettings;


                lcid = languageSettings.get_LanguageID(MsoAppLanguageID.msoLanguageIDUI);

                object missing = Type.Missing;
                thisApplication.Quit(ref missing, ref missing, ref missing);//dispose the winword
                thisApplication = null;
                GC.Collect();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                //System.Windows.Forms.MessageBox.show("lcid=" + lcid);
            }
            
            return lcid;
        }


        static private void doReplace(string fileFullName, string replacedBy, string findPattern)
        {
            string result = string.Empty;
            string inputText = string.Empty;
            string replacement = replacedBy;
            string pat = findPattern;

            //Regex r = new Regex(@"\[TARGETDIR\]");
            Regex r = new Regex(pat, RegexOptions.IgnoreCase);



            try
            {
                using (StreamReader sr = new StreamReader(fileFullName))
                {
                    inputText = sr.ReadToEnd();
                    Console.WriteLine("the inputText is: " + inputText);
                    Console.ReadLine();
                }

                // Compile the regular expression.
                if (r.IsMatch(inputText))
                {
                    result = r.Replace(inputText, replacement);

                    Console.WriteLine("\nthe result is: " + result);
                    Console.ReadLine();
                    // Add some text to the file.
                    using (StreamWriter sw = new StreamWriter(fileFullName))
                    {
                        sw.Write(result);
                    }
                }
                Console.WriteLine(fileFullName);
            }
            catch (Exception e)
            {

                Console.WriteLine("The process failed: {0}", e.ToString());
            }


        }
    }
}

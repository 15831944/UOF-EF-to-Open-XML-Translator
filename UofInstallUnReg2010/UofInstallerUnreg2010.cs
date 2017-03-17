using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Runtime.InteropServices;


namespace UofInstallUnReg2010
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:          卸载 COM Server
    ///     Author:             朴鑫 SYBASE v-xipia 
    ///     Create Date:        2009-02-05
    ///     Modified by:        
    ///     Last Modified Date: 2009-05-25
    /// </summary>
    [RunInstaller(true)]
    public partial class UofInstallerUnreg2010 : Installer
    {
        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        public UofInstallerUnreg2010()
        {
            InitializeComponent();
        }
        public override void  Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            string targetDir = this.Context.Parameters["TARGETDIR"].ToString();
            //System.Windows.Forms.MessageBox.Show(targetDir);


            string targetFile = targetDir + "UofTranslator2010.exe";
            //System.Windows.Forms.MessageBox.Show(targetFile);
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(@targetFile);
            //psi.RedirectStandardOutput = true; 
            psi.Arguments = "/unregserver";
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

            SHChangeNotify(0x8000000, 0, IntPtr.Zero, IntPtr.Zero);
        }
    }
}

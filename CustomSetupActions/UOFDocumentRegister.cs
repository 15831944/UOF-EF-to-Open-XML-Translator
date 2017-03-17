using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using System.Runtime.InteropServices;

namespace CustomSetupActions
{
    /// <summary>
    /// Register UOF file type in Open/Save As dialog of MS Office
    /// </summary>
    public class UOFDocumentRegister
    {
        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        [CustomAction]
        public static ActionResult RegisterFileTypeInOffice(Session session)
        {
            session.Log("Begin RegisterFileTypeInOffice");
            Record record = new Record(0);
            try
            {
                string targetDir = session.CustomActionData["InstallFolder"].ToString();
                string targetFile = targetDir + "UofTranslator2010.exe";
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(@targetFile);
                //psi.RedirectStandardOutput = true; 
                psi.Arguments = "/regserver";
                psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                psi.UseShellExecute = false;
                System.Diagnostics.Process listFiles;
                listFiles = System.Diagnostics.Process.Start(psi);
                listFiles.WaitForExit();

                SHChangeNotify(0x8000000, 0, IntPtr.Zero, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                record[0] = ex.Message;
                session.Message(InstallMessage.Error, record);
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        [CustomAction]
        public static ActionResult UnRegisterFileTypeInOffice(Session session)
        {
            session.Log("Begin RegisterFileTypeInOffice");
            Record record = new Record(0);
            
            try
            {
                string targetDir = session.CustomActionData["InstallFolder"].ToString();
                string targetFile = targetDir + "UofTranslator2010.exe";
                System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo(@targetFile);
                psi.Arguments = "/unregserver";
                psi.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                psi.UseShellExecute = false;
                System.Diagnostics.Process listFiles;
                listFiles = System.Diagnostics.Process.Start(psi);
                listFiles.WaitForExit();

                SHChangeNotify(0x8000000, 0, IntPtr.Zero, IntPtr.Zero);
                
            }
            catch (Exception ex)
            {
                record[0] = ex.Message;
                session.Message(InstallMessage.Error, record);
                return ActionResult.Failure;
            }

            return ActionResult.Success;
        }
    }
}

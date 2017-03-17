using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using System.Runtime.InteropServices;
using System.IO;

namespace CustomSetupActions
{
    /// <summary>
    /// Windows Shell Extension register
    /// </summary>
    public class ShellExtRegister
    {
        [CustomAction]
        public static ActionResult RegisterShellExtCOM(Session session)
        {
            session.Log("Begin RegisterShellExtCOM");
            Record record = new Record(0);

            try
            {
                string targetDir = session.CustomActionData["InstallFolder"].ToString();

                string runtimedir = Path.GetFullPath(Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), @"../.."));
                string to_exec = String.Concat(runtimedir, FrameworkLocation(), RuntimeEnvironment.GetSystemVersion(), @"\regasm.exe");
                          
                System.Diagnostics.Process proc = new System.Diagnostics.Process();

                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                proc.StandardInput.AutoFlush = true;
                proc.StandardInput.WriteLine(to_exec + " \"" + targetDir + "UofShellExt.dll\" /codebase");
                proc.StandardInput.WriteLine("exit");
                proc.WaitForExit();
                proc.Close();
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
        public static ActionResult UnRegisterShellExtCOM(Session session)
        {
            session.Log("Begin UnRegisterShellExtCOM");
            Record record = new Record(0);

            try
            {
                string targetDir = session.CustomActionData["InstallFolder"].ToString();

                string runtimedir = Path.GetFullPath(Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), @"../.."));
                string to_exec = String.Concat(runtimedir, FrameworkLocation(), RuntimeEnvironment.GetSystemVersion(), @"\regasm.exe");
                
                System.Diagnostics.Process proc = new System.Diagnostics.Process();

                proc.EnableRaisingEvents = false;
                proc.StartInfo.FileName = "cmd.exe";
                proc.StartInfo.UseShellExecute = false;
                proc.StartInfo.RedirectStandardInput = true;
                proc.StartInfo.RedirectStandardOutput = true;
                proc.StartInfo.RedirectStandardError = true;
                proc.StartInfo.CreateNoWindow = true;
                proc.Start();
                proc.StandardInput.AutoFlush = true;
                proc.StandardInput.WriteLine(to_exec + " \"" + targetDir + "UofShellExt.dll\" /u");
                proc.StandardInput.WriteLine("exit");
                proc.WaitForExit();
                proc.Close();
                
            }
            catch (Exception ex)
            {
                record[0] = ex.Message;
                session.Message(InstallMessage.Error, record);
                return ActionResult.Failure;
            }
            return ActionResult.Success;
        }

        public static string FrameworkLocation()
        {
            if (Environment.Is64BitOperatingSystem)
            {
                return "\\Framework64\\";
            }
            return "\\Framework\\";
        }
    }
}

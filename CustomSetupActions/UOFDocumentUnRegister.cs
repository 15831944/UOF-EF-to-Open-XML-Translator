using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Deployment.WindowsInstaller;
using System.Runtime.InteropServices;

namespace CustomSetupActions
{
    /// <summary>
    /// Unregister UOF file type in Open/Save As dialog of MS Office
    /// </summary>
    public class UOFDocumentUnRegister
    {
        [DllImport("shell32.dll")]
        public static extern void SHChangeNotify(uint wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);

        
    }
}

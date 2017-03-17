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
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.ComponentModel;
using UofAddinLib;

namespace UofCmdConverter
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:      ÃüÁîÐÐ×ª»»Æ÷
    ///     Author:         µË×·
    ///     Create Date:    2007-07-19
    /// </summary>
    public class Program
    {
        public enum ShowCommands : int
        {
            SW_HIDE = 0,
            SW_SHOWNORMAL = 1,
            SW_NORMAL = 1,
            SW_SHOWMINIMIZED = 2,
            SW_SHOWMAXIMIZED = 3,
            SW_MAXIMIZE = 3,
            SW_SHOWNOACTIVATE = 4,
            SW_SHOW = 5,
            SW_MINIMIZE = 6,
            SW_SHOWMINNOACTIVE = 7,
            SW_SHOWNA = 8,
            SW_RESTORE = 9,
            SW_SHOWDEFAULT = 10,
            SW_FORCEMINIMIZE = 11,
            SW_MAX = 11
        }

        [DllImport("shell32.dll")]
        public static extern IntPtr ShellExecute(
            IntPtr hwnd,
            string lpOperation,
            string lpFile,
            string lpParameters,
            string lpDirectory,
            ShowCommands nShowCmd);

        [STAThread]
        public static void Main(string[] args)
        {
            string[] outputFileName = null;
            string[] fileNames = null;
            MainDialog frmMain = null;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length == 2)
            {
                // Check The Param
                if (args[0].Equals("-o", StringComparison.OrdinalIgnoreCase))
                {
                    // Open an UOF File
                    // First Convert it.

                    // Check if the file exist
                    if (!File.Exists(args[1]))
                    {
                        Console.WriteLine(CmdConverterRes.ResourceManager.GetString("FileNotFoundPrompt"));
                        return;
                    }

                    // All ok, start the translation
                    // Get the Temp File Name
                    outputFileName = new string[1];
                    outputFileName[0] = System.IO.Path.GetTempPath()
                                        + System.IO.Path.GetFileNameWithoutExtension(args[1]);
                    FileState fState = new FileState();
                    fState.SourceFileName = args[1];
                    outputFileName[0] = MakeFileName(outputFileName[0], fState.getDestFileExtension());

                    fileNames = new string[1];
                    fileNames[0] = args[1];

                    string ext = Path.GetExtension(fileNames[0]);
                    if (ext.Contains("pptx") || ext.Contains("xlsx") || ext.Contains("docx"))
                    {
                        Console.WriteLine(CmdConverterRes.ResourceManager.GetString("OpenUOFFilePrompt"));
                        return;
                    }

                    // All ok, Invoke the Translator
                    frmMain = new MainDialog(null, DialogType.UofToOox, DocumentType.Unknown, fileNames, null, outputFileName);

                    Application.Run(frmMain);

                    if (frmMain.CurrentFileStates[0].IsSucceed)
                    {
                        // Succeed Translation
                        if (File.Exists(outputFileName[0]))
                        {
                            // Change the File to ReadOnly
                            File.SetAttributes(outputFileName[0], FileAttributes.ReadOnly);

                            // File Exist, Start Word
                            int execRet = ShellExecute(IntPtr.Zero, "open", outputFileName[0], "", "", ShowCommands.SW_SHOWDEFAULT).ToInt32();

                            if (execRet <= 32)
                            {
                                // Failed Open
                                string errorMessage = new Win32Exception(Marshal.GetLastWin32Error()).Message;
                                Console.WriteLine(errorMessage);
                            }

                            // Change the File to Not ReadOnly
                            File.SetAttributes(outputFileName[0], FileAttributes.Normal);
                        }
                        else
                        {

                            Console.WriteLine(CmdConverterRes.ResourceManager.GetString("TranslationFailedPrompt"));
                        }
                    }
                    else
                    {
                        Console.WriteLine(CmdConverterRes.ResourceManager.GetString("TranslationFailedPrompt"));
                    }

                    frmMain.Dispose();

                    return;
                }
                else
                {
                    PrintUsage();
                    return;
                }
            }
            else if (args.Length == 3)
            {
                // Check The Param
                if (args[0].Equals("-u2d", StringComparison.OrdinalIgnoreCase) || args[0].Equals("-d2u", StringComparison.OrdinalIgnoreCase))
                {
                    // Convert File

                    // Check if the file exist
                    if (!File.Exists(args[1]))
                    {
                        Console.WriteLine(CmdConverterRes.ResourceManager.GetString("FileNotFoundPrompt"));
                        return;
                    }

                    if (File.Exists(args[2]))
                    {
                        Console.WriteLine(CmdConverterRes.ResourceManager.GetString("FileAlreadyExistsPrompt"));
                        bool isDone = false;

                        while (!isDone)
                        {
                            string isOverwrite = Console.ReadLine();
                            if (isOverwrite.Equals("n", StringComparison.OrdinalIgnoreCase) ||
                                isOverwrite.Equals("no", StringComparison.OrdinalIgnoreCase) ||
                                isOverwrite.Equals("·ñ", StringComparison.OrdinalIgnoreCase))
                            {
                                return;
                            }
                            else if (isOverwrite.Equals("y", StringComparison.OrdinalIgnoreCase) ||
                                     isOverwrite.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                                     isOverwrite.Equals("ÊÇ", StringComparison.OrdinalIgnoreCase))
                            {
                                isDone = true;
                            }
                            else
                            {
                                Console.WriteLine(CmdConverterRes.ResourceManager.GetString("InputIllegalPrompt"));
                            }
                        }
                    }

                    // All ok, start the translation
                    outputFileName = new string[1];
                    outputFileName[0] = args[2];

                    fileNames = new string[1];
                    fileNames[0] = args[1];

                    if (args[0].Equals("-u2d", StringComparison.OrdinalIgnoreCase))
                    {
                        // All ok, Invoke the Translator
                        frmMain = new MainDialog(null, DialogType.UofToOox, DocumentType.Unknown, fileNames, null, outputFileName);
                    }
                    else
                    {
                        // All ok, Invoke the Translator
                        frmMain = new MainDialog(null, DialogType.OoxToUof, DocumentType.Unknown, fileNames, null, outputFileName);
                    }

                    Application.Run(frmMain);

                    if (frmMain.CurrentFileStates[0].IsSucceed)
                    {
                        // Succeed Translation
                        if (File.Exists(outputFileName[0]))
                        {
                            Console.WriteLine(CmdConverterRes.ResourceManager.GetString("TranslationSucceedPrompt"));
                        }
                        else
                        {
                            Console.WriteLine(CmdConverterRes.ResourceManager.GetString("TranslationFailedPrompt"));
                        }
                    }
                    else
                    {
                        Console.WriteLine(CmdConverterRes.ResourceManager.GetString("TranslationFailedPrompt"));
                    }

                    frmMain.Dispose();

                    return;
                }
                else
                {
                    PrintUsage();
                    return;
                }
            }
            else
            {
                PrintUsage();
                return;
            }
        }

        /// <summary>
        ///     Make A Docx File Name Based on the Original File Name
        /// </summary>
        /// <param name="baseFileName">
        ///     The Original File Name. For example, it can be "Test"
        ///     And the return will be like "Test.xxx"
        /// </param>
        /// <param name="extendFileName">
        ///     The Extend File Name. For example, it can be ".docx"
        ///     And the return will be like "xxx.docx"
        /// </param>
        /// <returns>
        ///     The File Name That is Made.
        /// </returns>
        ///     Author:         µË×·
        ///     Create Date:    2007-07-19
        private static string MakeFileName(string baseFileName, string extendFileName)
        {
            string outputFileName = baseFileName;
            string extFileName = extendFileName;
            List<string> currentOpening = new List<string>();
            UInt64 count = 1;
            bool isOK = false;

            while (!isOK)
            {
                isOK = true;
                if (isOK)
                {
                    try
                    {
                        if (System.IO.File.Exists(outputFileName + "(" + count.ToString() + ")" + extFileName))
                        {
                            // If the file already exist,test if it can be opened exclusively
                            System.IO.FileStream fs = System.IO.File.Open(outputFileName + "(" + count.ToString() + ")" + extFileName,
                                                                        System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite,
                                                                        System.IO.FileShare.None);
                            fs.Close();
                        }
                    }
                    catch (Exception)
                    {
                        // Oh We'll Change a name
                        count++;
                        isOK = false;
                    }
                }
            }

            outputFileName = outputFileName + "(" + count.ToString() + ")" + extFileName;
            return outputFileName;
        }

        /// <summary>
        ///     Print the Usage of this CmdLine Converter
        /// </summary>
        ///     Author:         µË×·
        ///     Create Date:    2007-07-20
        private static void PrintUsage()
        {
            Console.WriteLine(CmdConverterRes.ResourceManager.GetString("UsagePrompt",System.Globalization.CultureInfo.CurrentCulture));
  
        }
    }
}

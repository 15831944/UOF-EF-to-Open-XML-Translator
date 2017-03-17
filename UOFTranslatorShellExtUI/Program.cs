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
using System.Globalization;
using System.Windows.Forms;
using System.IO;
using System.Runtime.InteropServices;
using System.ComponentModel;
using UofAddinLib;

namespace UofStdConverter
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:      Explorer外壳扩展用户界面类
    ///     Author:         邓追
    ///     Create Date:    2007-06-07
    /// </summary>
    public static class Program
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
            DialogType translateType;
            string[] fileNames;
            string[] outputFileName = null;
            MainDialog frmMain = null;

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            if (args.Length < 1)
            {
                
                return;
            }
            else if (args.Length == 1)
            {
                try
                {
                    // Open the file for reading arguments
                    string argFileName = args[0];
                    FileStream argFileStream = File.Open(argFileName, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                    StreamReader argFileReader = new StreamReader(argFileStream);

                    // Get All arguments
                    string translateTypeString = argFileReader.ReadLine();
                    translateType = DialogType.OoxToUof;
                    if (String.Equals(translateTypeString, "OoxToUof"))
                    {
                        translateType = DialogType.OoxToUof;
                    }
                    else if (String.Equals(translateTypeString, "UofToOox"))
                    {
                        translateType = DialogType.UofToOox;
                    }
                    else
                    {
                        return;
                    }

                    // Get Params from Cmdline Args
                    fileNames = new string[Int32.Parse(argFileReader.ReadLine())];

                    for (int i = 0; i < fileNames.Length; i++)
                    {
                        fileNames[i] = argFileReader.ReadLine();
                    }

                    // All Done, Delete the Temp File
                    argFileReader.Close();
                    argFileStream.Close();
                    //File.Delete(argFileName);

                    ShowShellExtUI(translateType, fileNames);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
                finally
                {
                    
                }
            }
            else if (args.Length == 2)
            {
                // Check The Param
                if (args[0].Equals("-o", StringComparison.OrdinalIgnoreCase))
                {
                    // Open an UOF File
                    // First Convert it.

                    // Check if the file exist
                    if (!File.Exists(args[1]))
                    {
                        
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
                            
                            return;
                        }
                    }
                    else
                    {
                        
                        return;
                    }

                    frmMain.Dispose();
                    
                    return;
                }
            }
        }

        /// <summary>
        ///     Call This Function to Show Shell Extension UI
        /// </summary>
        /// <param name="translateType">
        ///     If Translate from Uof to Oox, this should be DialogType.UofToOox
        ///     If Translate from Oox to Uof, this should be DialogType.OoxToUof
        /// </param>
        /// <param name="fileNames">
        ///     The Full Path name(s) of the file(s) to be translated
        /// </param>
        ///     Author:         邓追
        ///     Create Date:    2007-06-10
        public static void ShowShellExtUI(DialogType translateType, string[] fileNames)
        {
            MainDialog frmMain = null;
            int i = 0;
            string[] targetFileNames = null;

            if (fileNames == null)
            {

                return;
            }

            if (((translateType != DialogType.OoxToUof) && (translateType != DialogType.UofToOox)) || (fileNames.Length == 0))
            {

                return;
            }

            //// Adjust Culture
            //if (StdConverterRes.ResourceManager.GetString("SaveAsUofDialogTitle") == null)
            //{
            //    // Not Support Language, Use English
            //    System.Threading.Thread.CurrentThread.CurrentUICulture = new CultureInfo("en");
            //}
            System.Threading.Thread.CurrentThread.CurrentUICulture = new CultureInfo("en");

            targetFileNames = new string[fileNames.Length];


            if (fileNames.Length == 1)
            {
                // Only One File, Let User Select Target File
                SaveFileDialog dialogSave = new SaveFileDialog();


                FileState fState = new FileState();
                fState.SourceFileName = fileNames[0];

                // Set The Title And The File Filter
                if (fState.TransType == TranslationType.UofToOox)
                {
                    if (fState.DocType == DocumentType.Word)
                    {
                       // MessageBox.Show("word document");
                        dialogSave.Title = StdConverterRes.ResourceManager.GetString("SaveAsOoxDialogTitle");
                        dialogSave.Filter = StdConverterRes.ResourceManager.GetString("OoxFileTypeDesc") + "(" + Extensions.OOX_WORD_SEARCH + ")|" + Extensions.OOX_WORD_SEARCH;
                        dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(fileNames[0]) + System.IO.Path.DirectorySeparatorChar +
                                                                       System.IO.Path.GetFileNameWithoutExtension(fileNames[0]), Extensions.OOX_WORD));
                    }
                    else if (fState.DocType == DocumentType.Excel)
                    {
                        dialogSave.Title = StdConverterRes.ResourceManager.GetString("SaveAsXlsxDialogTitle");
                        dialogSave.Filter = StdConverterRes.ResourceManager.GetString("XlsxFileTypeDesc") + "(" + Extensions.OOX_EXCEL_SEARCH + ")|" + Extensions.OOX_EXCEL_SEARCH;
                        dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(fileNames[0]) + System.IO.Path.DirectorySeparatorChar +
                                                                       System.IO.Path.GetFileNameWithoutExtension(fileNames[0]), Extensions.OOX_EXCEL));
                    }
                    else if (fState.DocType == DocumentType.Powerpnt)
                    {
                        dialogSave.Title = StdConverterRes.ResourceManager.GetString("SaveAsPptxDialogTitle");
                        dialogSave.Filter = StdConverterRes.ResourceManager.GetString("PptxFileTypeDesc") + "(" + Extensions.OOX_POWERPNT_SEARCH + ")|" + Extensions.OOX_POWERPNT_SEARCH;
                        dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(fileNames[0]) + System.IO.Path.DirectorySeparatorChar +
                                                                       System.IO.Path.GetFileNameWithoutExtension(fileNames[0]), Extensions.OOX_POWERPNT));
                    }
                    else
                    {

                        return;
                    }
                }
                else if (fState.TransType == TranslationType.OoxToUof)
                {
                    //dialogSave.Title = StdConverterRes.ResourceManager.GetString("SaveAsUofDialogTitle");
                    //dialogSave.Filter = StdConverterRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_SEARCH + ")|" + Extensions.UOF_SEARCH;
                    //dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(fileNames[0]) + System.IO.Path.DirectorySeparatorChar +
                    //                                                System.IO.Path.GetFileNameWithoutExtension(fileNames[0]), Extensions.UOF));
                   // MessageBox.Show("oox to uof");
                    if (fState.DocType == DocumentType.Word)
                    {
                      //  MessageBox.Show("word document");
                        dialogSave.Title = StdConverterRes.ResourceManager.GetString("SaveAsOoxDialogTitle");
                        dialogSave.Filter = StdConverterRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_WORD_SEARCH + ")|" + Extensions.UOF_WORD_SEARCH;
                        dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(fileNames[0]) + System.IO.Path.DirectorySeparatorChar +
                                                                       System.IO.Path.GetFileNameWithoutExtension(fileNames[0]), Extensions.UOF_WORD));
                    }
                    else if (fState.DocType == DocumentType.Excel)
                    {
                        dialogSave.Title = StdConverterRes.ResourceManager.GetString("SaveAsXlsxDialogTitle");
                        dialogSave.Filter = StdConverterRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_EXCEL_SEARCH + ")|" + Extensions.UOF_EXCEL_SEARCH;
                        dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(fileNames[0]) + System.IO.Path.DirectorySeparatorChar +
                                                                       System.IO.Path.GetFileNameWithoutExtension(fileNames[0]), Extensions.UOF_EXCEL));
                    }
                    else if (fState.DocType == DocumentType.Powerpnt)
                    {
                        dialogSave.Title = StdConverterRes.ResourceManager.GetString("SaveAsPptxDialogTitle");
                        dialogSave.Filter = StdConverterRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_POWERPNT_SEARCH + ")|" + Extensions.UOF_POWERPNT_SEARCH;
                        dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(fileNames[0]) + System.IO.Path.DirectorySeparatorChar +
                                                                       System.IO.Path.GetFileNameWithoutExtension(fileNames[0]), Extensions.UOF_POWERPNT));
                    }
                    else
                    {

                        return;
                    }
                }
                else
                {

                    return;
                }

                // Set Default Path and FileName
                dialogSave.InitialDirectory = Path.GetDirectoryName(fileNames[0]);

                // Show The Dialog and Get the result
                if (dialogSave.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    targetFileNames[0] = dialogSave.FileNames[0];

                    // All ok, Invoke the Translator
                    frmMain = new MainDialog(null, translateType, DocumentType.Unknown, fileNames, null, targetFileNames);

                    Application.Run(frmMain);

                    frmMain.Dispose();
                }

                dialogSave.Dispose();

            }
            else
            {
                // Many Files, Let User Select Target Directory
                FolderBrowserDialog dialogSaveFolder = new FolderBrowserDialog();

                dialogSaveFolder.ShowNewFolderButton = true;

                dialogSaveFolder.Description = StdConverterRes.ResourceManager.GetString("FolderSelectDialogTitle");
                dialogSaveFolder.SelectedPath = Path.GetDirectoryName(fileNames[0]);

                // Show The Dialog and Get the result
                if (dialogSaveFolder.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    // OK Now, Set the target Filenames
                    if (translateType == DialogType.UofToOox)
                    {
                        for (i = 0; i < fileNames.Length; i++)
                        {
                            FileState fState = new FileState();
                            fState.SourceFileName = fileNames[i];

                            targetFileNames[i] = dialogSaveFolder.SelectedPath +
                                                    Path.DirectorySeparatorChar +
                                                    Path.GetFileNameWithoutExtension(fileNames[i]) + fState.getDestFileExtension();
                        }
                    }
                    else
                    {
                        for (i = 0; i < fileNames.Length; i++)
                        {
                            //targetFileNames[i] = dialogSaveFolder.SelectedPath +
                            //                        Path.DirectorySeparatorChar +
                            //                        Path.GetFileNameWithoutExtension(fileNames[i]) + Extensions.UOF;
                            FileState fState = new FileState();
                            fState.SourceFileName = fileNames[i];
                            targetFileNames[i] = dialogSaveFolder.SelectedPath +
                                                    Path.DirectorySeparatorChar +
                                                    Path.GetFileNameWithoutExtension(fileNames[i]) +fState.getDestFileExtension() ;
                        }
                    }

                    // All ok, Invoke the Translator
                    frmMain = new MainDialog(null, translateType, DocumentType.Unknown, fileNames, null, targetFileNames);

                    Application.Run(frmMain);

                    frmMain.Dispose();
                }

                dialogSaveFolder.Dispose();

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
        ///     Author:         邓追
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

    }
}

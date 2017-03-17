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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Threading;

namespace UofAddinLib
{
    /// <summary>
    ///     Module ID:      
    ///     Depiction:      选择同名文件操作界面类
    ///                     其DialogResult返回值代表选择的操作
    ///                     DialogResult.Ignore 为 覆盖当前文件
    ///                     DialogResult.OK     为 覆盖所有重名文件
    ///                     DialogResult.Abort  为 跳过当前文件
    ///                     DialogResult.Cancel 为 跳过所有重命文件
    ///                     DialogResult.Retry  为 重命名当前文件
    ///     Author:         邓追
    ///     Create Date:    2007-07-04
    /// </summary>
    public partial class RenameOptionForm : Form
    {
        /// <summary>
        ///     The File Name that is conflicted
        /// </summary>
        private string conflictfileName;

        /// <summary>
        ///     The Translation Type
        /// </summary>
        private TranslationType translateType;

        private DocumentType docType;

        /// <summary>
        ///     The File Name Returned by this Form
        /// </summary>
        public string FileName
        {
            get { return conflictfileName; }
        }

        /// <summary>
        ///     Constructor of RenameOptionForm
        /// </summary>
        /// <param name="fileName">
        ///     The Name of the file that is conflicted with current one
        /// </param>
        /// <param name="dlgType">
        ///     If Translate from Uof to Oox, this should be DialogType.UofToOox
        ///     If Translate from Oox to Uof, this should be DialogType.OoxToUof
        /// </param>
        /// Author:         邓追
        /// Create Date:    2007-07-04
        public RenameOptionForm(string fileName, DocumentType docType, TranslationType transType)
        {

            InitializeComponent();

            // Store the fileName and Type
            conflictfileName = fileName;
            translateType = transType;
            this.docType = docType;
            
            // Set Icon of File
            picbIcon.Image = GetLargeIcon(fileName).ToBitmap();

            // Get File Info
            System.IO.FileInfo conflictFileInfo = new System.IO.FileInfo(fileName);

            lblFileName.Text = conflictFileInfo.Name;
            lblFileSize.Text += conflictFileInfo.Length + " " + TranslatorRes.ResourceManager.GetString("ByteDesc");
            lblFileCreatedTime.Text += conflictFileInfo.CreationTime.ToString();
            lblFileLastModified.Text += conflictFileInfo.CreationTime.ToString();
        }

        /// <summary>
        ///     Create Parameters, make this form no way to close
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                int CS_NOCLOSE = 0x200;
                CreateParams parameters = base.CreateParams;
                parameters.ClassStyle |= CS_NOCLOSE;

                return parameters;
            }
        }

        /// <summary>  
        ///     Get the Icon of a Path or file
        /// </summary>  
        /// <param name="path">
        ///     The Path of Directory or File
        /// </param>  
        /// <returns>
        ///     Icon Retrieved
        /// </returns>
        /// Author:         http://www.cnblogs.com/BG5SBK/articles/GetFileInfo.aspx
        /// Create Date:    2007-07-01
        public static Icon GetLargeIcon(string path)
        {
            FileInfomation info = new FileInfomation();
            FileInfo.GetFileInfo(path, 0, ref info, Marshal.SizeOf(info),
                            (int)(GetFileInfoFlags.SHGFI_ICON | GetFileInfoFlags.SHGFI_LARGEICON));
            try
            {
                return Icon.FromHandle(info.hIcon);
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        ///     Event Handler for the Rename Button
        /// </summary>
        /// <param name="sender">Event Sender</param>
        /// <param name="e">Event Parameter</param>
        /// Author:         邓追
        /// Create Date:    2007-07-04
        private void btnRename_Click(object sender, EventArgs e)
        {
            // Only One File, Let User Select Target File
            SaveFileDialog dialogSave = new SaveFileDialog();

            // Set The Title And The File Filter
            if (translateType == TranslationType.UofToOox)
            {
                if (docType == DocumentType.Word)
                {
                    dialogSave.Title = TranslatorRes.ResourceManager.GetString("SaveAsOoxDialogTitle");
                    dialogSave.Filter = TranslatorRes.ResourceManager.GetString("OoxFileTypeDesc") + "(" + Extensions.OOX_WORD_SEARCH + ")|" + Extensions.OOX_WORD_SEARCH;
                    dialogSave.FileName = System.IO.Path.GetFileName(FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(conflictfileName) + System.IO.Path.DirectorySeparatorChar +
                                                                   System.IO.Path.GetFileNameWithoutExtension(conflictfileName), Extensions.OOX_WORD_SEARCH));
                }
                else if (docType == DocumentType.Excel)
                {
                    dialogSave.Title = TranslatorRes.ResourceManager.GetString("SaveAsXlsxDialogTitle");
                    dialogSave.Filter = TranslatorRes.ResourceManager.GetString("XlsxFileTypeDesc") + "(" + Extensions.OOX_EXCEL_SEARCH + ")|" + Extensions.OOX_EXCEL_SEARCH;
                    dialogSave.FileName = System.IO.Path.GetFileName(FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(conflictfileName) + System.IO.Path.DirectorySeparatorChar +
                                                                   System.IO.Path.GetFileNameWithoutExtension(conflictfileName), Extensions.OOX_EXCEL_SEARCH));
                }
                else if (docType == DocumentType.Powerpnt)
                {
                    dialogSave.Title = TranslatorRes.ResourceManager.GetString("SaveAsPptxDialogTitle");
                    dialogSave.Filter = TranslatorRes.ResourceManager.GetString("PptxFileTypeDesc") + "(" + Extensions.OOX_POWERPNT_SEARCH + ")|" + Extensions.OOX_POWERPNT_SEARCH;
                    dialogSave.FileName = System.IO.Path.GetFileName(FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(conflictfileName) + System.IO.Path.DirectorySeparatorChar +
                                                                   System.IO.Path.GetFileNameWithoutExtension(conflictfileName), Extensions.OOX_POWERPNT_SEARCH));
                }
                else
                {
                    // Set the DialogResult and return
                    this.DialogResult = DialogResult.Abort;
                    return;
                }
            }
            else if (translateType == TranslationType.OoxToUof)
            {
                //dialogSave.Title = TranslatorRes.ResourceManager.GetString("SaveAsUofDialogTitle");
                //dialogSave.Filter = TranslatorRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_SEARCH + ")|" + Extensions.UOF_SEARCH;
                //dialogSave.FileName = System.IO.Path.GetFileName(FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(conflictfileName) + System.IO.Path.DirectorySeparatorChar +
                //                                                System.IO.Path.GetFileNameWithoutExtension(conflictfileName), Extensions.UOF_SEARCH));

                if (docType == DocumentType.Word)
                {
                   // MessageBox.Show("word document");
                    dialogSave.Title = TranslatorRes.ResourceManager.GetString("SaveAsOoxDialogTitle");
                    dialogSave.Filter = TranslatorRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_WORD_SEARCH + ")|" + Extensions.UOF_WORD_SEARCH;
                    dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(conflictfileName) + System.IO.Path.DirectorySeparatorChar +
                                                                   System.IO.Path.GetFileNameWithoutExtension(conflictfileName), Extensions.UOF_WORD));

                    MessageBox.Show(dialogSave.FileName);
                }
                else if (docType == DocumentType.Excel)
                {
                    dialogSave.Title = TranslatorRes.ResourceManager.GetString("SaveAsXlsxDialogTitle");
                    dialogSave.Filter = TranslatorRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_EXCEL_SEARCH + ")|" + Extensions.UOF_EXCEL_SEARCH;
                    dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(conflictfileName) + System.IO.Path.DirectorySeparatorChar +
                                                                   System.IO.Path.GetFileNameWithoutExtension(conflictfileName), Extensions.UOF_EXCEL));
                }
                else if (docType == DocumentType.Powerpnt)
                {
                    dialogSave.Title = TranslatorRes.ResourceManager.GetString("SaveAsPptxDialogTitle");
                    dialogSave.Filter = TranslatorRes.ResourceManager.GetString("UOFFileTypeDesc") + "(" + Extensions.UOF_POWERPNT_SEARCH + ")|" + Extensions.UOF_POWERPNT_SEARCH;
                    dialogSave.FileName = System.IO.Path.GetFileName(UofAddinLib.FileInfo.MakeFileName(System.IO.Path.GetDirectoryName(conflictfileName) + System.IO.Path.DirectorySeparatorChar +
                                                                   System.IO.Path.GetFileNameWithoutExtension(conflictfileName), Extensions.UOF_POWERPNT));
                }
            }
            else
            {
                this.DialogResult = DialogResult.Abort;
                return;
            }

            // Set Default Path and FileName
            dialogSave.InitialDirectory = System.IO.Path.GetDirectoryName(conflictfileName);

            // Show The Dialog and Get the result
            if (dialogSave.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                conflictfileName = dialogSave.FileNames[0];

                // Set the DialogResult and return
                this.DialogResult = DialogResult.Retry;
            }
        }

        private void RenameOptionForm_Load(object sender, EventArgs e)
        {
        }

        private void RenameOptionForm_FormClosed(object sender, FormClosedEventArgs e)
        {
        }

        private void btnOverwrite_Click(object sender, EventArgs e)
        {

        }

        private void btnOverwriteAll_Click(object sender, EventArgs e)
        {

        }

        private void btnSkipAll_Click(object sender, EventArgs e)
        {

        }
    }
}
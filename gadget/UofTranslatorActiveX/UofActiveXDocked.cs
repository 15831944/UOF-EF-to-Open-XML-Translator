using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;
using UofAddinLib;
using Act.UofTranslator.UofTranslatorLib;

namespace UofTranslatorActiveX
{
    public partial class UofActiveXDocked : UserControl
    {
        #region delegates

        public delegate void SetProgressBarMoveDelegate();
        public SetProgressBarMoveDelegate mySetProgressBarMove;
        public SetProgressBarMoveDelegate mySetProgressBarStop;

        public delegate void SetProgress(int totalfiles, int currentfile);
        public SetProgress mySetProgress;

        public delegate void IncProgressBlock();
        public IncProgressBlock myIncProgressBlock;

        public delegate void WriteLog(string log, bool clear);
        public WriteLog myWriteLog;

        public delegate void TransferState(FileState state);
        public TransferState myTransferState;

        private EventHandler progressListener;

        private EventHandler feedbackListenerForLog;

        #endregion

        # region members

        private Bitmap fileIcon;

        public Bitmap FileIcon
        {
            get { return fileIcon; }
            set { fileIcon = value; }
        }

        private FileState totalState;

        public FileState TotalState
        {
            get { return totalState; }
        }

        private int totalFiles;
        private int currentFile;
        private int currentBlockProgress;
        private string logText;

        public string LogText
        {
            get { return logText; }
        }

        private Thread workingThread;

        public Thread WorkingThread
        {
            get { return workingThread; }
            set { workingThread = value; }
        }

        # endregion

        #region For Getting File Icon

        [DllImport("shell32.dll", EntryPoint = "SHGetFileInfo", CharSet = CharSet.Auto)]
        public static extern int GetFileInfo(string pszPath, int dwFileAttributes,
         ref  FileInfomation psfi, int cbFileInfo, int uFlags);

        [StructLayout(LayoutKind.Sequential)]
        public struct FileInfomation
        {
            public IntPtr hIcon;
            public int iIcon;
            public int dwAttributes;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
        }

        private enum GetFileInfoFlags : int
        {
            SHGFI_ICON = 0x000000100,               //  get icon 
            SHGFI_DISPLAYNAME = 0x000000200,        //  get display name 
            SHGFI_TYPENAME = 0x000000400,           //  get type name 
            SHGFI_ATTRIBUTES = 0x000000800,         //  get attributes 
            SHGFI_ICONLOCATION = 0x000001000,       //  get icon location 
            SHGFI_EXETYPE = 0x000002000,            //  return exe type 
            SHGFI_SYSICONINDEX = 0x000004000,       //  get system icon index 
            SHGFI_LINKOVERLAY = 0x000008000,        //  put a link overlay on icon 
            SHGFI_SELECTED = 0x000010000,           //  show icon in selected state 
            SHGFI_ATTR_SPECIFIED = 0x000020000,     //  get only specified attributes 
            SHGFI_LARGEICON = 0x000000000,          //  get large icon 
            SHGFI_SMALLICON = 0x000000001,          //  get small icon 
            SHGFI_OPENICON = 0x000000002,           //  get open icon 
            SHGFI_SHELLICONSIZE = 0x000000004,      //  get shell size icon 
            SHGFI_PIDL = 0x000000008,               //  pszPath is a pidl 
            SHGFI_USEFILEATTRIBUTES = 0x000000010,  //  use passed dwFileAttribute 
            SHGFI_ADDOVERLAYS = 0x000000020,        //  apply the appropriate overlays 
            SHGFI_OVERLAYINDEX = 0x000000040        //  Get the index of the overlay 
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
        private Icon GetIcon(string path)
        {
            FileInfomation info = new FileInfomation();
            GetFileInfo(path, 0, ref info, Marshal.SizeOf(info),
                        (int)(GetFileInfoFlags.SHGFI_ICON));
            try
            {
                return Icon.FromHandle(info.hIcon);
            }
            catch
            {
                return null;
            }
        }

        #endregion

        public UofActiveXDocked()
        {
            InitializeComponent();
            mySetProgress = new SetProgress(SetProgressCore);
            myWriteLog = new WriteLog(WriteLogCore);
            myIncProgressBlock = new IncProgressBlock(IncProgressBlockCore);
            progressListener = new EventHandler(setProgressBlock);
            feedbackListenerForLog = new EventHandler(WriteConverterLog);
            mySetProgressBarMove = new SetProgressBarMoveDelegate(SetProgressMoveCore);
            mySetProgressBarStop = new SetProgressBarMoveDelegate(SetProgressStopCore);

            workingThread = null;
            logText = "";
        }

        #region Handle DragDrop

        private void listView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            totalState = new FileState();
            totalState.StateValProgress = StateVal.STATE_READY;
            totalState.SourceFileName = s[0];

            if (s.Length > 1)
            {
                MessageBox.Show("You just can drag one item into this box", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (s.Length == 1)
            {
                listView1.AllowDrop = false;
                this.Transfer1.Enabled = true;

                string filename = Path.GetFileName(s[0]);
                Icon icon = GetIcon(s[0]);
                if (icon == null)
                {
                    icon = IconResource.Default;
                }

                fileIcon = icon.ToBitmap();
                this.imageList1.Images.Add(fileIcon);
                this.listView1.Items.Add(new ListViewItem(filename, 0));
            }

            else
                return;
        }

        private void listView1_DragEnter(object sender, DragEventArgs e)
        {
            // Only Files can be drag into here
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void listView1_ItemDrag(object sender, ItemDragEventArgs e)
        {

            List<string> filelist = new List<string>();
            ListViewItem item = listView1.SelectedItems[0];

            if (this.totalState.StateValProgress == StateVal.STATE_DONE)
            {
                if (File.Exists(this.totalState.DestinationFileName) || Directory.Exists(this.totalState.DestinationFileName))
                {
                    filelist.Add(this.totalState.DestinationFileName);
                }
            }

            if (filelist.Count == 0)
            {
                return;
            }

            string[] files = filelist.ToArray();
            DragDropEffects effect = DoDragDrop(new DataObject(DataFormats.FileDrop, files), DragDropEffects.Copy);

            onRemove(item);
            listView1.Items.Remove(item);
            this.imageList1.Images.RemoveAt(0);
            this.listView1.AllowDrop = true;
            this.Invoke(myWriteLog, new object[] { "", true });
            this.Transfer1.Image = IconResource.Transfer1;
            this.Transfer1.Enabled = false;
            this.Info1.Enabled = false;

            
            this.progressBar1.Value = 1;
            totalState.StateValProgress = StateVal.STATE_READY;

            return;
        }

        #endregion

        #region  Handle Delete Key Press

        private void listView1_KeyUp(object sender, KeyEventArgs e)
        {
            if ((listView1.SelectedItems.Count != 0) && (e.KeyCode == Keys.Delete))
            {
                ListViewItem item = listView1.SelectedItems[0];
                onRemove(item);
                listView1.Items.Remove(item);
                this.imageList1.Images.RemoveAt(0);
                listView1.AllowDrop = true;
            }
            Transfer1.Image = IconResource.Transfer1;
            this.Transfer1.Enabled = false;
            this.Info1.Enabled = false;
            this.Invoke(myWriteLog, new object[] { "", true });
            
            this.progressBar1.Value = 1;
            totalState.StateValProgress = StateVal.STATE_READY;

        }

        private void onRemove(ListViewItem item)
        {
            if (this.totalState.StateValProgress == StateVal.STATE_DONE)
            {
                if (!String.IsNullOrEmpty(this.totalState.DestinationFileName))
                {
                      string parentPath;
                      parentPath = Path.GetDirectoryName(this.totalState.DestinationFileName);
                      if (Directory.Exists(parentPath))
                      {
                          Directory.Delete(parentPath, true);
                      }
                  }
              }
          }

        #endregion

        private void Transfer1_Click(object sender, EventArgs e)
        {
            try
            {
                this.Transfer1.Enabled = false;
                workingThread = new Thread(doDockWork);
                workingThread.Start();

                totalState.StateValProgress = StateVal.STATE_WORKING;

                progressBar1.Value = 1;
            }
            catch (Exception excep)
            {
                MessageBox.Show(excep.ToString());
            }
        }

        private void Info1_Click(object sender, EventArgs e)
        {
            MessageBox.Show(logText,"Transfer Error Information", MessageBoxButtons.OK);
        }

        private void doDockWork()
        {
            int filesCount = 0;
            int currentCount = 0;
            int successFiles = 0;

            try
            {
                IUOFTranslator converter = null;

                totalState.IsSucceed = false;
                totalState.HasWarning = false;

                Invoke(mySetProgressBarMove);

                if (Directory.Exists(totalState.SourceFileName))
                {
                    // Is Directory, Translate Each File

                    // Use UUID to create top dir
                    string srcroot = totalState.SourceFileName;
                    string topdir = Guid.NewGuid().ToString() + Path.DirectorySeparatorChar + new DirectoryInfo(srcroot).Name;

                    totalState.DestinationFileName = Path.GetTempPath() + topdir;

                    List<string> files = new List<string>();

                  //  files.AddRange(Directory.GetFiles(srcroot,Extensions.UOF_SEARCH_PRECISION,SearchOption.AllDirectories));
                    files.AddRange(Directory.GetFiles(srcroot, Extensions.UOF_EXCEL_SEARCH, SearchOption.AllDirectories));
                    files.AddRange(Directory.GetFiles(srcroot, Extensions.UOF_POWERPNT_SEARCH, SearchOption.AllDirectories));
                    files.AddRange(Directory.GetFiles(srcroot, Extensions.UOF_WORD_SEARCH, SearchOption.AllDirectories));
                    files.AddRange(Directory.GetFiles(srcroot,Extensions.OOX_WORD_SEARCH,SearchOption.AllDirectories));
                    files.AddRange(Directory.GetFiles(srcroot,Extensions.OOX_EXCEL_SEARCH,SearchOption.AllDirectories));
                    files.AddRange(Directory.GetFiles(srcroot,Extensions.OOX_POWERPNT_SEARCH,SearchOption.AllDirectories));

                    if (files.Count == 0)
                    {
                        goto Finish;
                    }

                    filesCount = files.Count;
                    currentCount = 0;
                    this.Invoke(mySetProgress, new object[] { filesCount, currentCount });

                    foreach (string file in files)
                    {
                        FileState fileState = new FileState();
                        fileState.SourceFileName = Path.GetFullPath(file);
                        fileState.DestinationFileName = totalState.DestinationFileName +
                                                        file.Substring(srcroot.Length);
                        fileState.adjustDestFileExtension();
                        converter = initConverter(fileState.SourceFileName, fileState.DocType);
                        doTranslateFile(fileState, converter);

                        if (fileState.IsSucceed)
                        {
                            totalState.IsSucceed = true;
                            successFiles++;
                        }
                        else
                        {
                            totalState.HasWarning = true;
                            this.Invoke(myWriteLog, new object[] { "错误:文件" + fileState.SourceFileName + "转换失败." , false});
                        }

                        currentCount++;
                        this.Invoke(mySetProgress, new object[] { filesCount, currentCount });
                    }
                }
                else
                {
                    filesCount = 1;
                    currentCount = 0;

                    if (totalState.DocType == DocumentType.Unknown)
                    {
                        this.Invoke(myWriteLog, new object[] { "错误:文件" + totalState.SourceFileName + "转换失败.", false });
                        goto Finish;
                    }

                    this.Invoke(mySetProgress, new object[] { filesCount, currentCount });

                    string topdir = Guid.NewGuid().ToString();

                    totalState.DestinationFileName = Path.GetTempPath() + topdir + Path.DirectorySeparatorChar + Path.GetFileName(totalState.SourceFileName);
                    totalState.adjustDestFileExtension();
                    converter = initConverter(totalState.SourceFileName, totalState.DocType);
                    doTranslateFile(totalState, converter);

                    currentCount ++;

                    if (totalState.IsSucceed == true)
                        successFiles ++;

                    this.Invoke(mySetProgress, new object[] { filesCount, currentCount });
                }

            Finish:
                this.Invoke(mySetProgressBarStop);
                this.Invoke(mySetProgress, new object[] { filesCount, currentCount });
                this.Transfer1.Enabled = true;
                totalState.StateValProgress = StateVal.STATE_DONE;
            }
            catch (ThreadAbortException ex)
            {
                this.Invoke(mySetProgressBarStop);
                MessageBox.Show(ex.ToString());
            }
            if (successFiles == filesCount)
            {
                Transfer1.Image = IconResource.OK1;
            }
            else if (this.totalState.HasWarning == true && successFiles > 0)
            {
                Transfer1.Image = IconResource.Warning1;
                Info1.Enabled = true;
            }
            else
            {
                Transfer1.Image = IconResource.Error1;
                Info1.Enabled = true;
            }
        }

        private IUOFTranslator initConverter(string inputFileName, DocumentType type)
        {
            IUOFTranslator Converter;
            if (type == DocumentType.Word)
            {
                Converter = TranslatorFactory.CheckFileType(DocType.Word);
                Converter.AddProgressMessageListener(this.progressListener);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                return Converter;
            }
            else if (type == DocumentType.Excel)
            {
                Converter = TranslatorFactory.CheckFileType(DocType.Excel);
                Converter.AddProgressMessageListener(this.progressListener);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                return Converter;
            }
            else if (type == DocumentType.Powerpnt)
            {
                Converter = TranslatorFactory.CheckFileType(DocType.Powerpoint);
                Converter.AddProgressMessageListener(this.progressListener);
                Converter.AddFeedbackMessageListener(this.feedbackListenerForLog);
                return Converter;
            }

            return null;
        }

        private void doTranslateFile(FileState state, IUOFTranslator converter)
        {
            if (converter == null)
            {
                state.IsSucceed = false;
                this.Invoke(myWriteLog, new object[] { "错误:文件" + state.SourceFileName + "转换失败." , false});
                return;
            }

            try
            {
                if (File.Exists(state.SourceFileName))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(state.DestinationFileName));
                    if (state.TransType == TranslationType.UofToOox)
                    {
                        //TODO:现在的ComputeSize不好用
                        //converter.UofToOoxComputeSize(state.SourceFileName);
                        this.Invoke(myWriteLog, new object[] { "开始转换文件" + state.SourceFileName, false });
                        converter.UofToOox(state.SourceFileName, state.DestinationFileName);
                    }
                    else if (state.TransType == TranslationType.OoxToUof)
                    {
                        //TODO:现在的ComputeSize不好用
                        //converter.OoxToUofComputeSize(state.SourceFileName);
                        this.Invoke(myWriteLog, new object[] { "开始转换文件" + state.SourceFileName, false });
                        converter.OoxToUof(state.SourceFileName, state.DestinationFileName);
                    }
                    else
                    {
                        state.IsSucceed = false;
                        return;
                    }
                }
                state.IsSucceed = true;

                //TODO: Set Warning
                state.HasWarning = false;
            }
            catch (Exception e)
            {
                this.Invoke(myWriteLog, new object[] { "转换中出现异常:" + e.ToString(), false });
                state.IsSucceed = false;
            }
        }

        #region delegates implementation

        private void adjustProgressBar()
        {
            double totalPer = 0;

            if (totalFiles != 0)
            {
                double currentBlockPer = (double)this.currentBlockProgress / 100.0;
                totalPer = ((double)this.currentFile + currentBlockPer) / (double)totalFiles;
            }

            this.progressBar1.Value = this.progressBar1.Minimum + (int)Math.Floor(totalPer * (double)(this.progressBar1.Maximum - this.progressBar1.Minimum));
        }

        public void SetProgressMoveCore()
        {
            this.progressBar1.MarqueeAnimationSpeed = 100;
        }

        public void SetProgressStopCore()
        {
            this.progressBar1.MarqueeAnimationSpeed = 0;
        }

        public void IncProgressBlockCore()
        {
            this.currentBlockProgress++;

            adjustProgressBar();
        }

        public void SetProgressCore(int totalfiles, int currentfile)
        {
            this.totalFiles = totalfiles;
            this.currentFile = currentfile;
            this.currentBlockProgress = 0;

            adjustProgressBar();
        }

        public void WriteLogCore(string log, bool clear)
        {
            if (clear)
            {
                logText = "";
            }
            else
            {
                logText += log;
                logText += System.Environment.NewLine;
            }
        }

        #endregion

        #region Event Handler For Converter

        private void setProgressBlock(object sender, EventArgs e)
        {
            this.Invoke(myIncProgressBlock);
        }

        private void WriteConverterLog(object sender, EventArgs e)
        {
            object[] param = new object[1];
            param[0] = ((UofEventArgs)e).Message;

            if (((UofEventArgs)e).Message != null)
            {
                this.Invoke(myWriteLog, new object[] { param, false });
            }
        }

        #endregion

        private void GadgetHelp_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.MessageBox.Show("Drag&Drop MS Office 2010 or UOF2.0 file to the blank, then click start button","UOFTranslator4 Gadget Help",MessageBoxButtons.OK,MessageBoxIcon.Information);
            
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.Win32;
using UofAddinLib;
using Act.UofTranslator.UofTranslatorLib;

namespace UofTranslatorActiveX
{
    public class WorkingListViewItem : ListViewItem
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

        #region members

        private FileState state;

        public FileState State
        {
            get { return state; }
        }

        private string displayName;

        public string DisplayName
        {
            get { return displayName; }
            set { displayName = value;}
        }

        private Bitmap fileIcon;

        public Bitmap FileIcon
        {
            get { return fileIcon; }
            set { fileIcon = value; }
        }

        private UofProgressBar pgbItem;

        public UofProgressBar PgbItem
        {
            get { return pgbItem; }
            set { pgbItem = value; }
        }

        private UofStatusButton btnItem;

        public UofStatusButton BtnItem
        {
            get { return btnItem; }
            set { btnItem = value; }
        }

        private Thread workingThread;

        public Thread WorkingThread
        {
            get { return workingThread; }
            set { workingThread = value; }
        }

        private int totalFiles;
        private int currentFile;
        private int currentBlockProgress;
        private string logText;

        public string LogText
        {
            get { return logText; }
        }

        private TextBox txtbxShow;

        #endregion

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
        private Icon GetSmallIcon(string path)
        {
            FileInfomation info = new FileInfomation();
            GetFileInfo(path, 0, ref info, Marshal.SizeOf(info),
                        (int)(GetFileInfoFlags.SHGFI_ICON | GetFileInfoFlags.SHGFI_SMALLICON));
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

        public WorkingListViewItem(string rootfile, ListView parent, TextBox txtbtShow)
        {
            //initializing delegates
            mySetProgress = new SetProgress(SetProgressCore);
            myWriteLog = new WriteLog(WriteLogCore);
            myTransferState = new TransferState(TransferStateCore);
            myIncProgressBlock = new IncProgressBlock(IncProgressBlockCore);
            progressListener = new EventHandler(setProgressBlock);
           // feedbackListenerForLog = new EventHandler(WriteConverterLog);
            mySetProgressBarMove = new SetProgressBarMoveDelegate(SetProgressMoveCore);
            mySetProgressBarStop = new SetProgressBarMoveDelegate(SetProgressStopCore);

            state = new FileState();
            state.StateValProgress = StateVal.STATE_READY;

            this.state.SourceFileName = rootfile;
            this.ToolTipText = rootfile;
            this.displayName = Path.GetFileName(rootfile);
            Icon icon = GetSmallIcon(rootfile);
            if (icon == null)
            {
                icon = IconResource.Default;
            }

            fileIcon = icon.ToBitmap();

            //Create Working Thread
            workingThread = null;

            //For Status Button
            ListViewSubItem statusSub = this.SubItems[0];
            btnItem = new UofStatusButton(this);
            btnItem.Tag = this.ListView;
            parent.Controls.Add(btnItem);
            ((UofListView)parent).AddControlToSubItem(btnItem, statusSub, true);

            //For name & Progress Bar
            ListViewSubItem nameSub = this.SubItems.Add("");
            pgbItem = new UofProgressBar(this);
            pgbItem.Minimum = 1;
            pgbItem.Maximum = 1000;
            pgbItem.Step = 1;
            pgbItem.Tag = this.ListView;
            pgbItem.Showtext = displayName;
            parent.Controls.Add(pgbItem);
            ((UofListView)parent).AddControlToSubItem(pgbItem, nameSub, false);

            totalFiles = 0;
            currentFile = 0;
            logText = "";

           // this.txtbxShow = txtbxShow;
        }

        private void doWork(object tState)
        {
            try
            {
                FileState totalState = (FileState)tState;
                IUOFTranslator converter = null;


                totalState.IsSucceed = false;
                totalState.HasWarning = false;

                int filesCount = 0;
                int currentCount = 0;

                ListView.Invoke(mySetProgressBarMove);

                if (Directory.Exists(totalState.SourceFileName))
                {
                    // Is Directory, Translate Each File

                    // Use UUID to create top dir
                    string srcroot = totalState.SourceFileName;
                    string topdir = Guid.NewGuid().ToString() + Path.DirectorySeparatorChar + new DirectoryInfo(srcroot).Name;

                    totalState.DestinationFileName = Path.GetTempPath() + topdir;

                    List<string> files = new List<string>();
                    bool isAll = false;
                    string searchpattern = Extensions.ALL_SEARCH;
                    if (totalState.TransType == TranslationType.UofToOox)
                    {
                        //searchpattern = Extensions.UOF_SEARCH_PRECISION;
                        if (totalState.DocType == DocumentType.Word)
                        {
                            searchpattern = Extensions.UOF_WORD_SEARCH;
                        }
                        else if (totalState.DocType == DocumentType.Excel)
                        {
                            searchpattern = Extensions.UOF_EXCEL_SEARCH;
                        }
                        else if (totalState.DocType == DocumentType.Powerpnt)
                        {
                            searchpattern = Extensions.UOF_POWERPNT_SEARCH;
                        }
                        else
                        {
                            files.AddRange(Directory.GetFiles(srcroot, Extensions.UOF_POWERPNT_SEARCH, SearchOption.AllDirectories));
                            files.AddRange(Directory.GetFiles(srcroot, Extensions.UOF_EXCEL_SEARCH, SearchOption.AllDirectories));
                            files.AddRange(Directory.GetFiles(srcroot, Extensions.UOF_WORD_SEARCH, SearchOption.AllDirectories));
                            isAll = true;
                        }
                    }
                    else if (totalState.TransType == TranslationType.OoxToUof)
                    {
                        if (totalState.DocType == DocumentType.Word)
                        {
                            searchpattern = Extensions.OOX_WORD_SEARCH;
                        }
                        else if (totalState.DocType == DocumentType.Excel)
                        {
                            searchpattern = Extensions.OOX_EXCEL_SEARCH;
                        }
                        else if (totalState.DocType == DocumentType.Powerpnt)
                        {
                            searchpattern = Extensions.OOX_POWERPNT_SEARCH;
                        }
                        else
                        {
                            files.AddRange(Directory.GetFiles(srcroot, Extensions.OOX_WORD_SEARCH, SearchOption.AllDirectories));
                            files.AddRange(Directory.GetFiles(srcroot, Extensions.OOX_EXCEL_SEARCH, SearchOption.AllDirectories));
                            files.AddRange(Directory.GetFiles(srcroot, Extensions.OOX_POWERPNT_SEARCH, SearchOption.AllDirectories));
                            isAll = true;
                        }
                    }
                    else
                    {
                        goto Finish;
                    }

                    if (!isAll)
                    {
                        files.AddRange(Directory.GetFiles(srcroot, searchpattern, SearchOption.AllDirectories));
                    }

                    if (files.Count == 0)
                    {
                        goto Finish;
                    }

                    filesCount = files.Count;
                    currentCount = 0;
                    ListView.Invoke(mySetProgress, new object[] { filesCount, currentCount });

                    foreach (string file in files)
                    {                        
                        FileState fileState = new FileState();
                        fileState.SourceFileName = Path.GetFullPath(file);
                        fileState.DestinationFileName = totalState.DestinationFileName +
                                                        file.Substring(srcroot.Length);
                        fileState.adjustDestFileExtension();

                        if (totalState.DocType != DocumentType.All)
                        {
                            if ((totalState.TransType == TranslationType.UofToOox) && (totalState.DocType != fileState.DocType))
                            {
                                //跳过这个UOF文件
                                continue;
                            }
                        }

                        converter = initConverter(fileState.SourceFileName, fileState.DocType);
                        doTranslateFile(fileState, converter);

                        if (fileState.IsSucceed)
                        {
                            totalState.IsSucceed = true;
                        }
                        else
                        {
                            totalState.HasWarning = true;
                            ListView.Invoke(myWriteLog, new object[] { "Error:File" + fileState.SourceFileName + "translate false." , false});
                        }

                        currentCount++;
                        ListView.Invoke(mySetProgress, new object[] { filesCount, currentCount });
                    }
                }
                else
                {
                    filesCount = 1;
                    currentCount = 0;

                    ListView.Invoke(mySetProgress, new object[] { filesCount, currentCount });

                    string topdir = Guid.NewGuid().ToString();

                    totalState.DestinationFileName = Path.GetTempPath() + topdir + Path.DirectorySeparatorChar + Path.GetFileName(totalState.SourceFileName);
                    totalState.adjustDestFileExtension();
                    converter = initConverter(totalState.SourceFileName, totalState.DocType);
                    doTranslateFile(totalState, converter);

                    currentCount ++;

                    ListView.Invoke(mySetProgress, new object[] { filesCount, currentCount });

                    if (!totalState.IsSucceed)
                    {
                        ListView.Invoke(myWriteLog, new object[] { "Error:File" + totalState.SourceFileName + "translate false.", false });
                    }
                }

            Finish:
                ListView.Invoke(mySetProgressBarStop);
                ListView.Invoke(mySetProgress, new object[] { filesCount, currentCount });
                totalState.StateValProgress = StateVal.STATE_DONE;
                ListView.Invoke(myTransferState, new object[] { totalState.clone() });
            }
            catch (ThreadAbortException ex)
            {
                ListView.Invoke(mySetProgressBarStop);
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
                return;
            }

            try
            {
                if (File.Exists(state.SourceFileName))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(state.DestinationFileName)); 
                    if(state.TransType == TranslationType.UofToOox)
                    {
                        //TODO:现在的ComputeSize不好用
                        //converter.UofToOoxComputeSize(state.SourceFileName);
                        ListView.Invoke(myWriteLog, new object[] { "Translating " + state.SourceFileName, false });
                        converter.UofToOox(state.SourceFileName, state.DestinationFileName);
                    }
                    else if (state.TransType == TranslationType.OoxToUof)
                    {
                        //TODO:现在的ComputeSize不好用
                        //converter.OoxToUofComputeSize(state.SourceFileName);
                        ListView.Invoke(myWriteLog, new object[] { "Translating " + state.SourceFileName, false });
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
               // ListView.Invoke(myWriteLog, new object[] { "转换中出现异常:" + e.ToString(), false });
                state.IsSucceed = false;
            }
        }

        public void start()
        {
            if (workingThread == null)
            {
                //button
                this.State.StateValProgress = StateVal.STATE_WORKING;
                btnItem.toogleIcon();

                //progress bar
                pgbItem.Value = 1;

                workingThread = new Thread(new ParameterizedThreadStart(doWork));
                workingThread.Start(this.State.clone());
            }
        }

        public void stop()
        {
            if (workingThread.IsAlive)
            {
                workingThread.Abort();

                this.State.StateValProgress = StateVal.STATE_READY;
                btnItem.toogleIcon();
                ListView.Invoke(myWriteLog, new object[] { "User stop translating  " + state.SourceFileName+".", false });
                workingThread = null;
            }
        }

        private void adjustProgressBar()
        {
            double totalPer = 0;

            if (totalFiles != 0)
            {
                double currentBlockPer = (double)this.currentBlockProgress / 100.0;
                totalPer = ((double)this.currentFile + currentBlockPer) / (double)totalFiles;
            }

            pgbItem.Value = pgbItem.Minimum + (int)Math.Floor(totalPer * (double)(pgbItem.Maximum - pgbItem.Minimum));
            /* no need to show percentage any more
            pgbItem.Showtext = String.Format("{0} ({1:P})", this.DisplayName, totalPer);
            */
            pgbItem.Showtext = String.Format("{0}", this.DisplayName);
        }

        #region delegates implementation

        public void SetProgressMoveCore()
        {
            this.pgbItem.MarqueeAnimationSpeed = 100;
        }

        public void SetProgressStopCore()
        {
            this.pgbItem.MarqueeAnimationSpeed = 0;
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

            if (ListView.SelectedItems.Count > 0)
            {
                if (ListView.SelectedItems[0] == this)
                {
                    txtbxShow.Text = logText;
                }
            }
        }

        public void TransferStateCore(FileState state)
        {
            lock (this.state)
            {
                this.state = state;
            }

            btnItem.toogleIcon();
        }

        #endregion

        #region Event Handler For Converter

        private void setProgressBlock(object sender, EventArgs e)
        {
            ListView.Invoke(myIncProgressBlock);
        }

        private void WriteConverterLog(object sender, EventArgs e)
        {
            object[] param = new object[1];
            param[0] = ((UofEventArgs)e).Message;

            if (((UofEventArgs)e).Message != null)
            {
                ListView.Invoke(myWriteLog, new object[] { param, false });
            }
        }

        #endregion
    }
}

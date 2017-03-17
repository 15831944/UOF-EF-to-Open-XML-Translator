using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Threading;
using Microsoft.Win32;
using Act.UofTranslator.UofTranslatorLib;
using UofAddinLib;

namespace UofTranslatorActiveX
{
    public partial class UofActiveXUndocked : UserControl
    {
        public static UofContextMenu contextMenu = new UofContextMenu();

        public UofActiveXUndocked()
        {
            Thread.CurrentThread.CurrentUICulture = Thread.CurrentThread.CurrentCulture;
            InitializeComponent();
        }

        #region Handle Drag Drop

        void lstFiles_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            int i;
            for (i = 0; i < s.Length; i++)
            {
                string filename = Path.GetFileName(s[i]);
                lstFiles.Items.Add(new WorkingListViewItem(s[i], lstFiles, txtbxResult));
            }
        }

        void lstFiles_DragEnter(object sender, DragEventArgs e)
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


        private void lstFiles_ItemDrag(object sender, ItemDragEventArgs e)
        {

            List<string> filelist = new List<string>();

            foreach (ListViewItem item in lstFiles.SelectedItems)
            {
                WorkingListViewItem workingItem = (WorkingListViewItem)item;
                if (workingItem.State.StateValProgress == StateVal.STATE_DONE)
                {
                    if (File.Exists(workingItem.State.DestinationFileName) || Directory.Exists(workingItem.State.DestinationFileName))
                    {
                        filelist.Add(workingItem.State.DestinationFileName);
                    }
                }
            }

            if (filelist.Count == 0)
            {
                return;
            }

            string[] files = filelist.ToArray();
            DragDropEffects effect = DoDragDrop(new DataObject(DataFormats.FileDrop, files), DragDropEffects.Copy);

            return;
        }

        #endregion

        #region Handle Delete Key Press

        private void lstFiles_KeyUp(object sender, KeyEventArgs e)
        {
            if ((lstFiles.SelectedItems.Count != 0) && (e.KeyCode == Keys.Delete))
            {
                for (int i = 0; i < lstFiles.SelectedItems.Count; i++)
                {
                    ListViewItem item = lstFiles.SelectedItems[i];
                    onRemove((WorkingListViewItem)item);
                    lstFiles.Items.Remove(item);
                }
            }
        }

        private void onRemove(WorkingListViewItem item)
        {
            if (item.State.StateValProgress == StateVal.STATE_DONE)
            {
                if (!String.IsNullOrEmpty(item.State.DestinationFileName))
                {
                    string parentPath;
                    parentPath = Path.GetDirectoryName(item.State.DestinationFileName);
                    if (Directory.Exists(parentPath))
                    {
                        Directory.Delete(parentPath, true);
                    }
                }
            }
            lstFiles.Controls.Remove(item.PgbItem);
        }

        #endregion

        private void lstFiles_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            if(lstFiles.SelectedItems.Count == 0)
            {
                txtbxResult.Text = "";
            }
            else
            {
                WorkingListViewItem item = lstFiles.SelectedItems[0] as WorkingListViewItem;
                if (item != null)
                {
                    txtbxResult.Text = item.LogText;
                }
            }
        }

        private void btnStopAll_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in lstFiles.Items)
            {
                WorkingListViewItem workItem = item as WorkingListViewItem;
                if (workItem != null)
                {
                    if (workItem.State.StateValProgress == StateVal.STATE_WORKING)
                    {
                        workItem.stop();
                    }
                }
            }
        }

        private void btnStartAllUOF_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in lstFiles.Items)
            {
                WorkingListViewItem workItem = item as WorkingListViewItem;
                if (workItem != null)
                {
                    if (workItem.State.StateValProgress == StateVal.STATE_READY)
                    {
                        if (workItem.State.TransType == TranslationType.UofToOox)
                        {
                            workItem.start();
                        }
                        else
                        {
                            if (Directory.Exists(workItem.State.SourceFileName))
                            {
                                workItem.State.setDirTranslationType(DocumentType.All, TranslationType.UofToOox);
                                workItem.start();
                            }
                        }
                    }
                }
            }
        }

        private void btnStartAllOOX_Click(object sender, EventArgs e)
        {
            foreach (ListViewItem item in lstFiles.Items)
            {
                WorkingListViewItem workItem = item as WorkingListViewItem;
                if (workItem != null)
                {
                    if (workItem.State.StateValProgress == StateVal.STATE_READY)
                    {
                        if (workItem.State.TransType == TranslationType.OoxToUof)
                        {
                            workItem.start();
                        }
                        else
                        {
                            if (Directory.Exists(workItem.State.SourceFileName))
                            {
                                workItem.State.setDirTranslationType(DocumentType.All, TranslationType.OoxToUof);
                                workItem.start();
                            }
                        }
                    }
                }
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using UofAddinLib;

namespace UofTranslatorActiveX
{
    public class UofStatusButton : Button
    {
        private WorkingListViewItem item;

        public WorkingListViewItem Item
        {
            get { return item; }
            set { item = value; }
        }

        public UofStatusButton(WorkingListViewItem item)
        {
            this.item = item;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);

            InitializeComponent();
            toogleIcon();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // UofStatusButton
            // 
            this.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.UseVisualStyleBackColor = true;
            this.Click += new System.EventHandler(this.UofStatusButton_Click);
            this.ResumeLayout(false);

        }

        private void UofStatusButton_Click(object sender, EventArgs e)
        {
            int stateProgress = Item.State.StateValProgress;

            if (stateProgress == StateVal.STATE_READY)
            {
                if (Directory.Exists(this.Item.State.SourceFileName))
                {
                    UofActiveXUndocked.contextMenu.CallbackButton = this;
                    UofActiveXUndocked.contextMenu.ContextMenuForDir.Show(this, new Point(0, 0));
                }
                else
                {
                    this.item.start();
                }
            }
            else if (stateProgress == StateVal.STATE_WORKING)
            {
                this.Item.stop();
            }
            else if (stateProgress == StateVal.STATE_DONE)
            {
            }
        }

        /// <summary>
        ///     Change icon based on current state
        /// </summary>
        public void toogleIcon()
        {
            int stateProgress = Item.State.StateValProgress;

            if (stateProgress == StateVal.STATE_READY)
            {
                this.Image = StatusIcons.Ready;
            }
            else if (stateProgress == StateVal.STATE_WORKING)
            {
                this.Image = StatusIcons.Stop;
            }
            else if (stateProgress == StateVal.STATE_DONE)
            {
                if (Item.State.hasMask(StateVal.STATE_MASK_ERROR))
                {
                    this.Image = StatusIcons.Error;
                }
                else
                {
                    if (Item.State.hasMask(StateVal.STATE_MASK_WARNING))
                    {
                        this.Image = StatusIcons.Warning;
                    }
                    else
                    {
                        this.Image = StatusIcons.OK;
                    }
                }
            }
        }
    }
}

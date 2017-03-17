namespace UofTranslatorActiveX
{
    partial class UofActiveXUndocked
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UofActiveXUndocked));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.lstFiles = new UofTranslatorActiveX.UofListView();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.btnStartAllOOX = new System.Windows.Forms.Button();
            this.btnStartAllUOF = new System.Windows.Forms.Button();
            this.btnStopAll = new System.Windows.Forms.Button();
            this.txtbxResult = new System.Windows.Forms.TextBox();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            resources.ApplyResources(this.splitContainer1, "splitContainer1");
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.lstFiles);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.tableLayoutPanel1);
            // 
            // lstFiles
            // 
            this.lstFiles.AllowDrop = true;
            this.lstFiles.BackColor = System.Drawing.Color.White;
            resources.ApplyResources(this.lstFiles, "lstFiles");
            this.lstFiles.FullRowSelect = true;
            this.lstFiles.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lstFiles.IsControlPressed = false;
            this.lstFiles.IsShiftPressed = false;
            this.lstFiles.ItemFont = new System.Drawing.Font("Arial", 9F);
            this.lstFiles.Name = "lstFiles";
            this.lstFiles.OwnerDraw = true;
            this.lstFiles.ShowItemToolTips = true;
            this.lstFiles.UseCompatibleStateImageBehavior = false;
            this.lstFiles.View = System.Windows.Forms.View.Details;
            this.lstFiles.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.lstFiles_ItemDrag);
            this.lstFiles.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.lstFiles_ItemSelectionChanged);
            this.lstFiles.DragDrop += new System.Windows.Forms.DragEventHandler(this.lstFiles_DragDrop);
            this.lstFiles.DragEnter += new System.Windows.Forms.DragEventHandler(this.lstFiles_DragEnter);
            this.lstFiles.KeyUp += new System.Windows.Forms.KeyEventHandler(this.lstFiles_KeyUp);
            // 
            // tableLayoutPanel1
            // 
            resources.ApplyResources(this.tableLayoutPanel1, "tableLayoutPanel1");
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.txtbxResult, 0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            // 
            // tableLayoutPanel2
            // 
            resources.ApplyResources(this.tableLayoutPanel2, "tableLayoutPanel2");
            this.tableLayoutPanel2.Controls.Add(this.btnStartAllOOX, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.btnStartAllUOF, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.btnStopAll, 1, 0);
            this.tableLayoutPanel2.GrowStyle = System.Windows.Forms.TableLayoutPanelGrowStyle.FixedSize;
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            // 
            // btnStartAllOOX
            // 
            resources.ApplyResources(this.btnStartAllOOX, "btnStartAllOOX");
            this.btnStartAllOOX.Name = "btnStartAllOOX";
            this.btnStartAllOOX.UseVisualStyleBackColor = true;
            this.btnStartAllOOX.Click += new System.EventHandler(this.btnStartAllOOX_Click);
            // 
            // btnStartAllUOF
            // 
            resources.ApplyResources(this.btnStartAllUOF, "btnStartAllUOF");
            this.btnStartAllUOF.Name = "btnStartAllUOF";
            this.btnStartAllUOF.UseVisualStyleBackColor = true;
            this.btnStartAllUOF.Click += new System.EventHandler(this.btnStartAllUOF_Click);
            // 
            // btnStopAll
            // 
            resources.ApplyResources(this.btnStopAll, "btnStopAll");
            this.btnStopAll.Name = "btnStopAll";
            this.btnStopAll.UseVisualStyleBackColor = true;
            this.btnStopAll.Click += new System.EventHandler(this.btnStopAll_Click);
            // 
            // txtbxResult
            // 
            this.txtbxResult.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            resources.ApplyResources(this.txtbxResult, "txtbxResult");
            this.txtbxResult.Name = "txtbxResult";
            this.txtbxResult.ReadOnly = true;
            // 
            // UofActiveXUndocked
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.Controls.Add(this.splitContainer1);
            this.Name = "UofActiveXUndocked";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private UofListView lstFiles;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Button btnStartAllUOF;
        private System.Windows.Forms.Button btnStopAll;
        private System.Windows.Forms.TextBox txtbxResult;
        private System.Windows.Forms.Button btnStartAllOOX;
    }
}

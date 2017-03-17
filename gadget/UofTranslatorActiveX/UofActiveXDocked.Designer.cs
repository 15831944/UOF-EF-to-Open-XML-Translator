namespace UofTranslatorActiveX
{
    partial class UofActiveXDocked
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
            this.components = new System.ComponentModel.Container();
            this.listView1 = new System.Windows.Forms.ListView();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.Info1 = new System.Windows.Forms.Button();
            this.Transfer1 = new System.Windows.Forms.Button();
            this.GadgetHelp = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.AllowDrop = true;
            this.listView1.BackColor = System.Drawing.SystemColors.Control;
            this.listView1.Font = new System.Drawing.Font("SimSun", 8F);
            this.listView1.LargeImageList = this.imageList1;
            this.listView1.Location = new System.Drawing.Point(0, 0);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(130, 120);
            this.listView1.TabIndex = 2;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this.listView1_ItemDrag);
            this.listView1.DragDrop += new System.Windows.Forms.DragEventHandler(this.listView1_DragDrop);
            this.listView1.DragEnter += new System.Windows.Forms.DragEventHandler(this.listView1_DragEnter);
            this.listView1.KeyUp += new System.Windows.Forms.KeyEventHandler(this.listView1_KeyUp);
            // 
            // imageList1
            // 
            this.imageList1.ColorDepth = System.Windows.Forms.ColorDepth.Depth32Bit;
            this.imageList1.ImageSize = new System.Drawing.Size(68, 68);
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 117);
            this.progressBar1.MarqueeAnimationSpeed = 0;
            this.progressBar1.Maximum = 1000;
            this.progressBar1.Minimum = 1;
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(130, 23);
            this.progressBar1.Step = 1;
            this.progressBar1.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            this.progressBar1.TabIndex = 3;
            this.progressBar1.Value = 1;
            // 
            // Info1
            // 
            this.Info1.Enabled = false;
            this.Info1.Image = global::UofTranslatorActiveX.IconResource.info1;
            this.Info1.Location = new System.Drawing.Point(65, 135);
            this.Info1.Name = "Info1";
            this.Info1.Size = new System.Drawing.Size(65, 65);
            this.Info1.TabIndex = 1;
            this.Info1.UseVisualStyleBackColor = false;
            this.Info1.Click += new System.EventHandler(this.Info1_Click);
            // 
            // Transfer1
            // 
            this.Transfer1.BackColor = System.Drawing.SystemColors.Control;
            this.Transfer1.Enabled = false;
            this.Transfer1.Image = global::UofTranslatorActiveX.IconResource.Transfer1;
            this.Transfer1.Location = new System.Drawing.Point(0, 135);
            this.Transfer1.Name = "Transfer1";
            this.Transfer1.Size = new System.Drawing.Size(65, 65);
            this.Transfer1.TabIndex = 0;
            this.Transfer1.UseVisualStyleBackColor = false;
            this.Transfer1.Click += new System.EventHandler(this.Transfer1_Click);
            // 
            // GadgetHelp
            // 
            this.GadgetHelp.Font = new System.Drawing.Font("SimSun", 12F);
            this.GadgetHelp.Location = new System.Drawing.Point(104, 92);
            this.GadgetHelp.Name = "GadgetHelp";
            this.GadgetHelp.Size = new System.Drawing.Size(25, 25);
            this.GadgetHelp.TabIndex = 4;
            this.GadgetHelp.Text = "?";
            this.GadgetHelp.UseVisualStyleBackColor = true;
            this.GadgetHelp.Click += new System.EventHandler(this.GadgetHelp_Click);
            // 
            // UofActiveXDocked
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.GadgetHelp);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.Info1);
            this.Controls.Add(this.Transfer1);
            this.Margin = new System.Windows.Forms.Padding(0);
            this.Name = "UofActiveXDocked";
            this.Size = new System.Drawing.Size(130, 200);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button Transfer1;
        private System.Windows.Forms.Button Info1;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button GadgetHelp;

    }
}

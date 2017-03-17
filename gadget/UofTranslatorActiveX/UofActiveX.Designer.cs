namespace UofTranslatorActiveX
{
    partial class UofActiveX
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
            this.uofActiveXDocked1 = new UofTranslatorActiveX.UofActiveXDocked();
            this.uofActiveXUndocked1 = new UofTranslatorActiveX.UofActiveXUndocked();
            this.SuspendLayout();
            // 
            // uofActiveXDocked1
            // 
            this.uofActiveXDocked1.FileIcon = null;
            this.uofActiveXDocked1.Location = new System.Drawing.Point(0, 0);
            this.uofActiveXDocked1.Margin = new System.Windows.Forms.Padding(0);
            this.uofActiveXDocked1.Name = "uofActiveXDocked1";
            this.uofActiveXDocked1.Size = new System.Drawing.Size(150, 150);
            this.uofActiveXDocked1.TabIndex = 0;
            this.uofActiveXDocked1.WorkingThread = null;
            // 
            // uofActiveXUndocked1
            // 
            this.uofActiveXUndocked1.BackColor = System.Drawing.SystemColors.Control;
            this.uofActiveXUndocked1.Font = new System.Drawing.Font("Calibri", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.uofActiveXUndocked1.Location = new System.Drawing.Point(12, 27);
            this.uofActiveXUndocked1.Name = "uofActiveXUndocked1";
            this.uofActiveXUndocked1.Size = new System.Drawing.Size(120, 218);
            this.uofActiveXUndocked1.TabIndex = 1;
            // 
            // UofActiveX
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.uofActiveXUndocked1);
            this.Controls.Add(this.uofActiveXDocked1);
            this.Name = "UofActiveX";
            this.ResumeLayout(false);

        }

        #endregion

        private UofActiveXDocked uofActiveXDocked1;
        private UofActiveXUndocked uofActiveXUndocked1;
    }
}


namespace ConfigManager
{
    partial class CfgManager
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.sBtn = new System.Windows.Forms.RadioButton();
            this.rBtn = new System.Windows.Forms.RadioButton();
            this.confirm = new System.Windows.Forms.Button();
            this.cancel = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.sBtn);
            this.groupBox1.Controls.Add(this.rBtn);
            this.groupBox1.Location = new System.Drawing.Point(32, 30);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(200, 86);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "OOXML Version";
            // 
            // sBtn
            // 
            this.sBtn.AutoSize = true;
            this.sBtn.Location = new System.Drawing.Point(28, 44);
            this.sBtn.Name = "sBtn";
            this.sBtn.Size = new System.Drawing.Size(119, 16);
            this.sBtn.TabIndex = 1;
            this.sBtn.TabStop = true;
            this.sBtn.Text = "ISO 29500 Strict";
            this.sBtn.UseVisualStyleBackColor = true;
            this.sBtn.Checked = false;
            // 
            // rBtn
            // 
            this.rBtn.AutoSize = true;
            this.rBtn.Location = new System.Drawing.Point(28, 21);
            this.rBtn.Name = "rBtn";
            this.rBtn.Size = new System.Drawing.Size(155, 16);
            this.rBtn.TabIndex = 0;
            this.rBtn.TabStop = true;
            this.rBtn.Text = "ISO 29500 Transitional";
            this.rBtn.UseVisualStyleBackColor = true;
            this.rBtn.Checked = false;
            // 
            // confirm
            // 
            this.confirm.Location = new System.Drawing.Point(72, 122);
            this.confirm.Name = "confirm";
            this.confirm.Size = new System.Drawing.Size(75, 23);
            this.confirm.TabIndex = 1;
            this.confirm.Text = "确定";
            this.confirm.UseVisualStyleBackColor = true;
            this.confirm.Click += new System.EventHandler(this.confirm_Click);
            // 
            // cancel
            // 
            this.cancel.Location = new System.Drawing.Point(157, 122);
            this.cancel.Name = "cancel";
            this.cancel.Size = new System.Drawing.Size(75, 23);
            this.cancel.TabIndex = 2;
            this.cancel.Text = "取消";
            this.cancel.UseVisualStyleBackColor = true;
            this.cancel.Click += new System.EventHandler(this.cancel_Click);
            // 
            // CfgManager
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(260, 167);
            this.Controls.Add(this.cancel);
            this.Controls.Add(this.confirm);
            this.Controls.Add(this.groupBox1);
            this.MaximizeBox = false;
            this.Name = "CfgManager";
            this.Text = "ConfigManager";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton sBtn;
        private System.Windows.Forms.RadioButton rBtn;
        private System.Windows.Forms.Button confirm;
        private System.Windows.Forms.Button cancel;
    }
}


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

namespace UofAddinLib
{
    partial class RenameOptionForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RenameOptionForm));
            this.panel2 = new System.Windows.Forms.Panel();
            this.lblFileLastModified = new System.Windows.Forms.Label();
            this.lblFileCreatedTime = new System.Windows.Forms.Label();
            this.lblFileSize = new System.Windows.Forms.Label();
            this.lblFileName = new System.Windows.Forms.Label();
            this.picbIcon = new System.Windows.Forms.PictureBox();
            this.lblPrompt = new System.Windows.Forms.Label();
            this.buttonPanel = new System.Windows.Forms.Panel();
            this.btnSkipAll = new System.Windows.Forms.Button();
            this.btnRename = new System.Windows.Forms.Button();
            this.btnSkip = new System.Windows.Forms.Button();
            this.btnOverwriteAll = new System.Windows.Forms.Button();
            this.btnOverwrite = new System.Windows.Forms.Button();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picbIcon)).BeginInit();
            this.buttonPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel2
            // 
            resources.ApplyResources(this.panel2, "panel2");
            this.panel2.Controls.Add(this.lblFileLastModified);
            this.panel2.Controls.Add(this.lblFileCreatedTime);
            this.panel2.Controls.Add(this.lblFileSize);
            this.panel2.Controls.Add(this.lblFileName);
            this.panel2.Controls.Add(this.picbIcon);
            this.panel2.Controls.Add(this.lblPrompt);
            this.panel2.Name = "panel2";
            // 
            // lblFileLastModified
            // 
            resources.ApplyResources(this.lblFileLastModified, "lblFileLastModified");
            this.lblFileLastModified.Name = "lblFileLastModified";
            // 
            // lblFileCreatedTime
            // 
            resources.ApplyResources(this.lblFileCreatedTime, "lblFileCreatedTime");
            this.lblFileCreatedTime.Name = "lblFileCreatedTime";
            // 
            // lblFileSize
            // 
            resources.ApplyResources(this.lblFileSize, "lblFileSize");
            this.lblFileSize.Name = "lblFileSize";
            // 
            // lblFileName
            // 
            resources.ApplyResources(this.lblFileName, "lblFileName");
            this.lblFileName.Name = "lblFileName";
            // 
            // picbIcon
            // 
            resources.ApplyResources(this.picbIcon, "picbIcon");
            this.picbIcon.Name = "picbIcon";
            this.picbIcon.TabStop = false;
            // 
            // lblPrompt
            // 
            resources.ApplyResources(this.lblPrompt, "lblPrompt");
            this.lblPrompt.Name = "lblPrompt";
            // 
            // buttonPanel
            // 
            resources.ApplyResources(this.buttonPanel, "buttonPanel");
            this.buttonPanel.Controls.Add(this.btnSkipAll);
            this.buttonPanel.Controls.Add(this.btnRename);
            this.buttonPanel.Controls.Add(this.btnSkip);
            this.buttonPanel.Controls.Add(this.btnOverwriteAll);
            this.buttonPanel.Controls.Add(this.btnOverwrite);
            this.buttonPanel.Name = "buttonPanel";
            // 
            // btnSkipAll
            // 
            resources.ApplyResources(this.btnSkipAll, "btnSkipAll");
            this.btnSkipAll.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnSkipAll.Name = "btnSkipAll";
            this.btnSkipAll.UseVisualStyleBackColor = true;
            this.btnSkipAll.Click += new System.EventHandler(this.btnSkipAll_Click);
            // 
            // btnRename
            // 
            resources.ApplyResources(this.btnRename, "btnRename");
            this.btnRename.Name = "btnRename";
            this.btnRename.UseVisualStyleBackColor = true;
            this.btnRename.Click += new System.EventHandler(this.btnRename_Click);
            // 
            // btnSkip
            // 
            resources.ApplyResources(this.btnSkip, "btnSkip");
            this.btnSkip.DialogResult = System.Windows.Forms.DialogResult.Abort;
            this.btnSkip.Name = "btnSkip";
            this.btnSkip.UseVisualStyleBackColor = true;
            // 
            // btnOverwriteAll
            // 
            resources.ApplyResources(this.btnOverwriteAll, "btnOverwriteAll");
            this.btnOverwriteAll.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOverwriteAll.Name = "btnOverwriteAll";
            this.btnOverwriteAll.UseVisualStyleBackColor = true;
            this.btnOverwriteAll.Click += new System.EventHandler(this.btnOverwriteAll_Click);
            // 
            // btnOverwrite
            // 
            resources.ApplyResources(this.btnOverwrite, "btnOverwrite");
            this.btnOverwrite.DialogResult = System.Windows.Forms.DialogResult.Ignore;
            this.btnOverwrite.Name = "btnOverwrite";
            this.btnOverwrite.UseVisualStyleBackColor = true;
            this.btnOverwrite.Click += new System.EventHandler(this.btnOverwrite_Click);
            // 
            // RenameOptionForm
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.buttonPanel);
            this.Controls.Add(this.panel2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "RenameOptionForm";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.RenameOptionForm_FormClosed);
            this.Load += new System.EventHandler(this.RenameOptionForm_Load);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picbIcon)).EndInit();
            this.buttonPanel.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel buttonPanel;
        private System.Windows.Forms.Label lblPrompt;
        private System.Windows.Forms.PictureBox picbIcon;
        private System.Windows.Forms.Label lblFileName;
        private System.Windows.Forms.Button btnRename;
        private System.Windows.Forms.Button btnSkip;
        private System.Windows.Forms.Button btnOverwriteAll;
        private System.Windows.Forms.Button btnOverwrite;
        private System.Windows.Forms.Label lblFileLastModified;
        private System.Windows.Forms.Label lblFileCreatedTime;
        private System.Windows.Forms.Label lblFileSize;
        private System.Windows.Forms.Button btnSkipAll;
    }
}
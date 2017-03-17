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

    partial class MainDialog
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainDialog));
            this.btnCancel = new System.Windows.Forms.Button();
            this.pnlButton = new System.Windows.Forms.Panel();
            this.btnSeeLog = new System.Windows.Forms.Button();
            this.pnlTop = new System.Windows.Forms.Panel();
            this.lblToLabel = new System.Windows.Forms.Label();
            this.lblCurrentProcessing = new System.Windows.Forms.Label();
            this.lblCurrentState = new System.Windows.Forms.Label();
            this.lblToFileName = new System.Windows.Forms.Label();
            this.lblFileOverAll = new System.Windows.Forms.Label();
            this.chkbxIgnoreError = new System.Windows.Forms.CheckBox();
            this.pbrCurrentFile = new System.Windows.Forms.ProgressBar();
            this.lblCurrentProgress = new System.Windows.Forms.Label();
            this.txtLog = new System.Windows.Forms.TextBox();
            this.pnlButton.SuspendLayout();
            this.pnlTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnCancel
            // 
            resources.ApplyResources(this.btnCancel, "btnCancel");
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            // 
            // pnlButton
            // 
            resources.ApplyResources(this.pnlButton, "pnlButton");
            this.pnlButton.Controls.Add(this.btnSeeLog);
            this.pnlButton.Controls.Add(this.btnCancel);
            this.pnlButton.Name = "pnlButton";
            // 
            // btnSeeLog
            // 
            resources.ApplyResources(this.btnSeeLog, "btnSeeLog");
            this.btnSeeLog.Name = "btnSeeLog";
            this.btnSeeLog.UseVisualStyleBackColor = true;
            this.btnSeeLog.Click += new System.EventHandler(this.btnSeeLog_Click);
            // 
            // pnlTop
            // 
            resources.ApplyResources(this.pnlTop, "pnlTop");
            this.pnlTop.Controls.Add(this.lblToLabel);
            this.pnlTop.Controls.Add(this.lblCurrentProcessing);
            this.pnlTop.Controls.Add(this.lblCurrentState);
            this.pnlTop.Controls.Add(this.lblToFileName);
            this.pnlTop.Controls.Add(this.lblFileOverAll);
            this.pnlTop.Controls.Add(this.chkbxIgnoreError);
            this.pnlTop.Controls.Add(this.pbrCurrentFile);
            this.pnlTop.Controls.Add(this.lblCurrentProgress);
            this.pnlTop.Controls.Add(this.pnlButton);
            this.pnlTop.Name = "pnlTop";
            this.pnlTop.Paint += new System.Windows.Forms.PaintEventHandler(this.pnlTop_Paint);
            // 
            // lblToLabel
            // 
            resources.ApplyResources(this.lblToLabel, "lblToLabel");
            this.lblToLabel.Name = "lblToLabel";
            // 
            // lblCurrentProcessing
            // 
            resources.ApplyResources(this.lblCurrentProcessing, "lblCurrentProcessing");
            this.lblCurrentProcessing.Name = "lblCurrentProcessing";
            // 
            // lblCurrentState
            // 
            resources.ApplyResources(this.lblCurrentState, "lblCurrentState");
            this.lblCurrentState.Name = "lblCurrentState";
            // 
            // lblToFileName
            // 
            resources.ApplyResources(this.lblToFileName, "lblToFileName");
            this.lblToFileName.Name = "lblToFileName";
            // 
            // lblFileOverAll
            // 
            resources.ApplyResources(this.lblFileOverAll, "lblFileOverAll");
            this.lblFileOverAll.Name = "lblFileOverAll";
            // 
            // chkbxIgnoreError
            // 
            resources.ApplyResources(this.chkbxIgnoreError, "chkbxIgnoreError");
            this.chkbxIgnoreError.Name = "chkbxIgnoreError";
            this.chkbxIgnoreError.UseVisualStyleBackColor = true;
            this.chkbxIgnoreError.CheckedChanged += new System.EventHandler(this.chkbxIgnoreError_CheckedChanged);
            // 
            // pbrCurrentFile
            // 
            resources.ApplyResources(this.pbrCurrentFile, "pbrCurrentFile");
            this.pbrCurrentFile.MarqueeAnimationSpeed = 0;
            this.pbrCurrentFile.Name = "pbrCurrentFile";
            this.pbrCurrentFile.Style = System.Windows.Forms.ProgressBarStyle.Marquee;
            // 
            // lblCurrentProgress
            // 
            resources.ApplyResources(this.lblCurrentProgress, "lblCurrentProgress");
            this.lblCurrentProgress.Name = "lblCurrentProgress";
            // 
            // txtLog
            // 
            resources.ApplyResources(this.txtLog, "txtLog");
            this.txtLog.Name = "txtLog";
            // 
            // MainDialog
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.pnlTop);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.Name = "MainDialog";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainDialog_FormClosed);
            this.Load += new System.EventHandler(this.MainDialog_Load);
            this.pnlButton.ResumeLayout(false);
            this.pnlTop.ResumeLayout(false);
            this.pnlTop.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Panel pnlButton;
        private System.Windows.Forms.Button btnSeeLog;
        private System.Windows.Forms.Panel pnlTop;
        private System.Windows.Forms.Label lblCurrentProgress;
        private System.Windows.Forms.ProgressBar pbrCurrentFile;
        private System.Windows.Forms.TextBox txtLog;
        private System.Windows.Forms.CheckBox chkbxIgnoreError;
        private System.Windows.Forms.Label lblFileOverAll;
        private System.Windows.Forms.Label lblToFileName;
        private System.Windows.Forms.Label lblCurrentState;
        private System.Windows.Forms.Label lblCurrentProcessing;
        private System.Windows.Forms.Label lblToLabel;
    }
}
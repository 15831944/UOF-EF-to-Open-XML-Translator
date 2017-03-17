using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using UofAddinLib;

namespace UofTranslatorActiveX
{
    public class UofContextMenu
    {
        private UofStatusButton callbackButton;

        public UofStatusButton CallbackButton
        {
            get { return callbackButton; }
            set { callbackButton = value; }
        }

        public System.Windows.Forms.ContextMenuStrip ContextMenuForDir;
        private System.Windows.Forms.ToolStripMenuItem OOX_UOFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem DOCX_UOFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem XLSX_UOFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem PPTX_UOFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem UOF_OOXToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem UOF_DOCXToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem UOF_XLSXToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem UOF_PPTXToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ALL_UOFToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem ALL_OOXToolStripMenuItem;

        public UofContextMenu()
        {
            callbackButton = null;

            this.ContextMenuForDir = new System.Windows.Forms.ContextMenuStrip();
            this.OOX_UOFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.DOCX_UOFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.XLSX_UOFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.PPTX_UOFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.UOF_OOXToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.UOF_DOCXToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.UOF_XLSXToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.UOF_PPTXToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ALL_UOFToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ALL_OOXToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.ContextMenuForDir.SuspendLayout();

            // 
            // contextMenuForDir
            // 
            this.ContextMenuForDir.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.OOX_UOFToolStripMenuItem,
            this.UOF_OOXToolStripMenuItem});
            this.ContextMenuForDir.Name = "contextMenuForDir";
            this.ContextMenuForDir.Size = new System.Drawing.Size(153, 70);
            this.ContextMenuForDir.Closed += new System.Windows.Forms.ToolStripDropDownClosedEventHandler(ContextMenuForDir_Closed);
            // 
            // oOXUOFToolStripMenuItem
            // 
            this.OOX_UOFToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.DOCX_UOFToolStripMenuItem,
            this.XLSX_UOFToolStripMenuItem,
            this.PPTX_UOFToolStripMenuItem,
            this.ALL_OOXToolStripMenuItem});
            this.OOX_UOFToolStripMenuItem.Name = "OOX_UOFToolStripMenuItem";
            this.OOX_UOFToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.OOX_UOFToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuOOXToUOF;
            // 
            // pPTXUOFToolStripMenuItem
            // 
            this.DOCX_UOFToolStripMenuItem.Name = "DOCX_UOFToolStripMenuItem";
            this.DOCX_UOFToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.DOCX_UOFToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuDOCXToUOF;
            this.DOCX_UOFToolStripMenuItem.Click += new EventHandler(DOCX_UOFToolStripMenuItem_Click);
            // 
            // xLSXUOFToolStripMenuItem
            // 
            this.XLSX_UOFToolStripMenuItem.Name = "XLSX_UOFToolStripMenuItem";
            this.XLSX_UOFToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.XLSX_UOFToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuXLSXToUOF;
            this.XLSX_UOFToolStripMenuItem.Click += new EventHandler(XLSX_UOFToolStripMenuItem_Click);
            // 
            // dOCXToolStripMenuItem
            // 
            this.PPTX_UOFToolStripMenuItem.Name = "PPTX_UOFToolStripMenuItem";
            this.PPTX_UOFToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.PPTX_UOFToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuPPTXToUOF;
            this.PPTX_UOFToolStripMenuItem.Click += new EventHandler(PPTX_UOFToolStripMenuItem_Click);
            // 
            // uOFOOXToolStripMenuItem
            // 
            this.UOF_OOXToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.UOF_DOCXToolStripMenuItem,
            this.UOF_XLSXToolStripMenuItem,
            this.UOF_PPTXToolStripMenuItem,
            this.ALL_UOFToolStripMenuItem});
            this.UOF_OOXToolStripMenuItem.Name = "UOF_OOXToolStripMenuItem";
            this.UOF_OOXToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.UOF_OOXToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuUOFToOOX;
            // 
            // uOFPPTXToolStripMenuItem
            // 
            this.UOF_DOCXToolStripMenuItem.Name = "UOF_DOCXToolStripMenuItem";
            this.UOF_DOCXToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.UOF_DOCXToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuUOFToDOCX;
            this.UOF_DOCXToolStripMenuItem.Click += new EventHandler(UOF_DOCXToolStripMenuItem_Click);
            // 
            // uOFXLSXToolStripMenuItem
            // 
            this.UOF_XLSXToolStripMenuItem.Name = "UOF_XLSXToolStripMenuItem";
            this.UOF_XLSXToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.UOF_XLSXToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuUOFToXLSX;
            this.UOF_XLSXToolStripMenuItem.Click += new EventHandler(UOF_XLSXToolStripMenuItem_Click);
            // 
            // uOFPPTXToolStripMenuItem1
            // 
            this.UOF_PPTXToolStripMenuItem.Name = "UOF_PPTXToolStripMenuItem";
            this.UOF_PPTXToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.UOF_PPTXToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuUOFToPPTX;
            this.UOF_PPTXToolStripMenuItem.Click += new EventHandler(UOF_PPTXToolStripMenuItem_Click);
            // 
            // uOFOOXToolStripMenuItem1
            // 
            this.ALL_UOFToolStripMenuItem.Name = "ALL_UOFToolStripMenuItem";
            this.ALL_UOFToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.ALL_UOFToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuALLUOFTOOOX;
            this.ALL_UOFToolStripMenuItem.Click += new EventHandler(ALL_UOFToolStripMenuItem_Click);
            // 
            // aLLOOXToolStripMenuItem
            // 
            this.ALL_OOXToolStripMenuItem.Name = "ALL_OOXToolStripMenuItem";
            this.ALL_OOXToolStripMenuItem.Size = new System.Drawing.Size(158, 22);
            this.ALL_OOXToolStripMenuItem.Text = UofTranslatorActiveXRes.ContextMenuALLOOXTOUOF;
            this.ALL_OOXToolStripMenuItem.Click += new EventHandler(ALL_OOXToolStripMenuItem_Click);


            this.ContextMenuForDir.ResumeLayout(false);
        }

        void ALL_OOXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.All, TranslationType.OoxToUof);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }

        void ALL_UOFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.All, TranslationType.UofToOox);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }

        void UOF_PPTXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.Powerpnt, TranslationType.UofToOox);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }

        void UOF_XLSXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.Excel, TranslationType.UofToOox);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }

        void UOF_DOCXToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.Word, TranslationType.UofToOox);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }

        void PPTX_UOFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.Powerpnt, TranslationType.OoxToUof);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }

        void XLSX_UOFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.Excel, TranslationType.OoxToUof);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }

        void ContextMenuForDir_Closed(object sender, System.Windows.Forms.ToolStripDropDownClosedEventArgs e)
        {
        }

        void DOCX_UOFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (callbackButton != null)
            {
                if (callbackButton.Item != null)
                {
                    callbackButton.Item.State.setDirTranslationType(DocumentType.Word, TranslationType.OoxToUof);
                    callbackButton.Item.start();
                }
            }
            this.ContextMenuForDir.Close(ToolStripDropDownCloseReason.ItemClicked);
            this.callbackButton = null;
        }
    }
}

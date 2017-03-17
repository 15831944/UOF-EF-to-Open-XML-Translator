using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Runtime.InteropServices;

namespace UofTranslatorActiveX
{
    [Guid("101F2FC5-ADEB-4f7f-8D38-7F84589B043C")]
    [ComVisible(true)]
    public partial class UofActiveX : UserControl
    {
        public UofActiveX()
        {
            InitializeComponent();
            setDock();
        }

        public void setDock()
        {
            this.uofActiveXUndocked1.Visible = false;
            this.uofActiveXDocked1.Visible = true;
            this.uofActiveXDocked1.Dock = DockStyle.Fill;
        }

        public void setUndock()
        {
            this.uofActiveXUndocked1.Visible = true;
            this.uofActiveXDocked1.Visible = false;
            this.uofActiveXUndocked1.Dock = DockStyle.Fill;
        }
    }
}

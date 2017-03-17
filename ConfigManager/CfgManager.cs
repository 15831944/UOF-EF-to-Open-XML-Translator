using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using UofAddinLib;

namespace ConfigManager
{
    public partial class CfgManager : Form
    {
        public CfgManager()
        {
            InitializeComponent();
        }

        private void confirm_Click(object sender, EventArgs e)
        {

            UofAddinLib.ConfigManager cfg = new UofAddinLib.ConfigManager(System.IO.Path.GetDirectoryName(typeof(MainDialog).Assembly.Location) + @"\conf\config.xml");

            cfg.LoadConfig();

            // transitional
            if (rBtn.Checked)
            { 
                cfg.IsUofToTransitioanlOOX = true;
                cfg.IsUofToStrictOOX = false;
                cfg.SaveConfig();
            }
            else // strict
            {
                cfg.IsUofToStrictOOX = true;
                cfg.IsUofToTransitioanlOOX = false;
                cfg.SaveConfig();
            }
            this.Close();
        }

        private void cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Windows.Forms;

namespace UofTranslatorActiveX
{
    public class UofProgressBar : ProgressBar
    {
        private string _showtext;

        public string Showtext
        {
            get { return _showtext; }
            set { _showtext = value; this.Invalidate(); }
        }

        private WorkingListViewItem item;

        private const int WM_PAINT = 0x000F;

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == WM_PAINT)
            {
                //Draw Text;
                if(_showtext != null)
                {
                    Graphics g = this.CreateGraphics();

                    StringFormat sf = new StringFormat();
                    sf.LineAlignment = StringAlignment.Center;
                    sf.Trimming = StringTrimming.EllipsisCharacter;
                    sf.Alignment = StringAlignment.Near;
                    sf.FormatFlags = StringFormatFlags.NoWrap;
                    g.DrawString(_showtext, ((UofListView)item.ListView).ItemFont, Brushes.Black, this.ClientRectangle, sf);
                }
            }
        }

        public UofProgressBar(WorkingListViewItem item)
        {
            this.item = item;
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            this.Style = ProgressBarStyle.Marquee;
            this.MarqueeAnimationSpeed = 0;

            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();
            // 
            // UofProgressBar
            // 
            this.Click += new System.EventHandler(this.UofProgressBar_Click);
            this.ResumeLayout(false);
        }

        private void UofProgressBar_Click(object sender, EventArgs e)
        {
            UofListView listView = ((UofListView)item.ListView);
            if (!listView.IsControlPressed)
            {
                foreach(ListViewItem subitem in listView.SelectedItems)
                {
                    if (subitem != item)
                    {
                        subitem.Selected = false;
                    }
                }
            }
            item.Selected = !item.Selected;
            item.Focused = true;
        }
    }
}

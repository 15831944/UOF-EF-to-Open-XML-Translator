using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;
using System.Runtime.InteropServices;

namespace UofTranslatorActiveX
{
    public class UofListView : ListView
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct MEASUREITEMSTRUCT
        {
            public int CtlType;
            public int CtlID;
            public int itemID;
            public int itemWidth;
            public int itemHeight;
            public IntPtr itemData;
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct WINDOWPOS
        {
            public System.IntPtr hwnd;
            public System.IntPtr hwndInsertAfter;
            public int x;
            public int y;
            public int cx;
            public int cy;
            public System.UInt32 flags;
        }


        private const UInt32 LVM_FIRST = 0x1000;
        private const UInt32 LVM_SCROLL = (LVM_FIRST + 20);
        private const int WM_HSCROLL = 0x114;
        private const int WM_VSCROLL = 0x115;
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_PAINT = 0x000F;

        private const int padding = 4;
        private const int buttonpadding = 2;
        private const int itemHeight = 14;
        private bool isControlPressed;

        public bool IsControlPressed
        {
            get { return isControlPressed; }
            set { isControlPressed = value; }
        }
        private bool isShiftPressed;

        public bool IsShiftPressed
        {
            get { return isShiftPressed; }
            set { isShiftPressed = value; }
        }

        private struct EmbeddedControl
        {
            public Control MyControl;
            public ListViewItem.ListViewSubItem SubItem;
            public bool isSystemColumn;
        }

        private ColumnHeader status;
        private ColumnHeader name;

        private List<EmbeddedControl> embeddedcontrols;
        private int headerHeight;
        private Font itemFont;

        public Font ItemFont
        {
            get { return itemFont; }
            set { if (value != null) { itemFont = value; this.Invalidate(); } }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_PAINT)
            {
                foreach (EmbeddedControl c in embeddedcontrols)
                {
                    Rectangle r;
                    if (c.MyControl is ProgressBar)
                    {
                        r = getStringRectWithFileIcon(c.SubItem.Bounds);
                    }
                    else
                    {
                        r = c.SubItem.Bounds;
                    }

                    if (c.isSystemColumn)
                    {
                        r = new Rectangle(r.Left, r.Top, status.Width, r.Height);
                    }

                    if (r.Y >= headerHeight && r.Y < this.ClientRectangle.Height)
                    {
                        c.MyControl.Visible = true;
                        if (c.MyControl is Button)
                        {
                            c.MyControl.Bounds = new Rectangle(r.X + buttonpadding, r.Y + buttonpadding, r.Width - (2 * buttonpadding), r.Height - (2 * buttonpadding));
                        }
                        else
                        {
                            c.MyControl.Bounds = new Rectangle(r.X + padding, r.Y + padding, r.Width - (2 * padding), r.Height - (2 * padding));
                        }
                    }
                    else
                    {
                        c.MyControl.Visible = false;
                    }
                }
            }
            switch (m.Msg)
            {
                case WM_HSCROLL:
                case WM_VSCROLL:
                case WM_MOUSEWHEEL:
                    this.Focus();
                    break;
            }
            base.WndProc(ref m);
        }

        public void AddControlToSubItem(Control control, ListViewItem.ListViewSubItem subitem, bool isSystemColumn)
        {
            EmbeddedControl ec;
            ec.MyControl = control;
            ec.SubItem = subitem;
            ec.isSystemColumn = isSystemColumn;
            this.embeddedcontrols.Add(ec);
        }

        public UofListView()
        {
            headerHeight = 0;
            embeddedcontrols = new List<EmbeddedControl>();
            this.ResizeRedraw = true;
            this.DoubleBuffered = true;
            this.isControlPressed = false;
            this.isShiftPressed = false;
            this.Font = new Font("Arial", itemHeight);
            this.ItemFont = new Font("Arial", 9);
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.status = new System.Windows.Forms.ColumnHeader();
            this.name = new System.Windows.Forms.ColumnHeader();
            this.SuspendLayout();
            // 
            // status
            // 
            this.status.Text = UofTranslatorActiveXRes.ListColumnOne;
            this.status.Width = 90;
            // 
            // name
            // 
            this.name.Text = UofTranslatorActiveXRes.ListColumnTwo;
            this.name.Width = 200;
            // 
            // UofListView
            // 
            this.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.status,
            this.name});
            this.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.OwnerDraw = true;
            this.View = System.Windows.Forms.View.Details;
            this.DrawColumnHeader += new System.Windows.Forms.DrawListViewColumnHeaderEventHandler(this.UofListView_DrawColumnHeader);
            this.KeyUp += new System.Windows.Forms.KeyEventHandler(this.UofListView_KeyUp);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.UofListView_KeyDown);
            this.DrawSubItem += new System.Windows.Forms.DrawListViewSubItemEventHandler(this.UofListView_DrawSubItem);
            this.ResumeLayout(false);

        }

        private void UofListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            Rectangle subRect = e.Bounds;
            WorkingListViewItem item = ((WorkingListViewItem)e.Item);
            if (e.ColumnIndex == status.Index)
            {
                //High Light
                bool bHighLight = e.Item.Selected;
                if (bHighLight)
                {
                    e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Item.Bounds);
                }

                //Draw Focus
                if (e.Item.Focused)
                {
                    e.DrawFocusRectangle(e.Item.Bounds);
                }
            }
            else if (e.ColumnIndex == name.Index)
            {
                //Draw Name
                DrawIcon(subRect, e.Graphics, item.FileIcon);
            }
        }

        private void DrawIcon(Rectangle rect, Graphics g, Bitmap fileIcon)
        {
            //Draw Icon
            Rectangle iconRect = new Rectangle(rect.Left, rect.Top, rect.Height, rect.Height);
            g.DrawImage(fileIcon, iconRect);
        }

        private Rectangle getStringRectWithFileIcon(Rectangle rect)
        {
            int shift = rect.Height + 3;
            return new Rectangle(rect.Left + shift, rect.Top, rect.Width - shift, rect.Height);
        }

        private void DrawHeader(Rectangle rect, Graphics g, string name, HorizontalAlignment align)
        {
            StringFormat sf = new StringFormat();
            sf.LineAlignment = StringAlignment.Center;
            sf.Trimming = StringTrimming.EllipsisCharacter;
            switch (align)
            {
                case HorizontalAlignment.Center:
                    sf.Alignment = StringAlignment.Center;
                    break;
                case HorizontalAlignment.Left:
                    sf.Alignment = StringAlignment.Near;
                    break;
                case HorizontalAlignment.Right:
                    sf.Alignment = StringAlignment.Far;
                    break;
            }
            sf.FormatFlags = StringFormatFlags.NoWrap;
            g.DrawString(name, this.ItemFont, Brushes.Black, rect, sf);
        }

        private void UofListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.DrawBackground();
            Rectangle rect = e.Bounds;
            this.headerHeight = rect.Height;
            DrawHeader(rect, e.Graphics, e.Header.Text, e.Header.TextAlign);
        }

        private void UofListView_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control)
            {
                this.isControlPressed = true;
            }
            if (e.Shift)
            {
                this.isShiftPressed = true;
            }
        }

        private void UofListView_KeyUp(object sender, KeyEventArgs e)
        {
            if ((e.KeyData & Keys.ControlKey) != 0)
            {
                this.isControlPressed = false;
            }
            if ((e.KeyData & Keys.ShiftKey) != 0)
            {
                this.isShiftPressed = false;
            }
        }
    }
}

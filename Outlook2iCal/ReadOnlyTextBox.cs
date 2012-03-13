using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;

namespace Outlook2iCal
{
    public partial class ReadOnlyTextBox : TextBox
    {
        public new bool ReadOnly
        {
            get
            {
                return true;
            }
            set
            {
                if (value == false)
                {
                    throw new NotImplementedException();
                }
            }
        }

        public new bool Enabled
        {
            get
            {
                return false;
            }
            set
            {
                if (value == true)
                {
                    throw new NotImplementedException();
                }
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            SolidBrush drawBrush = new SolidBrush(ForeColor); //Use the ForeColor property
            // Draw string to screen.
            e.Graphics.DrawString(Text, Font, drawBrush, 0f, 0f); //Use the Font property
        }

        public ReadOnlyTextBox()
        {
            SetStyle(ControlStyles.UserPaint,true);
            InitializeComponent();
            base.ReadOnly = true;
            base.Enabled = false;
        }

        /*public ReadOnlyTextBox(IContainer container)
        {
            container.Add(this);
            InitializeComponent();
        }*/
    }
}

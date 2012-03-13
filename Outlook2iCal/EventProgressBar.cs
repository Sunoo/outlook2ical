using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Outlook2iCal
{
    public partial class EventProgressBar : ProgressBar
    {
        [Category("Property Changed")]
        [Description("Event raised when the value of the Value property is changed on Control.")]
        public event OnValueChanged ValueChanged;
        public delegate void OnValueChanged(object sender, EventArgs e);

        public new int Value
        {
            get
            {
                return base.Value;
            }
            set
            {
                int oldValue = base.Value;
                base.Value = value;
                if (base.Value != oldValue)
                {
                    if (ValueChanged != null)
                    {
                        ValueChanged(this, new EventArgs());
                    }
                }
            }
        }

        public EventProgressBar()
        {
            InitializeComponent();
        }

        /*public EventProgressBar(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }*/
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FillExcel
{
    public partial class ProgressForm : Form
    {

        public ProgressForm()
        {
            InitializeComponent();
        }

        public void UpdateProgress(int value, string message)
        {
            progressBar1.Value = value;
            label1.Text = message;
        }
    }
}

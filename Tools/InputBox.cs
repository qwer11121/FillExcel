using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Tools
{
    public partial class InputBox : Form
    {      

        public string Title
        {
            get { return this.Text; }
            set { this.Text = value; }
        }

        public string Message
        {
            get { return label1.Text; }
            set { label1.Text = value; }
        }

        public string Value
        {
            get { return textBox1.Text; }
            set { textBox1.Text = value; }
        }
        
        public string OKButtonText
        {
            get { return bt_OK.Text; }
            set { bt_OK.Text = value; }
        }

        public string CancelButtonText
        {
            get { return bt_Cancel.Text; }
            set { bt_Cancel.Text = value; }
        }

        public string Key;

        public delegate void ButtonEventHandler(string key, string value);
        public event ButtonEventHandler OKButtonEvent;

        public InputBox()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="title">窗口标题</param>
        /// <param name="message">textbox 提示信息</param>
        /// <param name="defVal">textbox 值</param>
        /// <param name="OKButtonText">确定按钮的文字</param>
        /// <param name="cancelButtonText">取消按钮文字</param>
        public InputBox(string message, string defVal, string title = "设置",
            string OKButtonText = "确定", string cancelButtonText = "取消")
        {
            InitializeComponent();
            this.Title = title;
            this.Message = message;
            this.OKButtonText = OKButtonText;
            this.CancelButtonText = cancelButtonText;
            this.Value = defVal;
        }

        private void bt_OK_Click(object sender, EventArgs e)
        {            
            OKButtonEvent(this.Key, this.Value);
            this.Close();
        }
    }
}

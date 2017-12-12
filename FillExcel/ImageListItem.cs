using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace FillExcel
{
    public partial class ImageListItem : UserControl
    {
        public enum ControlList
        {
            Title,
            Pic,
            MarkAsRed
        };

        public string Title
        {
            get
            {
                return c_Title.Text;
            }
            set
            {
                c_Title.Text = value;
            }
        }
        public string FileName
        {
            get
            {
                return c_FileName.Text;
            }
            set
            {
                c_FileName.Text = value;
            }
        }
        public string ImageLocation
        {
            get
            {
                return c_PicBox.ImageLocation;
            }
            set
            {
                c_PicBox.ImageLocation = value;
                FileName = ImageLocation.Split('\\')[ImageLocation.Split('\\').Count() - 1];
            }
        }
        public bool MarkAsRed
        {
            get
            {
                return c_MarkAsRed.Checked;
            }
            set
            {
                c_MarkAsRed.Checked = value;
            }
        }

        public ImageListItem()
        {
            InitializeComponent();            
        }

        private void c_ChangePicButton_Click(object sender, EventArgs e)
        {
/*            ChangePicButton_Click(this, e)*/;
        }

        internal void SetFocus(ControlList title)
        {
            switch (title)
            {
                case ControlList.Title:
                    c_Title.Focus();
                    break;
                case ControlList.MarkAsRed:
                    c_MarkAsRed.Focus();
                    break;
            }
        }

        internal void HideMoveUpButton()
        {
            c_MoveUpButton.Visible = false;
            c_MoveDownButton.Top -= 12;
        }

        internal void HideMoveDownButton()
        {
            c_MoveDownButton.Visible = false;
            c_MoveUpButton.Top += 12;
        }

        internal void HideBothMoveButton()
        {
            c_MoveUpButton.Visible = c_MoveDownButton.Visible = false;
        }
    }
}

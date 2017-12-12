using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace FillExcel
{
    public partial class Form1 : Form
    {
        ImageListItem[] c_imageListItems;        

        string[] tempPics;
        string[] tempTitles;
        bool[] tempMarkAsRed;

        string FormText = "图片填充工具";

        public Form1()
        {
            InitializeComponent();
            this.Text = FormText;
        }

        private void c_FillExcel_Click(object sender, EventArgs e)
        {
            // pre-fill check
            if (c_imageListItems.Count() == 0)
            {
                MessageBox.Show("没有选中图片");
                return;
            }
            if (string.IsNullOrWhiteSpace(c_DocTitle.Text))
            {
                MessageBox.Show("请输入标题");
                c_DocTitle.Focus();
                return;
            }

            string inspectionTime = Tools.Tools.DateToStringEn(c_InspectionTime.Value);

            string defaultPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string defaultName = c_DocTitle.Text;
            string excelName = defaultPath + "\\" + defaultName + ".xlsx";
            string pdfName = defaultPath + "\\" + defaultName + ".pdf";
            
            if (!c_UseDefaultFileName.Checked)
            {
                // user input filename
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                DialogResult r = saveFileDialog.ShowDialog();
                if (r != DialogResult.OK)
                {
                    return;
                }

                excelName = saveFileDialog.FileName + ".xlsx";
                pdfName = saveFileDialog.FileName + ".pdf";
            }

            string[] pics = new string[c_imageListItems.Count()];
            string[] titles = new string[c_imageListItems.Count()];
            for (int i = 0; i < c_imageListItems.Count(); i++)
            {
                pics[i] = c_imageListItems[i].ImageLocation;
                titles[i] = c_imageListItems[i].Title;
            }

            // check if file name is used
            if(File.Exists(excelName)||File.Exists(pdfName))
            {
                DialogResult r = MessageBox.Show("同名文件已存在，是否覆盖？", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (r == DialogResult.No)
                {
                    return;
                }
            }

            this.Text = FormText + " 处理中...";

            // start creating excel file
            FillExcel fillExcel = new FillExcel();
            fillExcel.LoadConfig("config.ini");
            bool[] markAsRed = new bool[pics.Count()];
            for(int i=0;i<pics.Count();i++)
            {
                markAsRed[i] = c_imageListItems[i].MarkAsRed;
            }

            string result = fillExcel.Fill(excelName, pdfName, c_DocTitle.Text, c_ForthLine.Text,
                pics, titles, markAsRed,
                c_EndText.Text, c_MarkEndTextRed.Checked, c_ItemNo.Text, inspectionTime);
            fillExcel = null;
            if(result=="0")
            {
                result = "文件生成完成, 保存于 " + pdfName;
            }
            else
            {
                result = "发生错误： " + result;
            }
            MessageBox.Show(result);
            this.Text = FormText;
        }

        private void c_SelectImage_Click(object sender, EventArgs e)
        {
            string[] pics;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            DialogResult r = openFileDialog.ShowDialog();
            if (r == DialogResult.OK)
            {
                pics = openFileDialog.FileNames;
                string[] titles = new string[pics.Count()];
                bool[] markAsRed = new bool[pics.Count()];
                PlaceImage(pics, titles, markAsRed);
            }
        }

        private void c_ChangePic_Click(object sender, EventArgs e)
        {
            int i = FindControl((Button)sender);
            OpenFileDialog openFileDialog = new OpenFileDialog();
            DialogResult r = openFileDialog.ShowDialog();
            if (r == DialogResult.OK)
            {
                c_imageListItems[i].ImageLocation = openFileDialog.FileName;
                //c_imageListItems[i].FileName = c_imageListItems[i].ImageLocation.Split('\\')[c_imageListItems[i].ImageLocation.Split('\\').Count() - 1];

            }
            c_imageListItems[i].SetFocus(ImageListItem.ControlList.Title);
        }

        private void c_Title_GotFocus(object sender, EventArgs e)
        {
            int i = FindControl((TextBox)sender);
            c_LargeImage.ImageLocation = c_imageListItems[i].ImageLocation;
        }

        private void c_picBoxes_Click(object sender, EventArgs e)
        {
            int i = FindControl((PictureBox)sender);
            c_LargeImage.ImageLocation = c_imageListItems[i].ImageLocation;
            c_imageListItems[i].SetFocus(ImageListItem.ControlList.Title);
        }

        private void c_MoveUpButton_Click(object sender, EventArgs e)
        {
            int i = FindControl((Button)sender);
            if (i == 0)
            {
                return;
            }

            SwitchImageListItem(c_imageListItems[i], c_imageListItems[i - 1]);
            c_imageListItems[i - 1].SetFocus(ImageListItem.ControlList.Title);

        }

        private void c_MoveDownButton_Click(object sender, EventArgs e)
        {
            int i = FindControl((Button)sender);
            if (i == c_imageListItems.Count() - 1)
            {
                return;
            }

            SwitchImageListItem(c_imageListItems[i], c_imageListItems[i + 1]);
            c_imageListItems[i + 1].SetFocus(ImageListItem.ControlList.Title);
        }

        private void c_DeletePic_Click(object sender, EventArgs e)
        {
            int i = FindControl((Button)sender);
            tempPics = new string[c_imageListItems.Count()-1];
            tempTitles = new string[c_imageListItems.Count() - 1];
            tempMarkAsRed = new bool[c_imageListItems.Count() - 1];
            
            int k = 0;
            for(int j=0;j<c_imageListItems.Count();j++)
            {
                if(j!=i)
                {
                    tempPics[k] = c_imageListItems[j].ImageLocation;
                    tempTitles[k] = c_imageListItems[j].Title;
                    tempMarkAsRed[k] = c_imageListItems[j].MarkAsRed;
                    k++;
                }
            }
            PlaceImage(tempPics, tempTitles, tempMarkAsRed);
            tempPics = null;
            tempTitles = null;
            tempMarkAsRed = null;
            GC.Collect();
        }

        private void c_AddPic_Click(object sender, EventArgs e)
        {
            if(c_imageListItems==null)
            {
                MessageBox.Show("请先选择图片");
                return;
            }
            OpenFileDialog openFileDialog = new OpenFileDialog();
            DialogResult r = openFileDialog.ShowDialog();
            if(r==DialogResult.OK)
            {
                string newPic = openFileDialog.FileName;
                tempPics = new string[c_imageListItems.Count() + 1];
                tempTitles = new string[c_imageListItems.Count() + 1];
                tempMarkAsRed = new bool[c_imageListItems.Count() + 1];
                
                for(int i=0;i<c_imageListItems.Count();i++)
                {
                    tempPics[i] = c_imageListItems[i].ImageLocation;
                    tempTitles[i] = c_imageListItems[i].Title;
                    tempMarkAsRed[i] = c_imageListItems[i].MarkAsRed;
                }
                tempPics[c_imageListItems.Count()] = newPic;
                tempTitles[c_imageListItems.Count()] = "";
                tempMarkAsRed[c_imageListItems.Count()] = false;

                PlaceImage(tempPics, tempTitles, tempMarkAsRed);
                tempPics = null;
                tempTitles = null;
                tempMarkAsRed = null;
                GC.Collect();
            }
            
        }
                
        private void C_Title_KeyDown(object sender, KeyEventArgs e)
        {
            int i = FindControl((Control)sender);
            if (e.KeyCode==Keys.Up)
            {                
                if (i > 0)
                    c_imageListItems[i - 1].SetFocus(ImageListItem.ControlList.Title);
            }
            if (e.KeyCode == Keys.Down)
            {
                if (i < c_imageListItems.Count() - 1)
                    c_imageListItems[i + 1].SetFocus(ImageListItem.ControlList.Title);                
            }
        }

        private void C_Title_KeyPress(object sender, KeyPressEventArgs e)
        {
            int i = FindControl((Control)sender);
            if (e.KeyChar== Convert.ToChar(13))
            {
                if (i < c_imageListItems.Count() - 1)
                    c_imageListItems[i + 1].SetFocus(ImageListItem.ControlList.Title);
                else
                    c_EndText.Focus();

                e.Handled = true;
            }
        }

        void PlaceImage(string[] pics, string[] titles, bool[] markAsRed)
        {
            panel1.Controls.Clear();
            c_imageListItems = null;
            c_imageListItems = new ImageListItem[pics.Count()];

            for (int i = 0; i < pics.Count(); i++)
            {
                c_imageListItems[i] = new ImageListItem();
                c_imageListItems[i].Top = i * (c_imageListItems[i].Height + 10);
                c_imageListItems[i].Left = 10;
                panel1.Controls.Add(c_imageListItems[i]);
                c_imageListItems[i].ImageLocation = pics[i];
                c_imageListItems[i].Title = titles[i];
                c_imageListItems[i].MarkAsRed = markAsRed[i];

                c_imageListItems[i].c_PicBox.Click += c_picBoxes_Click;
                c_imageListItems[i].c_Title.GotFocus += c_Title_GotFocus;
                c_imageListItems[i].c_MoveUpButton.Click += c_MoveUpButton_Click;
                c_imageListItems[i].c_MoveDownButton.Click += c_MoveDownButton_Click;
                c_imageListItems[i].c_ChangePicButton.Click += c_ChangePic_Click;
                c_imageListItems[i].c_DeletePicButton.Click += c_DeletePic_Click;
                c_imageListItems[i].c_Title.KeyDown += C_Title_KeyDown;
                c_imageListItems[i].c_Title.KeyPress += C_Title_KeyPress;

                c_LargeImage.Image = null;
            }

            if (c_imageListItems.Count() > 0)
            {
                c_imageListItems[0].HideMoveUpButton();
                c_imageListItems[c_imageListItems.Count() - 1].HideMoveDownButton();
            }
        }

        void SwitchImageListItem(ImageListItem a, ImageListItem b)
        {
            string tempTitle = a.Title;
            string tempImageLocation = a.ImageLocation;
            bool tempMarkAsRed = a.MarkAsRed;

            a.Title = b.Title;
            a.ImageLocation = b.ImageLocation;
            a.MarkAsRed = b.MarkAsRed;

            b.Title = tempTitle;
            b.ImageLocation = tempImageLocation;
            b.MarkAsRed = tempMarkAsRed;
        }

        int FindControl(Control C)
        {
            for(int i=0;i<c_imageListItems.Count();i++)
            {
                if (c_imageListItems[i].Contains(C))
                {
                    return i;
                }
            }
            return -1;
        }

        private void C_About_Click(object sender, EventArgs e)
        {
            AboutBox about = new AboutBox();
            about.Show();
        }   

        /// <summary>
        /// 取标题的后2个单词填充第四行
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void c_DocTitle_Leave(object sender, EventArgs e)
        {
            if (String.IsNullOrWhiteSpace(c_DocTitle.Text))
            {
                return;
            }
            List<string> words = c_DocTitle.Text.Trim().Split(' ').ToList<string>();
            for(int i = words.Count() - 1; i >= 0; i--)
            {
                if (String.IsNullOrWhiteSpace(words[i]))
                {
                    words.RemoveAt(i);
                }
            }

            if (words.Count() < 2)
            {
                return;
            }

            string forthLine = words[words.Count() - 2] + " " + words[words.Count() - 1];
            c_ForthLine.Text = forthLine;
        }
    }
}

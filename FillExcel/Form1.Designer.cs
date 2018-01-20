namespace FillExcel
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.c_FillExcel = new System.Windows.Forms.Button();
            this.c_SelectImage = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.c_InspectionTime = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.c_ItemNo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.c_EndText = new System.Windows.Forms.TextBox();
            this.c_LargeImage = new System.Windows.Forms.PictureBox();
            this.c_UseDefaultFileName = new System.Windows.Forms.CheckBox();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.c_AddPic = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.c_DocTitle = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.c_ForthLine = new System.Windows.Forms.TextBox();
            this.c_MarkEndTextRed = new System.Windows.Forms.CheckBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.设置ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.行高ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.图片缩放ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            ((System.ComponentModel.ISupportInitialize)(this.c_LargeImage)).BeginInit();
            this.menuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // c_FillExcel
            // 
            this.c_FillExcel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.c_FillExcel.Location = new System.Drawing.Point(509, 32);
            this.c_FillExcel.Name = "c_FillExcel";
            this.c_FillExcel.Size = new System.Drawing.Size(75, 23);
            this.c_FillExcel.TabIndex = 3;
            this.c_FillExcel.Text = "开始填充";
            this.c_FillExcel.UseVisualStyleBackColor = true;
            this.c_FillExcel.Click += new System.EventHandler(this.c_FillExcel_Click);
            // 
            // c_SelectImage
            // 
            this.c_SelectImage.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.c_SelectImage.Location = new System.Drawing.Point(428, 32);
            this.c_SelectImage.Name = "c_SelectImage";
            this.c_SelectImage.Size = new System.Drawing.Size(75, 23);
            this.c_SelectImage.TabIndex = 2;
            this.c_SelectImage.Text = "选择图片";
            this.c_SelectImage.UseVisualStyleBackColor = true;
            this.c_SelectImage.Click += new System.EventHandler(this.c_SelectImage_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.AutoScroll = true;
            this.panel1.Location = new System.Drawing.Point(0, 89);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(347, 354);
            this.panel1.TabIndex = 2;
            // 
            // c_InspectionTime
            // 
            this.c_InspectionTime.CustomFormat = "MMM dd, yyyy";
            this.c_InspectionTime.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.c_InspectionTime.Location = new System.Drawing.Point(12, 63);
            this.c_InspectionTime.Name = "c_InspectionTime";
            this.c_InspectionTime.Size = new System.Drawing.Size(157, 20);
            this.c_InspectionTime.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(175, 67);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(55, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "ITEM NO:";
            // 
            // c_ItemNo
            // 
            this.c_ItemNo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.c_ItemNo.Location = new System.Drawing.Point(234, 63);
            this.c_ItemNo.Name = "c_ItemNo";
            this.c_ItemNo.Size = new System.Drawing.Size(188, 20);
            this.c_ItemNo.TabIndex = 1;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(93, 452);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "结束文本:";
            // 
            // c_EndText
            // 
            this.c_EndText.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.c_EndText.Location = new System.Drawing.Point(157, 449);
            this.c_EndText.Name = "c_EndText";
            this.c_EndText.Size = new System.Drawing.Size(485, 20);
            this.c_EndText.TabIndex = 5;
            // 
            // c_LargeImage
            // 
            this.c_LargeImage.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.c_LargeImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.c_LargeImage.Location = new System.Drawing.Point(353, 89);
            this.c_LargeImage.Name = "c_LargeImage";
            this.c_LargeImage.Size = new System.Drawing.Size(345, 354);
            this.c_LargeImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.c_LargeImage.TabIndex = 8;
            this.c_LargeImage.TabStop = false;
            // 
            // c_UseDefaultFileName
            // 
            this.c_UseDefaultFileName.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.c_UseDefaultFileName.AutoSize = true;
            this.c_UseDefaultFileName.Checked = true;
            this.c_UseDefaultFileName.CheckState = System.Windows.Forms.CheckState.Checked;
            this.c_UseDefaultFileName.Location = new System.Drawing.Point(590, 36);
            this.c_UseDefaultFileName.Name = "c_UseDefaultFileName";
            this.c_UseDefaultFileName.Size = new System.Drawing.Size(110, 17);
            this.c_UseDefaultFileName.TabIndex = 9;
            this.c_UseDefaultFileName.Text = "以标题做文件名";
            this.c_UseDefaultFileName.UseVisualStyleBackColor = true;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Arrow_Up.png");
            this.imageList1.Images.SetKeyName(1, "Arrow_Down.png");
            this.imageList1.Images.SetKeyName(2, "DeleteRed.png");
            // 
            // c_AddPic
            // 
            this.c_AddPic.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.c_AddPic.Location = new System.Drawing.Point(12, 447);
            this.c_AddPic.Name = "c_AddPic";
            this.c_AddPic.Size = new System.Drawing.Size(75, 23);
            this.c_AddPic.TabIndex = 10;
            this.c_AddPic.Text = "增加图片";
            this.c_AddPic.UseVisualStyleBackColor = true;
            this.c_AddPic.Click += new System.EventHandler(this.c_AddPic_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 36);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(34, 13);
            this.label3.TabIndex = 11;
            this.label3.Text = "标题:";
            // 
            // c_DocTitle
            // 
            this.c_DocTitle.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.c_DocTitle.Location = new System.Drawing.Point(52, 33);
            this.c_DocTitle.Name = "c_DocTitle";
            this.c_DocTitle.Size = new System.Drawing.Size(370, 20);
            this.c_DocTitle.TabIndex = 0;
            this.c_DocTitle.Leave += new System.EventHandler(this.c_DocTitle_Leave);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(428, 67);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 13);
            this.label4.TabIndex = 13;
            this.label4.Text = "第四行:";
            // 
            // c_ForthLine
            // 
            this.c_ForthLine.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.c_ForthLine.Location = new System.Drawing.Point(477, 63);
            this.c_ForthLine.Name = "c_ForthLine";
            this.c_ForthLine.Size = new System.Drawing.Size(100, 20);
            this.c_ForthLine.TabIndex = 99;
            // 
            // c_MarkEndTextRed
            // 
            this.c_MarkEndTextRed.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.c_MarkEndTextRed.AutoSize = true;
            this.c_MarkEndTextRed.Location = new System.Drawing.Point(648, 451);
            this.c_MarkEndTextRed.Name = "c_MarkEndTextRed";
            this.c_MarkEndTextRed.Size = new System.Drawing.Size(50, 17);
            this.c_MarkEndTextRed.TabIndex = 100;
            this.c_MarkEndTextRed.Text = "标红";
            this.c_MarkEndTextRed.UseVisualStyleBackColor = true;
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.设置ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(710, 24);
            this.menuStrip1.TabIndex = 101;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 设置ToolStripMenuItem
            // 
            this.设置ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.行高ToolStripMenuItem,
            this.图片缩放ToolStripMenuItem});
            this.设置ToolStripMenuItem.Name = "设置ToolStripMenuItem";
            this.设置ToolStripMenuItem.Size = new System.Drawing.Size(45, 20);
            this.设置ToolStripMenuItem.Text = "设置";
            // 
            // 行高ToolStripMenuItem
            // 
            this.行高ToolStripMenuItem.Name = "行高ToolStripMenuItem";
            this.行高ToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.行高ToolStripMenuItem.Text = "行高";
            this.行高ToolStripMenuItem.Click += new System.EventHandler(this.行高ToolStripMenuItem_Click);
            // 
            // 图片缩放ToolStripMenuItem
            // 
            this.图片缩放ToolStripMenuItem.Name = "图片缩放ToolStripMenuItem";
            this.图片缩放ToolStripMenuItem.Size = new System.Drawing.Size(152, 22);
            this.图片缩放ToolStripMenuItem.Text = "图片缩放";
            this.图片缩放ToolStripMenuItem.Click += new System.EventHandler(this.图片缩放ToolStripMenuItem_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(710, 474);
            this.Controls.Add(this.c_MarkEndTextRed);
            this.Controls.Add(this.c_ForthLine);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.c_DocTitle);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.c_AddPic);
            this.Controls.Add(this.c_UseDefaultFileName);
            this.Controls.Add(this.c_LargeImage);
            this.Controls.Add(this.c_EndText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.c_ItemNo);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.c_InspectionTime);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.c_SelectImage);
            this.Controls.Add(this.c_FillExcel);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "Form1";
            this.Text = "图片填充工具";
            ((System.ComponentModel.ISupportInitialize)(this.c_LargeImage)).EndInit();
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button c_FillExcel;
        private System.Windows.Forms.Button c_SelectImage;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DateTimePicker c_InspectionTime;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox c_ItemNo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox c_EndText;
        private System.Windows.Forms.PictureBox c_LargeImage;
        private System.Windows.Forms.CheckBox c_UseDefaultFileName;
        private System.Windows.Forms.ImageList imageList1;
        private System.Windows.Forms.Button c_AddPic;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox c_DocTitle;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox c_ForthLine;
        private System.Windows.Forms.CheckBox c_MarkEndTextRed;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 设置ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 行高ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 图片缩放ToolStripMenuItem;
    }
}


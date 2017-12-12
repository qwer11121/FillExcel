namespace FillExcel
{
    partial class ImageListItem
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ImageListItem));
            this.c_PicBox = new System.Windows.Forms.PictureBox();
            this.c_FileName = new System.Windows.Forms.Label();
            this.c_Title = new System.Windows.Forms.TextBox();
            this.c_MarkAsRed = new System.Windows.Forms.CheckBox();
            this.c_MoveUpButton = new System.Windows.Forms.Button();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.c_MoveDownButton = new System.Windows.Forms.Button();
            this.c_ChangePicButton = new System.Windows.Forms.Button();
            this.c_DeletePicButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.c_PicBox)).BeginInit();
            this.SuspendLayout();
            // 
            // c_PicBox
            // 
            this.c_PicBox.Location = new System.Drawing.Point(36, 6);
            this.c_PicBox.Name = "c_PicBox";
            this.c_PicBox.Size = new System.Drawing.Size(50, 50);
            this.c_PicBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.c_PicBox.TabIndex = 0;
            this.c_PicBox.TabStop = false;
            // 
            // c_FileName
            // 
            this.c_FileName.AutoSize = true;
            this.c_FileName.Location = new System.Drawing.Point(92, 10);
            this.c_FileName.Name = "c_FileName";
            this.c_FileName.Size = new System.Drawing.Size(49, 13);
            this.c_FileName.TabIndex = 1;
            this.c_FileName.Text = "file name";
            // 
            // c_Title
            // 
            this.c_Title.Location = new System.Drawing.Point(94, 29);
            this.c_Title.Name = "c_Title";
            this.c_Title.Size = new System.Drawing.Size(100, 20);
            this.c_Title.TabIndex = 2;
            this.c_Title.Text = "title";
            // 
            // c_MarkAsRed
            // 
            this.c_MarkAsRed.AutoSize = true;
            this.c_MarkAsRed.Location = new System.Drawing.Point(259, 9);
            this.c_MarkAsRed.Name = "c_MarkAsRed";
            this.c_MarkAsRed.Size = new System.Drawing.Size(50, 17);
            this.c_MarkAsRed.TabIndex = 3;
            this.c_MarkAsRed.Text = "标红";
            this.c_MarkAsRed.UseVisualStyleBackColor = true;
            // 
            // c_MoveUpButton
            // 
            this.c_MoveUpButton.ImageIndex = 1;
            this.c_MoveUpButton.ImageList = this.imageList1;
            this.c_MoveUpButton.Location = new System.Drawing.Point(8, 8);
            this.c_MoveUpButton.Name = "c_MoveUpButton";
            this.c_MoveUpButton.Size = new System.Drawing.Size(22, 22);
            this.c_MoveUpButton.TabIndex = 4;
            this.c_MoveUpButton.TabStop = false;
            this.c_MoveUpButton.UseVisualStyleBackColor = true;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "Arrow_Down.png");
            this.imageList1.Images.SetKeyName(1, "Arrow_Up.png");
            this.imageList1.Images.SetKeyName(2, "DeleteRed.png");
            // 
            // c_MoveDownButton
            // 
            this.c_MoveDownButton.ImageAlign = System.Drawing.ContentAlignment.TopLeft;
            this.c_MoveDownButton.ImageIndex = 0;
            this.c_MoveDownButton.ImageList = this.imageList1;
            this.c_MoveDownButton.Location = new System.Drawing.Point(8, 32);
            this.c_MoveDownButton.Name = "c_MoveDownButton";
            this.c_MoveDownButton.Size = new System.Drawing.Size(22, 22);
            this.c_MoveDownButton.TabIndex = 5;
            this.c_MoveDownButton.TabStop = false;
            this.c_MoveDownButton.UseVisualStyleBackColor = true;
            // 
            // c_ChangePicButton
            // 
            this.c_ChangePicButton.Location = new System.Drawing.Point(202, 28);
            this.c_ChangePicButton.Name = "c_ChangePicButton";
            this.c_ChangePicButton.Size = new System.Drawing.Size(75, 23);
            this.c_ChangePicButton.TabIndex = 6;
            this.c_ChangePicButton.TabStop = false;
            this.c_ChangePicButton.Text = "更改图片";
            this.c_ChangePicButton.UseVisualStyleBackColor = true;
            this.c_ChangePicButton.Click += new System.EventHandler(this.c_ChangePicButton_Click);
            // 
            // c_DeletePicButton
            // 
            this.c_DeletePicButton.ImageIndex = 2;
            this.c_DeletePicButton.ImageList = this.imageList1;
            this.c_DeletePicButton.Location = new System.Drawing.Point(284, 29);
            this.c_DeletePicButton.Name = "c_DeletePicButton";
            this.c_DeletePicButton.Size = new System.Drawing.Size(22, 22);
            this.c_DeletePicButton.TabIndex = 7;
            this.c_DeletePicButton.TabStop = false;
            this.c_DeletePicButton.UseVisualStyleBackColor = true;
            // 
            // ImageListItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.c_DeletePicButton);
            this.Controls.Add(this.c_ChangePicButton);
            this.Controls.Add(this.c_MoveDownButton);
            this.Controls.Add(this.c_MoveUpButton);
            this.Controls.Add(this.c_MarkAsRed);
            this.Controls.Add(this.c_Title);
            this.Controls.Add(this.c_FileName);
            this.Controls.Add(this.c_PicBox);
            this.Name = "ImageListItem";
            this.Size = new System.Drawing.Size(314, 61);
            ((System.ComponentModel.ISupportInitialize)(this.c_PicBox)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label c_FileName;
        private System.Windows.Forms.ImageList imageList1;
        public System.Windows.Forms.PictureBox c_PicBox;
        public System.Windows.Forms.TextBox c_Title;
        public System.Windows.Forms.Button c_MoveUpButton;
        public System.Windows.Forms.Button c_MoveDownButton;
        public System.Windows.Forms.Button c_ChangePicButton;
        public System.Windows.Forms.Button c_DeletePicButton;
        public System.Windows.Forms.CheckBox c_MarkAsRed;
    }
}

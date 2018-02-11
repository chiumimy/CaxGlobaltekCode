namespace FixInsUpload
{
    partial class FixInsUploadDlg
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
            this.styleManager1 = new DevComponents.DotNetBar.StyleManager(this.components);
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.FixInsNo = new DevComponents.DotNetBar.LabelX();
            this.ERPNo = new DevComponents.DotNetBar.LabelX();
            this.Desc = new DevComponents.DotNetBar.LabelX();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.PicPath = new DevComponents.DotNetBar.LabelX();
            this.SelPic = new DevComponents.DotNetBar.ButtonX();
            this.Close = new DevComponents.DotNetBar.ButtonX();
            this.Upload = new DevComponents.DotNetBar.ButtonX();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.ExportPDF = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // styleManager1
            // 
            this.styleManager1.ManagerStyle = DevComponents.DotNetBar.eStyle.Windows7Blue;
            this.styleManager1.MetroColorParameters = new DevComponents.DotNetBar.Metro.ColorTables.MetroColorGeneratorParameters(System.Drawing.Color.White, System.Drawing.Color.FromArgb(((int)(((byte)(43)))), ((int)(((byte)(87)))), ((int)(((byte)(154))))));
            // 
            // labelX1
            // 
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX1.Location = new System.Drawing.Point(12, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(174, 23);
            this.labelX1.TabIndex = 3;
            this.labelX1.Text = "模、檢、治具編號：";
            // 
            // labelX2
            // 
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX2.Location = new System.Drawing.Point(12, 46);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(106, 23);
            this.labelX2.TabIndex = 4;
            this.labelX2.Text = "ERP編號：";
            // 
            // labelX3
            // 
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX3.Location = new System.Drawing.Point(12, 80);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(106, 23);
            this.labelX3.TabIndex = 5;
            this.labelX3.Text = "品名：";
            // 
            // FixInsNo
            // 
            // 
            // 
            // 
            this.FixInsNo.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.FixInsNo.Location = new System.Drawing.Point(165, 12);
            this.FixInsNo.Name = "FixInsNo";
            this.FixInsNo.Size = new System.Drawing.Size(161, 23);
            this.FixInsNo.TabIndex = 6;
            // 
            // ERPNo
            // 
            // 
            // 
            // 
            this.ERPNo.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.ERPNo.Location = new System.Drawing.Point(92, 47);
            this.ERPNo.Name = "ERPNo";
            this.ERPNo.Size = new System.Drawing.Size(234, 23);
            this.ERPNo.TabIndex = 7;
            // 
            // Desc
            // 
            // 
            // 
            // 
            this.Desc.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.Desc.Location = new System.Drawing.Point(67, 80);
            this.Desc.Name = "Desc";
            this.Desc.Size = new System.Drawing.Size(259, 23);
            this.Desc.TabIndex = 8;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // PicPath
            // 
            this.PicPath.BackColor = System.Drawing.Color.Silver;
            // 
            // 
            // 
            this.PicPath.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.PicPath.Location = new System.Drawing.Point(80, 92);
            this.PicPath.Name = "PicPath";
            this.PicPath.Size = new System.Drawing.Size(222, 23);
            this.PicPath.TabIndex = 11;
            // 
            // SelPic
            // 
            this.SelPic.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.SelPic.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.SelPic.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.SelPic.Image = global::FixInsUpload.Properties.Resources.photo_24px;
            this.SelPic.Location = new System.Drawing.Point(31, 85);
            this.SelPic.Name = "SelPic";
            this.SelPic.Size = new System.Drawing.Size(43, 37);
            this.SelPic.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.SelPic.TabIndex = 12;
            this.SelPic.Click += new System.EventHandler(this.SelPic_Click);
            // 
            // Close
            // 
            this.Close.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.Close.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.Close.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Close.Image = global::FixInsUpload.Properties.Resources.close_24px;
            this.Close.Location = new System.Drawing.Point(238, 247);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(88, 29);
            this.Close.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.Close.TabIndex = 2;
            this.Close.Text = "關閉";
            this.Close.Click += new System.EventHandler(this.Close_Click);
            // 
            // Upload
            // 
            this.Upload.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.Upload.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.Upload.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Upload.Image = global::FixInsUpload.Properties.Resources.upload_24px;
            this.Upload.Location = new System.Drawing.Point(144, 247);
            this.Upload.Name = "Upload";
            this.Upload.Size = new System.Drawing.Size(88, 29);
            this.Upload.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.Upload.TabIndex = 1;
            this.Upload.Text = "上傳";
            this.Upload.Click += new System.EventHandler(this.Upload_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.labelX4);
            this.groupBox1.Controls.Add(this.ExportPDF);
            this.groupBox1.Controls.Add(this.SelPic);
            this.groupBox1.Controls.Add(this.PicPath);
            this.groupBox1.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.groupBox1.Location = new System.Drawing.Point(12, 109);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(314, 132);
            this.groupBox1.TabIndex = 13;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "選項操作";
            // 
            // ExportPDF
            // 
            // 
            // 
            // 
            this.ExportPDF.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.ExportPDF.Location = new System.Drawing.Point(10, 26);
            this.ExportPDF.Name = "ExportPDF";
            this.ExportPDF.Size = new System.Drawing.Size(176, 23);
            this.ExportPDF.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.ExportPDF.TabIndex = 0;
            this.ExportPDF.Text = "1.輸出製程圖PDF";
            // 
            // labelX4
            // 
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Location = new System.Drawing.Point(27, 56);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(185, 23);
            this.labelX4.TabIndex = 13;
            this.labelX4.Text = "2.可指定圖片檔(選項)";
            // 
            // FixInsUploadDlg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(338, 288);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.Desc);
            this.Controls.Add(this.ERPNo);
            this.Controls.Add(this.FixInsNo);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.Close);
            this.Controls.Add(this.Upload);
            this.DoubleBuffered = true;
            this.Name = "FixInsUploadDlg";
            this.Load += new System.EventHandler(this.FixInsUploadDlg_Load);
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.StyleManager styleManager1;
        private DevComponents.DotNetBar.ButtonX Upload;
        private DevComponents.DotNetBar.ButtonX Close;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.LabelX FixInsNo;
        private DevComponents.DotNetBar.LabelX ERPNo;
        private DevComponents.DotNetBar.LabelX Desc;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private DevComponents.DotNetBar.LabelX PicPath;
        private DevComponents.DotNetBar.ButtonX SelPic;
        private System.Windows.Forms.GroupBox groupBox1;
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.DotNetBar.Controls.CheckBoxX ExportPDF;
    }
}
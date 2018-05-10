namespace MEOpenTask
{
    partial class MEOpenTaskDlg
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
            this.comboBoxCus = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.comboBoxPartNo = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.comboBoxCusVer = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.superGridPanel = new DevComponents.DotNetBar.SuperGrid.SuperGridControl();
            this.gridColumn1 = new DevComponents.DotNetBar.SuperGrid.GridColumn();
            this.gridColumn2 = new DevComponents.DotNetBar.SuperGrid.GridColumn();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.comboBoxOpVer = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX5 = new DevComponents.DotNetBar.LabelX();
            this.Close = new DevComponents.DotNetBar.ButtonX();
            this.buttonOpenTask = new DevComponents.DotNetBar.ButtonX();
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
            this.labelX1.Location = new System.Drawing.Point(13, 38);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(69, 23);
            this.labelX1.TabIndex = 1;
            this.labelX1.Text = "客戶：";
            // 
            // comboBoxCus
            // 
            this.comboBoxCus.DisplayMember = "Text";
            this.comboBoxCus.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxCus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxCus.FormattingEnabled = true;
            this.comboBoxCus.ItemHeight = 16;
            this.comboBoxCus.Location = new System.Drawing.Point(67, 36);
            this.comboBoxCus.Name = "comboBoxCus";
            this.comboBoxCus.Size = new System.Drawing.Size(238, 22);
            this.comboBoxCus.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboBoxCus.TabIndex = 2;
            this.comboBoxCus.SelectedIndexChanged += new System.EventHandler(this.comboBoxCus_SelectedIndexChanged);
            // 
            // labelX2
            // 
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX2.Location = new System.Drawing.Point(13, 67);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(69, 23);
            this.labelX2.TabIndex = 3;
            this.labelX2.Text = "料號：";
            // 
            // comboBoxPartNo
            // 
            this.comboBoxPartNo.DisplayMember = "Text";
            this.comboBoxPartNo.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxPartNo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxPartNo.FormattingEnabled = true;
            this.comboBoxPartNo.ItemHeight = 16;
            this.comboBoxPartNo.Location = new System.Drawing.Point(67, 65);
            this.comboBoxPartNo.Name = "comboBoxPartNo";
            this.comboBoxPartNo.Size = new System.Drawing.Size(238, 22);
            this.comboBoxPartNo.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboBoxPartNo.TabIndex = 4;
            this.comboBoxPartNo.SelectedIndexChanged += new System.EventHandler(this.comboBoxPartNo_SelectedIndexChanged);
            // 
            // labelX3
            // 
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX3.Location = new System.Drawing.Point(13, 96);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(93, 23);
            this.labelX3.TabIndex = 5;
            this.labelX3.Text = "客戶版本：";
            // 
            // comboBoxCusVer
            // 
            this.comboBoxCusVer.DisplayMember = "Text";
            this.comboBoxCusVer.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxCusVer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxCusVer.FormattingEnabled = true;
            this.comboBoxCusVer.ItemHeight = 16;
            this.comboBoxCusVer.Location = new System.Drawing.Point(97, 94);
            this.comboBoxCusVer.Name = "comboBoxCusVer";
            this.comboBoxCusVer.Size = new System.Drawing.Size(208, 22);
            this.comboBoxCusVer.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboBoxCusVer.TabIndex = 6;
            this.comboBoxCusVer.SelectedIndexChanged += new System.EventHandler(this.comboBoxCusVer_SelectedIndexChanged);
            // 
            // superGridPanel
            // 
            this.superGridPanel.FilterExprColors.SysFunction = System.Drawing.Color.DarkRed;
            this.superGridPanel.Location = new System.Drawing.Point(13, 151);
            this.superGridPanel.Name = "superGridPanel";
            // 
            // 
            // 
            this.superGridPanel.PrimaryGrid.Columns.Add(this.gridColumn1);
            this.superGridPanel.PrimaryGrid.Columns.Add(this.gridColumn2);
            this.superGridPanel.PrimaryGrid.MultiSelect = false;
            this.superGridPanel.PrimaryGrid.ShowRowHeaders = false;
            this.superGridPanel.Size = new System.Drawing.Size(292, 137);
            this.superGridPanel.TabIndex = 9;
            this.superGridPanel.Text = "superGridControl1";
            this.superGridPanel.CellValueChanged += new System.EventHandler<DevComponents.DotNetBar.SuperGrid.GridCellValueChangedEventArgs>(this.superGridPanel_CellValueChanged);
            // 
            // gridColumn1
            // 
            this.gridColumn1.CellStyles.Default.Alignment = DevComponents.DotNetBar.SuperGrid.Style.Alignment.MiddleCenter;
            this.gridColumn1.ColumnSortMode = DevComponents.DotNetBar.SuperGrid.ColumnSortMode.None;
            this.gridColumn1.EditorType = typeof(DevComponents.DotNetBar.SuperGrid.GridCheckBoxXEditControl);
            this.gridColumn1.FillWeight = 50;
            this.gridColumn1.HeaderText = "#";
            this.gridColumn1.Name = "製程序";
            this.gridColumn1.Width = 30;
            // 
            // gridColumn2
            // 
            this.gridColumn2.AutoSizeMode = DevComponents.DotNetBar.SuperGrid.ColumnAutoSizeMode.Fill;
            this.gridColumn2.ColumnSortMode = DevComponents.DotNetBar.SuperGrid.ColumnSortMode.None;
            this.gridColumn2.FillWeight = 50;
            this.gridColumn2.HeaderText = "檔案名";
            this.gridColumn2.Name = "檔案名";
            this.gridColumn2.ReadOnly = true;
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // labelX4
            // 
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX4.Location = new System.Drawing.Point(13, 9);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(149, 23);
            this.labelX4.TabIndex = 10;
            this.labelX4.Text = "工程師職稱：ME";
            // 
            // comboBoxOpVer
            // 
            this.comboBoxOpVer.DisplayMember = "Text";
            this.comboBoxOpVer.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboBoxOpVer.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBoxOpVer.FormattingEnabled = true;
            this.comboBoxOpVer.ItemHeight = 16;
            this.comboBoxOpVer.Location = new System.Drawing.Point(97, 122);
            this.comboBoxOpVer.Name = "comboBoxOpVer";
            this.comboBoxOpVer.Size = new System.Drawing.Size(208, 22);
            this.comboBoxOpVer.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.comboBoxOpVer.TabIndex = 11;
            this.comboBoxOpVer.SelectedIndexChanged += new System.EventHandler(this.comboBoxOpVer_SelectedIndexChanged);
            // 
            // labelX5
            // 
            this.labelX5.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX5.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX5.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelX5.Location = new System.Drawing.Point(13, 122);
            this.labelX5.Name = "labelX5";
            this.labelX5.Size = new System.Drawing.Size(93, 23);
            this.labelX5.TabIndex = 12;
            this.labelX5.Text = "製程版次：";
            // 
            // Close
            // 
            this.Close.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.Close.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.Close.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.Close.Image = global::MEOpenTask.Properties.Resources.close_16px;
            this.Close.Location = new System.Drawing.Point(230, 294);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(75, 32);
            this.Close.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.Close.TabIndex = 13;
            this.Close.Text = "關閉";
            this.Close.Click += new System.EventHandler(this.Close_Click);
            // 
            // buttonOpenTask
            // 
            this.buttonOpenTask.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonOpenTask.ColorTable = DevComponents.DotNetBar.eButtonColor.Flat;
            this.buttonOpenTask.Font = new System.Drawing.Font("標楷體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.buttonOpenTask.Image = global::MEOpenTask.Properties.Resources.confirm_16px;
            this.buttonOpenTask.Location = new System.Drawing.Point(149, 294);
            this.buttonOpenTask.Name = "buttonOpenTask";
            this.buttonOpenTask.Size = new System.Drawing.Size(75, 32);
            this.buttonOpenTask.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.buttonOpenTask.TabIndex = 8;
            this.buttonOpenTask.Text = "開啟";
            this.buttonOpenTask.Click += new System.EventHandler(this.buttonOpenTask_Click);
            // 
            // MEOpenTaskDlg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(322, 337);
            this.Controls.Add(this.Close);
            this.Controls.Add(this.comboBoxOpVer);
            this.Controls.Add(this.labelX5);
            this.Controls.Add(this.labelX4);
            this.Controls.Add(this.superGridPanel);
            this.Controls.Add(this.buttonOpenTask);
            this.Controls.Add(this.comboBoxCusVer);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.comboBoxPartNo);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.comboBoxCus);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "MEOpenTaskDlg";
            this.Text = "ME開啟料號";
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.StyleManager styleManager1;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboBoxCus;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboBoxPartNo;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboBoxCusVer;
        private DevComponents.DotNetBar.ButtonX buttonOpenTask;
        private DevComponents.DotNetBar.SuperGrid.SuperGridControl superGridPanel;
        private DevComponents.DotNetBar.SuperGrid.GridColumn gridColumn1;
        private DevComponents.DotNetBar.SuperGrid.GridColumn gridColumn2;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.DotNetBar.Controls.ComboBoxEx comboBoxOpVer;
        private DevComponents.DotNetBar.LabelX labelX5;
        private DevComponents.DotNetBar.ButtonX Close;
    }
}
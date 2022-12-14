namespace OmnicellBlueprintingTool
{
	partial class MainForm
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
			this.lbl_Title = new System.Windows.Forms.Label();
			this.btn_Quit = new System.Windows.Forms.Button();
			this.MainTabControl = new System.Windows.Forms.TabControl();
			this.BuildVisioFromExcel = new System.Windows.Forms.TabPage();
			this.t1_btn_Submit = new System.Windows.Forms.Button();
			this.t1_tb_ExcelDataFile = new System.Windows.Forms.TextBox();
			this.t1_btn_ReadExcelfile = new System.Windows.Forms.Button();
			this.t1_lbl_SelectExcelDataFile = new System.Windows.Forms.Label();
			this.BuildExcelFromVisio = new System.Windows.Forms.TabPage();
			this.t2_btn_Submit = new System.Windows.Forms.Button();
			this.t2_tb_BuildVisioFilePath = new System.Windows.Forms.TextBox();
			this.t2_tb_BuildExcelPath = new System.Windows.Forms.TextBox();
			this.t2_btn_VisioFileToRead = new System.Windows.Forms.Button();
			this.t2_btn_SetExcelPath = new System.Windows.Forms.Button();
			this.t2_lbl_VisioFileToRead = new System.Windows.Forms.Label();
			this.t2_tb_BuildExcelFileName = new System.Windows.Forms.TextBox();
			this.t2_lbl_ExcelFileName = new System.Windows.Forms.Label();
			this.t2_lbl_ExcelfilePath = new System.Windows.Forms.Label();
			this.BuildExcelFromOIS = new System.Windows.Forms.TabPage();
			this.t3_btn_Submit = new System.Windows.Forms.Button();
			this.t3_tb_BuildOISFilePath = new System.Windows.Forms.TextBox();
			this.t3_tb_BuildExcelPath = new System.Windows.Forms.TextBox();
			this.t3_btn_OISFileToRead = new System.Windows.Forms.Button();
			this.t3_btn_SetExcelPath = new System.Windows.Forms.Button();
			this.t3_lbl_OISFileToRead = new System.Windows.Forms.Label();
			this.t3_tb_BuildExcelFileName = new System.Windows.Forms.TextBox();
			this.t3_lbl_ExcelFileName = new System.Windows.Forms.Label();
			this.t3_lbl_ExcelfilePath = new System.Windows.Forms.Label();
			this.MainTabControl.SuspendLayout();
			this.BuildVisioFromExcel.SuspendLayout();
			this.BuildExcelFromVisio.SuspendLayout();
			this.BuildExcelFromOIS.SuspendLayout();
			this.SuspendLayout();
			// 
			// lbl_Title
			// 
			this.lbl_Title.AutoSize = true;
			this.lbl_Title.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lbl_Title.Location = new System.Drawing.Point(110, 17);
			this.lbl_Title.Name = "lbl_Title";
			this.lbl_Title.Size = new System.Drawing.Size(538, 24);
			this.lbl_Title.TabIndex = 20;
			this.lbl_Title.Text = "Omnicell Blueprinting Tool for building standard Visio Diagrams";
			// 
			// btn_Quit
			// 
			this.btn_Quit.Location = new System.Drawing.Point(623, 483);
			this.btn_Quit.Name = "btn_Quit";
			this.btn_Quit.Size = new System.Drawing.Size(75, 23);
			this.btn_Quit.TabIndex = 3;
			this.btn_Quit.Text = "Quit";
			this.btn_Quit.UseVisualStyleBackColor = true;
			this.btn_Quit.Click += new System.EventHandler(this.btn_Quit_Click);
			// 
			// MainTabControl
			// 
			this.MainTabControl.Controls.Add(this.BuildVisioFromExcel);
			this.MainTabControl.Controls.Add(this.BuildExcelFromVisio);
			this.MainTabControl.Controls.Add(this.BuildExcelFromOIS);
			this.MainTabControl.Location = new System.Drawing.Point(5, 58);
			this.MainTabControl.Name = "MainTabControl";
			this.MainTabControl.SelectedIndex = 0;
			this.MainTabControl.Size = new System.Drawing.Size(730, 407);
			this.MainTabControl.TabIndex = 25;
			// 
			// BuildVisioFromExcel
			// 
			this.BuildVisioFromExcel.BackColor = System.Drawing.SystemColors.GradientInactiveCaption;
			this.BuildVisioFromExcel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.BuildVisioFromExcel.Controls.Add(this.t1_btn_Submit);
			this.BuildVisioFromExcel.Controls.Add(this.t1_tb_ExcelDataFile);
			this.BuildVisioFromExcel.Controls.Add(this.t1_btn_ReadExcelfile);
			this.BuildVisioFromExcel.Controls.Add(this.t1_lbl_SelectExcelDataFile);
			this.BuildVisioFromExcel.Location = new System.Drawing.Point(4, 22);
			this.BuildVisioFromExcel.Name = "BuildVisioFromExcel";
			this.BuildVisioFromExcel.Padding = new System.Windows.Forms.Padding(3);
			this.BuildVisioFromExcel.Size = new System.Drawing.Size(722, 381);
			this.BuildVisioFromExcel.TabIndex = 0;
			this.BuildVisioFromExcel.Text = "Build Visio from Excel data";
			// 
			// t1_btn_Submit
			// 
			this.t1_btn_Submit.Location = new System.Drawing.Point(326, 174);
			this.t1_btn_Submit.Name = "t1_btn_Submit";
			this.t1_btn_Submit.Size = new System.Drawing.Size(75, 23);
			this.t1_btn_Submit.TabIndex = 36;
			this.t1_btn_Submit.Text = "Submit";
			this.t1_btn_Submit.UseVisualStyleBackColor = true;
			this.t1_btn_Submit.Click += new System.EventHandler(this.t1_btn_Submit_Click);
			// 
			// t1_tb_ExcelDataFile
			// 
			this.t1_tb_ExcelDataFile.Location = new System.Drawing.Point(145, 84);
			this.t1_tb_ExcelDataFile.Name = "t1_tb_ExcelDataFile";
			this.t1_tb_ExcelDataFile.ReadOnly = true;
			this.t1_tb_ExcelDataFile.Size = new System.Drawing.Size(507, 20);
			this.t1_tb_ExcelDataFile.TabIndex = 35;
			// 
			// t1_btn_ReadExcelfile
			// 
			this.t1_btn_ReadExcelfile.Location = new System.Drawing.Point(658, 82);
			this.t1_btn_ReadExcelfile.Name = "t1_btn_ReadExcelfile";
			this.t1_btn_ReadExcelfile.Size = new System.Drawing.Size(31, 23);
			this.t1_btn_ReadExcelfile.TabIndex = 34;
			this.t1_btn_ReadExcelfile.Text = "...";
			this.t1_btn_ReadExcelfile.UseVisualStyleBackColor = true;
			this.t1_btn_ReadExcelfile.Click += new System.EventHandler(this.t1_btn_readExcelfile_Click);
			// 
			// t1_lbl_SelectExcelDataFile
			// 
			this.t1_lbl_SelectExcelDataFile.AutoSize = true;
			this.t1_lbl_SelectExcelDataFile.Location = new System.Drawing.Point(32, 86);
			this.t1_lbl_SelectExcelDataFile.Name = "t1_lbl_SelectExcelDataFile";
			this.t1_lbl_SelectExcelDataFile.Size = new System.Drawing.Size(111, 13);
			this.t1_lbl_SelectExcelDataFile.TabIndex = 33;
			this.t1_lbl_SelectExcelDataFile.Text = "Select Excel Data File";
			// 
			// BuildExcelFromVisio
			// 
			this.BuildExcelFromVisio.BackColor = System.Drawing.SystemColors.GradientActiveCaption;
			this.BuildExcelFromVisio.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.BuildExcelFromVisio.Controls.Add(this.t2_btn_Submit);
			this.BuildExcelFromVisio.Controls.Add(this.t2_tb_BuildVisioFilePath);
			this.BuildExcelFromVisio.Controls.Add(this.t2_tb_BuildExcelPath);
			this.BuildExcelFromVisio.Controls.Add(this.t2_btn_VisioFileToRead);
			this.BuildExcelFromVisio.Controls.Add(this.t2_btn_SetExcelPath);
			this.BuildExcelFromVisio.Controls.Add(this.t2_lbl_VisioFileToRead);
			this.BuildExcelFromVisio.Controls.Add(this.t2_tb_BuildExcelFileName);
			this.BuildExcelFromVisio.Controls.Add(this.t2_lbl_ExcelFileName);
			this.BuildExcelFromVisio.Controls.Add(this.t2_lbl_ExcelfilePath);
			this.BuildExcelFromVisio.Location = new System.Drawing.Point(4, 22);
			this.BuildExcelFromVisio.Name = "BuildExcelFromVisio";
			this.BuildExcelFromVisio.Padding = new System.Windows.Forms.Padding(3);
			this.BuildExcelFromVisio.Size = new System.Drawing.Size(722, 381);
			this.BuildExcelFromVisio.TabIndex = 1;
			this.BuildExcelFromVisio.Text = " Build Excel Data from Visio Diagram";
			// 
			// t2_btn_Submit
			// 
			this.t2_btn_Submit.Location = new System.Drawing.Point(326, 174);
			this.t2_btn_Submit.Name = "t2_btn_Submit";
			this.t2_btn_Submit.Size = new System.Drawing.Size(75, 23);
			this.t2_btn_Submit.TabIndex = 33;
			this.t2_btn_Submit.Text = "Submit";
			this.t2_btn_Submit.UseVisualStyleBackColor = true;
			this.t2_btn_Submit.Click += new System.EventHandler(this.t2_btn_Submit_Click);
			// 
			// t2_tb_BuildVisioFilePath
			// 
			this.t2_tb_BuildVisioFilePath.Location = new System.Drawing.Point(132, 115);
			this.t2_tb_BuildVisioFilePath.Name = "t2_tb_BuildVisioFilePath";
			this.t2_tb_BuildVisioFilePath.ReadOnly = true;
			this.t2_tb_BuildVisioFilePath.Size = new System.Drawing.Size(507, 20);
			this.t2_tb_BuildVisioFilePath.TabIndex = 32;
			// 
			// t2_tb_BuildExcelPath
			// 
			this.t2_tb_BuildExcelPath.Location = new System.Drawing.Point(133, 46);
			this.t2_tb_BuildExcelPath.Name = "t2_tb_BuildExcelPath";
			this.t2_tb_BuildExcelPath.ReadOnly = true;
			this.t2_tb_BuildExcelPath.Size = new System.Drawing.Size(506, 20);
			this.t2_tb_BuildExcelPath.TabIndex = 31;
			// 
			// t2_btn_VisioFileToRead
			// 
			this.t2_btn_VisioFileToRead.Location = new System.Drawing.Point(645, 113);
			this.t2_btn_VisioFileToRead.Name = "t2_btn_VisioFileToRead";
			this.t2_btn_VisioFileToRead.Size = new System.Drawing.Size(31, 23);
			this.t2_btn_VisioFileToRead.TabIndex = 30;
			this.t2_btn_VisioFileToRead.Text = "...";
			this.t2_btn_VisioFileToRead.UseVisualStyleBackColor = true;
			this.t2_btn_VisioFileToRead.Click += new System.EventHandler(this.t2_btn_VisioFileToRead_Click);
			// 
			// t2_btn_SetExcelPath
			// 
			this.t2_btn_SetExcelPath.Location = new System.Drawing.Point(645, 44);
			this.t2_btn_SetExcelPath.Name = "t2_btn_SetExcelPath";
			this.t2_btn_SetExcelPath.Size = new System.Drawing.Size(31, 23);
			this.t2_btn_SetExcelPath.TabIndex = 29;
			this.t2_btn_SetExcelPath.Text = "...";
			this.t2_btn_SetExcelPath.UseVisualStyleBackColor = true;
			this.t2_btn_SetExcelPath.Click += new System.EventHandler(this.t2_btn_openExcelPath_Click);
			// 
			// t2_lbl_VisioFileToRead
			// 
			this.t2_lbl_VisioFileToRead.AutoSize = true;
			this.t2_lbl_VisioFileToRead.Location = new System.Drawing.Point(44, 118);
			this.t2_lbl_VisioFileToRead.Name = "t2_lbl_VisioFileToRead";
			this.t2_lbl_VisioFileToRead.Size = new System.Drawing.Size(86, 13);
			this.t2_lbl_VisioFileToRead.TabIndex = 28;
			this.t2_lbl_VisioFileToRead.Text = "Visio file to Read";
			// 
			// t2_tb_BuildExcelFileName
			// 
			this.t2_tb_BuildExcelFileName.Location = new System.Drawing.Point(133, 80);
			this.t2_tb_BuildExcelFileName.Name = "t2_tb_BuildExcelFileName";
			this.t2_tb_BuildExcelFileName.ReadOnly = true;
			this.t2_tb_BuildExcelFileName.Size = new System.Drawing.Size(333, 20);
			this.t2_tb_BuildExcelFileName.TabIndex = 27;
			this.t2_tb_BuildExcelFileName.TextChanged += new System.EventHandler(this.t2_tb_buildExcelFileName_TextChanged);
			// 
			// t2_lbl_ExcelFileName
			// 
			this.t2_lbl_ExcelFileName.AutoSize = true;
			this.t2_lbl_ExcelFileName.Location = new System.Drawing.Point(52, 84);
			this.t2_lbl_ExcelFileName.Name = "t2_lbl_ExcelFileName";
			this.t2_lbl_ExcelFileName.Size = new System.Drawing.Size(78, 13);
			this.t2_lbl_ExcelFileName.TabIndex = 26;
			this.t2_lbl_ExcelFileName.Text = "Excel file name";
			// 
			// t2_lbl_ExcelfilePath
			// 
			this.t2_lbl_ExcelfilePath.AutoSize = true;
			this.t2_lbl_ExcelfilePath.Location = new System.Drawing.Point(56, 50);
			this.t2_lbl_ExcelfilePath.Name = "t2_lbl_ExcelfilePath";
			this.t2_lbl_ExcelfilePath.Size = new System.Drawing.Size(74, 13);
			this.t2_lbl_ExcelfilePath.TabIndex = 25;
			this.t2_lbl_ExcelfilePath.Text = "Excel file Path";
			// 
			// BuildExcelFromOIS
			// 
			this.BuildExcelFromOIS.BackColor = System.Drawing.Color.SkyBlue;
			this.BuildExcelFromOIS.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.BuildExcelFromOIS.Controls.Add(this.t3_btn_Submit);
			this.BuildExcelFromOIS.Controls.Add(this.t3_tb_BuildOISFilePath);
			this.BuildExcelFromOIS.Controls.Add(this.t3_tb_BuildExcelPath);
			this.BuildExcelFromOIS.Controls.Add(this.t3_btn_OISFileToRead);
			this.BuildExcelFromOIS.Controls.Add(this.t3_btn_SetExcelPath);
			this.BuildExcelFromOIS.Controls.Add(this.t3_lbl_OISFileToRead);
			this.BuildExcelFromOIS.Controls.Add(this.t3_tb_BuildExcelFileName);
			this.BuildExcelFromOIS.Controls.Add(this.t3_lbl_ExcelFileName);
			this.BuildExcelFromOIS.Controls.Add(this.t3_lbl_ExcelfilePath);
			this.BuildExcelFromOIS.Location = new System.Drawing.Point(4, 22);
			this.BuildExcelFromOIS.Name = "BuildExcelFromOIS";
			this.BuildExcelFromOIS.Padding = new System.Windows.Forms.Padding(3);
			this.BuildExcelFromOIS.Size = new System.Drawing.Size(722, 381);
			this.BuildExcelFromOIS.TabIndex = 2;
			this.BuildExcelFromOIS.Text = "Build Excel Data from OIS data";
			// 
			// t3_btn_Submit
			// 
			this.t3_btn_Submit.Location = new System.Drawing.Point(326, 174);
			this.t3_btn_Submit.Name = "t3_btn_Submit";
			this.t3_btn_Submit.Size = new System.Drawing.Size(75, 23);
			this.t3_btn_Submit.TabIndex = 41;
			this.t3_btn_Submit.Text = "Submit";
			this.t3_btn_Submit.UseVisualStyleBackColor = true;
			this.t3_btn_Submit.Click += new System.EventHandler(this.t3_btn_Submit_Click);
			// 
			// t3_tb_BuildOISFilePath
			// 
			this.t3_tb_BuildOISFilePath.Location = new System.Drawing.Point(132, 115);
			this.t3_tb_BuildOISFilePath.Name = "t3_tb_BuildOISFilePath";
			this.t3_tb_BuildOISFilePath.ReadOnly = true;
			this.t3_tb_BuildOISFilePath.Size = new System.Drawing.Size(507, 20);
			this.t3_tb_BuildOISFilePath.TabIndex = 40;
			// 
			// t3_tb_BuildExcelPath
			// 
			this.t3_tb_BuildExcelPath.Location = new System.Drawing.Point(133, 46);
			this.t3_tb_BuildExcelPath.Name = "t3_tb_BuildExcelPath";
			this.t3_tb_BuildExcelPath.ReadOnly = true;
			this.t3_tb_BuildExcelPath.Size = new System.Drawing.Size(506, 20);
			this.t3_tb_BuildExcelPath.TabIndex = 39;
			// 
			// t3_btn_OISFileToRead
			// 
			this.t3_btn_OISFileToRead.Location = new System.Drawing.Point(645, 113);
			this.t3_btn_OISFileToRead.Name = "t3_btn_OISFileToRead";
			this.t3_btn_OISFileToRead.Size = new System.Drawing.Size(31, 23);
			this.t3_btn_OISFileToRead.TabIndex = 38;
			this.t3_btn_OISFileToRead.Text = "...";
			this.t3_btn_OISFileToRead.UseVisualStyleBackColor = true;
			this.t3_btn_OISFileToRead.Click += new System.EventHandler(this.t3_btn_OISFileToRead_Click);
			// 
			// t3_btn_SetExcelPath
			// 
			this.t3_btn_SetExcelPath.Location = new System.Drawing.Point(645, 44);
			this.t3_btn_SetExcelPath.Name = "t3_btn_SetExcelPath";
			this.t3_btn_SetExcelPath.Size = new System.Drawing.Size(31, 23);
			this.t3_btn_SetExcelPath.TabIndex = 37;
			this.t3_btn_SetExcelPath.Text = "...";
			this.t3_btn_SetExcelPath.UseVisualStyleBackColor = true;
			this.t3_btn_SetExcelPath.Click += new System.EventHandler(this.t3_btn_openExcelPath_Click);
			// 
			// t3_lbl_OISFileToRead
			// 
			this.t3_lbl_OISFileToRead.AutoSize = true;
			this.t3_lbl_OISFileToRead.Location = new System.Drawing.Point(17, 118);
			this.t3_lbl_OISFileToRead.Name = "t3_lbl_OISFileToRead";
			this.t3_lbl_OISFileToRead.Size = new System.Drawing.Size(113, 13);
			this.t3_lbl_OISFileToRead.TabIndex = 36;
			this.t3_lbl_OISFileToRead.Text = "OIS Setup file to Read";
			// 
			// t3_tb_BuildExcelFileName
			// 
			this.t3_tb_BuildExcelFileName.Location = new System.Drawing.Point(133, 80);
			this.t3_tb_BuildExcelFileName.Name = "t3_tb_BuildExcelFileName";
			this.t3_tb_BuildExcelFileName.ReadOnly = true;
			this.t3_tb_BuildExcelFileName.Size = new System.Drawing.Size(333, 20);
			this.t3_tb_BuildExcelFileName.TabIndex = 35;
			this.t3_tb_BuildExcelFileName.TextChanged += new System.EventHandler(this.t3_tb_buildExcelFileName_TextChanged);
			// 
			// t3_lbl_ExcelFileName
			// 
			this.t3_lbl_ExcelFileName.AutoSize = true;
			this.t3_lbl_ExcelFileName.Location = new System.Drawing.Point(52, 84);
			this.t3_lbl_ExcelFileName.Name = "t3_lbl_ExcelFileName";
			this.t3_lbl_ExcelFileName.Size = new System.Drawing.Size(78, 13);
			this.t3_lbl_ExcelFileName.TabIndex = 34;
			this.t3_lbl_ExcelFileName.Text = "Excel file name";
			// 
			// t3_lbl_ExcelfilePath
			// 
			this.t3_lbl_ExcelfilePath.AutoSize = true;
			this.t3_lbl_ExcelfilePath.Location = new System.Drawing.Point(55, 50);
			this.t3_lbl_ExcelfilePath.Name = "t3_lbl_ExcelfilePath";
			this.t3_lbl_ExcelfilePath.Size = new System.Drawing.Size(74, 13);
			this.t3_lbl_ExcelfilePath.TabIndex = 33;
			this.t3_lbl_ExcelfilePath.Text = "Excel file Path";
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.BackColor = System.Drawing.SystemColors.ControlLight;
			this.ClientSize = new System.Drawing.Size(740, 522);
			this.Controls.Add(this.MainTabControl);
			this.Controls.Add(this.lbl_Title);
			this.Controls.Add(this.btn_Quit);
			this.Name = "MainForm";
			this.Text = "Omnicell Blueprinting Tool";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.MainTabControl.ResumeLayout(false);
			this.BuildVisioFromExcel.ResumeLayout(false);
			this.BuildVisioFromExcel.PerformLayout();
			this.BuildExcelFromVisio.ResumeLayout(false);
			this.BuildExcelFromVisio.PerformLayout();
			this.BuildExcelFromOIS.ResumeLayout(false);
			this.BuildExcelFromOIS.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.Windows.Forms.Label lbl_Title;
		private System.Windows.Forms.Button btn_Quit;
		private System.Windows.Forms.TabControl MainTabControl;
		private System.Windows.Forms.TabPage BuildVisioFromExcel;
		private System.Windows.Forms.TabPage BuildExcelFromVisio;
		private System.Windows.Forms.TabPage BuildExcelFromOIS;
		private System.Windows.Forms.TextBox t1_tb_ExcelDataFile;
		private System.Windows.Forms.Button t1_btn_ReadExcelfile;
		private System.Windows.Forms.Label t1_lbl_SelectExcelDataFile;
		private System.Windows.Forms.TextBox t2_tb_BuildVisioFilePath;
		private System.Windows.Forms.TextBox t2_tb_BuildExcelPath;
		private System.Windows.Forms.Button t2_btn_VisioFileToRead;
		private System.Windows.Forms.Button t2_btn_SetExcelPath;
		private System.Windows.Forms.Label t2_lbl_VisioFileToRead;
		private System.Windows.Forms.TextBox t2_tb_BuildExcelFileName;
		private System.Windows.Forms.Label t2_lbl_ExcelFileName;
		private System.Windows.Forms.Label t2_lbl_ExcelfilePath;
		private System.Windows.Forms.TextBox t3_tb_BuildOISFilePath;
		private System.Windows.Forms.TextBox t3_tb_BuildExcelPath;
		private System.Windows.Forms.Button t3_btn_OISFileToRead;
		private System.Windows.Forms.Button t3_btn_SetExcelPath;
		private System.Windows.Forms.Label t3_lbl_OISFileToRead;
		private System.Windows.Forms.TextBox t3_tb_BuildExcelFileName;
		private System.Windows.Forms.Label t3_lbl_ExcelFileName;
		private System.Windows.Forms.Label t3_lbl_ExcelfilePath;
		private System.Windows.Forms.Button t1_btn_Submit;
		private System.Windows.Forms.Button t2_btn_Submit;
		private System.Windows.Forms.Button t3_btn_Submit;
	}
}


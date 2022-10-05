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
			this.label4 = new System.Windows.Forms.Label();
			this.gb_BuildFromExceldata = new System.Windows.Forms.GroupBox();
			this.tb_excelDataFile = new System.Windows.Forms.TextBox();
			this.btn_readExcelfile = new System.Windows.Forms.Button();
			this.label6 = new System.Windows.Forms.Label();
			this.gb_BuildDataFile = new System.Windows.Forms.GroupBox();
			this.tb_buildVisioFilePath = new System.Windows.Forms.TextBox();
			this.tb_buildExcelPath = new System.Windows.Forms.TextBox();
			this.btn_VisioFileToRead = new System.Windows.Forms.Button();
			this.btn_SetExcelPath = new System.Windows.Forms.Button();
			this.label3 = new System.Windows.Forms.Label();
			this.tb_buildExcelFileName = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label1 = new System.Windows.Forms.Label();
			this.gb_SelectOperation = new System.Windows.Forms.GroupBox();
			this.rb_buildExcelFileFromVisio = new System.Windows.Forms.RadioButton();
			this.rb_buildFromExcelFile = new System.Windows.Forms.RadioButton();
			this.btn_Submit = new System.Windows.Forms.Button();
			this.btn_Quit = new System.Windows.Forms.Button();
			this.gb_BuildFromExceldata.SuspendLayout();
			this.gb_BuildDataFile.SuspendLayout();
			this.gb_SelectOperation.SuspendLayout();
			this.SuspendLayout();
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label4.Location = new System.Drawing.Point(195, 24);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(421, 24);
			this.label4.TabIndex = 20;
			this.label4.Text = "Build a new Visio Diagram or Build Excel Data file";
			// 
			// gb_BuildFromExceldata
			// 
			this.gb_BuildFromExceldata.Controls.Add(this.tb_excelDataFile);
			this.gb_BuildFromExceldata.Controls.Add(this.btn_readExcelfile);
			this.gb_BuildFromExceldata.Controls.Add(this.label6);
			this.gb_BuildFromExceldata.Location = new System.Drawing.Point(70, 146);
			this.gb_BuildFromExceldata.Name = "gb_BuildFromExceldata";
			this.gb_BuildFromExceldata.Size = new System.Drawing.Size(686, 89);
			this.gb_BuildFromExceldata.TabIndex = 22;
			this.gb_BuildFromExceldata.TabStop = false;
			this.gb_BuildFromExceldata.Text = "Build Visio Diagram from an Excel Data file";
			// 
			// tb_excelDataFile
			// 
			this.tb_excelDataFile.Location = new System.Drawing.Point(132, 39);
			this.tb_excelDataFile.Name = "tb_excelDataFile";
			this.tb_excelDataFile.ReadOnly = true;
			this.tb_excelDataFile.Size = new System.Drawing.Size(507, 20);
			this.tb_excelDataFile.TabIndex = 32;
			// 
			// btn_readExcelfile
			// 
			this.btn_readExcelfile.Location = new System.Drawing.Point(645, 37);
			this.btn_readExcelfile.Name = "btn_readExcelfile";
			this.btn_readExcelfile.Size = new System.Drawing.Size(31, 23);
			this.btn_readExcelfile.TabIndex = 31;
			this.btn_readExcelfile.Text = "...";
			this.btn_readExcelfile.UseVisualStyleBackColor = true;
			this.btn_readExcelfile.Click += new System.EventHandler(this.btn_readExcelfile_Click);
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(19, 41);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(111, 13);
			this.label6.TabIndex = 29;
			this.label6.Text = "Select Excel Data File";
			// 
			// gb_BuildDataFile
			// 
			this.gb_BuildDataFile.Controls.Add(this.tb_buildVisioFilePath);
			this.gb_BuildDataFile.Controls.Add(this.tb_buildExcelPath);
			this.gb_BuildDataFile.Controls.Add(this.btn_VisioFileToRead);
			this.gb_BuildDataFile.Controls.Add(this.btn_SetExcelPath);
			this.gb_BuildDataFile.Controls.Add(this.label3);
			this.gb_BuildDataFile.Controls.Add(this.tb_buildExcelFileName);
			this.gb_BuildDataFile.Controls.Add(this.label2);
			this.gb_BuildDataFile.Controls.Add(this.label1);
			this.gb_BuildDataFile.Location = new System.Drawing.Point(70, 247);
			this.gb_BuildDataFile.Name = "gb_BuildDataFile";
			this.gb_BuildDataFile.Size = new System.Drawing.Size(686, 146);
			this.gb_BuildDataFile.TabIndex = 23;
			this.gb_BuildDataFile.TabStop = false;
			this.gb_BuildDataFile.Text = "Build Data file from a Visio diagram";
			// 
			// tb_buildVisioFilePath
			// 
			this.tb_buildVisioFilePath.Location = new System.Drawing.Point(132, 98);
			this.tb_buildVisioFilePath.Name = "tb_buildVisioFilePath";
			this.tb_buildVisioFilePath.ReadOnly = true;
			this.tb_buildVisioFilePath.Size = new System.Drawing.Size(507, 20);
			this.tb_buildVisioFilePath.TabIndex = 11;
			// 
			// tb_buildExcelPath
			// 
			this.tb_buildExcelPath.Location = new System.Drawing.Point(133, 29);
			this.tb_buildExcelPath.Name = "tb_buildExcelPath";
			this.tb_buildExcelPath.ReadOnly = true;
			this.tb_buildExcelPath.Size = new System.Drawing.Size(506, 20);
			this.tb_buildExcelPath.TabIndex = 10;
			// 
			// btn_VisioFileToRead
			// 
			this.btn_VisioFileToRead.Location = new System.Drawing.Point(645, 96);
			this.btn_VisioFileToRead.Name = "btn_VisioFileToRead";
			this.btn_VisioFileToRead.Size = new System.Drawing.Size(31, 23);
			this.btn_VisioFileToRead.TabIndex = 9;
			this.btn_VisioFileToRead.Text = "...";
			this.btn_VisioFileToRead.UseVisualStyleBackColor = true;
			this.btn_VisioFileToRead.Click += new System.EventHandler(this.btn_VisioFileToRead_Click);
			// 
			// btn_SetExcelPath
			// 
			this.btn_SetExcelPath.Location = new System.Drawing.Point(645, 27);
			this.btn_SetExcelPath.Name = "btn_SetExcelPath";
			this.btn_SetExcelPath.Size = new System.Drawing.Size(31, 23);
			this.btn_SetExcelPath.TabIndex = 8;
			this.btn_SetExcelPath.Text = "...";
			this.btn_SetExcelPath.UseVisualStyleBackColor = true;
			this.btn_SetExcelPath.Click += new System.EventHandler(this.btn_openExcelPath_Click);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(44, 101);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(86, 13);
			this.label3.TabIndex = 6;
			this.label3.Text = "Visio file to Read";
			// 
			// tb_buildExcelFileName
			// 
			this.tb_buildExcelFileName.Location = new System.Drawing.Point(133, 63);
			this.tb_buildExcelFileName.Name = "tb_buildExcelFileName";
			this.tb_buildExcelFileName.ReadOnly = true;
			this.tb_buildExcelFileName.Size = new System.Drawing.Size(333, 20);
			this.tb_buildExcelFileName.TabIndex = 5;
			this.tb_buildExcelFileName.TextChanged += new System.EventHandler(this.tb_buildExcelFileName_TextChanged);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(52, 67);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(78, 13);
			this.label2.TabIndex = 4;
			this.label2.Text = "Excel file name";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(56, 33);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(74, 13);
			this.label1.TabIndex = 2;
			this.label1.Text = "Excel file Path";
			// 
			// gb_SelectOperation
			// 
			this.gb_SelectOperation.Controls.Add(this.rb_buildExcelFileFromVisio);
			this.gb_SelectOperation.Controls.Add(this.rb_buildFromExcelFile);
			this.gb_SelectOperation.Location = new System.Drawing.Point(70, 59);
			this.gb_SelectOperation.Name = "gb_SelectOperation";
			this.gb_SelectOperation.Size = new System.Drawing.Size(686, 72);
			this.gb_SelectOperation.TabIndex = 24;
			this.gb_SelectOperation.TabStop = false;
			this.gb_SelectOperation.Text = "Select desired operation";
			// 
			// rb_buildExcelFileFromVisio
			// 
			this.rb_buildExcelFileFromVisio.AutoSize = true;
			this.rb_buildExcelFileFromVisio.Location = new System.Drawing.Point(370, 32);
			this.rb_buildExcelFileFromVisio.Name = "rb_buildExcelFileFromVisio";
			this.rb_buildExcelFileFromVisio.Size = new System.Drawing.Size(216, 17);
			this.rb_buildExcelFileFromVisio.TabIndex = 3;
			this.rb_buildExcelFileFromVisio.TabStop = true;
			this.rb_buildExcelFileFromVisio.Text = "Build Excel Data file from a Visio diagram";
			this.rb_buildExcelFileFromVisio.UseVisualStyleBackColor = true;
			this.rb_buildExcelFileFromVisio.CheckedChanged += new System.EventHandler(this.rb_buildDataFileFromVisio_CheckedChanged);
			// 
			// rb_buildFromExcelFile
			// 
			this.rb_buildFromExcelFile.AutoSize = true;
			this.rb_buildFromExcelFile.Location = new System.Drawing.Point(87, 32);
			this.rb_buildFromExcelFile.Name = "rb_buildFromExcelFile";
			this.rb_buildFromExcelFile.Size = new System.Drawing.Size(232, 17);
			this.rb_buildFromExcelFile.TabIndex = 2;
			this.rb_buildFromExcelFile.TabStop = true;
			this.rb_buildFromExcelFile.Text = "Build new Visio Diagram from Excel Data file";
			this.rb_buildFromExcelFile.UseVisualStyleBackColor = true;
			this.rb_buildFromExcelFile.CheckedChanged += new System.EventHandler(this.rb_buildFromDataFile_CheckedChanged);
			// 
			// btn_Submit
			// 
			this.btn_Submit.Location = new System.Drawing.Point(570, 410);
			this.btn_Submit.Name = "btn_Submit";
			this.btn_Submit.Size = new System.Drawing.Size(75, 23);
			this.btn_Submit.TabIndex = 2;
			this.btn_Submit.Text = "Submit";
			this.btn_Submit.UseVisualStyleBackColor = true;
			this.btn_Submit.Click += new System.EventHandler(this.btn_Submit_Click);
			// 
			// btn_Quit
			// 
			this.btn_Quit.Location = new System.Drawing.Point(680, 410);
			this.btn_Quit.Name = "btn_Quit";
			this.btn_Quit.Size = new System.Drawing.Size(75, 23);
			this.btn_Quit.TabIndex = 3;
			this.btn_Quit.Text = "Quit";
			this.btn_Quit.UseVisualStyleBackColor = true;
			this.btn_Quit.Click += new System.EventHandler(this.btn_Quit_Click);
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(800, 455);
			this.Controls.Add(this.gb_SelectOperation);
			this.Controls.Add(this.gb_BuildDataFile);
			this.Controls.Add(this.gb_BuildFromExceldata);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.btn_Quit);
			this.Controls.Add(this.btn_Submit);
			this.Name = "MainForm";
			this.Text = "Omnicell Diagram Creator";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.gb_BuildFromExceldata.ResumeLayout(false);
			this.gb_BuildFromExceldata.PerformLayout();
			this.gb_BuildDataFile.ResumeLayout(false);
			this.gb_BuildDataFile.PerformLayout();
			this.gb_SelectOperation.ResumeLayout(false);
			this.gb_SelectOperation.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.GroupBox gb_BuildFromExceldata;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.GroupBox gb_BuildDataFile;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox tb_buildExcelFileName;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox gb_SelectOperation;
		private System.Windows.Forms.RadioButton rb_buildExcelFileFromVisio;
		private System.Windows.Forms.RadioButton rb_buildFromExcelFile;
		private System.Windows.Forms.Button btn_VisioFileToRead;
		private System.Windows.Forms.Button btn_SetExcelPath;
		private System.Windows.Forms.TextBox tb_buildVisioFilePath;
		private System.Windows.Forms.TextBox tb_buildExcelPath;
		private System.Windows.Forms.Button btn_readExcelfile;
		private System.Windows.Forms.TextBox tb_excelDataFile;
		private System.Windows.Forms.Button btn_Submit;
		private System.Windows.Forms.Button btn_Quit;
	}
}


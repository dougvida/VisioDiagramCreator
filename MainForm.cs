using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using VisioDiagramCreator.Models;
using VisioDiagramCreator.Visio;
using VisioDiagramCreator.Extensions;
using VisioDiagramCreator.Visio.Models;
using VisioDiagramCreator.ExcelHelpers;

namespace VisioDiagramCreator
{
	public partial class MainForm : Form
	{
		DiagramData diagramData = null;
		Boolean _bBuildVisioFromExcelDataFile = true;

		VisioHelper visHlp = new VisioHelper();

		string excelDataPath = _baseWorkingDir + "\\ExcelData\\";

		static string _baseWorkingDir = @"C:\Omnicell_Diagram_Creator";
		static string _scriptDataPath = _baseWorkingDir + @"\ScriptData";

		public MainForm()
		{
			InitializeComponent();

			diagramData = new DiagramData();
			//diagramData.TemplateFilePath = @"C:\Omnicell_Diagram_Creator\Templates\OC_ArchitectDiagramTemplate.vstx";
			//diagramData.StencilFilePath = @"C:\Omnicell_Diagram_Creator\Stencils\OC_ArchitectStencils.vssx";
		}

		private void MainForm_Load(object sender, EventArgs e)
		{
			///////////////////////////////////////////////////////////////////////
			// this section is for Building the excel data file from a Visio file
			tb_buildExcelFileName.Text = "NewExcelData.xlxs";
			tb_buildExcelPath.Text = excelDataPath;
			//////////////////////////////////////////////////////////////////////

			_bBuildVisioFromExcelDataFile = true;

			btn_Submit.Enabled = false;

			btn_SetExcelPath.Enabled = false;
			btn_VisioFileToRead.Enabled = false;

			rb_buildFromExcelFile.Checked = true;
			tb_buildExcelFileName.Enabled = false;
			tb_buildExcelPath.Enabled = false;
			tb_buildVisioFilePath.Enabled = false;
		}

		private void btn_Quit_Click(object sender, EventArgs e)
		{
			visHlp.VisioForceCloseAll();
			this.Close();
		}

		private void btn_Submit_Click(object sender, EventArgs e)
		{
			// parse the data file and draw the visio diagram
			try
			{
				if (_bBuildVisioFromExcelDataFile)
				{
					// build visio file form data file
					Console.WriteLine(String.Format("MainForm - Build Visio file from an excel data file:{0}", tb_excelDataFile.Text));
					diagramData = new ProcessExcelDataFile().ParseData(tb_excelDataFile.Text.Trim(), diagramData);
					if (diagramData == null)
					{
						MessageBox.Show("MainForm - ERROR: _parseData returned null");
						visHlp.VisioForceCloseAll();
						this.Close();
					}

					visHlp.DrawAllShapes(diagramData, VisioVariables.ShowDiagram.Show);

					// build the shape connection map to be used to establish connections between shapes on the diagrams
					diagramData.ShapeConnectionsMap = new ProcessVisioShapeConnections().BuildShapeConnections(diagramData);

					// Lets make the connections 
					bool bAns = visHlp.ConnectShapes(diagramData);

					// we need to close everything
//					visHlp.VisioForceCloseAll();

				}
				else
				{
					// for testing to view all the stencils in the document
					//visHlp.ListDocumentStencils(diagramData, VisioVariables.ShowDiagram.Show);

					// buid data file from existing Visio file
					Console.WriteLine("build excel data file from a Visio file");
					Dictionary<int, ShapeInformation> shapeMap = new ProcessVisioDiagramShapes().GetAllShapesProperties(tb_buildVisioFilePath.Text.Trim(), VisioVariables.ShowDiagram.Show);
					Console.WriteLine("\n");
					foreach (var allShp in shapeMap)
					{
						int nKey = allShp.Key;
						ShapeInformation shpInf = allShp.Value;
						Console.WriteLine(string.Format("MainForm - ID:{0}; UniqueKey:{1}; Image:{2}, ConnectToID:{3}; ConnectTo:{4}; ConnectFromID:{5}; ConnectFrom:{6}", shpInf.ID, shpInf.UniqueKey, shpInf.StencilImage, shpInf.ConnectToID, shpInf.ConnectTo, shpInf.ConnectFromID, shpInf.ConnectFrom));
					}

					// we are dont so we can close the visio document(s)
					visHlp.VisioForceCloseAll();
				}
				diagramData.Reset();
			}
			catch( IOException ioe)
			{
				MessageBox.Show(string.Format("Exception::MainForm - {0}",ioe.Message), "Warning File Access Exception", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			catch (Exception ex)
			{
				MessageBox.Show(string.Format("Exception::MainForm - {0}\n{1}", ex.Message,ex.StackTrace), "Exception");
			}
		}

		private void rb_buildFromDataFile_CheckedChanged(object sender, EventArgs e)
		{
			if(rb_buildFromExcelFile.Checked)
			{
				_bBuildVisioFromExcelDataFile = true;
				tb_excelDataFile.Enabled = true;

				rb_buildExcelFileFromVisio.Checked = false;
				tb_buildExcelFileName.Enabled = false;

				btn_readExcelfile.Enabled = true;

				btn_SetExcelPath.Enabled = false;
				btn_VisioFileToRead.Enabled = false;
			}
		}

		private void rb_buildDataFileFromVisio_CheckedChanged(object sender, EventArgs e)
		{
			if (rb_buildExcelFileFromVisio.Checked)
			{
				_bBuildVisioFromExcelDataFile = false;

				tb_buildExcelFileName.Enabled = true;

				tb_excelDataFile.Enabled = false; 
				rb_buildFromExcelFile.Checked = false;

				btn_SetExcelPath.Enabled = true;
				btn_VisioFileToRead.Enabled = true;

				btn_readExcelfile.Enabled = false;
			}
		}

		private void tb_buildExcelFileName_TextChanged(object sender, EventArgs e)
		{
			// check if a valid file name
			if (IsValidFileName(tb_buildExcelFileName.Text) && IsFormValidated(_bBuildVisioFromExcelDataFile))
			{
				btn_Submit.Enabled = true;
			}
			else
			{
				btn_Submit.Enabled = false;
			}
		}

		private void btn_openExcelPath_Click(object sender, EventArgs e)
		{
			string folder = string.Empty;
			folder = FileExtension.getFolder(excelDataPath, "Select the Excel output path");
			if (string.IsNullOrEmpty(folder))
			{
				// Cancel was pressed.  filePath will be empty
				Console.WriteLine("Cancel button pressed.  No folder selected");
			}
			else
			{
				// this will contain the folder path
				excelDataPath = folder;
				tb_buildExcelPath.Text = folder;

				// check if valid is so enable submit button
				if (IsValidFileName(tb_buildExcelFileName.Text) && IsFormValidated(_bBuildVisioFromExcelDataFile))
				{
					btn_Submit.Enabled = true;
				}
				else
				{
					btn_Submit.Enabled = false;
				}
			}
		}

		private void btn_VisioFileToRead_Click(object sender, EventArgs e)
		{
			string filePath = string.Empty;
			filePath = FileExtension.getFilePath(_baseWorkingDir + "\\VisioFiles\\", "vsdx files (*.vsdx)|*.vsdx", "Select a Visio file to process into an Excel data file");
			if (string.IsNullOrEmpty(filePath))
			{
				// Cancel was pressed.  filePath will be empty
				Console.WriteLine("Cancel button pressed.  No file was selected");
			}
			else
			{
				tb_buildVisioFilePath.Text = filePath;

				// check if valid is so enable submit button
				if (IsValidFileName(tb_buildExcelFileName.Text) && IsFormValidated(_bBuildVisioFromExcelDataFile))
				{
					btn_Submit.Enabled = true;
				}
				else
				{
					btn_Submit.Enabled = false;
				}
			}
		}

		private void btn_readExcelfile_Click(object sender, EventArgs e)
		{
			string filePath = string.Empty;
			filePath = FileExtension.getFilePath(_baseWorkingDir + "\\ScriptData\\", "vsdx files (*.xls)|*.xlsx", "Select the Excel data file to build a Visio diagram");
			if (string.IsNullOrEmpty(filePath))
			{
				// Cancel was pressed.  filePath will be empty
				Console.WriteLine("Cancel button pressed.  No file was selected");
			}
			else
			{
				//Get the path of specified file
				tb_excelDataFile.Text = filePath;

				// check if valid is so enable submit button
				if (IsFormValidated(_bBuildVisioFromExcelDataFile))
				{
					btn_Submit.Enabled = true;
				}
				else
				{
					btn_Submit.Enabled = false;
				}
			}
		}


		/// <summary>
		/// validate_field
		/// validate the form fields based on the buildMode checkbox optoin
		/// </summary>
		/// <param name="buildMode">true (Build Visio diagram from Excel data file)  false (Build Excel data file from a Visio diagram)</param>
		/// <returns>true - Valid / false - Not Valid</returns>
		private bool IsFormValidated(bool buildMode)
		{
			//_bBuildVisioFromExcelDataFile
			if (buildMode == true)
			{
				// we are using the Excel data file to build the visio diagram
				if (string.IsNullOrEmpty(tb_excelDataFile.Text))
				{
					return false;
				}
			}
			else
			{
				// we are building an excel data file from a Visio diagram file
				// check required fields and paths
				if (string.IsNullOrEmpty(tb_buildExcelPath.Text))
				{
					return false;
				}
				if (string.IsNullOrEmpty(tb_buildExcelFileName.Text))
				{
					return false;
				}
				if (string.IsNullOrEmpty(tb_buildVisioFilePath.Text))
				{
					return false;
				}			
			}
			return true;
		}

		/// <summary>
		/// IsValidFileName
		/// validate the filename is a valid file name
		/// </summary>
		/// <param name="testName">filename</param>
		/// <returns>true - Valid / false - Not valid</returns>
		bool IsValidFileName(string testName)
		{
			Regex containsABadCharacter = new Regex("[" + Regex.Escape(new string(System.IO.Path.GetInvalidPathChars())) + "]");
			if (containsABadCharacter.IsMatch(testName)) 
			{ 
				return false; 
			}
			// other checks for UNC, drive-path format, etc

			return true;
		}

	}
}

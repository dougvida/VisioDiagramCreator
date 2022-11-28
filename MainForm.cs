using Microsoft.Office.Interop.Excel;
using OmnicellBlueprintingTool.Configuration;
using OmnicellBlueprintingTool.ExcelHelpers;
using OmnicellBlueprintingTool.Extensions;
using OmnicellBlueprintingTool.Models;
using OmnicellBlueprintingTool.Visio;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;

namespace OmnicellBlueprintingTool
{
	public partial class MainForm : Form
	{
		DiagramData diagramData = null;
		Boolean _bBuildVisioFromExcelDataFile = true;
		VisioHelper visioHelper = new VisioHelper();

		string ExcelDataFileName = string.Empty;  // if this is populated it will single the close operation to prompt the user to save the visio file

#if DEBUG
		private static string _baseWorkingDir = @"C:\Omnicell_Blueprinting_tool";
#else
		private static string _baseWorkingDir = Application.StartupPath;  // System.IO.Directory.GetCurrentDirectory();
#endif

		private readonly string _sJsonConfigFile = string.Format(@"{0}\OmnicellBlueprintingTool.json", _baseWorkingDir);

		private static string _scriptDataPath = _baseWorkingDir + @"\data\ScriptData\";
		private static string _visioTemplateFilesPath = _baseWorkingDir + @"\data\Templates\";
		private static string _visioStencilFilesPath = _baseWorkingDir + @"\data\Stencils\";

		private static string _excelDataPath = _baseWorkingDir + @"\ExcelData\";
		private static string _visioFilesPath = _baseWorkingDir + @"\VisioFiles\";

		public MainForm()
		{
			InitializeComponent();
			this.Text = "Omnicell Blueprinting Tool";
		}

		private void MainForm_Load(object sender, EventArgs e)
		{
			///////////////////////////////////////////////////////////////////////
			// this section is for Building the excel data file from a Visio file
			// use todays date as part of the file name
			tb_buildExcelFileName.Text = string.Empty; // String.Format("_{0}.xlsx", DateTime.Now.ToString("MMddyyyy"));
			tb_buildExcelPath.Text = _excelDataPath;
			//////////////////////////////////////////////////////////////////////

			_bBuildVisioFromExcelDataFile = true;
			rb_buildFromExcelFile.Checked = true;

			rb_buildExcelFileFromVisio.Visible = false;     // underconstruction so don't enable this option
			rb_buildExcelFileFromVisio.Enabled = false;     // underconstruction so don't enable this option

			btn_Submit.Enabled = false;

			// turn all this off.   Under construction at this time
			gb_BuildDataFile.Visible = false;
			gb_SelectOperation.Visible = false;
			btn_SetExcelPath.Visible = false;
			btn_VisioFileToRead.Visible = false;
			tb_buildExcelFileName.Visible = false;
			tb_buildExcelPath.Visible = false;
			tb_buildVisioFilePath.Visible = false;

			// debug mode lets turn on this additional stuff for testing
			rb_buildExcelFileFromVisio.Visible = true;
			rb_buildExcelFileFromVisio.Enabled = true;

			gb_BuildDataFile.Visible = true;
			gb_SelectOperation.Visible = true;
			btn_SetExcelPath.Visible = true;
			btn_VisioFileToRead.Visible = true;
			tb_buildExcelFileName.Visible = true;
			tb_buildExcelPath.Visible = true;
			tb_buildVisioFilePath.Visible = true;

			btn_SetExcelPath.Enabled = false;
			btn_VisioFileToRead.Enabled = false;

			tb_buildExcelFileName.Enabled = false;
			tb_buildExcelPath.Enabled = false;
			tb_buildVisioFilePath.Enabled = false;

			/*
			 * Lets read the application json file
			 * this file contains the variables used in the Excel Data file "Tables" sheet
			 * it allows the application to be more dynamic
			 * these entries are the dropdowns used in the Excel data file
			 * Stencils names, Colors, Arrows etc.
			 */
			if (ReadJsonFile.ReadJSONFile(_sJsonConfigFile, ref visioHelper))
			{
				string sTmp = "MainForm - ERROR\n\nOmnicellBlueprintingTool.json file not found";
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine(sTmp);

				// The user wants to exit the application. Close everything down.
				Application.Exit();

				// close the application
				// this.Close();
				return;
			}
			// parse the data file and draw the visio diagram
			diagramData = new DiagramData();
		}

		private void btn_Quit_Click(object sender, EventArgs e)
		{
			// close the Visio diagram if open
			if (!string.IsNullOrEmpty(ExcelDataFileName))
			{
				SaveVisioDiagram(string.Format(@"{0}{1}.vsdx", diagramData.VisioFilesPath, ExcelDataFileName));		
				//SaveVisioDiagram(string.Format(@"{0}{1}.vsdx", _visioFilesPath, ExcelDataFileName));

				visioHelper.VisioForceCloseAll();
			}
			ExcelDataFileName = string.Empty;

			// The user wants to exit the application. Close everything down.
			Application.Exit();

			// close the application
			//this.Close();
		}

		private void btn_Submit_Click(object sender, EventArgs e)
		{
			// close the Visio diagram if open
			// after the run will keep it open to allow the user to work on the diagram before saving
			// if the quit button is pressed the close Visio document will be call
			if (!string.IsNullOrEmpty(ExcelDataFileName))
			{
				SaveVisioDiagram(string.Format(@"{0}{1}.vsdx", diagramData.VisioFilesPath, ExcelDataFileName));
				//SaveVisioDiagram(string.Format(@"{0}{1}.vsdx", _visioFilesPath, ExcelDataFileName));
				visioHelper.VisioForceCloseAll();

				ExcelDataFileName = string.Empty;
			}

			// parse the data file and draw the visio diagram
			diagramData = new DiagramData();

			diagramData.BaseWorkingDir = _baseWorkingDir;
			diagramData.ScriptDataPath = _scriptDataPath;
			diagramData.ExcelDataPath = _excelDataPath;
			if (string.IsNullOrEmpty(diagramData.VisioFilesPath))
			{	
				// value has not been set so use the defaiult
				diagramData.VisioFilesPath = _visioFilesPath;
			}

			string sTmp = string.Empty;

			try
			{
				if (_bBuildVisioFromExcelDataFile)
				{
					// Set cursor as hourglass
					Cursor.Current = Cursors.WaitCursor;

					// build visio file form data file
					ConsoleOut.writeLine(String.Format("MainForm - Build Visio file from an excel data file:{0}", tb_excelDataFile.Text));
					ExcelDataFileName = FileExtension.GetFileNameOnly(tb_excelDataFile.Text);
					diagramData = new ProcessExcelDataFile().parseExcelFile(tb_excelDataFile.Text.Trim(), diagramData, visioHelper);
					if (diagramData == null)
					{
						sTmp = "MainForm - ERROR\n\nReturn from ProcessExcelDataFile returned null\nNo shapes will be drawn";
						//MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						Console.WriteLine(sTmp);
						visioHelper.VisioForceCloseAll();

						// The user wants to exit the application. Close everything down.
						Application.Exit();

						//this.Close();
					}

					// Set cursor as default arrow
					Cursor.Current = Cursors.Default;

					// for testing to view all the stencils in the document
					// visioHelper.ListDocumentStencils(diagramData, VisioVariables.ShowDiagram.Show);

					if (diagramData != null)
					{
						if (!visioHelper.DrawAllShapes(diagramData, VisioVariables.ShowDiagram.Show))
						{
							// build the shape connection map to be used to establish connections between shapes on the diagrams
							diagramData.ShapeConnectionsMap = new ProcessVisioShapeConnections().BuildShapeConnections(diagramData);

							// Lets make the connections 
							bool bAns = visioHelper.ConnectShapes(diagramData);
							if (bAns)
							{
								// there was an exception so we need to get out
								visioHelper.VisioForceCloseAll(true);

								// The user wants to exit the application. Close everything down.
								Application.Exit();

								// close the application
								//this.Close();
								return;
							}


							// set focus to first page
							int maxPages = visioHelper.GetNumberOfPages();
							visioHelper.SetActivePage(1);
						}
					}
				}
				else
				{
					diagramData.VisioTemplateFilePath = _visioTemplateFilesPath;
					diagramData.VisioStencilFilePaths.Add(_visioStencilFilesPath);

					// sTmp = string.Format("This process will build an Excel data file from a Visio file.\n\n" +
					//	"Note: This process may take a few minutes so please be patient.\n\n" +
					//	"Use the excel data file with this tool to rebuild the Visio diagram.\nWhen making modifications / additions make it to the Excel Data file.\n\n" +
					//	"You may need to modify stencils positions as well as connections");
					//MessageBox.Show(this, sTmp, "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);

					// Set cursor as hourglass
					Cursor.Current = Cursors.WaitCursor;

					// for testing to view all the stencils in the document
					//visioHelper.ListDocumentStencils(diagramData, VisioVariables.ShowDiagram.Show);

					// buid data file from existing Visio file
					ConsoleOut.writeLine("build excel data file from a Visio file");
					Dictionary<int, ShapeInformation> shapesMap = new ProcessVisioDiagramShapes().GetAllShapesProperties(visioHelper, tb_buildVisioFilePath.Text.Trim(), VisioVariables.ShowDiagram.Show);
					if (shapesMap == null)
					{
						sTmp = "MainForm\n\nNo shapes found on the Visio Diagram";
						MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					else
					{
						foreach (var allShp in shapesMap)
						{
							int nKey = allShp.Key;
							ShapeInformation shpInf = allShp.Value;
							ConsoleOut.writeLine(string.Format("MainForm - ID:{0}; UniqueKey:{1}; Image:{2}, ConnectToID:{3}; ConnectTo:{4}; ToLabel:{5}; ConnectFromID:{6}; ConnectFrom:{7}; FromLabel:{8}", shpInf.ID, shpInf.UniqueKey, shpInf.StencilImage, shpInf.ConnectToID, shpInf.ConnectTo, shpInf.ToLineLabel, shpInf.ConnectFromID, shpInf.ConnectFrom, shpInf.FromLineLabel));
						}
						CreateExcelDataFile createExcelDataFile = new CreateExcelDataFile();
						string sPath = string.Format(@"{0}{1}", tb_buildExcelPath.Text.Trim(), tb_buildExcelFileName.Text.Trim());
						if (createExcelDataFile.PopulateExcelDataFile(diagramData, visioHelper, shapesMap, sPath))
						{
							sTmp = String.Format("MainForm - Error\n\nFailed to create excel data file:{0}", sPath);
							MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
				// Set cursor as default arrow
				Cursor.Current = Cursors.Default;
				if (diagramData != null)
				{
					diagramData.Reset();
					//diagramData = null;
				}
			}
			catch (IOException ioe)
			{
				sTmp = string.Format("MainForm - IOEException\n{0}\n", ioe.Message, ioe.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
				sTmp = string.Format("MainForm - Exception\n{0}\n{1}", ex.Message, ex.StackTrace);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Set cursor as default arrow
				Cursor.Current = Cursors.Default;
			}
		}
		private void rb_buildFromDataFile_CheckedChanged(object sender, EventArgs e)
		{
			if (rb_buildFromExcelFile.Checked)
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
			folder = FileExtension.getFolder(_excelDataPath, "Select the Excel output path");
			if (string.IsNullOrEmpty(folder))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No folder selected");
			}
			else
			{
				// this will contain the folder path
				_excelDataPath = folder;
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
			if (string.IsNullOrEmpty(diagramData.VisioFilesPath))
			{
				filePath = FileExtension.getFilePath(_visioFilesPath, "Visio files (*.vsdx)|*.vsdx", "Select a Visio file to process into an Excel data file");
			}
			else
			{
				filePath = FileExtension.getFilePath(diagramData.VisioFilesPath, "Visio files (*.vsdx)|*.vsdx", "Select a Visio file to process into an Excel data file");
			}
			if (string.IsNullOrEmpty(filePath))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No file was selected");
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

				// need to break apart the file remove any date format and append current data
				string sName = Path.GetFileNameWithoutExtension(filePath);
				if (sName.Length > 8)
				{
					string sTmp = sName.Substring(sName.Length - 8);
					if (Regex.IsMatch(sTmp, @"^\d+$"))
					{
						sName = sName.Substring(0, sName.Length - 8);
					}
				}
				tb_buildExcelFileName.Text = String.Format("{0}_ExcelData_{1}.xlsx", sName, DateTime.Now.ToString("MMddyyyy"));

				// keep the folder that was selected for next time if needed
				diagramData.VisioFilesPath = Path.GetDirectoryName(filePath);
			}
		}
		private void btn_readExcelfile_Click(object sender, EventArgs e)
		{
			string filePath = string.Empty;
			filePath = FileExtension.getFilePath(_scriptDataPath, "Excel(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm;", "Select the Excel data file to build a Visio diagram");
			if (string.IsNullOrEmpty(filePath))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No file was selected");
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
		/// SaveVisioDiagram
		/// Save the visio diagram
		/// first prompt the user asking if they want to save the diagram
		/// return prompt answer
		/// </summary>
		/// <param name="fileNamePath"></param>
		/// <returns>bool - true (yes); false (no)</returns>
		private bool SaveVisioDiagram(string fileNamePath)
		{
			bool bSave = false;  // dont save

			// ask user if want to save diagram
			if (MessageBox.Show(string.Format("Do you want to save the Visio document ({0}) ?", fileNamePath), "Save Visio Diagram", MessageBoxButtons.YesNo) == DialogResult.Yes)
			{
				bSave = true;
			}
			// must always call SaveDocuments
			if (visioHelper.SaveDocument(fileNamePath, bSave))
			{
				string sTmp = string.Format("WARNING:: failed to save Visio diagram to the file:'{0}'", fileNamePath);
				MessageBox.Show(sTmp, "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return bSave;
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

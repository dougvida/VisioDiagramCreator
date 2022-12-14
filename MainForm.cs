using Microsoft.Office.Interop.Excel;
using OIS.Models;
using OmnicellBlueprintingTool.Configuration;
using OmnicellBlueprintingTool.ExcelHelpers;
using OmnicellBlueprintingTool.Extensions;
using OmnicellBlueprintingTool.Models;
using OmnicellBlueprintingTool.Visio;
using OmnicellOISNodes.Processing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;

namespace OmnicellBlueprintingTool
{
	public partial class MainForm : Form
	{
		DiagramData diagramData = null;
		VisioHelper visioHelper = new VisioHelper();

		string ExcelDataFileName = string.Empty;  // if this is populated it will single the close operation to prompt the user to save the visio file

#if DEBUG
		private static string _baseWorkingDir = @"C:\Omnicell_Blueprinting_tool";
#else
		private static string _baseWorkingDir = Application.StartupPath;  // System.IO.Directory.GetCurrentDirectory();
#endif

		private static string _appConfigurationJsonFile = string.Format(@"{0}\{1}", _baseWorkingDir, VisioVariables.DefaultAppConfigJsonFile);

		private static string _excelScriptDataPath = _baseWorkingDir + @"\data\ScriptData\";		// folder for the Excel data script files (read)
		private static string _excelFileDataPath = _baseWorkingDir + @"\ExcelData\";					// folder to save the Excel data script files (write from Visio)

		private static string _visioTemplateFilesPath = _baseWorkingDir + @"\data\Templates\";		// folder where Visio Templates are placed
		private static string _visioStencilFilesPath = _baseWorkingDir + @"\data\Stencils\";		// folder where Visio Stencils are placed
		private static string _visioFilesPath = _baseWorkingDir + @"\VisioFiles\";						// folder to save the Visio Diagram files

		private static string _oisFilesPath = _baseWorkingDir + @"\OISFiles\";							// folder where OIS setup CSV files exist

		enum BuildType
		{
			BuildVisioFromExcel,
			BuildExcelFromVisio,
			BuildExcelFromOIS
		}
		
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
			t2_tb_BuildExcelFileName.Text = string.Empty; // String.Format("_{0}.xlsx", DateTime.Now.ToString("MMddyyyy"));
			t2_tb_BuildExcelPath.Text = _excelFileDataPath;

			t3_tb_BuildExcelPath.Text = _excelFileDataPath;

			// create this folder if not already
			if (!Directory.Exists(_excelFileDataPath))
			{
				Directory.CreateDirectory(_excelFileDataPath);
			}

			// create this folder if not already
			if (!Directory.Exists(_oisFilesPath))
			{
				Directory.CreateDirectory(_oisFilesPath);
			}

			//////////////////////////////////////////////////////////////////////

			// tab 1
			t1_btn_Submit.Enabled = false;

			// tab 2
			t2_btn_Submit.Enabled = false;

			// tab 3 OIS
			t3_btn_Submit.Enabled = false;

			/*
			 * Lets read the application json file
			 * this file contains the variables used in the Excel Data file "Tables" sheet
			 * it allows the application to be more dynamic
			 * these entries are the dropdowns used in the Excel data file
			 * Stencils names, Colors, Arrows etc.
			 */
			if (ReadJsonFile.ReadJSONFile(_appConfigurationJsonFile, ref visioHelper))
			{
				// The user wants to exit the application. Close everything down.
				Application.Exit();
				return;
			}
			
			// parse the data file and draw the visio diagram
			diagramData = new DiagramData();
			diagramData.BaseWorkingDir = _baseWorkingDir;

			diagramData.ExcelScriptDataPath = _excelScriptDataPath;	// used for Excel Script data files Open dialog
			diagramData.ExcelDataFileSavePath = _excelFileDataPath;  // used for saving Excel script data files created from a Visio Diagram

			diagramData.VisioFilesPath = _visioFilesPath;
			diagramData.oisFilesPath = _oisFilesPath;
		}

		private void timeToWork(BuildType buildType)
		{
			// close the Visio diagram if open
			// after the run will keep it open to allow the user to work on the diagram before saving
			// if the quit button is pressed the close Visio document will be call
			if (!string.IsNullOrEmpty(ExcelDataFileName))
			{
				//SaveVisioDiagram(string.Format(@"{0}{1}.vsdx", diagramData.VisioFilesPath, ExcelDataFileName));
				string sTmp1 = Path.Combine(@diagramData.VisioFilesPath, @ExcelDataFileName) + ".vsdx";
				SaveVisioDiagram(sTmp1);
				visioHelper.VisioForceCloseAll();

				ExcelDataFileName = string.Empty;
			}

			string sTmp = string.Empty;

			try
			{
				switch (buildType)
				{
#region BuildVisioFromExcel
					case BuildType.BuildVisioFromExcel:
						// Set cursor as hourglass
						Cursor.Current = Cursors.WaitCursor;

						if (diagramData.visioStencilFilesPath != null)
						{
							diagramData.VisioStencilFilePaths.Clear();
						}

						// build visio file form data file
						ConsoleOut.writeLine(String.Format("MainForm - Build Visio file from an excel data file:{0}", t1_tb_ExcelDataFile.Text));
						ExcelDataFileName = FileExtension.GetFileNameOnly(t1_tb_ExcelDataFile.Text);
						diagramData = new ProcessExcelDataFile().parseExcelFile(t1_tb_ExcelDataFile.Text.Trim(), diagramData, ref visioHelper);
						if (diagramData == null)
						{
							sTmp = "MainForm - ERROR\n\nReturn from ProcessExcelDataFile returned null\nNo shapes will be drawn";
							//MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
							ConsoleOut.writeLine(sTmp);
							visioHelper.VisioForceCloseAll();

							// The user wants to exit the application. Close everything down.
							Application.Exit();
						}

						// Set cursor as default arrow
						Cursor.Current = Cursors.Default;

						// for testing to view all the stencilsList in the document
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

									return;
								}
								// set focus to first page
								int maxPages = visioHelper.GetNumberOfVisioPages();
								visioHelper.SetActivePage(1);
							}
						}
						break;
#endregion
#region BuildExcelFromVisio
					case BuildType.BuildExcelFromVisio:

						diagramData.VisioTemplateFilePath = _visioTemplateFilesPath;

						diagramData.VisioStencilFilePaths.Clear();
						diagramData.VisioStencilFilePaths.Add(_visioStencilFilesPath);

						// sTmp = string.Format("This process will build an Excel data file from a Visio file.\n\n" +
						//	"Note: This process may take a few minutes so please be patient.\n\n" +
						//	"Use the excel data file with this tool to rebuild the Visio diagram.\nWhen making modifications / additions make it to the Excel Data file.\n\n" +
						//	"You may need to modify stencilsList positions as well as connections");
						//MessageBox.Show(this, sTmp, "INFORMATION", MessageBoxButtons.OK, MessageBoxIcon.Information);

						// Set cursor as hourglass
						Cursor.Current = Cursors.WaitCursor;

						// for testing to view all the stencilsList in the document
						//visioHelper.ListDocumentStencils(diagramData, VisioVariables.ShowDiagram.Show);

						// buid data file from existing Visio file
						ConsoleOut.writeLine("build excel data file from a Visio file");
						Dictionary<string, ShapeInformation> shapesMap = new ProcessVisioDiagramShapes().GetAllShapesProperties(visioHelper, t2_tb_BuildVisioFilePath.Text.Trim(), VisioVariables.ShowDiagram.Show);
						if (shapesMap == null)
						{
							sTmp = "MainForm\n\nNo shapes found on the Visio Diagram";
							MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						else
						{
							// list all the shapes and connections
							//foreach (var allShp in shapesMap)
							//{
							//	string sKey = allShp.Key;
							//	ShapeInformation shpInf = allShp.Value;
							//	ConsoleOut.writeLine(string.Format("MainForm - GUID:{0}; UniqueKey:{1}; Image:{2}, ConnectToID:{3}; ConnectTo:{4}; ToLabel:{5}; ConnectFromID:{6}; ConnectFrom:{7}; FromLabel:{8}",
							//	shpInf.GUID.PadRight(40), shpInf.UniqueKey.PadRight(25), shpInf.StencilImage.PadRight(25), shpInf.ConnectToID.ToString().PadRight(10), shpInf.ConnectTo.PadRight(20), shpInf.ToLineLabel, shpInf.ConnectFromID, shpInf.ConnectFrom, shpInf.FromLineLabel));
							//}
							CreateExcelDataFile createExcelDataFile = new CreateExcelDataFile();
							string sPath = Path.Combine(@t2_tb_BuildExcelPath.Text.Trim(), @t2_tb_BuildExcelFileName.Text.Trim());
							if (createExcelDataFile.PopulateExcelDataFile(diagramData, visioHelper, shapesMap, sPath))
							{
								sTmp = String.Format("MainForm - Error\n\nFailed to create excel data file:\n'{0}'\n\nVerify the file namd and path is correct", sPath);
								MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
							}
						}
						break;
#endregion
#region BuildExcelFromOIS
					case BuildType.BuildExcelFromOIS:

						//diagramData.VisioTemplateFilePath = _visioTemplateFilesPath;
						diagramData.VisioStencilFilePaths.Clear();
						diagramData.VisioStencilFilePaths.Add(_visioStencilFilesPath);

						Dictionary<string, ShapeInformation> oisShapeInfoMap = null;

						// Set cursor as hourglass
						Cursor.Current = Cursors.WaitCursor;

						// buid data file from existing Visio file
						ConsoleOut.writeLine("Build Excel data file from OIS setup (CSV) file");

						Dictionary<string, List<OISSetupData>> oisDataMap = ParseOISSetup.ParseOISSetupFile(t3_tb_BuildOISFilePath.Text.Trim());
						if (oisDataMap == null)
						{
							sTmp = string.Format("MainForm\n\nNo OIS entries found in this file:\b'{0}'", t3_tb_BuildOISFilePath.Text.Trim());
							MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
						else
						{
							
							// translate the data from OIS Data to ShapeInformation map
							oisShapeInfoMap = TranslateData.ConvertData(oisDataMap);

							CreateExcelDataFile createExcelDataFile = new CreateExcelDataFile();
							string sPath = Path.Combine(@t3_tb_BuildExcelPath.Text.Trim(), @t3_tb_BuildExcelFileName.Text.Trim());
							if (createExcelDataFile.PopulateExcelDataFile(diagramData, visioHelper, oisShapeInfoMap, sPath))
							{
								sTmp = String.Format("MainForm - Error\n\nFailed to create excel data file:\n'{0}'\n\nVerify the file namd and path is correct", sPath);
								MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
							}
						}
						break;
#endregion
					default:
						break;
				}
				// Set cursor as default arrow
				Cursor.Current = Cursors.Default;
				if (diagramData != null)
				{
					diagramData.Reset();
				}
			}
			catch (IOException ioe)
			{
				sTmp = string.Format("MainForm - IOException\n{0}\n", ioe.Message, ioe.StackTrace);
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

		/// <summary>
		/// timeToQuit
		/// this function will prompt the user to save any file that has been created
		/// then exit the application
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void timeToQuit()
		{
			// close the Visio diagram if open
			if (!string.IsNullOrEmpty(ExcelDataFileName))
			{
				//SaveVisioDiagram(string.Format(@"{0}{1}.vsdx", diagramData.VisioFilesPath, ExcelDataFileName));
				string sTmp = Path.Combine(@diagramData.VisioFilesPath, @ExcelDataFileName) + ".vsdx";
				SaveVisioDiagram(sTmp);
				visioHelper.VisioForceCloseAll();
			}
			ExcelDataFileName = string.Empty;

			// The user wants to exit the application. Close everything down.
			Application.Exit();
		}

		/// <summary>
		/// btn_Quit_Click
		/// this is called when the user clicks on the Quite button
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void btn_Quit_Click(object sender, EventArgs e)
		{
			timeToQuit();
		}

		/// <summary>
		/// t1_btn_Submit_Click
		/// Tab control 1 submit button   
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t1_btn_Submit_Click(object sender, EventArgs e)
		{
			timeToWork(BuildType.BuildVisioFromExcel);
		}

		/// <summary>
		/// t2_btn_Submit_Click
		/// Tab control 2 submit button   
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t2_btn_Submit_Click(object sender, EventArgs e)
		{
			timeToWork(BuildType.BuildExcelFromVisio);
		}

		/// <summary>
		/// t3_btn_Submit_Click
		/// Tab control 3 submit button   
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t3_btn_Submit_Click(object sender, EventArgs e)
		{
			timeToWork(BuildType.BuildExcelFromOIS);
		}

		/// <summary>
		/// btn_readExcelfile_Click
		/// this function will prompt the user with a file dialog to select the 
		/// Excel file to open for read
		/// the variable "tb_excelDataFile.Text" will be set to the file name & path selected by the user
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t1_btn_readExcelfile_Click(object sender, EventArgs e)
		{
			string filePath = string.Empty;
			if (string.IsNullOrEmpty(diagramData.ExcelScriptDataPath))
			{
				diagramData.ExcelScriptDataPath = _excelFileDataPath;
			}

			filePath = FileExtension.getFilePath(diagramData.ExcelScriptDataPath, "Excel(*.xls;*.xlsx;*.xlsm)|*.xls;*.xlsx;*.xlsm;", "Select the Excel data file to uild a Visio diagram");
			if (string.IsNullOrEmpty(filePath))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No file was selected");
			}
			else
			{
				// lets save the selected Excel Data Path
				diagramData.ExcelScriptDataPath = Path.GetDirectoryName(filePath);

				//Get the path of specified file
				t1_tb_ExcelDataFile.Text = filePath;

				// check if valid is so enable submit button
				if (IsFormValidated(BuildType.BuildVisioFromExcel))
				{
					t1_btn_Submit.Enabled = true;
				}
				else
				{
					t1_btn_Submit.Enabled = false;
				}
			}
		}

		private void t2_tb_buildExcelFileName_TextChanged(object sender, EventArgs e)
		{
			// check if a valid file name
			//if (IsValidFileName(t2_tb_BuildExcelFileName.Text) && IsFormValidated(BuildType.BuildVisioFromExcel))
			if (IsValidFileName(t2_tb_BuildExcelFileName.Text) && IsFormValidated(BuildType.BuildExcelFromVisio))
			{
				t2_btn_Submit.Enabled = true;
			}
			else
			{
				t2_btn_Submit.Enabled = false;
			}
		}

		/// <summary>
		/// btn_openExcelPath_Clock
		/// this function will open a folder dialog 
		/// allowing the user to select a folder to save the newly created Excel file
		/// both variables "_excelFileDataPath" and "tb_buildExcelPath.Text" will be set to the folder path
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t2_btn_openExcelPath_Click(object sender, EventArgs e)
		{
			string folder = string.Empty;
			if (string.IsNullOrEmpty(diagramData.ExcelDataFileSavePath))
			{
				diagramData.ExcelDataFileSavePath = _excelFileDataPath;
			}

			//folder = FileExtension.getFolder(_excelFileDataPath, "Select the Excel output path");
			folder = FileExtension.getFolder(diagramData.ExcelDataFileSavePath, "Select the Excel output path");
			if (string.IsNullOrEmpty(folder))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No folder selected");
			}
			else
			{
				// this will contain the folder path
				diagramData.ExcelDataFileSavePath = folder;
				t2_tb_BuildExcelPath.Text = folder;

				// check if valid is so enable submit button
				if (IsValidFileName(t2_tb_BuildExcelPath.Text) && IsFormValidated(BuildType.BuildExcelFromVisio))
				{
					t2_btn_Submit.Enabled = true;
				}
				else
				{
					t2_btn_Submit.Enabled = false;
				}
			}
		}

		/// <summary>
		/// btn_VisioFileToRead_Click
		/// this function will open a file dialog allowing the user to select the Visio file to open and process
		/// the variable "tb_buildVisioFilePath.Text" will be populated with the path & file name selected by the user
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t2_btn_VisioFileToRead_Click(object sender, EventArgs e)
		{
			string filePath = string.Empty;
			if (string.IsNullOrEmpty(diagramData.VisioFilesPath))
			{
				diagramData.VisioFilesPath = _visioFilesPath;
			}

			filePath = FileExtension.getFilePath(diagramData.VisioFilesPath, "Visio files (*.vsdx)|*.vsdx", "Select a Visio file to process into an Excel data file");
			if (string.IsNullOrEmpty(filePath))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No file was selected");
			}
			else
			{
				t2_tb_BuildVisioFilePath.Text = filePath;

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
				t2_tb_BuildExcelFileName.Text = String.Format("{0}_ExcelData_{1}.xlsx", sName, DateTime.Now.ToString("MMddyyyy"));

				// check if valid is so enable submit button
				if (IsValidFileName(t2_tb_BuildExcelFileName.Text) && IsFormValidated(BuildType.BuildExcelFromVisio))
				{
					t2_btn_Submit.Enabled = true;
				}
				else
				{
					t2_btn_Submit.Enabled = false;
				}

				// keep the folder that was selected for next time if needed
				diagramData.VisioFilesPath = Path.GetDirectoryName(filePath);
			}
		}

		private void t3_tb_buildExcelFileName_TextChanged(object sender, EventArgs e)
		{
			// check if a valid file name
			if (IsValidFileName(t3_tb_BuildExcelFileName.Text) && IsFormValidated(BuildType.BuildExcelFromOIS))
			{
				t3_btn_Submit.Enabled = true;
			}
			else
			{
				t3_btn_Submit.Enabled = false;
			}
		}

		/// <summary>
		/// t3_btn_openExcelPath_Clock
		/// this function will open a folder dialog 
		/// allowing the user to select a folder to save the newly created Excel file
		/// both variables "_excelFileDataPath" and "t3_tb_buildExcelPath.Text" will be set to the folder path
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t3_btn_openExcelPath_Click(object sender, EventArgs e)
		{
			string folder = string.Empty;
			if (string.IsNullOrEmpty(diagramData.ExcelDataFileSavePath))
			{
				diagramData.ExcelDataFileSavePath = _excelFileDataPath;
			}

			//folder = FileExtension.getFolder(_excelFileDataPath, "Select the Excel output path");
			folder = FileExtension.getFolder(diagramData.ExcelDataFileSavePath, "Select the Excel output path");
			if (string.IsNullOrEmpty(folder))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No folder selected");
			}
			else
			{
				// this will contain the folder path
				diagramData.ExcelDataFileSavePath = folder;
				t3_tb_BuildExcelPath.Text = folder;

				// check if valid is so enable submit button
				if (IsValidFileName(t3_tb_BuildExcelPath.Text) && IsFormValidated(BuildType.BuildExcelFromOIS))
				{
					t3_btn_Submit.Enabled = true;
				}
				else
				{
					t3_btn_Submit.Enabled = false;
				}
			}
		}

		/// <summary>
		/// btn_VisioFileToRead_Click
		/// this function will open a file dialog allowing the user to select the Visio file to open and process
		/// the variable "tb_buildVisioFilePath.Text" will be populated with the path & file name selected by the user
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void t3_btn_OISFileToRead_Click(object sender, EventArgs e)
		{
			string filePath = string.Empty;
			if (string.IsNullOrEmpty(diagramData.oisFilesPath))
			{
				diagramData.oisFilesPath = _oisFilesPath;
			}

			filePath = FileExtension.getFilePath(diagramData.oisFilesPath, "Visio files (*.csv)|*.csv", "Select a CSV file to process into an Excel data file");
			if (string.IsNullOrEmpty(filePath))
			{
				// Cancel was pressed.  filePath will be empty
				ConsoleOut.writeLine("Cancel button pressed.  No file was selected");
			}
			else
			{
				t3_tb_BuildOISFilePath.Text = filePath;

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
				t3_tb_BuildExcelFileName.Text = String.Format("{0}_ExcelData_{1}.xlsx", sName, DateTime.Now.ToString("MMddyyyy"));

				// check if valid is so enable submit button
				if (IsValidFileName(t3_tb_BuildExcelFileName.Text) && IsFormValidated(BuildType.BuildExcelFromOIS))
				{
					t3_btn_Submit.Enabled = true;
				}
				else
				{
					t3_btn_Submit.Enabled = false;
				}

				// keep the folder that was selected for next time if needed
				diagramData.oisFilesPath = Path.GetDirectoryName(filePath);
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
				MessageBox.Show(sTmp, "WARNING", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			return bSave;
		}

		/// <summary>
		/// validate_field
		/// validate the form fields based on the buildMode checkbox optoin
		/// </summary>
		/// <param name="buildMode">true (Build Visio diagram from Excel data file)  false (Build Excel data file from a Visio diagram)</param>
		/// <returns>true - Valid / false - Not Valid</returns>
		private bool IsFormValidated(BuildType buildType)
		{
			switch (buildType)
			{
				case BuildType.BuildVisioFromExcel:
					// we are using the Excel data file to build the visio diagram
					if (string.IsNullOrEmpty(t1_tb_ExcelDataFile.Text))
					{
						return false;
					}
					break;

				case BuildType.BuildExcelFromVisio:
					// we are building an excel data file from a Visio diagram file
					// check required fields and paths
					if (string.IsNullOrEmpty(t2_tb_BuildExcelPath.Text))
					{
						return false;
					}
					if (string.IsNullOrEmpty(t2_tb_BuildExcelFileName.Text))
					{
						return false;
					}
					if (string.IsNullOrEmpty(t2_tb_BuildVisioFilePath.Text))
					{
						return false;
					}
					break;

				case BuildType.BuildExcelFromOIS:
					// we are building an excel data file from a Visio diagram file
					// check required fields and paths
					if (string.IsNullOrEmpty(t3_tb_BuildExcelPath.Text))
					{
						return false;
					}
					if (string.IsNullOrEmpty(t3_tb_BuildExcelFileName.Text))
					{
						return false;
					}
					if (string.IsNullOrEmpty(t3_tb_BuildOISFilePath.Text))
					{
						return false;
					}
					break;

				default:
					break;
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
			if (string.IsNullOrEmpty(testName)) 
			{
				return false;
			}
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

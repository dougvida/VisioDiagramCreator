using System;
using System.IO;
using System.Windows.Forms;

namespace OmnicellBlueprintingTool.Extensions
{

	public static class FileExtension
	{
		///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
		/// <summary>
		/// getFolder
		/// Open the FolderBrowserDialog allow the user to select the folder to use
		/// the defaultFolder argument is used as the defualt folder
		/// </summary>
		/// <param name="defaultFolder">default folder</param>
		/// <param name="desc">description to display in the dialog</param>
		/// <returns>Folder or null/empty (No selection)</returns>
		///
		public static string getFolder(string defaultFolder, string desc)
		{
			string folder = string.Empty;

			// Display the FolderBrowserDialog
			using (var dialog = new FolderBrowserDialog())
			{
				dialog.SelectedPath = defaultFolder;
				dialog.Description = desc;

				//dialog.RootFolder = Environment.SpecialFolder.MyComputer;
				dialog.ShowNewFolderButton = false;

				if (dialog.ShowDialog() == DialogResult.OK)
				{
					// this will contain the folder path
					folder = dialog.SelectedPath.Trim();
				}
			}
			return folder;
		}

		/// <summary>
		/// getFilePath
		/// Open the OpenFileDialog allowing the user to select a file
		/// </summary>
		/// <param name="defaultFolder">default folder</param>
		/// <param name="filter">Filter of extensions to show in Dialog</param>
		/// <param name="desc">description to display in the dialog</param>
		/// <returns>Full file name with Path or null/empty (No selection)</returns>
		///
		public static string getFilePath(string defaultFolder, string filter, string desc)
		{
			string filePath = string.Empty;

			// display normal OpenFileDialog
			using (OpenFileDialog dialog = new OpenFileDialog())
			{
				dialog.Title = desc;
				dialog.InitialDirectory = defaultFolder;
				dialog.Filter = filter;
				if (dialog.ShowDialog() == DialogResult.OK)
				{
					//Get the path of specified file
					filePath = dialog.FileName.Trim();
				}
			}
			return filePath;
		}

		/// <summary>
		/// FileExists
		/// does the file exist
		/// </summary>
		/// <param name="fileNamePath"></param>
		/// <returns>true Yes, false no</returns>
		public static bool FileExists(string fileNamePath)
		{
			return File.Exists(fileNamePath);
		}
	}
}

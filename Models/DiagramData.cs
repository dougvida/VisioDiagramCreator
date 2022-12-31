using System.Collections.Generic;
using OmnicellBlueprintingTool.Visio;

namespace OmnicellBlueprintingTool.Models
{
	public class DiagramData
	{
		public string BaseWorkingDir = string.Empty;
		public string ExcelScriptDataPath = string.Empty;

		public string visioTemplateFilesPath = string.Empty;
		public string visioStencilFilesPath = string.Empty;
		public string VisioFilesPath = string.Empty;

		public string ExcelDataFileSavePath = string.Empty;
		public string ExcelDataFileName = string.Empty;

		public string oisFilesPath = string.Empty;

		public VisioVariables.VisioPageOrientation VisioPageOrientation { get; set; } = VisioVariables.VisioPageOrientation.Portrait;
		public VisioVariables.VisioPageSize VisioPageSize { get; set; } = VisioVariables.VisioPageSize.Letter;
		public bool AutoSizeVisioPages { get; set; } = false;	// don't autosize the Visio pages

		// This map contains all the shapes from the Excel Data file
		// it will be used to build up the connections with other shapes if needed
		public Dictionary<string, Device> AllShapesMap = new Dictionary<string, Device>();

		// This map will contain all the connections to and from shaps based on the Excel Data 
		public Dictionary<int, ShapeConnection> ShapeConnectionsMap = new Dictionary<int, ShapeConnection>();

		public string VisioTemplateFilePath { get; set; }

		public List<string> VisioStencilFilePaths = new List<string>();

		public int MaxVisioPages { get; set; } = 1;     // default is 1 page visio document.

		public List<Device> Devices { get; set; }

		public void Reset()
		{
			MaxVisioPages = 1;
			AutoSizeVisioPages = false;

			visioTemplateFilesPath = string.Empty;
			visioStencilFilesPath = string.Empty;
			
			/// TODO-1 removed for testing.  I don't think we need to clear this variable
			//ExcelDataFileSavePath = string.Empty;

			if (AllShapesMap != null)
			{
				AllShapesMap.Clear();
			}
			if (ShapeConnectionsMap != null)
			{
				ShapeConnectionsMap.Clear();
			}
			if (Devices != null)
			{
				Devices.Clear();
			}
		}
	}
}

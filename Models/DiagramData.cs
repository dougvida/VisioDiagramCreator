using System.Collections.Generic;
using VisioDiagramCreator.Visio;

namespace VisioDiagramCreator.Models
{
	public class DiagramData
	{
		public string VisioPageOrientation { get; set; } = "Portrait";
		public string VisioPageSize { get; set; } = "Letter";

		// This map contains all the shapes from the Excel Data file
		// it will be used to build up the connections with other shapes if needed
		public Dictionary<string, Device> AllShapesMap = new Dictionary<string, Device>();

		// This map will contain all the connections to and from shaps based on the Excel Data 
		public Dictionary<int, ShapeConnection> ShapeConnectionsMap = new Dictionary<int, ShapeConnection>();

		public string visioTemplateFilePath { get; set; }
		public List<string> visioStencilFilePaths { get; set; }

		public int MaxVisioPages { get; set; } = 1;     // default is 1 page visio document

		public List<Device> Devices { get; set; }

		public void Reset()
		{
			MaxVisioPages = 1;

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

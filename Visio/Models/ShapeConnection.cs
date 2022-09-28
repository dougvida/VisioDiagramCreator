using VisioDiagramCreator.Models;
using Visio1 = Microsoft.Office.Interop.Visio;

namespace VisioDiagramCreator.Visio
{
	public class ShapeConnection
	{
		public string UniqueFromKey { get; set; }
		public Visio1.Shape ShpFromObj { get; set; }
		public string UniqueToKey { get; set; }
		public Visio1.Shape ShpToObj { get; set; }

		public Device device { get; set; }

		// this section is specific arrow settings for establishing shape connections
		public string LineLabel { get; set; } = string.Empty;
		public double LinePattern { get; set; } = VisioVariables.LINE_PATTERN_SOLID;  // solid line
		public string LineColor { get; set; } = "BLACK";
		public string ArrowType { get; set; } = VisioVariables.sARROW_NONE;

		public Visio1.Shape ShpObj { get; set; }     // this shape object
	}
}

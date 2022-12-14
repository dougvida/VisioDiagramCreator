using OmnicellBlueprintingTool.Models;
using System;
using System.Collections.Generic;
using System.Net.Http.Headers;

namespace OmnicellBlueprintingTool.Visio
{
	public class VisioVariables
	{
		public static string DefaultBlueprintingTemplateFile = "OC_BlueprintingTemplate.vstx";
		public static string DefaultBlueprintingStencilFile = "OC_BlueprintingStencils.vssx";
		public static string DefaultAppConfigJsonFile = "OmnicellBlueprintingTool.json";

		public VisioVariables()
		{
		}

		public const double HEIGHT = 0.25;

		// connector ends
		public const double BEGIN_ARROW = 4;      // Filled arrow
		public const double END_ARROW = 4;        // Filled arrow
		public const double ARROW_NONE = 0;       // None
		public const string sARROW_NONE = "None";
		public const string sARROW_START = "Start";
		public const string sARROW_END = "End";
		public const string sARROW_BOTH = "Both";

		// connector weight (default is LINE_WEIGHT_1
		public const string sLINE_WEIGHT_1 = "1 pt";

		// connector corner 
		public const double ROUNDING = 0.0625;

		// color string names
		public const string sCOLOR_BLACK = "Black";

		// connector line pattern
		public const double SHDW_PATTERN = 0;           // None

		public const string sLINE_PATTERN_SOLID = "Solid";
		public const double LINE_PATTERN_SOLID = 1;     // ______ solid line
		public const string sLINE_PATTERN_DASHED = "Dashed";
		public const double LINE_PATTERN_DASH = 2;      // _ _ _ _ dashed lined
		public const string sLINE_PATTERN_DOTTED = "Dotted";
		public const double LINE_PATTERN_DOTTED = 3;    // . . . . dotted line
		public const string sLINE_PATTERN_DASHDOT = "Dash_Dot";
		public const double LINE_PATTERN_DASHDOT = 4;   // _ . _ . _ dash/dot-(t)ed line

		public const string STINCEL_LABEL_POSITION_TOP = "Top";
		public const string STINCEL_LABEL_POSITION_BOTTOM = "Bottom";

		public const string RGB_COLOR_2SKIP = "RGB(77,77,77)";   // when reading a Visio diagram and writing an Excel file this is a color that is added.
																					// it's dark gray so we want to ignore this one

		public const string UNKNOWN_SHAPE = "OC_Unknown";        // value of the app default visio stencil name for an unknown shape when processing a Visio Diag.

		public const double STENCIL_SIZE_BUFFER = 0.080;	// buffer size.
																			// used in calculation to determine if the shape on the diagram has been resized
																			// If the shape width and height are with the margin based on the stencil width,height the
																			// shape has not been resized and the width and height can be set to 0

		/// <summary>
		/// ShowDiagram
		/// enum
		/// </summary>
		public enum ShowDiagram
		{
			NoShow = 0,
			Show = 1
		}

		/** ************************************************************************************** **/


		public enum VisioPageOrientation
		{
			Landscape,
			Portrait
		}

		/// <summary>
		/// GetVisioPageOrientation
		/// return the VisioPazgeOrientation based on the argument value
		/// </summary>
		/// <param name="pgOr"></param>
		/// <returns>VisioPageOrientation</returns>
		public static VisioPageOrientation GetVisioPageOrientation(string pgOr)
		{
			switch (pgOr.Trim().ToUpper())
			{
				case "Portrait":
					return VisioPageOrientation.Landscape;

				case "Landscape":
				default:
					return VisioPageOrientation.Portrait;
			}
		}

		/** ************************************************************************************** **/

		public enum VisioPageSize
		{
			Letter,
			Tabloid,
			Ledger,
			Legal,
			A3,
			A4
		}

		/// <summary>
		/// GetVisioPageSize
		/// return the supported VisioPageSize based on the value passed in
		/// </summary>
		/// <param name="pgSz"></param>
		/// <returns>VisioPageSize</returns>
		public static VisioPageSize GetVisioPageSize(string pgSz)
		{
			switch (pgSz.Trim().ToUpper())
			{
				case "TABLOID":
					return VisioPageSize.Tabloid;
				case "LEDGER":
					return VisioPageSize.Ledger;
				case "LEGAL":
					return VisioPageSize.Legal;
				case "A3":
					return VisioPageSize.A3;
				case "A4":
					return VisioPageSize.A4;

				case "LETTER":
				default:
					return VisioPageSize.Letter;
			}
		}

		/** ************************************************************************************** **/


	}
}

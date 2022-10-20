namespace OmnicellBlueprintingTool.Visio
{
	public class VisioVariables
	{
		public static string DefaultBlueprintingTemplateFile = "OC_BlueprintingTemplate.vstx";
		public static string DefaultBlueprintingStencilFile = "OC_BlueprintingStencils.vssx";

		//lngRed = RGB(255, 0, 0);
		//lngBlack = RGB(0, 0, 0);
		//lngYellow = RGB(255, 255, 0);
		//lngWhite = RGB(255, 255, 255);


		// Fill colors / connector colors
		public const string COLOR_BLACK = "RGB(0,0,0)";
		public const string COLOR_RED = "RGB(247,28,8)";
		public const string COLOR_GRAY = "RGB(128,128,128)";

		public const string COLOR_GREEN = "RGB(0,176,80)";
		//public const string COLOR_GREEN_SERVER = "RGB(146,208,80)";
		public const string COLOR_GREEN_LIGHT = "RGB(198,224,180)";

		public const string COLOR_ORANGE = "RGB(255,165,0)";
		public const string COLOR_ORANGE_SERVER = "RGB(255,192,0)";
		public const string COLOR_ORANGE_LIGHT = "RGB(253,147,8)";

		public const string COLOR_CYAN = "RGB(0,174,255)";

		public const string COLOR_BLUE = "RGB(51,102,255)";
		public const string COLOR_BLUE_SERVER = "RGB(0,170,255)";
		public const string COLOR_BLUE_LIGHT = "RGB(204,204,255)";

		public const string COLOR_PINK_LIGHT = "RGB(255,204,255)";
		public const string COLOR_WHITE = "RGB(255,255,255)";
		public const string COLOR_YELLOW = "RGB(246,222,75)";

		public const double HEIGHT = 0.25;

		public const double LINE_COLOR = 8; // Black
		public const double LINE_COLOR_MANY = 10;

		// connector ends
		public const double BEGIN_ARROW = 4;      // Filled arrow
		public const double END_ARROW = 4;        // Filled arrow
		public const double ARROW_NONE = 0;       // None
		public const string sARROW_NONE = "NONE";
		public const string sARROW_START = "START";
		public const string sARROW_END = "END";
		public const string sARROW_BOTH = "BOTH";

		// connector weight (default is LINE_WEIGHT_1
		public const string sLINE_WEIGHT_1 = "1.0 pt";
		public const string sLINE_WEIGHT_1_5 = "1.5 pt";
		public const string sLINE_WEIGHT_2 = "2 pt";

		// connector corner 
		public const double ROUNDING = 0.0625;

		// connector line pattern
		public const double SHDW_PATTERN = 0;           // None

		public const string sLINE_PATTERN_SOLID = "SOLID";
		public const double LINE_PATTERN_SOLID = 1;     // ______ solid line
		public const string sLINE_PATTERN_DASHED = "DASHED";
		public const double LINE_PATTERN_DASH = 2;      // _ _ _ _ dashed lined
		public const string sLINE_PATTERN_DOTTED = "DOTTED";
		public const double LINE_PATTERN_DOTTED = 3;    // . . . . dotted line
		public const string sLINE_PATTERN_DASHDOT = "DASH_DOT";
		public const double LINE_PATTERN_DASHDOT = 4;   // _ . _ . _ dash/dot-(t)ed line

		public const string STINCEL_LABEL_POSITION_TOP = "TOP";
		public const string STINCEL_LABEL_POSITION_BOTTOM = "BOTTOM";

		//public const string CATEGORY_PREFIX = "CAT_";
		//public const string FORMULA_PREFIX = "FOR_";
		//public const string LIST_PREFIX = "LIS_";
		//public const string LOOKUP_PREFIX = "LKP_";
		//public const short NAME_CHARACTER_SIZE = 10; // 10pt
		//public const string RULE_PREFIX = "RUL_";
		//public const short VISIO_SECTION_OJBECT_INDEX = 1;

		public enum FormulaUse
		{
			Value,
			Error,
			Hide,
			Match
		}

		public enum ShowDiagram
		{
			NoShow = 0,
			Show = 1
		}

		public enum VisioPageOrientation
		{
			Landscape,
			Portrait
		}

		public enum VisioPageSize
		{
			Letter,
			Tabloid,
			Ledger,
			Legal,
			A3, 
			A4
		}


		public static VisioPageOrientation GetVisioPageOrientation(string pgOr)
		{
			switch (pgOr.Trim().ToUpper())
			{
				case "Portrait":
					return VisioPageOrientation.Landscape;

				case "Landscape":
				default:
					return VisioPageOrientation.Portrait;
					break;
			}
		}

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
					break;
			}
		}

	}
}

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
		private static StringComparer comparer = StringComparer.OrdinalIgnoreCase;

		private static Dictionary<string, string> _visioColorsMap = null; // new Dictionary<string, string>(comparer); 
		//private static Dictionary<string, string> _visioColorsMap = null;

		public VisioVariables()
		{
			setupVisioColorsMap();
		}

		// notes
		//public const string COLOR_PEACH		// BD groupbox4 color
		//public const string COLOR_GREEN_SEA	// Omnicell header green
		//public const string COLOR_BLUE_STEEL // groupbox1 color

		public const double HEIGHT = 0.25;

		//public const double LINE_COLOR = 8; // Black
		//public const double LINE_COLOR_MANY = 10;

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
			}
		}

		private static void setupVisioColorsMap()
		{
			_visioColorsMap = new Dictionary<string, string>(comparer); 
			//_visioColorsMap = new Dictionary<string, string>();

			// Visio colors
			_visioColorsMap.Add("", "RGB(0,0,0)");
			_visioColorsMap.Add("Beige", "RGB(245,245,220)");
			_visioColorsMap.Add("Black", "RGB(0,0,0)");
			_visioColorsMap.Add("Blue", "RGB(30,144,255)");				// Dodger blue
			_visioColorsMap.Add("Blue Alice", "RGB(240,248,255)");
			_visioColorsMap.Add("Blue Server", "RGB(0,170,255)");
			_visioColorsMap.Add("Blue Steel", "RGB(176,196,222)");	// groupbox1 color
			_visioColorsMap.Add("Brown", "RGB(210,105,30)");
			_visioColorsMap.Add("Cyan", "RGB(0,255,255)");
			_visioColorsMap.Add("Gold", "RGB(255,215,0)");
			_visioColorsMap.Add("Gray", "RGB(128,128,128)");
			_visioColorsMap.Add("Green", "RGB(0,128,0)");
			_visioColorsMap.Add("Green Light", "RGB(154,205,50)");
			_visioColorsMap.Add("Green Lime", "RGB(50,205,50)");
			_visioColorsMap.Add("Green Sea", "RGB(60,179,113)");		// Omnicell header green
			_visioColorsMap.Add("Khaki", "RGB(240,230,140)");
			_visioColorsMap.Add("Dark Khaki", "RGB(189,183,107)");
			_visioColorsMap.Add("LIME", "RGB(0,255,0)");
			_visioColorsMap.Add("Magenta", "RGB(255,0,255)");
			_visioColorsMap.Add("Mint", "RGB(198,224,180)");
			_visioColorsMap.Add("Navy", "RGB(0,0,128)");
			_visioColorsMap.Add("Olive", "RGB(128,128,0)");
			_visioColorsMap.Add("Olive Drab", "RGB(107,142,35)");
			_visioColorsMap.Add("Orange", "RGB(255,165,0)");
			_visioColorsMap.Add("Orange Light", "RGB(255,192,0)");
			_visioColorsMap.Add("Peach", "RGB(255,242,204)");			// BD groupbox4 color
			_visioColorsMap.Add("Pink Light", "RGB(255,182,193)");
			_visioColorsMap.Add("Purple", "RGB(128,0,128)");
			_visioColorsMap.Add("Red", "RGB(255,0,0)");
			_visioColorsMap.Add("Salmon", "RGB(250,128,114)");
			_visioColorsMap.Add("Silver", "RGB(192,192,192)");
			_visioColorsMap.Add("Tan", "RGB(210,180,140)");
			_visioColorsMap.Add("Teal", "RGB(0,128.128)");
			_visioColorsMap.Add("White", "RGB(255,255,255)");
			_visioColorsMap.Add("White Smoke", "RGB(245,245,245)");
			_visioColorsMap.Add("Yellow", "RGB(255,255,0)");
		}

		/// <summary>
		/// GetRGBColor
		/// return the RGB color value based on the color string argument
		/// color argument "Black" will return "RGB(0,0,0)"
		/// </summary>
		/// <param name="color">search value</param>
		/// <returns>"RGB(???,???,???)"</returns>
		public static string GetRGBColor(string color)
		{
			string value = string.Empty;
			if (_visioColorsMap == null)
			{
				setupVisioColorsMap();
			}
			if (string.IsNullOrEmpty(color))
			{
				return null;
			}

			foreach(KeyValuePair<string, string> kvp in _visioColorsMap)
			{
				if (string.Equals(kvp.Key, color, StringComparison.OrdinalIgnoreCase))
				{
					return kvp.Value.Trim().ToString();
				}
			}

			//_visioColorsMap.TryGetValue(color, out value);
			//if (string.IsNullOrEmpty(value))
			//{
			//	return null;
			//}
			return value;
		}

		/// <summary>
		/// GetColorValueFromDB
		/// return the color string based on the rgb value argument
		/// search "RGB(0,0,0)" will return "BLACK
		/// </summary>
		/// <param name="rgb"></param>
		/// <returns>string</returns>
		/// <text>color name</text>
		public static string GetColorValueFromRGB(string rgb)
		{
			if (_visioColorsMap == null)
			{
				setupVisioColorsMap();
			}
			if (string.IsNullOrEmpty(rgb))
			{
				return null; // default to BLACK
			}
			foreach (KeyValuePair<string, string> item in _visioColorsMap)
			{
				if (item.Value.Equals(rgb.Trim()))
				{
					return item.Key;
				}
			}
			return null;
		}

		public static string[] GetAllKeyValues()
		{
			int nIdx = 0;
			string[] saTmp = new string[_visioColorsMap.Count];
			foreach (KeyValuePair<string, string> keyValue in _visioColorsMap)
			{
				saTmp[nIdx] = keyValue.Key.Trim();
			}
			return saTmp;
		}
	}
}

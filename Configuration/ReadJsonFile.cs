using OmnicellBlueprintingTool.Models;
using System.Collections.Generic;
using System.IO;
using System;
using System.Runtime.Remoting.Lifetime;
using System.Linq;
using Microsoft.Office.Interop.Excel;

using System.Text;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Web.Script.Serialization;
using OmnicellBlueprintingTool.Extensions;
using System.Windows.Forms;

namespace OmnicellBlueprintingTool.Configuration
{
	public static class ReadJsonFile
	{
		public static AppConfiguration ReadJSONFile(string fileNamePath)
		{
			AppConfiguration appConfig = new AppConfiguration();

			try
			{
				using (StreamReader r = new StreamReader(fileNamePath))
				{
					string json = r.ReadToEnd();
					JavaScriptSerializer serializer = new JavaScriptSerializer();
					dynamic jsonObject = serializer.Deserialize<dynamic>(json);

					dynamic root = jsonObject["Omnicell"]; // Root
					dynamic app = root["BlueprintingTool"]; // result is asdf
					dynamic tables = app["Tables"]; // result is asdf

					object[] values = tables["Colors"]; // result is asdf
					appConfig.Colors = values.Select(i => i.ToString()).ToList();

					values = tables["Arrows"]; // result is asdf
					appConfig.Arrows = values.Select(i => i.ToString()).ToList();

					values = tables["Shape Types"]; // result is asdf
					appConfig.ShapeTypes = values.Select(i => i.ToString()).ToList();

					values = tables["Line Patterns"]; // result is asdf
					appConfig.LinePatterns = values.Select(i => i.ToString()).ToList();

					values = tables["Stencil Label Positions"]; // result is asdf
					appConfig.StencilLabelPosition = values.Select(i => i.ToString()).ToList();

					values = tables["Label Font Sizes"]; // result is asdf
					appConfig.LabelFontSizes = values.Select(i => i.ToString()).ToList();
				}
			}
			catch(FileNotFoundException fne)
			{
				string sTmp = string.Format("ReadJsonFile - Exception\n\nApplication JSON configuration file not found:{0}\nPlease ensure the 'OmnicellBlueprintingTool.json' file is in the same folder as the application\n\n{1}", fileNamePath, fne.Message);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				appConfig = null;
			}
			return appConfig;
		}
	}
}

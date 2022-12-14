using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using OmnicellBlueprintingTool.Visio;

namespace OmnicellBlueprintingTool.Configuration
{
	public static class ReadJsonFile
	{
		/// <summary>
		/// ReadJSONFile
		/// Process the Application configuration JSON file to obtain app settings
		/// </summary>
		/// <param name="fileNamePath"></param>
		/// <param name="visioHelper"></param>
		/// <returns>bool false:success</returns>
		public static bool ReadJSONFile(string fileNamePath, ref VisioHelper visioHelper)
		{
			//AppConfiguration appConfig = new AppConfiguration();
			try
			{
				using (StreamReader r = new StreamReader(fileNamePath))
				{
					string json = r.ReadToEnd();
					JavaScriptSerializer serializer = new JavaScriptSerializer();
					dynamic jsonObject = serializer.Deserialize<dynamic>(json);

					dynamic root = jsonObject["Omnicell"];		// Root
					dynamic app = root["BlueprintingTool"];   // sub-root	

					// process the Json data for the Default Application configuration file

					dynamic tables = app["Tables"];           // Blueprinting tool Excel data file Tables sheet

					List<string> colors = new List<string>();
					Dictionary<string, string> colorMap = new Dictionary<string, string>();
					object[] values = tables["Colors"]; // result is asdf

					foreach (Dictionary<string, object> pair in values)
					{
						string key = pair.ElementAtOrDefault(0).Value.ToString().Trim();
						string value = pair.ElementAtOrDefault(1).Value.ToString().Trim();
						if (key.Equals("Blank", StringComparison.OrdinalIgnoreCase))
						{
							colors.Add("");
							colorMap.Add("\"\"", value);
						}
						else
						{
							colors.Add(key);
							colorMap.Add(key, value);
						}
					}
					visioHelper.SetColorsMap(colorMap);

					values = tables["Arrows"];
					visioHelper.SetConnectorArrowTypes(values.Select(i => i.ToString()).ToList());

					values = tables["Shape Types"];
					visioHelper.SetShapeTypes(values.Select(i => i.ToString()).ToList());

					values = tables["Line Patterns"];
					visioHelper.SetConnectorLinePatterns(values.Select(i => i.ToString()).ToList());

					values = tables["Line Weights"];
					visioHelper.SetConnectorLineWeights(values.Select(i => i.ToString()).ToList());

					values = tables["Stencil Label Positions"];
					visioHelper.SetStencilLabelPositions(values.Select(i => i.ToString()).ToList());

					values = tables["Shape Label Font Sizes"];
					visioHelper.SetStencilLabelFontSizes(values.Select(i => i.ToString()).ToList());

					values = tables["Stencil Image Names"];
					visioHelper.SetDefaultStencilNames(values.Select(i => i.ToString()).ToList());
				}
			}
			catch(FileNotFoundException)
			{
				string sTmp = string.Empty;
				sTmp = string.Format("ReadJsonFile - Exception\n\nThe Custom Config JSON file '{0}' was not found\n\nPlease ensure this file in the folder specified within the Excel Data file", fileNamePath);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return true;	// error
			}
			return false;	// success
		}
	}
}

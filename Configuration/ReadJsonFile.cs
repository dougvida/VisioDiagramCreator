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

					dynamic root = jsonObject["Omnicell"]; // Root
					dynamic app = root["BlueprintingTool"]; // result is asdf
					dynamic tables = app["Tables"]; // result is asdf

					List<string> colors = new List<string>();
					Dictionary<string, string> colorMap = new Dictionary<string, string>();
					object[] values = tables["Colors"]; // result is asdf
      
					foreach (Dictionary<string,object> pair in values)
					{
						string key = pair.ElementAtOrDefault(0).Value.ToString().Trim();
						string value = pair.ElementAtOrDefault(1).Value.ToString().Trim();
						if (key.Equals("Blank"))
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
					//visioHelper.Colors = colors;
					//var xxx = values.Select(i => i.ToString()).ToList();

					values = tables["Arrows"];
					visioHelper.SetConnectorArrowsMap(values.Select(i => i.ToString()).ToList());

					values = tables["Shape Types"];
					visioHelper.SetShapeTypesMap(values.Select(i => i.ToString()).ToList());

					values = tables["Line Patterns"];
					visioHelper.SetConnectorLinePatterns(values.Select(i => i.ToString()).ToList());

					values = tables["Line Weights"];
					visioHelper.SetConnectorLineWeightsMap(values.Select(i => i.ToString()).ToList());

					values = tables["Stencil Label Positions"];
					visioHelper.SetStencilLabelPositionsMap(values.Select(i => i.ToString()).ToList());

					values = tables["Shape Label Font Sizes"];
					visioHelper.SetStencilLabelFontSizeMap(values.Select(i => i.ToString()).ToList());

					values = tables["Stencil Image Names"];
					visioHelper.SetDefaultStencilNamesMap(values.Select(i => i.ToString()).ToList());
				}
			}
			catch(FileNotFoundException fne)
			{
				string sTmp = string.Format("ReadJsonFile - Exception\n\nApplication JSON configuration file not found:{0}\nPlease ensure the 'OmnicellBlueprintingTool.json' file is in the same folder as the application\n\n{1}", fileNamePath, fne.Message);
				MessageBox.Show(sTmp, "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return true;	// error
			}
			return false;	// success
		}
	}
}

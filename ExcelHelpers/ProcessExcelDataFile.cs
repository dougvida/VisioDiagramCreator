using System;
using System.Collections.Generic;
using System.Linq;
using IronXL;
using VisioDiagramCreator.Models;
using VisioDiagramCreator.Visio;

// https://ironsoftware.com/csharp/excel/tutorials/how-to-read-excel-file-csharp/#get-cell-range


namespace VisioDiagramCreator.Helpers
{
	public class ProcessExcelDataFile
	{
		private enum _cellIndex
		{
			// NOTE ****
			// the order of this enum must match the column order in the Excel file
			VisioPage = 0,			// Page indicator to place this shape
			ShapeType,				// key
			ShapeKey,				// device unique Key used for connecting visio shapes
			ShapeImage,				// device visio image name
			ShapeLabel,				// device label
			Mach_name,				// device machine Name
			Mach_id,					// device machine Id
			Site_id,					// device site Id
			Site_name,				// deivce site name
			Site_address,			// device site address
			Omnis_name,				// device name
			Omnis_id,            // device Id
			SiteIdOmniId,			// site_id+omni_id
			IP,						// device IP address
			Ports,               // device Ports
			DevicesCount,        // number of Devices for this type (part of a group)

			PosX,						// Shape position X
			PosY,						// shape position Y
			Width,					// shape width
			Height,					// shape height
			FillColor,				// color to fill stincel
			ConnectFrom,         // used to link this visio shape to another visio shape
			FromLineLabel,			// Arrow Text
			FromLinePattern,		// Line pattern solid = 1
			FromArrowType,			// Can contain one of these [None, Start, End, Both]
			FromLineColor,			// Arrow Color
			ConnectTo,           // used to link this visio shape to another visio shape
			ToLineLabel,			// Arrow Text
			ToLinePattern,       // Line pattern solid = 1
			ToArrowType,			// Can contain one of these [None, Start, End, Both]
			ToLineColor,			// Arrow Color
		}

		/// <summary>
		/// ParseData
		/// Parse the Excel data into a DiagramData class.
		/// this class will hold all the excel data that will be used to transfer into Visio diagram data
		/// </summary>
		/// <param name="file"></param>
		/// <param name="diagData">DiagramData</param>
		/// <returns>DiagramData</returns>
		/// <exception cref="ArgumentNullException"></exception>
		/// <exception cref="Exception"></exception>
		public DiagramData ParseData( string file, DiagramData diagData )
		{
			if (string.IsNullOrEmpty(file))
			{
				// Error file is empty
				throw new ArgumentNullException(nameof(file));
			}

			List<Device> devices = new List<Device>();
			Device device = null;

			WorkBook workbook = WorkBook.Load(file);
			WorkSheet sheet = workbook.WorkSheets.First();

			string[] saTmp1 = sheet.RangeAddressAsString.Split(':');
			int nIdx1 = saTmp1[1].IndexOfAny("0123456789".ToCharArray());	// if a digit is found get the index
			string endColumn = saTmp1[1].Substring(0, nIdx1);

			for (var row = 2; row <= sheet.RowCount; row++)
			{
				var cells = sheet[$"A{row}:{endColumn}{row}"].ToList();
				if (!cells[0].ToString().StartsWith(";")) // first row is a header so skip
				{
					try
					{
						if (cells[(int)_cellIndex.VisioPage].IntValue > diagData.MaxVisioPages)
						{
							diagData.MaxVisioPages = cells[(int)_cellIndex.VisioPage].IntValue;
						}
						switch (cells[(int)_cellIndex.ShapeType].ToString().Trim())
						{
							case "Template":
								diagData.TemplateFilePath = cells[(int)_cellIndex.ShapeKey].ToString().Trim().Substring(0, cells[(int)_cellIndex.ShapeKey].ToString().Trim().Length - 1);
								break;

							case "Stencil":
								diagData.StencilFilePath = cells[(int)_cellIndex.ShapeKey].ToString().Trim().Substring(0, cells[(int)_cellIndex.ShapeKey].ToString().Trim().Length - 1);
								break;

							case "Shape":
								device = _parseExcelData(cells);
								devices.Add(device);
								diagData.AllShapesMap.Add(device.ShapeInfo.UniqueKey, device);
								break;

							default:
								//throw new Exception("Unknown label in CSV file: " + saTmp[0].Trim().ToString());
								break;
						}
					}
					catch (Exception ex)
					{
						Console.WriteLine(ex.Message +" - "+ ex.StackTrace);	
						throw new Exception(String.Format("Exception: Duplicate key:({0}) found.\nPlease resolve this issue in the Excel Data file\n{1}", device.ShapeInfo.UniqueKey, ex.Message)); //, ex.StackTrace.ToString);
					}
				}
			}

			if (diagData != null)
			{
				diagData.Devices = devices;
			}
			if(workbook != null)
			{ 
				workbook.Close();
			}

			return diagData;
		}

		/// <summary>
		/// _parseExcelData
		/// parse the Excel data into a Device object
		/// </summary>
		/// <param name="data">List<cell></param>
		/// <returns>Device</returns>
		private Device _parseExcelData(List<Cell> data)
		{
			Device device = new Device();
			ShapeInformation visioInfo = new ShapeInformation();
			try
			{
				visioInfo.VisioPage = data[(int)_cellIndex.VisioPage].IntValue;
				visioInfo.UniqueKey = data[(int)_cellIndex.ShapeKey].ToString().Trim();      // unique key for this shape
				visioInfo.StencilImage = data[(int)_cellIndex.ShapeImage].ToString().Trim(); // must match exactly the name in the visio stencil
				visioInfo.StencilLabel = data[(int)_cellIndex.ShapeLabel].ToString().Trim(); // text to add to the stencil image

				device.MachineName = data[(int)_cellIndex.Mach_name ].ToString().Trim();
				device.MachineId = data[(int)_cellIndex.Mach_id].ToString().Trim();
				device.SiteId = data[(int)_cellIndex.Site_id].ToString().Trim();
				device.SiteName = data[(int)_cellIndex.Site_name].ToString().Trim();
				device.SiteAddress = data[(int)_cellIndex.Site_address].ToString().Trim();
				device.OmniName = data[(int)_cellIndex.Omnis_name].ToString().Trim();
				device.OmniId = data[(int)_cellIndex.Omnis_id].ToString().Trim();
				device.SiteId_OmniId = data[(int)_cellIndex.SiteIdOmniId].ToString().Trim();
				device.OmniIP = data[(int)_cellIndex.IP].ToString().Trim();
				device.OmniPorts = data[(int)_cellIndex.Ports].ToString().Trim();

				if (!string.IsNullOrEmpty(data[(int)_cellIndex.PosX].ToString().Trim()))         // position X to place the stencil image
				{
					visioInfo.Pos_x = double.Parse(data[(int)_cellIndex.PosX].ToString().Trim(), System.Globalization.CultureInfo.InvariantCulture);
				}
				if (!string.IsNullOrEmpty(data[(int)_cellIndex.PosY].ToString().Trim()))         // position Y to place the stencil image
				{
					visioInfo.Pos_y = double.Parse(data[(int)_cellIndex.PosY].ToString().Trim(), System.Globalization.CultureInfo.InvariantCulture);
				}
				if (!string.IsNullOrEmpty(data[(int)_cellIndex.Width].ToString().Trim()))        // Width of the stencil image
				{
					visioInfo.Width = double.Parse(data[(int)_cellIndex.Width].ToString().Trim(), System.Globalization.CultureInfo.InvariantCulture);
				}
				if (!string.IsNullOrEmpty(data[(int)_cellIndex.Height].ToString().Trim()))       // Height of the stencil image
				{
					visioInfo.Height = double.Parse(data[(int)_cellIndex.Height].ToString().Trim(), System.Globalization.CultureInfo.InvariantCulture);
				}
				if (!string.IsNullOrEmpty(data[(int)_cellIndex.DevicesCount].ToString()))        // number of cabinets/Devices for this object.   If (empty/null) don't append to the stencil text
				{
					visioInfo.StencilLabel += " / " + data[(int)_cellIndex.DevicesCount].ToString().Trim();
				}
				if(!string.IsNullOrEmpty(data[(int)_cellIndex.FillColor].ToString()))
				{
					// should be a string like
					visioInfo.FillColor = data[(int)_cellIndex.FillColor].ToString().Trim();
				}

				// Get the ShpFromObj section
				if (!string.IsNullOrEmpty(data[(int)_cellIndex.ConnectFrom].ToString().Trim()))  // unique key for the To shape identifier - will match another shape object field 2 or empty for no connection
				{
					visioInfo.ConnectFrom = data[(int)_cellIndex.ConnectFrom].ToString().Trim();
				}
				visioInfo.FromLineLabel = data[(int)_cellIndex.FromLineLabel].ToString().Trim();

				// Arrow type to use if enabled
				visioInfo.FromLinePattern = data[(int)_cellIndex.FromLinePattern].IntValue;
				if (visioInfo.FromLinePattern <= 0)
				{
					visioInfo.FromLinePattern = (int)VisioVariables.LINE_PATTERN_SOLID;
				}

				// set the ShpFromObj ArrowType
				string sTmp = data[(int)_cellIndex.FromArrowType].ToString().Trim().ToUpper();
				switch ( sTmp )
				{
					case VisioVariables.sARROW_START:
						visioInfo.FromArrowType = VisioVariables.sARROW_START;
						break;
					case VisioVariables.sARROW_END:
						visioInfo.FromArrowType = VisioVariables.sARROW_END;
						break;
					case VisioVariables.sARROW_BOTH:
						visioInfo.FromArrowType = VisioVariables.sARROW_BOTH;
						break;
					default:
						visioInfo.FromArrowType = VisioVariables.sARROW_NONE;
						break;
				}

				visioInfo.FromLineColor = data[(int)_cellIndex.FromLineColor].ToString().Trim();
				if (string.IsNullOrEmpty(visioInfo.FromLineColor))
				{
					visioInfo.FromLineColor = VisioVariables.COLOR_BLACK;
				}

				// Get the To section
				if (!string.IsNullOrEmpty(data[(int)_cellIndex.ConnectTo].ToString().Trim()))    // unique key for the To shape identifier - will match another shape object field 2 or empty for no connection
				{
					visioInfo.ConnectTo = data[(int)_cellIndex.ConnectTo].ToString().Trim();
				}
				visioInfo.ToLineLabel = data[(int)_cellIndex.ToLineLabel].ToString().Trim();

				// Arrow type to use if enabled
				visioInfo.ToLinePattern = data[(int)_cellIndex.ToLinePattern].DoubleValue;
				if(visioInfo.ToLinePattern <= 0)
				{
					visioInfo.ToLinePattern = VisioVariables.LINE_PATTERN_SOLID;
				}

				// do we want to have a start arrow
				// set the ShpFromObj ArrowType
				sTmp = data[(int)_cellIndex.ToArrowType].ToString().Trim().ToUpper();
				switch (sTmp)
				{
					case VisioVariables.sARROW_START:
						visioInfo.ToArrowType = VisioVariables.sARROW_START;
						break;
					case VisioVariables.sARROW_END:
						visioInfo.ToArrowType = VisioVariables.sARROW_END;
						break;
					case VisioVariables.sARROW_BOTH:
						visioInfo.ToArrowType = VisioVariables.sARROW_BOTH;
						break;
					default:
						visioInfo.ToArrowType = VisioVariables.sARROW_NONE;
						break;
				}

				visioInfo.ToLineColor = data[(int)_cellIndex.ToLineColor].ToString().Trim();
				if (string.IsNullOrEmpty(visioInfo.ToLineColor))
				{
					visioInfo.ToLineColor = VisioVariables.COLOR_BLACK;
				}
				device.ShapeInfo = visioInfo;
			}
			catch (Exception exp)
			{
				Console.WriteLine(exp.Message+" - "+exp.StackTrace);
				return null;
			}
			//Console.WriteLine("adding stencil:{0}",visioInfo.UniqueKey);
			return device;
		}
	}
}

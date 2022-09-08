using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Visio1 = Microsoft.Office.Interop.Visio;

namespace VisioDiagramCreator.Visio
{
	internal class ConnectAllShapes
	{
		partial class Form1
		{
			private const string VisioAppID = "visio.application";

			private void m_connectAllShapes()
			{
				// First, find a running instance of Visio:
				Visio1.Application visApp = m_getVisio();

				if (visApp == null)
				{
					System.Console.WriteLine(
					"Couldn't find a running instance of Visio!");
					return;
				}

				int UndoID = -1;

				try
				{
					// Create an undo-scope, so that we can undo all the
					// connections with just one Ctrl + Z:
					UndoID = visApp.BeginUndoScope("Connect All Shapes to Each Other");

					// This is where we really get the connecting done:
					// Get a Visio Page object:
					Visio1.Page pg = visApp.ActivePage;

					// Connect all shapes on the page:
					m_connect(pg);

					visApp.EndUndoScope(UndoID, true);
				}
				catch (System.Exception ex)
				{
					// Try to close the undo scope, but reject the changes:
					if ((UndoID == -1) && visApp != null)
					{
						visApp.EndUndoScope(UndoID, false);
					}
					System.Console.WriteLine(
						 "An error occurred!\n\n" + ex.Message);
				}
			}

			private void m_connect(Visio1.Page visPg)
			{
				Visio1.Shape shpFrom, shpTo;
				List< Visio1.Shape> collShapes;

				// Set the page-layout settings for routing-style,
				// jump-style, etc.
				m_setPageLayoutSettings(visPg);

				// Add all the non-connector shapes to a VB collection
				collShapes = m_getShapesToConnect(visPg);

				// Loop through the shapes in the shapes collection --
				// connect the ith shape to each jth shape, so to speak:
				for (int i = 0; i < collShapes.Count; i++)
				{
					shpFrom = collShapes[i];

					// Connect to all the other shapes:
					for (int j = i + 1; j < collShapes.Count; j++)
					{
						shpTo = collShapes[j];
						m_connectShapes(shpFrom, shpTo);
					}
				}
			}

			private void m_connectShapes(Visio1.Shape shpFrom, Visio1.Shape shpTo)
			{
				// Visio 2007 introduced a new method for connection
				// shapes. This proc looks at the Visio version and
				// decides whether to use the old way or the new way.

				Visio1.Page pg = shpFrom.ContainingPage;

				// Note: if you're not running Visio 2007, this might not
				// even compile -- you'll have to comment-out the first part
				// of the If-Then block...
				if (string.Compare(pg.Application.Version, "12.0", true) == 0)
				{
					shpFrom.AutoConnect(shpTo,(short)Visio1.VisAutoConnectDir.visAutoConnectDirNone, null);
				}
				else
				{
					// Drop the built-in connector object somewhere on the page:
					Visio1.Shape shpConn;
					shpConn = pg.Drop(pg.Application.ConnectorToolDataObject, 0, 0);

					// Connect its Begin to the 'From' shape:
					shpConn.get_CellsU("BeginX").GlueTo(shpFrom.get_CellsU("PinX"));

					// Connect its End to the 'To' shape:
					shpConn.get_CellsU("EndX").GlueTo(shpTo.get_CellsU("PinX"));
				}
			}

			private List<Visio1.Shape> m_getShapesToConnect(Visio1.Page visPg)
			{
				List< Visio1.Shape> collShapes = new List<Visio1.Shape>();

				// For this example, we will get all shapes on the page
				// that ARE NOT of these:
				//
				//  1. Connectors
				//  2. Foreign objects (like Buttons)
				//  3. Guides

				foreach (Visio1.Shape shp in visPg.Shapes)
				{
					if ((shp.OneD == 0) &&
						  (shp.Type != (short)Visio1.VisShapeTypes.visTypeForeignObject) &&
						  (shp.Type != (short)Visio1.VisShapeTypes.visTypeGuide))
					{
						collShapes.Add(shp);
					}
				}
				return collShapes;
			}

			private Visio1.Application m_getVisio()
			{
				Visio1.Application visApp;
				object objVis;

				objVis = System.Runtime.InteropServices.

				Marshal.GetActiveObject(VisioAppID);
				visApp = (Visio1.Application)objVis;

				return visApp;
			}

			private void m_setPageLayoutSettings(Visio1.Page visPg)
			{
				// We can set layout and routing options for the page by
				// accessing the ShapeSheet for the page, and setting cells
				// in the Page Layout section.
				//
				// You can see the PageSheet by deselecting all shapes on the
				// page, and choosing Window > Show ShapeSheet.
				// Set page routing style to center-to-center:

				visPg.PageSheet.get_CellsSRC(
						(short)Visio1.VisSectionIndices.visSectionObject,
						(short)Visio1.VisRowIndices.visRowPageLayout,
						(short)Visio1.VisCellIndices.visPLORouteStyle).ResultIUForce = 16;

				// Set to connector intersection to 'gap':
				visPg.PageSheet.get_CellsSRC(
						(short)Visio1.VisSectionIndices.visSectionObject,
						(short)Visio1.VisRowIndices.visRowPageLayout,
						(short)Visio1.VisCellIndices.visPLOJumpStyle).ResultIUForce = 2;

				// Note: another way to access the PageSheet cells is by name, ie:
				//visPg.PageSheet.get_Cells("RouteStyle").ResultIU = 16;
				//visPg.PageSheet.get_Cells("LineJumpStyle").ResultIU = 2;
			}
		}
	}
}

using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;

namespace TagLinkedRooms
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Access current selection
            Selection sel = uidoc.Selection;

            // Retrieve rooms from the document
            FilteredElementCollector col = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            // Iterate over the rooms
            foreach (Element e in col)
            {
                Room room = e as Room;
                if (room != null)
                {
                    try
                    {
                        // Get the centroid of the room
                        XYZ roomCentroid = GetRoomCentroid(room);
                        Debug.Print("Room Centroid: " + roomCentroid.ToString());
                    }
                    catch (Exception ex)
                    {
                        Debug.Print("An error occurred: " + ex.Message);
                    }
                }
            }

            return Result.Succeeded;
        }

        private XYZ GetRoomCentroid(Room room)
        {
            // Get the room boundary curve
            CurveLoop roomBoundary = GetRoomBoundary(room);

            // Calculate the centroid of the room boundary curve
            XYZ centroid = XYZ.Zero;
            int count = 0;

            // Iterate over the segments of the boundary curve
            foreach (var segment in roomBoundary)
            {
                // Accumulate the coordinates of the vertices
                centroid += segment.GetEndPoint(0);
                count++;
            }

            if (count > 0)
            {
                // Calculate the average centroid
                centroid /= count;
                return centroid;
            }
            else
            {
                // Handle cases where the room boundary is not valid
                throw new InvalidOperationException("Unable to calculate room centroid.");
            }
        }

        private CurveLoop GetRoomBoundary(Room room)
        {
            // Get the room geometry
            GeometryElement roomGeometry = room.get_Geometry(new Options());

            // Find the boundary curve of the room
            foreach (GeometryObject geomObj in roomGeometry)
            {
                Solid solid = geomObj as Solid;
                if (solid != null && solid.Volume > 0)
                {
                    // Iterate over the faces of the solid
                    foreach (Face face in solid.Faces)
                    {
                        // Check if the face is planar
                        if (face is PlanarFace planarFace)
                        {
                            // Get the outer loop of the face
                            IList<CurveLoop> loops = planarFace.GetEdgesAsCurveLoops();
                            foreach (CurveLoop loop in loops)
                            {
                                return loop;
                            }
                        }
                    }
                }
            }

            // Return null if the room boundary curve is not found
            return null;
        }
    }
}

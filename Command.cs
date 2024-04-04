using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

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

            //Get revit linked document
            var revitLinks = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_RvtLinks)
                .OfClass(typeof(RevitLinkInstance)).ToList();
            //Get revit arch model links
            var archLinks = revitLinks;//.Where(a => a.Document.Title.ToUpper().Contains("ARCH")).ToList();
            //Get first link available
            var firstArchLink = archLinks.FirstOrDefault();
            //Get document of first link available
            var firstArchDoc = firstArchLink.Document;

            // Retrieve rooms from the linked arch document
            FilteredElementCollector linkedArchRooms = new FilteredElementCollector(firstArchDoc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            // Retrieve rooms from the document
            FilteredElementCollector currentModelRooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();
            //message
            var roomCenters = string.Empty;
            Transaction transaction = new Transaction(doc, "Tran");
            transaction.Start();
            // Iterate over the rooms
            foreach (Element e in currentModelRooms)
            {
                Room room = e as Room;
                if (room != null)
                {
                    try
                    {
                        // Get the centroid of the room
                        var roomLocation = (room.Location as LocationPoint).Point;
                        doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(roomLocation.X, roomLocation.Y), doc.ActiveView.Id);
                        roomCenters = roomCenters + Environment.NewLine + "Room Location: " + roomLocation.ToString();
                    }
                    catch (Exception ex)
                    {
                        Debug.Print("An error occurred: " + ex.Message);
                    }
                }
            }
            // Iterate over the linked rooms
            foreach (Element e in linkedArchRooms)
            {
                Room room = e as Room;
                if (room != null)
                {
                    try
                    {
                        // Get the centroid of the room
                        var roomLocation = (room.Location as LocationPoint).Point;
                        doc.Create.NewRoomTag(new LinkElementId(room.Id), new UV(roomLocation.X, roomLocation.Y), doc.ActiveView.Id);
                        roomCenters = roomCenters + Environment.NewLine + "Room Location: " + roomLocation.ToString();
                    }
                    catch (Exception ex)
                    {
                        Debug.Print("An error occurred: " + ex.Message);
                    }
                }
            }
            transaction.Commit();
            TaskDialog.Show("Room Centroids", roomCenters);
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

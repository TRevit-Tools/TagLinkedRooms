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

            //Get revit link instances
            var revitLinks = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance)).ToList();
            //Get revit link instances that have ARCH in the name
            var archLinks = revitLinks.Where(a => a.Name.ToUpper().Contains("ARCH")).ToList();
            //Get first link available
            var firstArchLink = archLinks.FirstOrDefault() as RevitLinkInstance;
            //Get the linkdocument from the revit link instance
            var firstArchDoc = firstArchLink.GetLinkDocument();



            //Retrieve Levels from the Linked Arch Document
            var linkedArchlevels = new FilteredElementCollector(firstArchDoc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            //retrieve levels form the host model
            var hostLevels = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .ToList();

            //Compare Link levels VS host levels and return linked levels that match host levels

            var matchingLevels = linkedArchlevels
                .Where(archLevel => hostLevels.Any(hostLevel => hostLevel.Name == archLevel.Name))
                .ToList();

            var filteredRooms = new FilteredElementCollector(firstArchDoc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Where(room =>
                {
                    var roomLevel = room.get_Parameter(BuiltInParameter.LEVEL_NAME).AsString();

                    return matchingLevels.Any(level => level.Name.Equals(roomLevel));
                })
                .ToList();

          

            //Retrieve all floor and ceiling plans
            List<ViewPlan> floorPlans = GetFloorPlans(doc);
            List<ViewPlan> ceilingPlans = GetCeilingPlans(doc);
            var allPlans = floorPlans.Concat(ceilingPlans).ToList();

            //message
            var roomCenters = string.Empty;
            //start transaction for placing room tags
            Transaction transaction = new Transaction(doc, "Place Room Tags");
            transaction.Start();

            List<ElementId> taggedRoomIds = new List<ElementId>(); // Declaration of taggedRoomIds list



            // Iterate over the rooms
            try
            {
                foreach (ViewPlan plans in allPlans)
                {
                    foreach (Element element in filteredRooms)
                    {
                        Room room = element as Room; // Assign element to room if it's a Room

                        if (room != null && room.LevelId != ElementId.InvalidElementId) //&& !taggedRoomIds.Contains(room.Id)) // Check if room is not null and not already tagged
                        {
                            if (room.Location is LocationPoint locationPoint)
                            {
                                LinkElementId linkedRoom = new LinkElementId(firstArchLink.Id, room.Id);
                                var roomLocation = locationPoint.Point;
                                var modifiedRoomLocation = firstArchLink.GetTransform().Origin + roomLocation;
                                doc.Create.NewRoomTag(linkedRoom, new UV(modifiedRoomLocation.X, modifiedRoomLocation.Y), plans.Id);
                                //taggedRoomIds.Add(room.Id);
                            }
                            else
                            {
                                Debug.Print($"Room location is not a LocationPoint. Location type: {room.Location?.GetType().ToString() ?? "null"}");
                            }
                        }
                        else if (room == null)
                        {
                            Debug.Print("Element is not a Room.");
                        }
                        else
                        {
                            Debug.Print("Room is already tagged.");
                        }
                    }
                }
            }
            catch (InvalidOperationException invalidOpEx)
            {
                Debug.Print("An invalid operation occurred: " + invalidOpEx.Message);
            }
            catch (NullReferenceException nullRefEx)
            {
                Debug.Print("A null reference exception occurred: " + nullRefEx.Message);
            }
            catch (Exception ex)
            {
                Debug.Print("An error occurred: " + ex.Message);
            }


            //commit transaction for placing room tags
            transaction.Commit();
            //TaskDialog.Show("Room Centroids", roomCenters);
            return Result.Succeeded;
        }





        private List<ViewPlan> GetFloorPlans(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(v => v.ViewType == ViewType.FloorPlan)
                .ToList();
        }

        private List<ViewPlan> GetCeilingPlans(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(v => v.ViewType == ViewType.CeilingPlan)
                .ToList();
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

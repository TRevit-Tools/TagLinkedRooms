using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
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

            // Get all linked instances in the document
            FilteredElementCollector linkCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance));

            foreach (Element element in linkCollector)
            {
                if (element is RevitLinkInstance linkInstance)
                {
                    Document linkedDoc = linkInstance.GetLinkDocument();

                    if (linkedDoc == null)
                        continue;

                    // Get all floor plans and reflected ceiling plans
                    List<ViewPlan> floorPlans = GetFloorPlans(doc);
                    List<ViewPlan> ceilingPlans = GetCeilingPlans(doc);

                    // Filter rooms in the linked document
                    FilteredElementCollector roomCollector = new FilteredElementCollector(linkedDoc)
                        .OfCategory(BuiltInCategory.OST_Rooms)
                        .WhereElementIsNotElementType();

                    foreach (ViewPlan floorPlan in floorPlans)
                    {
                        Level hostLevel = GetHostLevel(doc, floorPlan);
                        if (hostLevel != null)
                        {
                            // Find the corresponding view plan for the room's level in the host model
                            ViewPlan associatedView = GetViewPlanByLevel(doc, hostLevel, floorPlans, ceilingPlans);

                            foreach (Element elem in roomCollector)
                            {
                                if (elem is Room room)
                                {
                                    try
                                    {
                                        // Get the room centroid
                                        XYZ roomCentroid = GetRoomCentroid(room);
                                        Debug.Print("Room Centroid: " + roomCentroid.ToString());

                                        // Convert XYZ to UV
                                        UV roomCentroidUV = new UV(roomCentroid.X, roomCentroid.Y);

                                        // Tag the room in the associated view
                                        using (Transaction transaction = new Transaction(doc))
                                        {
                                            transaction.Start("Tag Room");
                                            doc.Create.NewRoomTag(new LinkElementId(linkInstance.Id), roomCentroidUV, associatedView.Id);
                                            transaction.Commit();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.Print("An error occurred: " + ex.Message);
                                    }
                                }
                            }
                        }
                    }

                    foreach (ViewPlan ceilingPlan in ceilingPlans)
                    {
                        Level hostLevel = GetHostLevel(doc, ceilingPlan);
                        if (hostLevel != null)
                        {
                            // Find the corresponding view plan for the room's level in the host model
                            ViewPlan associatedView = GetViewPlanByLevel(doc, hostLevel, floorPlans, ceilingPlans);

                            foreach (Element elem in roomCollector)
                            {
                                if (elem is Room room)
                                {
                                    try
                                    {
                                        // Get the room centroid
                                        XYZ roomCentroid = GetRoomCentroid(room);
                                        Debug.Print("Room Centroid: " + roomCentroid.ToString());

                                        // Convert XYZ to UV
                                        UV roomCentroidUV = new UV(roomCentroid.X, roomCentroid.Y);

                                        // Tag the room in the associated view
                                        using (Transaction transaction = new Transaction(doc))
                                        {
                                            transaction.Start("Tag Room");
                                            doc.Create.NewRoomTag(new LinkElementId(linkInstance.Id), roomCentroidUV, associatedView.Id);
                                            transaction.Commit();
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Debug.Print("An error occurred: " + ex.Message);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return Result.Succeeded;
        }

        private Level GetHostLevel(Document doc, ViewPlan viewPlan)
        {
            // Get the level associated with the view plan
            Level viewPlanLevel = doc.GetElement(viewPlan.LevelId) as Level;

            // Get the host level by name
            string levelName = viewPlanLevel?.Name;
            if (levelName != null)
            {
                // Find the corresponding host level by name
                Level hostLevel = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .WhereElementIsNotElementType()
                    .Cast<Level>()
                    .FirstOrDefault(l => l.Name == levelName);

                return hostLevel;
            }

            return null;
        }

        private ViewPlan GetViewPlanByLevel(Document doc, Level hostLevel, List<ViewPlan> floorPlans, List<ViewPlan> ceilingPlans)
        {
            // Get the name of the host level
            string hostLevelName = hostLevel.Name;

            // Iterate over the floor plans to find the matching view plan
            foreach (ViewPlan floorPlan in floorPlans)
            {
                // Get the level associated with the floor plan
                Level floorPlanLevel = doc.GetElement(floorPlan.LevelId) as Level;

                // Check if the floor plan level is linked to a level with a matching name in the host model
                if (floorPlanLevel != null && floorPlanLevel.Name == hostLevelName)
                {
                    return floorPlan;
                }
            }

            // Iterate over the ceiling plans to find the matching view plan
            foreach (ViewPlan ceilingPlan in ceilingPlans)
            {
                // Get the level associated with the ceiling plan
                Level ceilingPlanLevel = doc.GetElement(ceilingPlan.LevelId) as Level;

                // Check if the ceiling plan level is linked to a level with a matching name in the host model
                if (ceilingPlanLevel != null && ceilingPlanLevel.Name == hostLevelName)
                {
                    return ceilingPlan;
                }
            }

            return null; // Return null if no matching view plan is found
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
            LocationPoint roomLocation = room.Location as LocationPoint;

            if (roomLocation != null)
            {
                XYZ roomPoint = roomLocation.Point;
                return roomPoint;
            }
            else
            {
                throw new InvalidOperationException("Room location is not a point.");
            }
        }
    }
}

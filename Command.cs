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

            // Get active view
            View activeView = uidoc.ActiveView;

            // Filter rooms in the active view
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc, activeView.Id)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType();

            foreach (Element elem in roomCollector)
            {
                if (elem is Room room)
                {
                    try
                    {
                        // Get the room centroid
                        XYZ roomCentroid = GetRoomCentroid(room);
                        Debug.Print("Room Centroid: " + roomCentroid.ToString());

                        // Tag the room in the active view
                        using (Transaction transaction = new Transaction(doc))
                        {
                            transaction.Start("Tag Room");
                            doc.Create.NewRoomTag(activeView, roomCentroid, false);
                            transaction.Commit();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.Print("An error occurred: " + ex.Message);
                    }
                }
            }

            // Get level of active view
            Level activeLevel = activeView.GenLevel;

            // Get all floor plans and reflected ceiling plans with the same level as the active view
            List<ViewPlan> floorPlans = GetFloorPlans(doc, activeLevel.Id);
            List<ViewPlan> ceilingPlans = GetCeilingPlans(doc, activeLevel.Id);

            foreach (ViewPlan floorPlan in floorPlans)
            {
                CopyRoomTags(activeView, floorPlan, doc);
            }

            foreach (ViewPlan ceilingPlan in ceilingPlans)
            {
                CopyRoomTags(activeView, ceilingPlan, doc);
            }

            return Result.Succeeded;
        }

        private void CopyRoomTags(View sourceView, View targetView, Document doc)
        {
            // Filter room tags in the source view
            FilteredElementCollector tagCollector = new FilteredElementCollector(doc, sourceView.Id)
                .OfCategory(BuiltInCategory.OST_RoomTags)
                .WhereElementIsNotElementType();

            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("Copy Room Tags");
                foreach (Element elem in tagCollector)
                {
                    if (elem is RoomTag roomTag)
                    {
                        // Copy room tags to the target view
                        doc.Create.NewRoomTag(targetView, roomTag.TagHeadPosition, false);
                    }
                }
                transaction.Commit();
            }
        }

        private List<ViewPlan> GetFloorPlans(Document doc, ElementId levelId)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(v => v.ViewType == ViewType.FloorPlan && v.GenLevel.Id == levelId)
                .ToList();
        }

        private List<ViewPlan> GetCeilingPlans(Document doc, ElementId levelId)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(ViewPlan))
                .Cast<ViewPlan>()
                .Where(v => v.ViewType == ViewType.CeilingPlan && v.GenLevel.Id == levelId)
                .ToList();
        }

        private XYZ GetRoomCentroid(Room room)
        {
            LocationPoint roomLocation = room.Location as LocationPoint;

            if (roomLocation != null)
            {
                return roomLocation.Point;
            }
            else
            {
                throw new InvalidOperationException("Room location is not a point.");
            }
        }
    }
}

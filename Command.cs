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

                    // Filter rooms in the linked document
                    FilteredElementCollector roomCollector = new FilteredElementCollector(linkedDoc)
                        .OfCategory(BuiltInCategory.OST_Rooms)
                        .WhereElementIsNotElementType();

                    // Get all floor plans and reflected ceiling plans
                    List<ViewPlan> floorPlans = GetFloorPlans(doc);
                    List<ViewPlan> ceilingPlans = GetCeilingPlans(doc);

                    foreach (ViewPlan floorPlan in floorPlans)
                    {
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

                                    // Tag the room in the floor plan view
                                    using (Transaction transaction = new Transaction(doc))
                                    {
                                        transaction.Start("Tag Room");
                                        doc.Create.NewRoomTag(new LinkElementId(linkInstance.Id), roomCentroidUV, floorPlan.Id);
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

                    foreach (ViewPlan ceilingPlan in ceilingPlans)
                    {
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

                                    // Tag the room in the ceiling plan view
                                    using (Transaction transaction = new Transaction(doc))
                                    {
                                        transaction.Start("Tag Room");
                                        doc.Create.NewRoomTag(new LinkElementId(linkInstance.Id), roomCentroidUV, ceilingPlan.Id);
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

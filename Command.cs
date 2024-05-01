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

            // Filter rooms by matching levels
            var filteredRoomsByLevel = new FilteredElementCollector(firstArchDoc)
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

            // Get the phases associated with the plans
            var planPhases = allPlans.Select(plan =>
            {
                var phaseParameterValue = plan.LookupParameter("Phase Created");
                return phaseParameterValue != null ? phaseParameterValue.AsString() : null;
            }).Where(phaseName => !string.IsNullOrEmpty(phaseName)).ToList();


            //phase id instead - phase ID integer

            // Filter the rooms to only include rooms with the same phase as any of the plan phases
            var filteredRooms = new FilteredElementCollector(firstArchDoc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Where(room =>
                {
                    var roomPhaseParam = room.get_Parameter(BuiltInParameter.ROOM_PHASE);
                    if (roomPhaseParam != null)
                    {
                        var roomPhaseName = roomPhaseParam.AsString();
                        return !string.IsNullOrEmpty(roomPhaseName);
                    }
                    return false;
                })
                .ToList();

            //message
            var roomCenters = string.Empty;
            //start transaction for placing room tags
            Transaction transaction = new Transaction(doc, "Place Room Tags");
            transaction.Start();

            List<ElementId> taggedRoomIds = new List<ElementId>(); // Declaration of taggedRoomIds list

            // Iterate over the plans
            try
            {
                foreach (ViewPlan plan in allPlans)
                {
                    // Get the name of the level associated with the plan
                    string planLevelName = plan.LookupParameter("Associated Level")?.AsString();
                    string planPhaseName = plan.get_Parameter(BuiltInParameter.VIEW_PHASE)?.AsString();

                    Debug.Print($"Plan: {plan.Name}, Plan Level: {planLevelName}, Plan Phase: {planPhaseName}");

                    // Filter the rooms to only include rooms on the current plan from the linked model
                    var roomsOnPlan = filteredRoomsByLevel
                        .Where(room =>
                        {
                            var roomLevelParam = room.get_Parameter(BuiltInParameter.LEVEL_NAME);
                            var roomPhaseParam = room.get_Parameter(BuiltInParameter.ROOM_PHASE);

                            // Check if both room level and phase parameters exist
                            if (roomLevelParam != null && roomPhaseParam != null)
                            {
                                var roomLevelName = roomLevelParam.AsString();
                                var roomPhaseName = roomPhaseParam.AsString();

                                Debug.Print($"Room Level: {roomLevelName}, Room Phase: {roomPhaseName}");

                                // Check if room level and phase names are not null or empty
                                if (!string.IsNullOrEmpty(roomLevelName) && !string.IsNullOrEmpty(roomPhaseName))
                                {
                                    // Check if both room phase and plan phase names match
                                    if (planPhaseName != null && roomPhaseName == planPhaseName)
                                    {
                                        // Use the planLevelName variable from the outer scope
                                        return roomLevelName == planLevelName;
                                    }
                                }
                            }

                            // If any parameter is null or empty, or if phase names don't match, exclude the room
                            return false;
                        })
                        .ToList();

                    // Only proceed if there are rooms on the current plan
                    if (roomsOnPlan.Any())
                    {
                        foreach (Element element in roomsOnPlan)
                        {
                            Room room = element as Room;
                            if (room != null && room.Location is LocationPoint locationPoint)
                            {
                                LinkElementId linkedRoom = new LinkElementId(firstArchLink.Id, room.Id);
                                var roomLocation = locationPoint.Point;
                                var modifiedRoomLocation = firstArchLink.GetTransform().Origin + roomLocation;
                                doc.Create.NewRoomTag(linkedRoom, new UV(modifiedRoomLocation.X, modifiedRoomLocation.Y), plan.Id);
                            }
                            else if (room == null)
                            {
                                Debug.Print("Element is not a Room.");
                            }
                            else
                            {
                                Debug.Print($"Room {room.Id} on plan {plan.Name} is already tagged.");
                            }
                        }
                    }
                    else
                    {
                        Debug.Print($"No rooms found on plan {plan.Name}. Skipping tagging.");
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
    }
}

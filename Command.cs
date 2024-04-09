using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
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

            // Get all linked models
            var revitLinks = new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();

            // Get the first linked model
            var firstLink = revitLinks.FirstOrDefault();
            if (firstLink == null)
            {
                TaskDialog.Show("Error", "No linked room instance found.");
                return Result.Failed;
            }

            // Get the link document from the linked model
            var linkDoc = firstLink.GetLinkDocument();
            if (linkDoc == null)
            {
                TaskDialog.Show("Error", "Failed to retrieve link document.");
                return Result.Failed;
            }

            // Retrieve rooms from the linked model document
            var linkedRooms = new FilteredElementCollector(linkDoc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .ToList();

            // Start transaction for placing room tags
            using (Transaction transaction = new Transaction(doc, "Place Room Tags"))
            {
                transaction.Start();

                // Iterate over the linked rooms
                foreach (Element e in linkedRooms)
                {
                    Room room = e as Room;
                    if (room != null)
                    {
                        // Get the room location point
                        var roomLocation = (room.Location as LocationPoint)?.Point;
                        if (roomLocation != null)
                        {
                            // Convert XYZ to UV
                            UV tagLocation = new UV(roomLocation.X, roomLocation.Y);

                            // Convert room Id to LinkElementId
                            LinkElementId roomId = new LinkElementId(room.Id);

                            // Place new room tag
                            doc.Create.NewRoomTag(roomId, tagLocation, uidoc.ActiveView.Id);
                            
                        }
                    }
                }

                // Commit transaction for placing room tags
                transaction.Commit();
                TaskDialog.Show("Success", "Transaction committed.");
            }

            return Result.Succeeded;
        }
    }
}

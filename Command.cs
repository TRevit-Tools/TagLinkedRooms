using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
using System.Collections;
using System.Collections.ObjectModel;
using Autodesk.Revit;

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
                .Cast<Room>()
                .ToList();

            // Start transaction for placing room tags
            using (Transaction transaction = new Transaction(doc, "Place Room Tags"))
            {
                transaction.Start();

                // Iterate over the linked rooms
                foreach (Room room in linkedRooms)
                {
                    // Get the room location point
                    LocationPoint roomLocation = room.Location as LocationPoint;
                    if (null != LocationPoint)
                    {
                        // Convert XYZ to UV
                        double u = roomLocation
                        UV tagLocation = new UV(roomLocation.X, roomLocation.Y);

                        // Place new room tag
                        uiapp.ActiveUIDocument.Document.Create.NewRoomTag(new LinkElementId(room.Id), tagLocation, uiapp.ActiveUIDocument.ActiveView.Id);
                      
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

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

            
            List<Room> m_rooms = new List<Room>();                   //a list to store all rooms in the project
            List<RoomTag> m_roomTags = new List<RoomTag>();          //a list to store all room tags in the project
            List<Room> m_roomsWithTag = new List<Room>();            //a list to store all rooms with tag
            List<Room> m_roomsWithoutTag = new List<Room>();         //a list to store all rooms without tag
           
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
         


            ElementFilter filterRoom = new ElementClassFilter(typeof(SpatialElement));
            FilteredElementCollector roomCollector = new FilteredElementCollector(uiapp.ActiveUIDocument.Document);
            roomCollector.WherePasses(filterRoom);

            foreach (Autodesk.Revit.DB.Element ee in roomCollector)
            {
                Room tmpRoom = ee as Room;
                if (null != tmpRoom)
                {
                    m_rooms.Add(tmpRoom);
                    continue;
                }
            }


            // Start transaction for placing room tags
            using (Transaction transaction = new Transaction(doc, "Place Room Tags"))
            {
                transaction.Start();

                foreach (Room tmpRoom in m_roomsWithoutTag)
                {
                    //get the location point of the room
                    LocationPoint locPoint = tmpRoom.Location as LocationPoint;
                    if (null != locPoint)
                    {
                        //create a instance of UV class
                        double u = locPoint.Point.X;
                        double v = locPoint.Point.Y;

                        UV point = uiapp.Application.Create.NewUV(u, v);

                        //create room tag

                        uiapp.ActiveUIDocument.Document.Create.NewRoomTag(new LinkElementId(tmpRoom.Id), point, null);
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

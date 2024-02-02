#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Markup;

#endregion

namespace Intermediate_Challenge_Module01
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //1. Get rooms
            FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
            roomCollector.OfCategory(BuiltInCategory.OST_Rooms);

            List<string> roomList = new List<string>();
            int counter = 0;

            foreach (Element room in roomCollector)
            {
                //get list of departments
                string name = GetParameterValueByName(room, "Department");
                roomList.Add(name);
                counter++;
            }
            List <string> uniqueRoomList = roomList.Distinct().ToList(); 
            
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create schedules");
                //.Create schedules

                foreach (String departmentName in uniqueRoomList)
                {
                    // create schedule
                    ElementId catId = new ElementId(BuiltInCategory.OST_Rooms);
                    ViewSchedule newSchedule = ViewSchedule.CreateSchedule(doc, catId);
                    newSchedule.Name = "Dept - " + departmentName;

                    //Get parameter for fields

                    Element roomInst = roomCollector.FirstElement();
                    Parameter roomNumber = roomInst.get_Parameter(BuiltInParameter.ROOM_NUMBER); 
                    Parameter roomName = roomInst.get_Parameter(BuiltInParameter.ROOM_NAME);
                    Parameter roomDepartment = roomInst.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT);
                    Parameter roomComments = roomInst.LookupParameter("Comments");
                    Parameter roomArea = roomInst.get_Parameter(BuiltInParameter.ROOM_AREA);
                    Parameter roomLevel = roomInst.LookupParameter("Level");

                    //Create fields
                    ScheduleField roomNumberField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomNumber.Id); 
                    ScheduleField roomNameField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomName.Id);
                    ScheduleField roomDepartmentField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomDepartment.Id);
                    ScheduleField roomCommentsField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomComments.Id);
                    ScheduleField roomAreaField = newSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomArea.Id);
                    ScheduleField roomLevelField = newSchedule.Definition.AddField(ScheduleFieldType.Instance, roomLevel.Id);

                    roomAreaField.DisplayType = ScheduleFieldDisplayType.Totals;
                    
                    //filter by department
                    ScheduleFilter departmentFilter = new ScheduleFilter(roomDepartmentField.FieldId, ScheduleFilterType.Equal, departmentName);
                    newSchedule.Definition.AddFilter(departmentFilter);

                    //group by level
                    ScheduleSortGroupField levelGroup = new ScheduleSortGroupField(roomLevelField.FieldId);
                    levelGroup.ShowBlankLine = true;
                    levelGroup.ShowHeader = true;
                    levelGroup.ShowFooter = true;
                    newSchedule.Definition.AddSortGroupField(levelGroup);

                    roomLevelField.IsHidden = true;

                    //sort by room name
                    ScheduleSortGroupField roomSort = new ScheduleSortGroupField(roomNameField.FieldId);
                    newSchedule.Definition.AddSortGroupField(roomSort);

                    //display area for each level group and total area
                    ScheduleSortGroupField areaGroup =  new ScheduleSortGroupField(roomAreaField.FieldId);
                    areaGroup.ShowBlankLine = true;
                    areaGroup.ShowFooter = true;

                    // total count for the department 
                    roomAreaField.Definition.ShowGrandTotal = true;
                    roomAreaField.Definition.ShowGrandTotalCount = true;
                    roomAreaField.Definition.ShowGrandTotalTitle = true;

                    newSchedule.Definition.ShowGrandTotal = true;
                    newSchedule.Definition.ShowGrandTotalTitle = true;
                    newSchedule.Definition.ShowGrandTotalCount = true;

                }



                t.Commit();
            }


            return Result.Succeeded;
        }

        private string GetParameterValueByName(Element element, string paramName)
        {
                
            IList<Parameter> paramList = element.GetParameters(paramName);
            Parameter myParam = paramList.First();

            return myParam.AsString();
        }

        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}

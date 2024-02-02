#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;

#endregion

namespace Intermediate_Challenge_Module01
{
    [Transaction(TransactionMode.Manual)]
    public class Command2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Your code goes here
            // create schedulecreate a schedule that lists all the departments. Name this schedule “All Departments”.
            // The schedule should list only the department and area. It should be sorted by department. The area
            // field should be set to display totals. Also, the schedule should not itemize each room. Lastly, show
            // the grand total and title.  
            using (Transaction t = new Transaction(doc))
            {
                t.Start("Create all departments schedule");
                ElementId cateId = new ElementId(BuiltInCategory.OST_Rooms);
                ViewSchedule newDeptSchedule = ViewSchedule.CreateSchedule(doc, cateId);
                newDeptSchedule.Name = "All Departments";

                FilteredElementCollector roomCollector = new FilteredElementCollector(doc);
                roomCollector.OfCategory(BuiltInCategory.OST_Rooms);

                Element roomDepInst = roomCollector.FirstElement();
                Parameter roomDep = roomDepInst.LookupParameter("Department");
                Parameter roomA = roomDepInst.LookupParameter("Area");


                ScheduleField roomDepField = newDeptSchedule.Definition.AddField(ScheduleFieldType.Instance, roomDep.Id);
                ScheduleField roomAField = newDeptSchedule.Definition.AddField(ScheduleFieldType.ViewBased, roomA.Id);

                ScheduleSortGroupField depSort = new ScheduleSortGroupField(roomDepField.FieldId);
                newDeptSchedule.Definition.AddSortGroupField(depSort);

                roomAField.DisplayType = ScheduleFieldDisplayType.Totals;
                newDeptSchedule.Definition.ShowGrandTotal = true;
                newDeptSchedule.Definition.ShowGrandTotalCount = true;
                newDeptSchedule.Definition.ShowGrandTotalTitle = true;

                roomDepField.Definition.IsItemized = false;

                t.Commit();
            }

            return Result.Succeeded;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand2";
            string buttonTitle = "Button 2";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 2");

            return myButtonData1.Data;
        }
    }
}

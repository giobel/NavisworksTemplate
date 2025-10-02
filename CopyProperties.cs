using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using swf = System.Windows.Forms;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using System.Runtime.CompilerServices;
using Autodesk.Navisworks.Api.Timeliner;

namespace LOR_FBA
{
    public class CopyProperties
    {
        public static int CopyTimeLinerProperties(params string[] parameters)
        {
            try
            {
                // current document (.NET)
                Document doc = Application.ActiveDocument;
                // current document (COM)
                InwOpState10 cdoc = ComApiBridge.State;
                // current selected items
                ModelItemCollection items = doc.CurrentSelection.SelectedItems;
                
                 DocumentTimeliner tlDoc = doc.Timeliner as DocumentTimeliner;

                //swf.MessageBox.Show(Application.Gui.MainWindow, items.Count.ToString());
                List<TimelinerTask> allSubTasks = new List<TimelinerTask>();
                foreach (TimelinerTask rootTask in tlDoc.Tasks)
                {
                    // Get all child tasks recursively
                    allSubTasks = GetAllChildTasks(rootTask);
                }

                if (items.Count > 0)
                    {

                        foreach (ModelItem item in items)
                        {

                            InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                            // Get item's PropertyCategoryCollection
                            InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                            string foundTasks = "";

                            foreach (TimelinerTask task in allSubTasks)
                            {
                                TimelinerSelection taskSel = task.Selection;

                                

                                if (taskSel == null || !taskSel.HasExplicitSelection)
                                    continue;

                                foreach (ModelItem taskItem in taskSel.ExplicitSelection)
                                {
                                    if (taskItem.Equals(item))
                                    {
                                        foundTasks += $"{task.DisplayName}\n";

                                        cpropcates.SetUserDefined(1, "TimeLiner +", "TimeLiner +", GetPropertiesFromTask(task));

                                        break;
                                    }
                                }
                            }
                        }
                    }
            }
            catch (Exception ex)
            {
                swf.MessageBox.Show(Application.Gui.MainWindow, ex.Message);
            }

            swf.MessageBox.Show(Application.Gui.MainWindow, "Done");

            return 0;
        }

        private static InwOaProperty AddProperty(InwOpState10 cdoc, string name, DateTime dt)
        {
            InwOaProperty prop = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
            prop.name = name;
            prop.UserName = name;

            var v = RuntimeHelpers.GetObjectValue(ToUtc(dt));
            prop.value = v;
            return prop;
        } 
        

        private static InwOaProperty AddProperty(InwOpState10 cdoc, string name, object dt)
        {
            InwOaProperty prop = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
            prop.name = name;
            prop.UserName = name;

            var v = dt;
            prop.value = v;
            return prop;
        }

        private static InwOaPropertyVec GetPropertiesFromTask(TimelinerTask task)
        {
            InwOpState10 cdoc = ComApiBridge.State;
            // a new propertycategory object
            InwOaPropertyVec category = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

            DateTime dt = DateTime.Now;  // or your target datetime

            category.Properties().Add(AddProperty(cdoc, "Dump Date", dt));

            category.Properties().Add(AddProperty(cdoc, "Tasks Count", 1));
            category.Properties().Add(AddProperty(cdoc, "ActivityID", task.User1));
            category.Properties().Add(AddProperty(cdoc, "Display Name", task.DisplayName));

            //swf.MessageBox.Show(task.PlannedStartDate);
            try{
                category.Properties().Add(AddProperty(cdoc, "Planned Start Date", task.PlannedStartDate));
                category.Properties().Add(AddProperty(cdoc, "Planned End Date", task.PlannedEndDate));
            }
            catch {}
            
            try{
                category.Properties().Add(AddProperty(cdoc, "Actual Start Date", task.ActualStartDate));
                category.Properties().Add(AddProperty(cdoc, "Actual End Date", task.ActualEndDate));
            }
            catch{}
            
            category.Properties().Add(AddProperty(cdoc, "Task Type", task.SimulationTaskTypeName));
            category.Properties().Add(AddProperty(cdoc, "Work Pack Number", ""));
            category.Properties().Add(AddProperty(cdoc, "Work Pack Id", ""));

            string taskName = Enum.GetName(typeof(TaskStatus), task.TaskStatus);
            category.Properties().Add(AddProperty(cdoc, "Task Status", taskName));
            
            return category;
        }


        public enum TaskStatus
        {
            None = 0,
            Same = 1,
            Before = 2,
            After = 3,
            EarlyStartLateFinish = 4,
            LateStartEarlyFinish = 5,
            EarlyStart = 6,
            LateFinish = 7,
            EarlyStartFinish = 8,
            LateStartFinish = 9,
            ActualOnly = 10,
            PlannedOnly = 11,
            EarlyFinish = 12,
            LateStart = 13
        }



        private static DateTime ToUtc(DateTime dt)
        {

                DateTime time = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second, dt.Millisecond, DateTimeKind.Local);
                return TimeZone.CurrentTimeZone.ToUniversalTime(time);
 
        }
    private static List<TimelinerTask> GetAllChildTasks(TimelinerTask task)
{
    List<TimelinerTask> result = new List<TimelinerTask>();

    foreach (TimelinerTask child in task.Children)
    {
        result.Add(child);

        // Recursively add grandchildren
        result.AddRange(GetAllChildTasks(child));
    }

    return result;
}

    }
}
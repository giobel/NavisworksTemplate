//------------------------------------------------------------------
// NavisWorks Sample code
//------------------------------------------------------------------

// (C) Copyright 2009 by Autodesk Inc.

// Permission to use, copy, modify, and distribute this software in
// object code form for any purpose and without fee is hereby granted,
// provided that the above copyright notice appears in all copies and
// that both that copyright notice and the limited warranty and
// restricted rights notice below appear in all supporting
// documentation.

// AUTODESK PROVIDES THIS PROGRAM "AS IS" AND WITH ALL FAULTS.
// AUTODESK SPECIFICALLY DISCLAIMS ANY IMPLIED WARRANTY OF
// MERCHANTABILITY OR FITNESS FOR A PARTICULAR USE.  AUTODESK
// DOES NOT WARRANT THAT THE OPERATION OF THE PROGRAM WILL BE
// UNINTERRUPTED OR ERROR FREE.
//------------------------------------------------------------------
//
// This is the 'Hello World' Sample illustrated in the User guide.
// It demonstrates the various aspects of a basic plugin.
//
//------------------------------------------------------------------
#region HelloWorld

using System;
using System.Collections.Generic;
using swf = System.Windows.Forms;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using Autodesk.Navisworks.Api.Timeliner;
using System.Runtime.CompilerServices;
using System.IO;
using System.Text;
using System.Linq;

namespace BasicPlugIn
{

    [PluginAttribute("Export.ATestPlugin",                   //Plugin name
                         "ADSK",                                       //4 character Developer ID or GUID
                         ToolTip = "Test.ATestPlugin tool tip",//The tooltip for the item in the ribbon
                         DisplayName = "Export")]          //Display name for the Plugin in the Ribbon

    public class Export : AddInPlugin                        //Derives from AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {

            // current document (.NET)
            Document doc = Application.ActiveDocument;
            
            ModelItemCollection selectedItems = doc.CurrentSelection.SelectedItems;

            InwOpState10 cdoc = ComApiBridge.State;

            var sb = new StringBuilder();


            // EXTRACT PROPERTY - DISABLED BECAUSE IS SLOW
            foreach (ModelItem item in selectedItems)
            {
                InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                // Get item's PropertyCategoryCollection
                InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                //string line = $"{item.InstanceGuid},";

                foreach (PropertyCategory category in item.PropertyCategories)
                {
                    if (category.DisplayName == "Element")
                        {
                            //System.Windows.Forms.MessageBox.Show(category.DisplayName);
                            foreach (DataProperty prop in category.Properties)
                            {
                            // if (prop.DisplayName == "LOR_UniqueID") // or use internal name
                            // {
                            string value = "";
                            try
                            {
                                var v = prop.Value;
                                // Check the data type and access the value accordingly
                                value = v.ToString();
        
                            }
                            catch (Exception ex)
                            {
                                value = ex.Message;
                            }
                            // Do something with value
                            //line += $"{prop.DisplayName}||{value},";
                            sb.AppendLine($"{item.InstanceGuid},{prop.DisplayName},{value}");


                                // System.Windows.Forms.MessageBox.Show(value);
                                // }
                            }
                        }
                }
                //remove last comma
                //line = line.Remove(line.Length - 1);
                //sb.AppendLine(line);

                InwOaPath comPath = ComApiBridge.ToInwOaPath(item);
                InwGUIPropertyNode2 propNode = (InwGUIPropertyNode2)ComApiBridge.State.GetGUIPropertyNode(comPath, true);

                //Get PropertyCategoryCollection data
                InwGUIAttributesColl propCol = cpropcates.GUIAttributes();


                InwOaPropertyVec newCategory = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                DateTime dt = DateTime.Now;  // or your target datetime

                newCategory.Properties().Add(AddProperty(cdoc, "Dump Date", dt));

                foreach (PropertyCategory category in item.PropertyCategories)
                {
                    if (category.DisplayName == "SRC_FBA")
                    {
                        //System.Windows.Forms.MessageBox.Show(category.DisplayName);
                        foreach (DataProperty prop in category.Properties)
                        {
                            // if (prop.DisplayName == "LOR_UniqueID") // or use internal name
                            // {
                            string value = prop.Value.ToDisplayString();
                            // Do something with value
                            sb.AppendLine($"{prop.DisplayName}, {value}");


                            newCategory.Properties().Add(AddProperty(cdoc, prop.DisplayName, value));
                            // System.Windows.Forms.MessageBox.Show(value);
                            // }
                        }
                    }
                }


                            try
                            {
                                cpropcates.SetUserDefined(1, "SRC FBA22", "SRC_FBA22", newCategory);
                            }
                            catch
                            {
                                cpropcates.SetUserDefined(0, "SRC FBA22", "SRC_FBA22", newCategory);
                            }

            }

            File.WriteAllText(@"C:\Temp\output.csv", sb.ToString());


            // Show the names in a message box
            System.Windows.Forms.MessageBox.Show("Done");
            return 0;
        }



        public InwOaPropertyVec GetPropertiesFromTask()
        {
            InwOpState10 cdoc = ComApiBridge.State;
            // a new propertycategory object
            InwOaPropertyVec category = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

            DateTime dt = DateTime.Now;  // or your target datetime

            category.Properties().Add(AddProperty(cdoc, "Dump Date", dt));

            return category;
        }


        private InwOaProperty AddProperty(InwOpState10 cdoc, string name, DateTime dt)
        {
            InwOaProperty prop = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
            prop.name = name;
            prop.UserName = name;

            var v = RuntimeHelpers.GetObjectValue(ToUtc(dt));
            prop.value = v;
            return prop;
        }


        private InwOaProperty AddProperty(InwOpState10 cdoc, string name, object dt)
        {
            InwOaProperty prop = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
            prop.name = name;
            prop.UserName = name;

            var v = dt;
            prop.value = v;
            return prop;
        }
        
                private DateTime ToUtc(DateTime dt)
        {

                DateTime time = new DateTime(dt.Year, dt.Month, dt.Day, dt.Hour, dt.Minute, dt.Second, dt.Millisecond, DateTimeKind.Local);
                return TimeZone.CurrentTimeZone.ToUniversalTime(time);
 
        }

    }
}
#endregion
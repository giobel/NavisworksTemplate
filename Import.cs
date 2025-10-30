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
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using Autodesk.Navisworks.Api.Timeliner;
using System.Runtime.CompilerServices;
using System.IO;
using System.Text;
using System.Linq;
using Microsoft.VisualBasic.FileIO;


namespace BasicPlugIn
{

    [PluginAttribute("Import.ATestPlugin",                   //Plugin name
                         "ADSK",                                       //4 character Developer ID or GUID
                         ToolTip = "Test.ATestPlugin tool tip",//The tooltip for the item in the ribbon
                         DisplayName = "Import")]          //Display name for the Plugin in the Ribbon

    public class Import : AddInPlugin                        //Derives from AddInPlugin
    {
        public override int Execute(params string[] parameters)
        {
            // current document (.NET)
            Document doc = Application.ActiveDocument;
            ModelItemEnumerableCollection moicc = doc.Models.CreateCollectionFromRootItems().DescendantsAndSelf;

            // Some guids
            // var targetGuid = new Guid("49f1f04b-4797-431a-84c3-0a78fdc54209");

            // IEnumerable<ModelItem> foundItems = moicc.WhereInstanceGuid(targetGuid);

            // foreach (ModelItem item in foundItems)
            // {
            //     swf.MessageBox.Show(item.InstanceGuid.ToString());

            // }


            ModelItemCollection selectedItems = doc.CurrentSelection.SelectedItems;

            InwOpState10 cdoc = ComApiBridge.State;



            string inputFile = @"C:\Temp\output.csv";

            var properties = new Dictionary<string, string>();


            foreach (var line in File.ReadAllLines(inputFile))
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Split by comma or tab
                var parts = line.Split(new[] { ',', '\t' }, 2, StringSplitOptions.None);
                if (parts.Length == 2)
                {
                    string key = parts[0].Trim();
                    string value = parts[1].Trim();
                    properties[key] = value;
                }
            }



            IEnumerable<ModelItem> foundItems = moicc.WhereInstanceGuid(targetGuid);

            // EXTRACT PROPERTY - DISABLED BECAUSE IS SLOW
            foreach (ModelItem item in selectedItems)
            {
                InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                // Get item's PropertyCategoryCollection
                InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                //Get PropertyCategoryCollection data
                InwGUIAttributesColl propCol = cpropcates.GUIAttributes();


                InwOaPropertyVec newCategory = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
                DateTime dt = DateTime.Now;  // or your target datetime

                newCategory.Properties().Add(AddProperty(cdoc, "Dump Date", dt));


                foreach (var kvp in properties)
                    {
                        newCategory.Properties().Add(AddProperty(cdoc, kvp.Key, kvp.Value));
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


            // Show the names in a message box
            System.Windows.Forms.MessageBox.Show("Done");
            return 0;
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
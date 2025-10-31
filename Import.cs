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

            DateTime dt = DateTime.Now;  // or your target datetime

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


            //ModelItemCollection selectedItems = doc.CurrentSelection.SelectedItems;

            InwOpState10 cdoc = ComApiBridge.State;



            string inputFile = @"C:\Temp\output.csv";

            List<ItemProperty> excelData = new List<ItemProperty>();

            foreach (var line in File.ReadAllLines(inputFile))
            {

                if (string.IsNullOrWhiteSpace(line))
                    continue;

                // Split by comma or tab
                var parts = line.Split(new[] { ',', '\t' }, 4, StringSplitOptions.None);

                if (parts.Length < 4) continue;

                string propType = parts[3].Trim();
                string propValue = parts[2].Trim();

                if (IsCsvQuoted(propValue))
                {
                    propValue = CsvUnquote(propValue);
                }

                var propValueObject = AssignObjectType(propValue, propType);

                excelData.Add(new ItemProperty
                {
                    Guid = parts[0].Trim(),
                    PropName = parts[1].Trim(),
                    PropValue = propValueObject,
                });

            }

            var grouped = excelData
                            .GroupBy(i => i.Guid)
                            .ToDictionary(
                                g => g.Key,
                                g => g.Select(x => new { x.PropName, x.PropValue }).ToList()
                            );


            foreach (var kvp in grouped)
            {
                string guid = kvp.Key;
                var props = kvp.Value;

                InwOaPropertyVec newCategory = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                newCategory.Properties().Add(AddProperty(cdoc, "Dump Date", dt));

                IEnumerable<ModelItem> foundItems = moicc.WhereInstanceGuid(new Guid(guid));

                foreach (var prop in props)
                {
                    //string name = prop.PropName;

                    // var parts = prop.PropValue.Split(new[] { ':' }, 2);
                    // string value = parts.Length == 2 ? parts[1] : prop.PropValue;

                    // add property to Navisworks item here
                    newCategory.Properties().Add(AddProperty(cdoc, prop.PropName, prop.PropValue));
                    //newCategory.Properties().Add(ApplyPropertyValueFromType(cdoc, prop.PropName, prop.PropValue, prop.PropType));
                }

                if (foundItems.Count() > 0)
                {
                    foreach (ModelItem item in foundItems)
                    {
                        InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                        // Get item's PropertyCategoryCollection
                        InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                        //Get PropertyCategoryCollection data
                        InwGUIAttributesColl propCol = cpropcates.GUIAttributes();



                        try
                        {
                            cpropcates.SetUserDefined(1, "Excel Data", "Excel Data", newCategory);
                        }
                        catch
                        {
                            cpropcates.SetUserDefined(0, "Excel Data", "Excel Data", newCategory);
                        }

                    }
                }
            }

            foreach (var data in excelData)
            {




            }




            //swf.MessageBox.Show(foundItems.Count().ToString());







            /*
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
            */

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


        private object AssignObjectType(string value, string type)
        {
            switch (type)
            {
                case "DisplayString":
                case "String":
                    return value;

                case "Double":
                case "AnyDouble":
                    if (double.TryParse(value, out double d))
                        return d;
                    return value;

                case "Int32":
                    if (int.TryParse(value, out int i32))
                        return i32;
                    return value;

                case "Boolean":
                    if (bool.TryParse(value, out bool b))
                        return b;
                    return value;

                // case "NamedConstant":
                //     // You saved the name, now map enum back
                //     var nc = FindNamedConstant(item, propName, value);
                //     if (nc != null)
                //         v.FromNamedConstant(nc);
                //     else
                //         v.FromString(value); // fallback
                //     break;

                case "DateTime":
                    if (DateTime.TryParse(value, out DateTime dt))
                        return dt;
                    return value;

                default:
                    return value;
            }
                }
private InwOaProperty  ApplyPropertyValueFromType(
    InwOpState10 cdoc,
    string propName,
    string value,
    string type)
{

            InwOaProperty prop = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);
            prop.name = propName;
            prop.UserName = propName;


    VariantData v = new VariantData();

            try
            {
                switch (type)
                {
                    case "DisplayString":
                    case "String":
                        v = VariantData.FromDisplayString(value);
                        break;

                    case "Double":
                    case "AnyDouble":
                        if (double.TryParse(value, out double d))
                            v = VariantData.FromDouble(d);
                        break;

                    case "Int32":
                        if (int.TryParse(value, out int i32))
                            v = VariantData.FromInt32(i32);
                        break;

                    case "Boolean":
                        if (bool.TryParse(value, out bool b))
                            v = VariantData.FromBoolean(b);
                        break;

                    // case "NamedConstant":
                    //     // You saved the name, now map enum back
                    //     var nc = FindNamedConstant(item, propName, value);
                    //     if (nc != null)
                    //         v.FromNamedConstant(nc);
                    //     else
                    //         v.FromString(value); // fallback
                    //     break;

                    case "DateTime":
                        if (DateTime.TryParse(value, out DateTime dt))
                            v = VariantData.FromDateTime(dt);
                        break;

                    default:
                        v = VariantData.FromDisplayString(value);
                        break;
                }

                // Write back
                prop.value = v;
                return prop;
    }
            catch
            {
                return null;
            }
}

        string CsvUnquote(string input)
{
    if (string.IsNullOrEmpty(input))
        return input;

    input = input.Trim();

    // If wrapped in quotes, remove outer quotes
    if (input.StartsWith("\"") && input.EndsWith("\""))
        input = input.Substring(1, input.Length - 2);

    // Replace doubled quotes ("") with actual quote (")
    input = input.Replace("\"\"", "\"");

    return input;
}
bool IsCsvQuoted(string value)
{
    return value.StartsWith("\"") && value.EndsWith("\"");
}
    }

    public class ItemProperty
    {
        public string Guid { get; set; }
        public string PropName { get; set; }
        public object PropValue { get; set; }
        public string PropType { get; set; }
}
}
#endregion
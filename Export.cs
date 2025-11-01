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
using System.Collections.ObjectModel;

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

            if (selectedItems.Count == 0)
            {
                swf.MessageBox.Show("Please select some items first.");
                return 0;
            }

            InwOpState10 cdoc = ComApiBridge.State;

            var sb = new StringBuilder();


            var categories = new HashSet<string>();

            // Loop through all the items
            foreach (var item in selectedItems)
            {
                // Access the item's property collections
                PropertyCategoryCollection propertyCategories = item.PropertyCategories;

                foreach (var category in propertyCategories)
                {
                    // Add each category to the HashSet (automatically handles duplicates)
                    categories.Add(category.DisplayName);
                }
            }

            ObservableCollection<string> observableCategoryList = new ObservableCollection<string>(categories);


            var form = new LOR_FBA.ExportPanel(observableCategoryList);


            form.ShowDialog();

            if (form.DialogResult == false)
            {
                return 0;
            }

            //swf.MessageBox.Show(form.SelectedCategory);

            // EXTRACT PROPERTY - DISABLED BECAUSE IS SLOW
            foreach (ModelItem item in selectedItems)
            {
                InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                // Get item's PropertyCategoryCollection
                InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                foreach (PropertyCategory category in item.PropertyCategories)
                {
                    if (form.SelectedCategories.Contains(category.DisplayName))
                    {
                        //System.Windows.Forms.MessageBox.Show(category.DisplayName);
                        foreach (DataProperty prop in category.Properties)
                        {

                            string value = GetPropertyValue(prop);

                            string propDisplayName = prop.DisplayName;

                            switch (prop.Value.DataType)
                            {
                                case VariantDataType.DoubleLength:
                                    propDisplayName += " (m)";
                                    break;
                                case VariantDataType.DoubleArea:
                                    propDisplayName += " (m²)";
                                    break;
                                case VariantDataType.DoubleVolume:
                                    propDisplayName += " (m³)";
                                    break;
                            }
                            
                            sb.AppendLine($"{item.InstanceGuid},{prop.Name},{propDisplayName},{value},{prop.Value.DataType}");


                            // System.Windows.Forms.MessageBox.Show(value);
                            // }
                        }
                    }
                }
                //remove last comma
                //line = line.Remove(line.Length - 1);
                //sb.AppendLine(line);

            //     InwOaPath comPath = ComApiBridge.ToInwOaPath(item);
            //     InwGUIPropertyNode2 propNode = (InwGUIPropertyNode2)ComApiBridge.State.GetGUIPropertyNode(comPath, true);

            //     //Get PropertyCategoryCollection data
            //     InwGUIAttributesColl propCol = cpropcates.GUIAttributes();


            //     InwOaPropertyVec newCategory = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);
            //     DateTime dt = DateTime.Now;  // or your target datetime

            //     newCategory.Properties().Add(AddProperty(cdoc, "Dump Date", dt));

            //     foreach (PropertyCategory category in item.PropertyCategories)
            //     {
            //         if (category.DisplayName == "SRC_FBA")
            //         {
            //             //System.Windows.Forms.MessageBox.Show(category.DisplayName);
            //             foreach (DataProperty prop in category.Properties)
            //             {
            //                 // if (prop.DisplayName == "LOR_UniqueID") // or use internal name
            //                 // {
            //                 string value = prop.Value.ToDisplayString();
            //                 // Do something with value
            //                 sb.AppendLine($"{prop.DisplayName}, {value}");


            //                 newCategory.Properties().Add(AddProperty(cdoc, prop.DisplayName, value));
            //                 // System.Windows.Forms.MessageBox.Show(value);
            //                 // }
            //             }
            //         }
            //     }


            //     try
            //     {
            //         cpropcates.SetUserDefined(1, "SRC FBA22", "SRC_FBA22", newCategory);
            //     }
            //     catch
            //     {
            //         cpropcates.SetUserDefined(0, "SRC FBA22", "SRC_FBA22", newCategory);
            //     }

            }

            File.WriteAllText(form.filePathTextBox.Text, sb.ToString());


            // Show the names in a message box
            System.Windows.Forms.MessageBox.Show("Done");
            return 0;
        }


        string GetPropertyValue(DataProperty prop)
        {
            if (prop == null || prop.Value == null)
                return "";

            var v = prop.Value;


            if (v.IsDisplayString)
                return v.ToDisplayString();

            if (v.DataType.ToString() == "DoubleVolume")
                return UnitConvert.Ft3ToMeters3(v.ToAnyDouble()).ToString();

            if (v.DataType.ToString() == "DoubleLength")
                return UnitConvert.FtToMeters(v.ToAnyDouble()).ToString();

            if (v.DataType.ToString() == "DoubleArea")
                return UnitConvert.Ft2ToMeters2(v.ToAnyDouble()).ToString();

            if (v.IsDouble)
                return v.ToDouble().ToString();

            if (v.IsAnyDouble)
                return v.ToAnyDouble().ToString();

            if (v.IsInt32)
                return v.ToInt32().ToString();

            if (v.IsBoolean)
                return v.ToBoolean().ToString();

            if (v.IsNamedConstant)
            {
                string nc = ExtractInsideParentheses(v.ToString());
                return $"{CsvQuote(nc)}";
            }

            if (v.IsDateTime)
                return v.ToDateTime().ToString("yyyy-MM-dd HH:mm:ss");

            // Fallback — safe catch-all
            return v.ToString();
        }

public static class UnitConvert
{
    private const double FtToM = 0.3048;
    private const double Ft2ToM2 = 0.09290304;
    private const double Ft3ToM3 = 0.028316846592;

    // === LENGTH ===
    public static double FtToMeters(double ft) => ft * FtToM;
    public static double MetersToFt(double m) => m / FtToM;

    // === AREA ===
    public static double Ft2ToMeters2(double ft2) => ft2 * Ft2ToM2;
    public static double Meters2ToFt2(double m2) => m2 / Ft2ToM2;

    // === VOLUME ===
    public static double Ft3ToMeters3(double ft3) => ft3 * Ft3ToM3;
    public static double Meters3ToFt3(double m3) => m3 / Ft3ToM3;
}

string GetPropertyValueAndType(DataProperty prop)
        {
            if (prop == null || prop.Value == null)
                return "";

            var v = prop.Value;

            if (v.IsDisplayString)
                return $"{v.ToDisplayString()},DisplayString";

            if (v.IsDouble)
                return $"{v.ToDouble()},Double";

            if (v.IsAnyDouble)
                return $"{v.ToAnyDouble()},AnyDouble";

            if (v.IsInt32)
                return $"{v.ToInt32()},Int32";

            if (v.IsBoolean)
                return $"{v.ToBoolean()},Boolean";

            if (v.IsNamedConstant)
            {
                //var nc = v.ToNamedConstant();
                string nc = ExtractInsideParentheses(v.ToString());
            return $"{CsvQuote(nc)},NamedConstant";
            }

            if (v.IsDateTime)
                return $"{v.ToDateTime():yyyy-MM-dd HH:mm:ss},DateTime";

            // Fallback — safe catch-all
            return $"{v.ToString()},Unknown";
        }

string ExtractInsideParentheses(string input)
{
    if (string.IsNullOrEmpty(input))
        return "";

    int start = input.IndexOf('(');
    int end = input.LastIndexOf(')');

    if (start < 0 || end < 0 || end <= start)
        return input;

    return input.Substring(start + 1, end - start - 1);
}
string CsvQuote(string value)
{
    if (value.Contains(",") || value.Contains("\""))
        return $"\"{value.Replace("\"", "\"\"")}\"";

    return value;
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
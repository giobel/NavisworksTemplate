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
using System.Text;

//Add two new namespaces
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using Autodesk.Navisworks.Api.Timeliner;


namespace BasicPlugIn
{
   [PluginAttribute("BasicPlugIn.ABasicPlugin",                   //Plugin name
                    "ADSK",                                       //4 character Developer ID or GUID
                    ToolTip = "BasicPlugIn.ABasicPlugin tool tip",//The tooltip for the item in the ribbon
                    DisplayName = "Add Property")]          //Display name for the Plugin in the Ribbon

   public class AddProperty : AddInPlugin                        //Derives from AddInPlugin
   {
      public override int Execute(params string[] parameters)
      {
            // current document (.NET)
            Document doc = Application.ActiveDocument;
            // current document (COM)
            InwOpState10 cdoc = ComApiBridge.State;

            // //timeliner
            // DocumentTimeliner timeliner = doc.GetTimeliner();

            // foreach (TimelinerTask task in timeliner.Tasks)
            // {
            //    swf.MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow, task.DisplayName);
            // }

            // current selected items
            ModelItemCollection items = doc.CurrentSelection.SelectedItems;

            if (items.Count > 0)
            {

               try{

                foreach (ModelItem item in items)
                {

                    // convert ModelItem to COM Path
                    InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                    // Get item's PropertyCategoryCollection
                    InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                     //Get PropertyCategoryCollection data
                    InwGUIAttributesColl propCol = cpropcates.GUIAttributes();

                    InwOaPropertyVec newCategory = (InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);


                        InwOaProperty newProp = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty,null, null);
                        newProp.UserName = "CustomLength";

                        double meters = 123.45;
                        VariantData vd = new VariantData(meters, VariantDataType.DoubleLength);
                     
                        newProp.value = VariantData.FromDisplayString("m");;

                        newCategory.Properties().Add(newProp);

                        // InwOaProperty unitProp = (InwOaProperty)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaProperty, null, null);

                        // unitProp.name = "CustomLength_Units";
                        // unitProp.value = VariantData.FromDisplayString("m");
                        //  newCategory.Properties().Add(unitProp);

                        cpropcates.SetUserDefined(1, "Excel Data", "Excel Data", newCategory);
                    


                }
               }
               catch(Exception ex){
                                          swf.MessageBox.Show(Application.Gui.MainWindow, ex.Message);

               }
            }



         swf.MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow, "Property added");
         return 0;
      }

 // add new property to existing category
        public InwOaPropertyVec AddNewPropertyToExtgCategory(InwGUIAttribute2 propertyCategory)
        {
            // COM state (document)
            InwOpState10 cdoc = ComApiBridge.State;
            // a new propertycategory object
            InwOaPropertyVec category = (InwOaPropertyVec)cdoc.ObjectFactory(
                nwEObjectType.eObjectType_nwOaPropertyVec, null, null);


            return category;
        }
      
   }
}
#endregion
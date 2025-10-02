using System;
using System.Collections.Generic;
using swf = System.Windows.Forms;
using System.Text;

//Add two new namespaces
using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;

namespace LOR_FBA
{

   public static class RenameProperties
   {
public static int RenamePropertiesFBA(){
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


                int counter = 0;

                foreach (ModelItem item in items)
                {

                    // convert ModelItem to COM Path
                    InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                    // Get item's PropertyCategoryCollection
                    InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                     //Get PropertyCategoryCollection data
                    InwGUIAttributesColl propCol = cpropcates.GUIAttributes();

                     foreach  (InwGUIAttribute2 i in propCol){
                        
                        //swf.MessageBox.Show(Application.Gui.MainWindow, i.ClassUserName);

                        //if (i.UserDefined && i.ClassUserName == "Cat Test"){

                        if (i.ClassUserName == "TimeLiner +"){
                            
                            //swf.MessageBox.Show(Application.Gui.MainWindow, i.ClassUserName);

                        
                            foreach (InwOaProperty property in i.Properties())
                                {
                                    if (property.UserName == "User1"){
                                        //swf.MessageBox.Show(Application.Gui.MainWindow, property.UserName);
                                        //property.name = "new_name";
                                        property.UserName = "ActivityID";
                                        counter ++;
                                    }
                                    //swf.MessageBox.Show(Application.Gui.MainWindow, property.name);
                                    if (property.UserName == "User2"){
                                        //swf.MessageBox.Show(Application.Gui.MainWindow, property.UserName);
                                        //property.name = "new_name";
                                        property.UserName = "Work Pack Number";
                                        counter ++;
                                    }
                                    else if (property.UserName == "User3"){
                                        //swf.MessageBox.Show(Application.Gui.MainWindow, property.UserName);
                                        //property.name = "new_name";
                                        property.UserName = "Work Pack Id";
                                        counter ++;
                                    }
                                    //property.value = "new value";
                                }

                        }


                     }
              


                }
               
               swf.MessageBox.Show(Application.Gui.MainWindow, counter.ToString() + " properties modified");
               
               }
               catch(Exception ex){
                    swf.MessageBox.Show(Application.Gui.MainWindow, ex.Message);
               }
            }

            return 0;
}
   }
}
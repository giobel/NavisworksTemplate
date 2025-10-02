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

   public static class AddCategory
   {
      public static int AddCategoryTab(params string[] parameters)
      {
            // current document (.NET)
            Document doc = Application.ActiveDocument;
            // current document (COM)
            InwOpState10 cdoc = ComApiBridge.State;
            // current selected items
            ModelItemCollection items = doc.CurrentSelection.SelectedItems;

            if (items.Count > 0)
            {
                foreach (ModelItem item in items)
                {

                    // convert ModelItem to COM Path
                    InwOaPath citem = (InwOaPath)ComApiBridge.ToInwOaPath(item);
                    // Get item's PropertyCategoryCollection
                    InwGUIPropertyNode2 cpropcates = (InwGUIPropertyNode2)cdoc.GetGUIPropertyNode(citem, true);

                     //Get PropertyCategoryCollection data
                    InwGUIAttributesColl propCol = cpropcates.GUIAttributes();

                     InwOaPropertyVec newcate =(InwOaPropertyVec)cdoc.ObjectFactory(nwEObjectType.eObjectType_nwOaPropertyVec, null, null);

                     cpropcates.SetUserDefined(0, "TimeLiner +", "TimeLiner +", newcate);
                    
                }

            }

         swf.MessageBox.Show(Autodesk.Navisworks.Api.Application.Gui.MainWindow, "Timeliner + Tab added");
         return 0;
      }


      
   }
}

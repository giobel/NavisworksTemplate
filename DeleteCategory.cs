using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using swf = System.Windows.Forms;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.DocumentParts;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;
using System.Runtime.CompilerServices;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace LOR_FBA
{
    public static class DeleteCategory
    {
        public static int DeleteCategoryTab(params string[] parameters)
        {
            try
            {
                // current document (.NET)
                Document doc = Application.ActiveDocument;
                // current document (COM)
                InwOpState10 cdoc = ComApiBridge.State;
                // current selected items
                ModelItemCollection items = doc.CurrentSelection.SelectedItems;

                //swf.MessageBox.Show(Application.Gui.MainWindow, items.Count.ToString());
                string categoryToDelete = "TimeLiner +";

                if (items.Count > 0)
                {

                    foreach (ModelItem item in items)
                    {

                        RemoveCategoryByName(item, categoryToDelete);
                    }
                }

                swf.MessageBox.Show(Application.Gui.MainWindow, "Timeliner + Tab Removed");
            }
            catch (Exception ex)
            {
                swf.MessageBox.Show(Application.Gui.MainWindow, ex.Message);
            }

            return 0;
        }
	private static void RemoveCategoryByName(ModelItem mi, string CategoryDisplayName)
	{
		InwGUIPropertyNode2 miNode = (InwGUIPropertyNode2)ComApiBridge.State.GetGUIPropertyNode(ComApiBridge.ToInwOaPath(mi), true);
		int userCategoryIndex = GetUserCategoryIndex(mi, CategoryDisplayName);
		if (userCategoryIndex > -1)
		{
			miNode.RemoveUserDefined(userCategoryIndex);
		}
	}

	private static int GetUserCategoryIndex(ModelItem mi, string CategoryDisplayName)
	{
		InwGUIPropertyNode2 miNode = (InwGUIPropertyNode2)ComApiBridge.State.GetGUIPropertyNode(ComApiBridge.ToInwOaPath(mi), true);
		int num = 0;
		foreach (InwGUIAttribute2 item in miNode.GUIAttributes())
		{
			InwGUIAttribute2 val2 = item;
			if (val2.UserDefined)
			{
				num = checked(num + 1);
				if (Operators.CompareString(val2.ClassUserName, CategoryDisplayName, TextCompare: false) == 0)
				{
					return num;
				}
			}
		}
		return -1;
	}
    }
}
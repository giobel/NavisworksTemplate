using System;
using System.Collections.Generic;
using swf = System.Windows.Forms;
using System.Text;

using Autodesk.Navisworks.Api;
using Autodesk.Navisworks.Api.Plugins;
using Autodesk.Navisworks.Api.Interop.ComApi;
using ComApiBridge = Autodesk.Navisworks.Api.ComApi.ComApiBridge;


namespace LOR_FBA
{
    [Plugin("LORFBAAddin","FBA",DisplayName ="LOR-FBA")]
    [RibbonLayout("LORFBAAddin.xaml")]
    [RibbonTab("ID_CustomTab_1",DisplayName ="LOR-FBA")]
    [Command("ID_Button_1", Icon = "add_16.png", LargeIcon = "add_32.png", ToolTip = "Add TimeLiner + Tab to selected items. This is necessary before copying properties from TimeLiner")]
    [Command("ID_Button_2", Icon = "delete_16.png", LargeIcon = "delete_32.png", ToolTip = "Remove TimeLiner + Tab from selected items.")]
    [Command("ID_Button_3", Icon = "3_16.png", LargeIcon = "3_32.png", ToolTip = "Copy TimeLiner properties in selected items.")]
    [Command("ID_Button_4", Icon = "2_16.png", LargeIcon = "2_16.png", ToolTip = "Rename FBA properties")]



    public class LORaddin : CommandHandlerPlugin
    {
        public override int ExecuteCommand(string name, params string[] parameters)
        {
            switch (name)
            {
                case "ID_Button_1":
                AddCategory.AddCategoryTab();
                    break;
                case "ID_Button_2":
                    DeleteCategory.DeleteCategoryTab();
                    break;
                case "ID_Button_3":
                    CopyProperties.CopyTimeLinerProperties();
                    break;
                case "ID_Button_4":
                    RenameProperties.RenamePropertiesFBA();
                    break;
            }
            return 0;
        }
   }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Navisworks.Api.Plugins;
namespace LOR_FBA
{
    [Plugin("LORFBAAddin","FBA",DisplayName ="LOR-FBA")]
    [RibbonLayout("LORFBAAddin.xaml")]
    [RibbonTab("ID_CustomTab_1",DisplayName ="LOR-FBA")]
    [Command("ID_Button_1",ToolTip = "Test")]
    public class LORaddin : CommandHandlerPlugin
    {
        public override int ExecuteCommand(string name, params string[] parameters)
        {
            switch (name)
            {
                case "ID_Button_1":
                    MessageBox.Show("Button 1");
                    break;
            }
            return 0;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace LOR_FBA
{
    /// <summary>
    /// Interaction logic for ExportPanel.xaml
    /// </summary>
    public partial class ExportPanel : Window
    {
        public string ExportFilePath {get; set; }

        public ObservableCollection<string> CategoryList { get; set; }
        public List<string> SelectedCategories { get; set; }
        public ExportPanel(ObservableCollection<string> categoryList)
        {
            InitializeComponent();
            this.DataContext = this;
            CategoryList = categoryList;
            ExportFilePath = "C:\\Temp\\navisExport.csv"; // Default path
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Clear previous selections
            SelectedCategories = new List<string>();

            

            // Add selected items to the list
            foreach (var item in categoryListBox.SelectedItems)
            {
                SelectedCategories.Add(item.ToString());
            }
            DialogResult = true;
        }
    }
}

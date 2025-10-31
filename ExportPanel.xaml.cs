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
        public ObservableCollection<string> CategoryList { get; set; }
        public string SelectedCategory { get; set; }
        public ExportPanel(ObservableCollection<string> categoryList)
        {
            InitializeComponent();
            this.DataContext = this;
            CategoryList = categoryList;
        }

        private void CategoryComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            SelectedCategory = CategoryComboBox.SelectedItem as string;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}

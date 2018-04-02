using System;
using System.Collections.Generic;
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

namespace spanish_nl_analyzer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void file_browse_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            //File extensions
            dlg.DefaultExt = ".txt";
            dlg.Filter = "*.docx|*.docx|*.doc|*.doc|*.txt|*.txt";

            //Display
            Nullable<bool> result = dlg.ShowDialog();

            //Process dialogue box results
            if (result == true)
            {
                //Open document
                string filename = dlg.FileName;
                file_contents.Text = filename;
            }
        }

        private void exit_menu_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}

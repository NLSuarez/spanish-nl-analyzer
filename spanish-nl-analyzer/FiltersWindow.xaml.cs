using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.IO.IsolatedStorage;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace spanish_nl_analyzer
{
    /// <summary>
    /// Interaction logic for FiltersWindow.xaml
    /// </summary>
    public partial class FiltersWindow : Window
    {
        public FiltersWindow()
        {
            InitializeComponent();
        }

        private void persistFilters()
        {
            IsolatedStorageFile storage = IsolatedStorageFile.GetUserStoreForDomain();
            using (IsolatedStorageFileStream stream = new IsolatedStorageFileStream("spanish-nl-analyzer-filters.txt", FileMode.Create, storage))
            using (StreamWriter writer = new StreamWriter(stream))
            {
                // Persist each application-scope property individually
                Dictionary<string, string> Filters = (Dictionary<string, string>)Application.Current.Properties["Filters"];
                foreach (KeyValuePair<string, string> entry in Filters)
                {
                    writer.WriteLine("{0},{1}", entry.Key, entry.Value);
                }
            }
        }

        private void Add_Word_Click(object sender, RoutedEventArgs e)
        {
            //Get words
            string words = new TextRange(newFilterWord.Document.ContentStart, newFilterWord.Document.ContentEnd).Text;
            //Update filters and properties
            Dictionary<string, string> Filters = (Dictionary<string, string>)Application.Current.Properties["Filters"];
            Regex rx = new Regex(@"['’\.]?[\p{L}\p{Nd}]+(?:\.[\p{L}\p{Nd}]+|-[\p{L}\p{Nd}]+|'[\p{L}\p{Nd}]+|’[\p{L}\p{Nd}]+)*['’]?");
            MatchCollection matches = rx.Matches(words);
            foreach (Match match in matches)
            {
                Filters.Add(match.Value, match.Value);
            }
            Application.Current.Properties["Filters"] = Filters;
            filterWordsList.ItemsSource = Filters.ToList();
            persistFilters();
        }

        private void Delete_Word_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}

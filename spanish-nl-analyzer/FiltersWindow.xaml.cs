using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.IO.IsolatedStorage;
using System.ComponentModel;
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
            try
            {
                Dictionary<string, string> Filters = (Dictionary<string, string>)Application.Current.Properties["Filters"];
                filterWordsList.ItemsSource = HelperFunctions.DictionaryToSortedFilterObjectsList(Filters);
                CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(filterWordsList.ItemsSource);
                view.SortDescriptions.Add(new SortDescription("Word", ListSortDirection.Ascending));
            } catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        private void persistFilters(Dictionary<string, string> Filters)
        {
            IsolatedStorageFile storage = IsolatedStorageFile.GetUserStoreForDomain();
            using (IsolatedStorageFileStream stream = new IsolatedStorageFileStream("spanish-nl-analyzer-filters.txt", FileMode.Create, storage))
            using (StreamWriter writer = new StreamWriter(stream))
            {
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
            //Create list to catch pre-existing words
            List<string> rejections = new List<string>(50); //Generous assumption of 50 rejections
            foreach (Match match in matches)
            {
				string word = match.Value.ToLower();//lowercase for case insensitivity
				if (!Filters.ContainsKey(word))
                    Filters.Add(word, word); 
                else
                    rejections.Add(word);
            }
            Application.Current.Properties["Filters"] = Filters;
            filterWordsList.ItemsSource = HelperFunctions.DictionaryToSortedFilterObjectsList(Filters); //Convert dictionary to filter objects for display
            persistFilters(Filters);
            if (rejections.Any())
            {
                displayRejections(rejections);
            }
        }

        private void Delete_Word_Click(object sender, RoutedEventArgs e)
        {
			//Update the view
			List<filterListItem> currentItems = (List<filterListItem>)filterWordsList.ItemsSource;
			currentItems.RemoveAll((filterListItem i) => i.IsSelected); //O(n)
			filterWordsList.ItemsSource = currentItems;
			ICollectionView view = CollectionViewSource.GetDefaultView(filterWordsList.ItemsSource);
			view.Refresh();
			//Convert to dictionary
			Dictionary<string, string> newItems = HelperFunctions.FilterObjectsListToDictionary(currentItems);
			//Update the application properties and persist
			Application.Current.Properties["Filters"] = newItems;
			persistFilters(newItems);
        }

        private void enter_new_word(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Return)
            {
                Add_Word_Click(this, new RoutedEventArgs());
            }
        }

        private void displayRejections(List<string> rejections)
        {
            string message = "The following words are already filtered and were ignored from your entry:\r\n";
            foreach (string rejection in rejections)
            {
                message += rejection + "\r\n";
            }
            MessageBox.Show(message, "Duplicate Filters", MessageBoxButton.OK, MessageBoxImage.Warning);
        }

        private void Select_All(object sender, RoutedEventArgs e)
        {
            foreach (filterListItem item in filterWordsList.Items)
            {
                item.IsSelected = true;
			}
        }

        private void Unselect_All(object sender, RoutedEventArgs e)
        {
            foreach (filterListItem item in filterWordsList.Items)
            {
                item.IsSelected = false;
            }
        }
    }
}

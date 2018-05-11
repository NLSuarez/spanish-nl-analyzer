using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace spanish_nl_analyzer
{
    class HelperFunctions
    {
		//O(N+C)
		public static Dictionary<string, string> FilterObjectsListToDictionary(List<filterListItem> filtersList)
		{
			int size = filtersList.Count(); //O(1)
			Dictionary<string, string> filtersDictionary = new Dictionary<string, string>(size);
			foreach (filterListItem fi in filtersList)
			{
				filtersDictionary.Add(fi.Word, fi.Word);
			}
			return filtersDictionary;
		}

		//O(NLog(N)+C)
		public static List<filterListItem> DictionaryToSortedFilterObjectsList(Dictionary<string, string> dict)
		{
			int size = dict.Count(); //O(1)
			List<filterListItem> filterObjects = new List<filterListItem>(size);
			foreach (KeyValuePair<string, string> kv in dict)
			{
				filterObjects.Add(new filterListItem(kv.Value));
			}
			return filterObjects.OrderBy(i => i.Word).ToList(); //O(N(Log(N))
		}
    }

    #region helper_objects
    public class filterListItem : INotifyPropertyChanged
    {
        private string word;
        private bool selected;
		public event PropertyChangedEventHandler PropertyChanged;

        public string Word
        {
            get {
                return word;
            }
            set
            {
                word = value;
				OnPropertyChanged("Word");
            }
        }

        public bool IsSelected
        {
            get
            {
                return selected;
            }
            set
            {
                selected = value;
				OnPropertyChanged("IsSelected");
            }
        }

		protected void OnPropertyChanged(string name)
		{
			PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
		}

        public filterListItem(string value, bool selected=false)
        {
            Word = value;
            IsSelected = selected;
        }
    }
    #endregion helper_objects
}

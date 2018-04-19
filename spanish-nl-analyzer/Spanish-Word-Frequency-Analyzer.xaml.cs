using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
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
        //System.Windows.Threading.DispatcherTimer _typingTimer;
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
            dlg.Filter = "*.docx, *.doc, *.txt|.docx;*.doc;*.txt|All Files|*.*";

            //Display
            Nullable<bool> result = dlg.ShowDialog();

            //Process dialogue box results
            if (result == true)
            {
                //Determine document extension and process accordingly
                string filename = dlg.FileName;
                string ext = System.IO.Path.GetExtension(filename);
                doc_name.Content = filename;
                if (ext == ".txt")
                {
                    //If plain text, just go ahead and put it in the box.
                    string text;
                    using (Stream stream = dlg.OpenFile())
                    {
                        StreamReader reader = new StreamReader(stream);
                        text = reader.ReadToEnd();
                    }
                    file_contents.Text = text;
                } else if (ext == ".docx")
                {

                } else if (ext == ".doc")
                {

                } else
                {
                    //Do nothing or say invalid.
                }
                
            }
        }

        /*
         * Function to exit application on clicking the exit option in the menu.
         */
        private void exit_menu_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void analyze_Click(object sender, RoutedEventArgs e)
        {
            //Disable the buttons and get the text.
            analyze_button.IsEnabled = false;
            save_button.IsEnabled = false;
            file_contents.IsEnabled = false;
            string inputText = file_contents.Text;
            /*
             * Input analysis and parse logic:
             * Before we can process, we'll need to parse the input.
             * 
             * Spanish has more delimiters than English. Some of them are special characters, so we need to pay extra attention to those language specifics. 
             * We will also need to ensure that empty values are ignored.
             * 
             * List of special characters:
             * space, tab, period, comma, colon, semi-colon, hyphen surrounded by spaces, parentheses, brackets, braces, ellipsis, 
             * question marks(both right side up and upside down varients), exclamation marks, quotation marks, dashes, angle quotes,
             * new lines
             * 
             * Where relevant, unicode variants have been included.
             */
            string[] delimiterTokens = { " ", "\t", ".", ",", ":", ";", " - ", "(", ")", "[", "]", "{", "}", "...", "?", "!", "\u00BF", "\u00A1", "'", "\"", "\u2018", "\u2019", "\u201C","\u201D", "\u2012", "\u2013", "\u2014", "\u2015", "\u2053", "\u00AB", "\u00BB", "<<", ">>", "\u2039", "\u203A", "<", ">", "\n", "\r","\r\n" };
            string[] words = inputText.Split(delimiterTokens, System.StringSplitOptions.RemoveEmptyEntries);
            /*
             * After getting words array, create a dictionary to contain them.
             * 
             * The dictionary will have integers as the values and strings as the key.
             */
            Dictionary<string, int> frequencyDict = new Dictionary<string, int>();
            foreach (string word in words)
            {
                string lowerWord = word.ToLower();
                if (frequencyDict.ContainsKey(lowerWord)) {
                    frequencyDict[lowerWord] += 1;
                } else
                {
                    frequencyDict.Add(lowerWord, 1);
                }
            }
            // With dictionary in hand, display list in results_box
            List<KeyValuePair<string, int>> results = frequencyDict.OrderByDescending(kv => kv.Value).ToList();
            individual_frequency_results_box.ItemsSource = results;
            // Reenable buttons
            analyze_button.IsEnabled = true;
            save_button.IsEnabled = true;
            file_contents.IsEnabled = true;
        }

        private void save_output_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();

            dlg.Filter = "text files (*.txt)|*.txt|All Files (*.*)|*.*";
            dlg.DefaultExt = ".txt";
            dlg.FileName = System.IO.Path.GetFileName((String)doc_name.Content) + " Frequency";
            dlg.RestoreDirectory = true;

            Nullable<bool> result = dlg.ShowDialog();

            if(result == true)
            {
                List<KeyValuePair<string, int>> input = (List<KeyValuePair<string, int>>)individual_frequency_results_box.ItemsSource;
                using (Stream output = dlg.OpenFile())
                {
                    using (StreamWriter wText = new StreamWriter(output))
                    {
                        wText.WriteLine("\r\n\"" + dlg.FileName + "\" Word Frequencies \r\n");
                        String currentFrequency;
                        wText.WriteLine("By word: \r\n");
                        foreach (KeyValuePair<string, int> item in input)
                        {
                            currentFrequency = item.Key + " : " + item.Value;
                            wText.WriteLine(currentFrequency);
                        }
                        wText.WriteLine("\r\n ###################### \r\n");
                        wText.Flush();
                    }    
                }
            }
        }
        /*
         * Timer to dispatch a command for analysis.
         * Originally, this function was meant to make sure that someone was finished typing before it processed. Based on conversations with the client, however,
         * that approach is no longer feasible based on the size of files he may pass through. To make sure that everything is completely done before analysis begins,
         * this function has been removed. However, I've preserved it in case I want to do something like this in the future.
         */

        //        private void file_contents_TextChanged(object sender, TextChangedEventArgs e)
        //        {
        //            if (_typingTimer == null)
        //            {
        //                _typingTimer = new System.Windows.Threading.DispatcherTimer();
        //                _typingTimer.Interval = TimeSpan.FromMilliseconds(2000);
        //                _typingTimer.Tick += new EventHandler(this.handleTypingTimerTimeout);
        //            }
        //            _typingTimer.Stop(); //Reset timer
        //            _typingTimer.Tag = (sender as TextBox).Text;
        //            _typingTimer.Start();
        //        }

        //        private void handleTypingTimerTimeout(object sender, EventArgs e)
        //        {
        //            var timer = sender as System.Windows.Threading.DispatcherTimer;
        //            if (timer == null)
        //            {
        //                return;
        //            }

        //            string textToAnalyze = timer.Tag.ToString();
        //            results_box.Text = textToAnalyze;
        //            timer.Stop(); //Stop the timer so we act only once per keystroke
        //        }
    }
}



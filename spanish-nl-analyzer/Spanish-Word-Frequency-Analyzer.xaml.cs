using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using WinForms = System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using NPOI.XWPF.UserModel;
using NPOI.XWPF.Extractor;
using NPOI.HWPF.UserModel;
using NPOI.HWPF.Extractor;

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


        /*
         * saveFileResults
         * 
         * When given the appropriate dictionaries and file path, save output as a .csv.
         * 
         * Note: Path should be the full path, including the filename. Make sure you plug in the right path for this function when calling it.
         */
        private void saveFileResults(string filePath, Dictionary<string, int> frequencyDict)
        {
            using (FileStream output = File.Open(filePath,FileMode.Create))
            {
                using (StreamWriter wText = new StreamWriter(output))
                {
                    wText.WriteLine("Word,Occurrences,");
                    String currentFrequency;
                    foreach (KeyValuePair<string, int> item in frequencyDict)
                    {
                        currentFrequency = item.Key + "," + item.Value;
                        wText.WriteLine(currentFrequency);
                    }
                    wText.WriteLine(",,");
                    wText.Flush();
                }
            }
        }

        /*
         * processFile
         * 
         * Converts a given file into a string and returns it 
         */
        private string processFile(string path)
        {
            string text;
            string ext = System.IO.Path.GetExtension(path);
            if (ext == ".txt")
            {
                //If plain text, just go ahead and return as is.
                using (Stream stream = File.Open(path, FileMode.Open))
                {
                    StreamReader reader = new StreamReader(stream);
                    text = reader.ReadToEnd();
                }
                return text;
            }
            else if (ext == ".docx")
            {
                XWPFDocument document;
                using (FileStream fs = new FileStream(path, FileMode.Open))
                {
                    document = new XWPFDocument(fs);
                    XWPFWordExtractor wordExtractor = new XWPFWordExtractor(document);
                    text = wordExtractor.Text;
                    return text;
                }
            }
            else if (ext == ".doc")
            {
                using (FileStream fs = new FileStream(path, FileMode.Open))
                {
                    WordExtractor wordExtractor = new WordExtractor(fs);
                    text = wordExtractor.Text;
                    return text;
                }
            }
            else
            {
                return "Invalid";
            }
        }

        /*
         * processDirectory
         * 
         * Recursive function used primarily when mass processing folders. This will take a source directory and a destination directory as paramaters.
         * 
         * Source directory: what directory to read
         * Destination directory: what directory to save the results to
         * 
         * Everything in the source directory that is relevant to the program will be iterated over recursively, subdirectories included.
         * 
         * Responsibility for analyzing the actual files is delegated to "processFile" for extracting the string, "individual_frequency_parse" for the
         * individual word frequency dictionary, and "saveFileResults" for save the appropriate csv file.
         */
        private void processDirectory(string sourcePath, string destinationPath)
        {
            string[] fileEntries = Directory.GetFiles(sourcePath);
            string extension;
            string newFilePath;
            string fileContents;
            Dictionary<string, int> individualFrequency;
            foreach (string fileName in fileEntries)
            {
                extension = System.IO.Path.GetExtension(fileName);
                if (extension == ".txt" || extension == ".doc" || extension == ".doc")
                {
                    //Convert to string
                    fileContents = processFile(fileName);
                    //Get individual word frequency dict
                    individualFrequency = individual_frequency_parse(fileContents);
                    //With dictionary in hand, call the save function.
                    newFilePath = System.IO.Path.Combine(destinationPath, System.IO.Path.GetFileName(fileName) + " Frequency Spreadsheet.csv");
                    saveFileResults(newFilePath, individualFrequency);
                }
            }

            string[] subdirectoryEntries = Directory.GetDirectories(sourcePath);
            string newDirectory;
            foreach (string subdirectory in subdirectoryEntries)
            {
                //For sub directory, get current source and append new directory name.
                newDirectory = System.IO.Path.Combine(destinationPath, System.IO.Path.GetDirectoryName(subdirectory));
                if (Directory.Exists(newDirectory))
                {
                    processDirectory(subdirectory, newDirectory);
                } else
                {
                    Directory.CreateDirectory(newDirectory);
                    processDirectory(subdirectory, newDirectory);
                }
            }
        }

        /*
         * individual_frequency_parse
         * 
         * When given a string, return a frequency dictionary for each word regardless of group.
         */
        private Dictionary<string, int> individual_frequency_parse(string input)
        {
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
             * new lines, back slashes, forward slashes.
             * 
             * Where relevant, unicode variants have been included.
             */
            string[] delimiterTokens = { " ", "\t", ".", ",", ":", ";", " - ", "--", "(", ")", "[", "]", "{", "}", "...",
                "\u2026", "?", "!", "\u00BF", "\u00A1", "'", "\"", "\u2018", "\u2019", "\u201C", "\u201D", "\u2012", "\u2013",
                "\u2014", "\u2015", "\u2053", "\u00AB", "\u00BB", "<<", ">>", "\u2039", "\u203A", "<", ">", "\n", "\r", "\r\n",
                "\\","/" };
            string[] words = input.Split(delimiterTokens, System.StringSplitOptions.RemoveEmptyEntries);
            /*
             * After getting words array, create a dictionary to contain them.
             * 
             * The dictionary will have integers as the values and strings as the key.
             */
            Dictionary<string, int> frequencyDict = new Dictionary<string, int>();
            foreach (string word in words)
            {
                string lowerWord = word.ToLower();
                if (frequencyDict.ContainsKey(lowerWord))
                {
                    frequencyDict[lowerWord] += 1;
                }
                else
                {
                    frequencyDict.Add(lowerWord, 1);
                }
            }
            return frequencyDict;
        }

        /*
         * file_browse_Click
         * 
         * Handle click event on open file menu.
         */
        private void file_browse_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();

            //File extensions
            dlg.DefaultExt = ".txt";
            dlg.Filter = "*.docx, *.doc, *.txt|*.docx;*.doc;*.txt";

            //Description
            dlg.Title = "Choose Your Text File";

            //Display
            Nullable<bool> result = dlg.ShowDialog();

            //Process dialogue box results
            if (result == true)
            {
                //Set window description for currently viewed document.
                string filename = System.IO.Path.GetFileName(dlg.FileName);
                doc_name.Content = filename;
                //Retrieve text from file and place in textbox.
                string text = processFile(dlg.FileName);
                file_contents.Text = text;
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
            Dictionary<string, int> frequencyDict = individual_frequency_parse(file_contents.Text);
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

            dlg.Filter = "spreadsheet files (*.csv)|*.csv";
            dlg.DefaultExt = ".csv";
            dlg.AddExtension = true;
            dlg.FileName = (String)doc_name.Content + " Frequency Spreadsheet";
            //Save to my documents by default
            string combinedPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),"Spanish Word Frequency Output Files", DateTime.Now.ToString("yyyy-MM-dd"));
            if (Directory.Exists(combinedPath))
            {
                //Set
                dlg.InitialDirectory = System.IO.Path.GetFullPath(combinedPath);
            } else
            {
                //Create
                System.IO.Directory.CreateDirectory(System.IO.Path.GetFullPath(combinedPath));
                //Set
                dlg.InitialDirectory = System.IO.Path.GetFullPath(combinedPath);
            }

            Nullable<bool> result = dlg.ShowDialog();

            if(result == true)
            {
                List<KeyValuePair<string, int>> input = (List<KeyValuePair<string, int>>)individual_frequency_results_box.ItemsSource;
                using (Stream output = dlg.OpenFile())
                {
                    using (StreamWriter wText = new StreamWriter(output))
                    {
                        wText.WriteLine("Word,Occurrences,");
                        String currentFrequency;
                        foreach (KeyValuePair<string, int> item in input)
                        {
                            currentFrequency = item.Key + "," + item.Value;
                            wText.WriteLine(currentFrequency);
                        }
                        wText.WriteLine(",,");
                        wText.Flush();
                    }    
                }
                //Open folder after save in explorer
                Process.Start(System.IO.Path.GetFullPath(combinedPath));
            }
        }

        private void folder_browse_Click(object sender, RoutedEventArgs e)
        {
            using(WinForms.FolderBrowserDialog dlg = new WinForms.FolderBrowserDialog())
            {
                dlg.Description = "Select The Folder Containing the Files You Want to Analyze";
                dlg.RootFolder = Environment.SpecialFolder.Desktop;
                WinForms.DialogResult result = dlg.ShowDialog();

                if (result == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
                {
                    //Create save folder in my documents
                    string combinedPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Spanish Word Frequency Output Files", DateTime.Now.ToString("yyyy-MM-dd"));
                    if (Directory.Exists(combinedPath))
                    {
                        //Pass path to process directory as second parameter
                        processDirectory(dlg.SelectedPath, combinedPath);
                    }
                    else
                    {
                        //Create destination directory
                        System.IO.Directory.CreateDirectory(System.IO.Path.GetFullPath(combinedPath));
                        //Pass path to process directory as second parameter
                        processDirectory(dlg.SelectedPath, combinedPath);
                    }
                    //Open folder after save in explorer
                    Process.Start(System.IO.Path.GetFullPath(combinedPath));
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



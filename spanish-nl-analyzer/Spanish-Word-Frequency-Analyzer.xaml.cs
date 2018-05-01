using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
using Xceed.Words.NET;
using System.Xml;
using System.Xml.Linq;
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
         * xmlToFrequencyText && xmlToFrequencyText
         * 
         * A trio of methods specifically for converting xml to frequency friendly text. DocX mindlessly parses lists without
         * concern for spacing. Therefore, I've created my own variant where I just use that library to get me the
         * xml and parse it in a way that my frequency reader can comprehend.
         */
        private string xmlToFrequencyText(XElement e)
        {
            StringBuilder sb = new StringBuilder();
            xmlToFrequencyTextRecursive(e, ref sb);
            return sb.ToString();
        }

        private void xmlToFrequencyTextRecursive(XElement Xml, ref StringBuilder sb)
        {
            //Convert current xml selection to a string
            string tempStr = ToText(Xml);
            //If it ends with a space, just append. If it doesn't, add the space and append.
            if (tempStr.EndsWith(" "))
            {
                sb.Append(tempStr);
            } else
            {
                sb.Append(tempStr + " ");
            }
            
            //Recursively iterate through the rest of the xml elements.
            if (Xml.HasElements)
                foreach (XElement e in Xml.Elements())
                    xmlToFrequencyTextRecursive(e, ref sb);
        }

        //Port of Xceed's ToText method.
        private string ToText(XElement e)
        {
            switch (e.Name.LocalName)
            {
                case "tab":
                    return "\t";
                case "br":
                    return "\n";
                case "t":
                    goto case "delText";
                case "delText":
                    {
                        if (e.Parent != null && e.Parent.Name.LocalName == "r")
                        {
                            XElement run = e.Parent;
                            var rPr = run.Elements().FirstOrDefault(a => a.Name.LocalName == "rPr");
                            if (rPr != null)
                            {
                                var caps = rPr.Elements().FirstOrDefault(a => a.Name.LocalName == "caps");

                                if (caps != null)
                                    return e.Value.ToUpper();
                            }
                        }

                        return e.Value;
                    }
                case "tr":
                    goto case "br";
                case "tc":
                    goto case "tab";
                default:
                    return "";
            }
        }

        /*
         * saveFileResults
         * 
         * When given the appropriate dictionaries and file path, save output as a .csv.
         * 
         * Note: Path should be the full path, including the filename. Make sure you plug in the right path for this function when calling it.
         */
        private void saveFileResults(string filePath, List<KeyValuePair<string, int>> frequencyDict)
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
                DocX doc = DocX.Load(path);
                text = xmlToFrequencyText(doc.Xml);
                return text;
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
         * Responsibility for analyzing the actual files is delegated to "processFile" for extracting the string, "descending_individual_frequency_parse" for the
         * individual word frequency dictionary, and "saveFileResults" for save the appropriate csv file.
         */
        private void processDirectory(string sourcePath, string destinationPath)
        {

            string[] fileEntries = System.IO.Directory.GetFiles(sourcePath);
            string extension;
            string newFilePath;
            List<KeyValuePair<string, int>> individualFrequency;
            foreach (string fileName in fileEntries)
            {
                extension = System.IO.Path.GetExtension(fileName).ToLower();
                if (extension == ".txt" || extension == ".doc" || extension == ".docx")
                {
                    //Get individual word frequency dict and sort into list
                    individualFrequency = descending_individual_frequency_parse(processFile(fileName));
                    //With list in hand, call the save function.
                    newFilePath = System.IO.Path.Combine(destinationPath, System.IO.Path.GetFileName(fileName) + " Frequency Spreadsheet.csv");
                    saveFileResults(newFilePath, individualFrequency);
                }
            }

            string[] subdirectoryEntries = System.IO.Directory.GetDirectories(sourcePath);
            string newDirectory;
            foreach (string subdirectory in subdirectoryEntries)
            {
                //For sub directory, get current source and append new directory name.
                newDirectory = System.IO.Path.Combine(destinationPath, new DirectoryInfo(subdirectory).Name);
                if (System.IO.Directory.Exists(newDirectory))
                {
                    processDirectory(subdirectory, newDirectory);
                } else
                {
                    System.IO.Directory.CreateDirectory(newDirectory);
                    processDirectory(subdirectory, newDirectory);
                }
            }
        }

        /*
         * descending_individual_frequency_parse
         * 
         * When given a string, return a frequency list for each word in descending order regardless of group.
         */
        private List<KeyValuePair<string, int>> descending_individual_frequency_parse(string input)
        {
            /*
             * Input analysis and parse logic:
             * Before we can process, we'll need to parse the input.
             * 
             * There are too many specific delimiters to keep track of, and a limited, regular pattern to normal words.
             * Therefore, we choose regex and just hunt for things we know are words.
             */

            //Regex rx = new Regex(@"['’]?[\w]+(?:-\w+|'\w+|’\w+)*['’]?"); //Optional apostrophe in the beginning, followed by one or more letters, followed by a sequence
            //of one or more instances of hyphens preceeding one or more letters or one of the apostrophes followed by letters, and then finally an optional apostrophe 
            //at the end.
            Regex rx = new Regex(@"['’\.]?[\p{L}\p{Nd}]+(?:\.[\p{L}\p{Nd}]+|-[\p{L}\p{Nd}]+|'[\p{L}\p{Nd}]+|’[\p{L}\p{Nd}]+)*['’]?"); 
            /*
             * Optional apostrophe, unicode or otherwise, or period followed by one or more letters and numbers, followed by a sequence of one or
             * more instances of a hyphen, period, or apostrophe preceeding anything in step 2, and then, finally, an optional apostrophe.
             * 
             * \p{L} (all unicode and ascii letters)
             * \p{Nd} (all numbers and decimal digits)
             * These patterns are .NET specific and not transferrable.
             */
            MatchCollection matches = rx.Matches(input);

            /*
             * After getting matches, create a dictionary to contain them.
             * 
             * The dictionary will have integers as the values and strings as the key.
             */
            Dictionary<string, int> frequencyDict = new Dictionary<string, int>();
            foreach (Match match in matches)
            {
                //string manipulations
                string lowerWord = match.Value.ToLower(); //Ignore case
                lowerWord = lowerWord.Replace('’', '\''); //Replace fancy quotes with regular quotes

                if (frequencyDict.ContainsKey(lowerWord))
                {
                    frequencyDict[lowerWord] += 1;
                }
                else
                {
                    frequencyDict.Add(lowerWord, 1);
                }
            }
            return frequencyDict.OrderByDescending(kv => kv.Value).ThenByDescending(kv => kv.Key).ToList();
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
                doc_name.Text = System.IO.Path.GetFileName(dlg.FileName);
                //Retrieve text from file and place in textbox.
                string text = processFile(dlg.FileName);
                file_contents.Document.Blocks.Clear();
                file_contents.Document.Blocks.Add(new System.Windows.Documents.Paragraph(new System.Windows.Documents.Run(text)));
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
            //Take text and parse it for a descending order frequency.
            string textToAnalyze = new TextRange(file_contents.Document.ContentStart, file_contents.Document.ContentEnd).Text;
            List<KeyValuePair<string, int>> results = descending_individual_frequency_parse(textToAnalyze);
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
            dlg.FileName = (String)doc_name.Text + " Frequency Spreadsheet";
            //Save to my documents by default
            string combinedPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),"Spanish Word Frequency Output Files", DateTime.Now.ToString("yyyy-MM-dd"));
            if (System.IO.Directory.Exists(combinedPath))
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
                WinForms.DialogResult result = dlg.ShowDialog();

                if (result == WinForms.DialogResult.OK && !string.IsNullOrWhiteSpace(dlg.SelectedPath))
                {
                    //Create save folder in my documents
                    string originalFolderName = new DirectoryInfo(dlg.SelectedPath).Name;
                    string combinedPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Spanish Word Frequency Output Files");
                    combinedPath = System.IO.Path.Combine(combinedPath, DateTime.Now.ToString("yyyy-MM-dd"));
                    combinedPath = System.IO.Path.Combine(combinedPath, originalFolderName);
                    if (System.IO.Directory.Exists(combinedPath))
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

        private void Window_PreviewMouseMove(object sender, MouseEventArgs e)
        {
            var element = Mouse.DirectlyOver as FrameworkElement;
            HoverToolTip = GetToolTip(element);
        }

        #region HoverToolTip Property
        public object HoverToolTip
        {
            get { return (object)GetValue(HoverToolTipProperty); }
            set { SetValue(HoverToolTipProperty, value); }
        }

        public static readonly DependencyProperty HoverToolTipProperty =
            DependencyProperty.Register(nameof(HoverToolTip), typeof(object), typeof(MainWindow),
                new PropertyMetadata(null));
        #endregion HoverToolTip Property

        protected static Object GetToolTip(FrameworkElement obj)
        {
            if (obj == null)
            {
                return null;
            }
            else if (obj.ToolTip != null)
            {
                return obj.ToolTip;
            }
            else
            {
                return GetToolTip(VisualTreeHelper.GetParent(obj) as FrameworkElement);
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



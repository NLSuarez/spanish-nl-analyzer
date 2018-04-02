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
        System.Windows.Threading.DispatcherTimer _typingTimer;
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

        private void exit_menu_Click(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void file_contents_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_typingTimer == null)
            {
                _typingTimer = new System.Windows.Threading.DispatcherTimer();
                _typingTimer.Interval = TimeSpan.FromMilliseconds(2000);
                _typingTimer.Tick += new EventHandler(this.handleTypingTimerTimeout);
            }
            _typingTimer.Stop(); //Reset timer
            _typingTimer.Tag = (sender as TextBox).Text;
            _typingTimer.Start();
        }

        private void handleTypingTimerTimeout(object sender, EventArgs e)
        {
            var timer = sender as System.Windows.Threading.DispatcherTimer;
            if (timer == null)
            {
                return;
            }

            string textToAnalyze = timer.Tag.ToString();
            results_box.Text = textToAnalyze;
            timer.Stop(); //Stop the timer so we act only once per keystroke
        }
    }
}

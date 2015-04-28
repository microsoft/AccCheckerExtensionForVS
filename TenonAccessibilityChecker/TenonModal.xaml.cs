using Microsoft.VisualStudio.PlatformUI;
using System;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Navigation;

namespace Microsoft.TenonAccessibilityChecker
{
    /// <summary>
    /// Interaction logic for TenonModal.xaml
    /// </summary>
    public sealed partial class TenonModal : DialogWindow
    {
        private string _providedTenonkey;
        private string _selectedLevel;
        private string _selectedCertainity;
        public static string RenderedContent { get; set; }
        public static bool InvokedFromBrowser { get; set; }

        public TenonModal()
        {
            InitializeComponent();

            //Get the TenonKey from stored in users's machine.
            TextBox1.Text = ReadTenonKey();
        }

        /// <summary>
        /// validte the specified tenon key
        /// </summary>
        /// <returns></returns>
        private bool IsValidTenonKey()
        {
            LoadingStack.Visibility = Visibility.Visible;

            if (string.IsNullOrEmpty(TextBox1.Text))
            {
                LoadingStack.Visibility = Visibility.Hidden;
                Spinner.Visibility = Visibility.Hidden;
                this.processingtext.Visibility = Visibility.Hidden;
                this.ErrorPlaceholder.Visibility = Visibility.Visible;
                this.ErrorPlaceholder.Text = contents.TenonKeyValidatioMessage;
                this.ErrorPlaceholder.Foreground = new SolidColorBrush(Colors.Red);
                return false;
            }
            return true;
        }

        /// <summary>
        /// set the user inputs
        /// </summary>
        private void SetUserInputs()
        {
            LoadingStack.Visibility = Visibility.Visible;
            Spinner.Visibility = Visibility.Visible;
            processingtext.Visibility = Visibility.Visible;
            ErrorPlaceholder.Visibility = Visibility.Hidden;
            _providedTenonkey = TextBox1.Text;
            _selectedCertainity = ComboBox1.Text;
            _selectedLevel = ComboBox2.Text;
            this.BtnValidate.IsEnabled = false;
        }

        /// <summary>
        /// Method invoked on click of validate button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnValidate_Click(object sender, RoutedEventArgs e)
        {
               //Tenon key mandatory validation
               if(!IsValidTenonKey()) return;

               //collect the user inputs
               SetUserInputs();

               //initate the process
               ProcessRequest();
        }

        private void ProcessRequest()
        {
            // Create a background task for your work
            var task = Task.Factory.StartNew(InvokeTenonService);

            // When it completes, have it hide (on the UI thread), Spinner element
            task.ContinueWith(t =>
                {
                    Spinner.Visibility = Visibility.Hidden;
                    this.BtnCancel.Content = contents.OKMessage;
                    this.BtnValidate.Visibility = Visibility.Hidden;
                    this.processingtext.Visibility = Visibility.Hidden;
                    this.ErrorPlaceholder.Visibility = Visibility.Visible;
                    this.ErrorPlaceholder.Text = TenonAccessibilityCheckerPackage.TenonErrorMessage;
                    this.ErrorPlaceholder.Foreground = TenonAccessibilityCheckerPackage.TenonStatusCode == contents.TenonApiResponseSuccess ? new SolidColorBrush(Colors.Green) : new SolidColorBrush(Colors.Red);

                    // Sets the focused element in focusScope1 focusScope1 is a StackPanel.
                    FocusManager.SetFocusedElement(ErrorPlaceholder, ErrorPlaceholder);

                    // Gets the focused element for focusScope 1
                    IInputElement focusedElement = FocusManager.GetFocusedElement(ErrorPlaceholder);
                    focusedElement.Focus();
                },
                  TaskScheduler.FromCurrentSynchronizationContext()
             );
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            //Close the window.
            Close();
        }

        /// <summary>
        /// Method which invokes the tenonapi service
        /// </summary>
        private void InvokeTenonService()
        {
                // Read the file contents
                var content = (InvokedFromBrowser ? RenderedContent : File.ReadAllText(TenonAccessibilityCheckerPackage.ItemFullPath));

                //Get json from tenon io
                var str = TenonAccessibilityCheckerPackage.GetJson(_providedTenonkey, content, _selectedCertainity, _selectedLevel);

                //Parse API response
                var errorcollection = TenonAccessibilityCheckerPackage.ParseJson(str);

                if (TenonAccessibilityCheckerPackage.TenonStatusCode == contents.TenonApiResponseSuccess)
                {
                    //write errors to visual studio error window
                    TenonAccessibilityCheckerPackage.WriteErrors(TenonAccessibilityCheckerPackage.ItemFullPath,
                        TenonAccessibilityCheckerPackage.Hierarchy, errorcollection);

                    //Store the tenon key in c drive of your machine
                    PersistTenonKey(_providedTenonkey);
                }
        }

        /// <summary>
        /// Read the Tenon Key stored in your locale file in C drive.
        /// </summary>
        /// <returns>string</returns>
        private static string ReadTenonKey()
        {
            string dataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            dataPath = Path.Combine(dataPath, "TenonAccessibilityChecker", "TenonKey.txt");
            return File.Exists(dataPath) ? File.ReadAllText(dataPath) : string.Empty;
        }

        /// <summary>
        /// Store the provided tenon Key in C Drive of your local machine.
        /// </summary>
        /// <param name="tenonkey"></param>
        private static void PersistTenonKey(string tenonkey)
        {
            string dataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            dataPath = Path.Combine(dataPath, "TenonAccessibilityChecker", "TenonKey.txt");

            // If the directory doesn't exist, create it.
            if (!Directory.Exists(Path.GetDirectoryName(dataPath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(dataPath));
            }

            //Write the tenon key to your local file
            File.WriteAllText(dataPath, tenonkey);
        }
        private void OnCloseCmdExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            this.Close();
        }

        private void Hyperlink_OnRequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
            e.Handled = true;
        }

    }
}

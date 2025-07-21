using ExceedanceFilterApp; // Use your backend class
using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;

namespace PhaseExceedanceFilterApp.Gui
{
    public partial class MainWindow : Window
    {
        // Timer for updating elapsed time in the log box
        private DispatcherTimer? timer;
        // Stopwatch to track elapsed time for the operation
        private Stopwatch? stopwatch;

        public MainWindow()
        {
            InitializeComponent();
            LogBox.Text = "Ready.\n";
        }

        // Handler for Browse Config button click
        private void BrowseConfig_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Select Configuration CSV",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                ConfigPathBox.Text = dlg.FileName;
            }
        }

        // Handler for Browse Data button click
        private void BrowseData_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Title = "Select Data CSV",
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
            };
            if (dlg.ShowDialog() == true)
            {
                DataPathBox.Text = dlg.FileName;
            }
        }

        // Handler for Exceedance button click
        // Handles Exceedance button click event with safe file handling and error reporting
        private async void Exceedance_Click(object sender, RoutedEventArgs e)
        {
            // Disable UI and show loader
            SetUiEnabled(false);
            LoaderBar.Visibility = Visibility.Visible;
            LogBox.Text = "Processing, please wait...\nTime elapsed: 0.0 s";

            // Get config and data paths from UI
            string configPath = ConfigPathBox.Text.Trim();
            string dataPath = DataPathBox.Text.Trim();

            // Check that required files exist
            if (!File.Exists(configPath))
            {
                LogBox.Text = "Config file not found.\n";
                LoaderBar.Visibility = Visibility.Collapsed;
                SetUiEnabled(true);
                return;
            }
            if (!File.Exists(dataPath))
            {
                LogBox.Text = "Data file not found.\n";
                LoaderBar.Visibility = Visibility.Collapsed;
                SetUiEnabled(true);
                return;
            }

            // Set output file path: always use Excel in same directory as data
            string outputFile = Path.Combine(
                Path.GetDirectoryName(dataPath) ?? Environment.CurrentDirectory,
                "FilteredOutput.xlsx"
            );

            // Prepare timer and stopwatch for elapsed time reporting
            stopwatch = Stopwatch.StartNew();
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromMilliseconds(200);
            timer.Tick += (s, args) =>
            {
                var logLines = LogBox.Text.Split('\n');
                if (logLines.Length > 0)
                    logLines[logLines.Length - 1] = $"Time elapsed: {stopwatch.Elapsed.TotalSeconds:F1} s";
                LogBox.Text = string.Join("\n", logLines);
                LogBox.CaretIndex = LogBox.Text.Length;
                LogBox.ScrollToEnd();
            };
            timer.Start();

            try
            {
                // Try to open file for writing or clear existing content
                bool canWrite = true;
                if (File.Exists(outputFile))
                {
                    try
                    {
                        // Try to open file for writing (truncate)
                        using (var fs = new FileStream(outputFile, FileMode.Truncate, FileAccess.Write, FileShare.None))
                        {
                            // File is now empty and ready for overwrite
                        }
                    }
                    catch (IOException)
                    {
                        // File is open in Excel or locked
                        canWrite = false;
                    }
                    catch (UnauthorizedAccessException)
                    {
                        canWrite = false;
                    }
                }

                if (!canWrite)
                {
                    throw new IOException("The output Excel file is open in another application. Please close it and try again.");
                }

                // Run the exceedance filter in a background thread
                await Task.Run(() =>
                {
                    ExceedanceFilterApp.Program.ExceedanceFilterToExcel(configPath, dataPath, outputFile);
                });

                stopwatch.Stop();
                timer.Stop();

                // Check if output file was created successfully
                if (!File.Exists(outputFile))
                {
                    LogBox.Text += $"\n❌ Output file was not created. Something went wrong.";
                }
                else
                {
                    LogBox.Text += $"\n✅ Done! Excel file created: {outputFile}";
                }
                LogBox.Text += $"\nTotal time taken: {stopwatch.Elapsed.TotalSeconds:F1} s";
            }
            catch (Exception ex)
            {
                stopwatch.Stop();
                timer.Stop();
                LogBox.Text += $"\n❌ Error: {ex.Message}";
                LogBox.Text += $"\nTotal time taken: {stopwatch.Elapsed.TotalSeconds:F1} s";
            }
            finally
            {
                LoaderBar.Visibility = Visibility.Collapsed;
                SetUiEnabled(true);
            }
        }

        // Helper method to enable or disable all input controls except loader
        private void SetUiEnabled(bool isEnabled)
        {
            ConfigPathBox.IsEnabled = isEnabled;
            DataPathBox.IsEnabled = isEnabled;
            PhaseBtn.IsEnabled = isEnabled;
            ExceedanceBtn.IsEnabled = isEnabled;
            SummaryBtn.IsEnabled = isEnabled;
            BrowseConfigBtn.IsEnabled = isEnabled;
            BrowseDataBtn.IsEnabled = isEnabled;
            LoaderBar.IsEnabled = true;
            // Optionally: disable/enable more controls here as needed
        }

        // Placeholder for Phase button
        private void Phase_Click(object sender, RoutedEventArgs e)
        {
            LogBox.Text = "Phase button pressed (not implemented yet).\n";
        }

        // Placeholder for Summary button
        private void Summary_Click(object sender, RoutedEventArgs e)
        {
            LogBox.Text = "Summary Report button pressed (not implemented yet).\n";
        }
    }
}

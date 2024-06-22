using Fluent;
using Microsoft.Win32;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using unvell.ReoGrid;
using unvell.ReoGrid.IO.OpenXML.Schema;
using AutoUpdaterDotNET;
using System.Text.RegularExpressions;

namespace Sheets
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : RibbonWindow
    {
        bool isDebug;

        public MainWindow()
        {
            InitializeComponent();
            string extendedUserName = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            string userName = Environment.UserName;
            user.Text = userName;
            isDebug = false;
        }

        private void showinsiderinfo(object sender, RoutedEventArgs e)
        {
            ToggleThemeTeachingTip1.IsOpen = true;
        }

        private async void Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                var value = File.ReadAllText(openFileDialog.FileName);
                value = Regex.Escape(value);
                await monaco.ExecuteScriptAsync($"editor.setValue('{value}');");
                UnsavedTextBlock.Visibility = Visibility.Collapsed;
                AppTitle.Text = openFileDialog.SafeFileName + " - PowerNotepad";
            }
        }

        private async void Save(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            if (saveFileDialog.ShowDialog() == true)
            {
                var value = await monaco.ExecuteScriptAsync("editor.getValue();");
                value = Regex.Unescape(value);
                value = value.Substring(1, value.Length - 2);
                File.WriteAllText(saveFileDialog.FileName, value);
                UnsavedTextBlock.Visibility = Visibility.Collapsed;
                AppTitle.Text = saveFileDialog.SafeFileName + " - PowerNotepad";
            }
        }

        private void update(object sender, RoutedEventArgs e)
        { 
            if (isDebug == true)
            {
                AutoUpdater.Start("https://raw.githubusercontent.com/jpbandroid/jpbOffice-Resources/main/Sheets/updateinfo_debug.xml");

            } else {
                AutoUpdater.Start("https://raw.githubusercontent.com/jpbandroid/jpbOffice-Resources/main/Sheets/updateinfo.xml");
            }
        }

        private void About(object sender, RoutedEventArgs e)
        {
            AboutWindow about = new AboutWindow();
            about.Show();
            about.Activate();
        }

        private void RibbonWindow_Loaded(object sender, RoutedEventArgs e)
        {
            this.monaco.Source =
new Uri(System.IO.Path.Combine(
System.AppDomain.CurrentDomain.BaseDirectory,
@"Monaco\index.html"));
        }
    }
}
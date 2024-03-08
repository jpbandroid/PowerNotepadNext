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
            isDebug = true;
        }

        private void showinsiderinfo(object sender, RoutedEventArgs e)
        {
            ToggleThemeTeachingTip1.IsOpen = true;
        }

        private void Open(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
            {
                sheet.Load(openFileDialog.FileName, unvell.ReoGrid.IO.FileFormat.Excel2007);
                UnsavedTextBlock.Visibility = Visibility.Collapsed;
                AppTitle.Text = openFileDialog.SafeFileName + " - Sheets";
            }
        }

        private void Save(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            if (saveFileDialog.ShowDialog() == true)
            {
                sheet.Save(saveFileDialog.FileName, unvell.ReoGrid.IO.FileFormat.Excel2007);
                UnsavedTextBlock.Visibility = Visibility.Collapsed;
                AppTitle.Text = saveFileDialog.SafeFileName + " - Sheets";
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
    }
}
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace JsonToExcel
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        internal MainDataContext DC => (MainDataContext)DataContext;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog()
            {
                Filter = $"{Properties.Resources.Excel_File}|*.xlsx",
                FileName = $"{System.IO.Path.GetFileNameWithoutExtension(DC.JsonPath)}.xlsx"
            };
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                try
                {
                    DC.Export(dialog.FileName);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message);
                }
            }
        }

        private void BtnSelectFile_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            if (dialog.ShowDialog(App.Current.MainWindow) == true)
            {
                DC.JsonPath = dialog.FileName;
            }
        }

        private void BtnExportFolder_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(DC.OutputPath))
                System.Diagnostics.Process.Start(DC.OutputPath);
        }
    }
}

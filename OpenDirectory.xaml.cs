using System;
using System.Diagnostics;
using System.Media;
using System.Threading;
using System.Windows;
using System.Windows.Input;

namespace ExcelPatcher
{
    /// <summary>
    /// Логика взаимодействия для OpenDirectory.xaml
    /// </summary>
    public partial class OpenDirectory : Window
    {
        readonly string DirrectoryPath;
        public OpenDirectory(string path)
        {
            DirrectoryPath = path;
            InitializeComponent();
        }

        private void Cancel_MouseEnter(object sender, MouseEventArgs e)
        {
            Sounds.PlaySound("select");
        }

        private void OK_MouseEnter(object sender, MouseEventArgs e)
        {
            Sounds.PlaySound("select");
        }

        private void Cancel_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Sounds.PlaySound("click");
            Close();
        }
        
        private void OK_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Sounds.PlaySound("click");
            Process.Start(DirrectoryPath);
            Close();
        }
    }
}

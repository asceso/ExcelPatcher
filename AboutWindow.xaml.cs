using System;
using System.Media;
using System.Threading;
using System.Windows;
using System.Windows.Input;

namespace ExcelPatcher
{
    /// <summary>
    /// Логика взаимодействия для AboutWindow.xaml
    /// </summary>
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
        }

        private void Image_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Sounds.PlaySound("click");
            Close();
        }

        private void Image_MouseEnter(object sender, MouseEventArgs e)
        {
            Sounds.PlaySound("select");
        }
    }
}

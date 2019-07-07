using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace ExcelPatcher
{
    /// <summary>
    /// Логика взаимодействия для AllertWindow.xaml
    /// </summary>
    public partial class AllertWindow : Window
    {
        public string DirPath { get; set; }
        public AllertWindow(string mess,string DirPath)
        {
            InitializeComponent();
            AllertText.Text = mess;
            this.DirPath = DirPath;
        }

        private void CloseAlert_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Sounds.PlaySound("click");
            Close();
        }

        private void CloseAlert_MouseEnter(object sender, MouseEventArgs e)
        {
            Sounds.PlaySound("select");
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            OpenDirectory od = new OpenDirectory(DirPath);
            od.Show();
        }
    }
}

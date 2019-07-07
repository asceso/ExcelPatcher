using Microsoft.Win32;
using System;
using System.IO;
using System.IO.Compression;
using System.Media;
using System.Windows;
using System.Windows.Input;
using System.Xml;
using System.Threading;

namespace ExcelPatcher
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void CloseApp_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            Sounds.PlaySound("click");
            Close();
        }
        private void WhatAbout_MouseEnter(object sender, MouseEventArgs e)
        {
            Sounds.PlaySound("select");
        }

        private void CloseApp_MouseEnter(object sender, MouseEventArgs e)
        {
            Sounds.PlaySound("select");
        }
        private void Button_MouseEnter(object sender, MouseEventArgs e)
        {
            Sounds.PlaySound("select");
        }
        private void WhatAbout_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            AboutWindow aw = new AboutWindow();
            Sounds.PlaySound("click");
            aw.Show();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            string FilePath=null;
            try
            {
                FilePath = GetPathToFile();
                Sounds.PlaySound("click");
                GetPathingExcel(FilePath);
            }
            catch (Exception ex)
            {
                string DirPath = FilePath.Remove(FilePath.LastIndexOf("\\"));
                AllertWindow aw = new AllertWindow(ex.Message,DirPath);
                aw.Show();
            }
            finally
            {
                Dispose(FilePath);
            }
        }

        string GetPathToFile()
        {
        string resultPath;
        OpenFileDialog openFile = new OpenFileDialog {Filter = "Excel files (*.xlsx)|*.xlsx"};
            if (openFile.ShowDialog() == false)
                return null;
            else
                resultPath = openFile.FileName;
            return resultPath;
        }

        void Dispose(string CurDir)
        {
            string NewPath = CurDir.Remove(CurDir.LastIndexOf("\\"));
            File.Delete(NewPath + "\\Copy.zip");
            Directory.Delete(NewPath + "\\ExcelPatcher", true);
        }

        void GetPathingExcel(string FilePath)
        {
            string NewPath = FilePath;
            string NewFileName = FilePath;
            //Получаем путь файла
            NewPath = NewPath.Remove(NewPath.LastIndexOf("\\"));
            //имя файла                                                   
            NewFileName = NewFileName.Remove(0, NewFileName.LastIndexOf("\\") + 1);
            NewFileName = NewFileName.Remove(NewFileName.LastIndexOf("."));
            //копируем Excel файл с расширением Zip
            File.Copy(FilePath, NewPath + "\\Copy.zip");
            //рабочая директория
            string workDirectory = NewPath + "\\ExcelPatcher";
            //извелечение архива
            ZipFile.ExtractToDirectory(NewPath + "\\Copy.zip", workDirectory);
            //Удаляем архив
            //File.Delete(NewPath + "\\Copy.zip");
            /****РАБОТА С ЛИСТАМИ EXCEL*****/
            string CurDir = workDirectory;
            CurDir += "\\xl\\worksheets";
            int CounterFiles = Directory.GetFiles(CurDir).Length;
            for (int i = 0; i < CounterFiles; i++)
            {
                string CurrentFile = CurDir + "\\sheet" + (i + 1) + ".xml";
                XmlDocument doc = new XmlDocument();
                doc.Load(CurrentFile);
                XmlNode xmlNode = doc.GetElementsByTagName("sheetProtection")[0];
                xmlNode.ParentNode.RemoveChild(xmlNode);
                doc.Save(CurrentFile);
            }
            /****************************/
            ZipFile.CreateFromDirectory(workDirectory, NewPath + "\\" + NewFileName + "(без защиты)" + ".xlsx");
            //Удаляем директорию с временными файлами
            //Directory.Delete(workDirectory, true);
            //Спрашиваем открытие директории
            OpenDirectory od = new OpenDirectory(NewPath);
            od.Show();
        }
    }
}

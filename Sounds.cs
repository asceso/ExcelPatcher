using System;
using System.Collections.Generic;
using System.Linq;
using System.Media;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace ExcelPatcher
{
    static class Sounds
    {
        static public void PlaySound(string NameOfSound)
        {
            try
            {
                switch (NameOfSound)
                {
                    case "select":
                        {
                            SoundPlayer player = new SoundPlayer(
                                Environment.CurrentDirectory +
                                "\\Resources\\" +
                                "select.wav");
                            player.Play();
                        }
                        break;
                    case "click":
                        {
                            SoundPlayer player = new SoundPlayer(
                                Environment.CurrentDirectory +
                                "\\Resources\\" +
                                "click.wav");
                            player.Play();
                            Thread.Sleep(300);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}

using Microsoft.Win32;
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
using System.Windows.Navigation;

using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Data;
using Syncfusion.XlsIO;
using System.Windows.Interop;

namespace Adressbuch
{
    /// <summary>
    /// Interaktionslogik für Kontaktebearbeiten.xaml
    /// </summary>
    public partial class Kontaktebearbeiten : UserControl
    {
        string bildJaNein = "0";
        public Kontaktebearbeiten()
        {
            InitializeComponent();
            UserKontakte userKontakte = new UserKontakte();
        }
        public BitmapSource kontaktBildImage { get; private set; }
        private void bildkontakt_MouseDown(object sender, MouseButtonEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
               

                if (vorname.Text != "" && name.Text != "")
                {
                   

                    if (openFileDialog.ShowDialog() == true)
                    {
                        bildkontakt.Source = null;
                        bildkontakt.Source = kontaktBildImage;
                        bildkontakt.Source = null;
                        Bitmap bitmap = new Bitmap(openFileDialog.FileName);
                        using (MemoryStream memory = new MemoryStream())
                        {
                            using (FileStream fs = new FileStream(Paths.GetFilePath( @"Bilder\") + vorname.Text + name.Text + ".png", FileMode.Create, FileAccess.ReadWrite))
                            { //speichern
                                bitmap.Save(memory, ImageFormat.Png);
                                byte[] bytes = memory.ToArray();
                                fs.Write(bytes, 0, bytes.Length);
                            }
                        }
                        kontaktBildImage = Imaging.CreateBitmapSourceFromHBitmap(bitmap.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
                        bildkontakt.Source = kontaktBildImage;


                        bildJaNein = "1";
                            
                           
                        
                    }
                }
                else
                {
                    MessageBox.Show("Bitte Vor- und Nachname eingeben");
                }

                workbook.Save();
                workbook.Close();


            }
        }








        private void baestaetigen_Click(object sender, RoutedEventArgs e)
        {
            if (vorname.Text != "" && name.Text != "")
            {
                int splateA = 2;
            using (ExcelEngine excelEngine = new ExcelEngine())
            {

                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //create a dokument

                IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
               
                while (true)
                {
                    if (worksheet.Range["A" + splateA].Text == vorname.Text && worksheet.Range["B" + splateA].Text == name.Text)
                    {
                        if (worksheet.Range["J"+splateA].Text == "1")
                        {
                            bildJaNein = "1";
                        }

                        worksheet.Range["A" + splateA].Text = vorname.Text;
                        worksheet.Range["B" + splateA].Text = name.Text;
                        worksheet.Range["K" + splateA].Text = datePicker.Text;
                        worksheet.Range["D" + splateA].Text = HinzufuegenStrasse.Text;
                        worksheet.Range["E" + splateA].Text = HinzufuegenHausnummer.Text;
                        worksheet.Range["F" + splateA].Text = HinzufuegenPostleizahl.Text;
                        worksheet.Range["G" + splateA].Text = HinzufuegenOrt.Text;
                        worksheet.Range["H" + splateA].Text = HinzufuegenTelefon.Text;
                        worksheet.Range["I" + splateA].Text = HinzufuegenEmail.Text;
                        string[] teilen = datePicker.Text.Split('.');
                        if (datePicker.Text == null)
                        {
                            worksheet.Range["C" + splateA].Text = teilen[0] + "." + teilen[1] + ".";
                        }
                        worksheet.Range["I" + splateA].Text = HinzufuegenEmail.Text;

                        if (bildJaNein == "1")
                        {
                            worksheet.Range["J" + splateA].Text = bildJaNein;
                        }
                        else
                        {
                            worksheet.Range["J" + splateA].Text = bildJaNein;
                        }
                        
                       

                        

                        MessageBox.Show("Kontakt erfolgreich Hinzugefügt, Bildänderungen werden erst Beim neustart Angezeigt");

                       

                        break;
                    }
                    else
                    {
                        splateA++;
                    }

                }



                // save

                workbook.Save();
                workbook.Close();

            }
            }
            else
            {
                MessageBox.Show("Bitte Vor- und Nachname eingeben");
            }
        }
       




        private void abbrechen_Click(object sender, RoutedEventArgs e)
        {
            UserKontakte userKontakte = new UserKontakte();
            Control.Content = userKontakte.Control.Content;
        }

        private void HinzufuegenHausnummer_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
           
        }

        private void CheckIsNumeric(TextCompositionEventArgs e)
        {
            int result;

            if (!(int.TryParse(e.Text, out result) || e.Text == "."))
            {
                e.Handled = true;
            }
        }

        private void HinzufuegenPostleizahl_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckIsNumeric(e);
        }

        private void HinzufuegenTelefon_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckIsNumeric(e);
        }
    }
}

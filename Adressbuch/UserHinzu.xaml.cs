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
using System.Windows.Shapes;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Data;
using Syncfusion.XlsIO;
namespace Adressbuch
{
    /// <summary>
    /// Interaktionslogik für UserHinzu.xaml
    /// </summary>
    public partial class UserHinzu : UserControl
    {
        string bildJaNein = "0";
       
        public UserHinzu()
        {
            InitializeComponent();
            datePicker.Text = "";

           
        }
        




        private void bildkontakt_MouseDown(object sender, MouseButtonEventArgs e)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];



                if (vorname.Text != "" && name.Text != "")
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    if (openFileDialog.ShowDialog() == true)
                    {
                        Uri fileUri = new Uri(openFileDialog.FileName);
                        bildkontakt.Source = new BitmapImage(fileUri);


                        Bitmap bitmap = new Bitmap(openFileDialog.FileName);
                        bitmap.Save(Paths.GetFilePath("Bilder\\") + vorname.Text + name.Text + ".png", ImageFormat.Png);
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
                    if (worksheet.Range["A" + splateA].Text == "" || worksheet.Range["A" + splateA].Text == null)
                    {

                        
                        worksheet.Range["A" + splateA].Text = vorname.Text;
                        worksheet.Range["B" + splateA].Text = name.Text;
                        
                        if (datePicker.SelectedDate != null)
                        {
                            worksheet.Range["K" + splateA].Text = datePicker.Text;
                        }
                        worksheet.Range["D" + splateA].Text = HinzufuegenStrasse.Text;
                        worksheet.Range["E" + splateA].Text = HinzufuegenHausnummer.Text;
                        worksheet.Range["F" + splateA].Text = HinzufuegenPostleizahl.Text;
                        worksheet.Range["G" + splateA].Text = HinzufuegenOrt.Text;
                        worksheet.Range["H" + splateA].Text = HinzufuegenTelefon.Text;
                        worksheet.Range["I" + splateA].Text = HinzufuegenEmail.Text;
                        
                        if (datePicker.SelectedDate != null)
                        {
                            string[] teilen = datePicker.Text.Split('.');

                            worksheet.Range["C" + splateA].Text = teilen[0] + "." + teilen[1] + ".";
                        }
                        worksheet.Range["I" + splateA].Text = HinzufuegenEmail.Text;

                        if(bildJaNein == "1")
                        {
                            worksheet.Range["J" + splateA].Text = bildJaNein;
                        }
                        else
                        {
                            worksheet.Range["J" + splateA].Text = bildJaNein;
                        }
                       


                        MessageBox.Show("Kontakt erfolgreich Hinzugefügt");

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

               vorname.Text ="";
               name.Text="";
               

               datePicker.Text="";

               HinzufuegenStrasse.Text="";
               HinzufuegenHausnummer.Text="";
               HinzufuegenPostleizahl.Text="";
               HinzufuegenOrt.Text="";
               HinzufuegenTelefon.Text="";
               HinzufuegenEmail.Text="";
                UserKontakte userKontakte = new UserKontakte();
                ControlUserHinzu.Content = userKontakte;

            }
            }
            else
            {
                MessageBox.Show("Bitte Vor- und Nachname eingeben");
            }
        }

        private void abbrechen_Click(object sender, RoutedEventArgs e)
        {
            MainWindow mainWindow = new MainWindow();
            ControlUserHinzu.Content = mainWindow.Kontakte.Content; 

        }

        private void bildkontakt_MouseEnter(object sender, MouseEventArgs e)
        {
            bildkontakt.Source = new BitmapImage(new Uri(Paths.GetFilePath( @"ressours\hinzu2.png"), UriKind.RelativeOrAbsolute));

        }

        private void bildkontakt_MouseLeave(object sender, MouseEventArgs e)
        {
            bildkontakt.Source = new BitmapImage(new Uri(Paths.GetFilePath( @"ressours\hinzu.png"), UriKind.RelativeOrAbsolute));

        }

        private void HinzufuegenPostleizahl_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckIsNumeric(e);
        }
        private void CheckIsNumeric(TextCompositionEventArgs e)
        {
            int result;

            if (!(int.TryParse(e.Text, out result) || e.Text == "."))
            {
                e.Handled = true;
            }
        }

        private void HinzufuegenHausnummer_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
           
        }

        private void HinzufuegenTelefon_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            CheckIsNumeric(e);
        }
    }
}

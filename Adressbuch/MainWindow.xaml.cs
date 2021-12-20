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

using Syncfusion.XlsIO;
using System.Drawing.Imaging;
using Microsoft.Win32;
using System.Drawing;
using System.IO;
using System.Windows.Interop;

namespace Adressbuch
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    
    public partial class MainWindow : Window
    {

       

        public MainWindow()
        {

            InitializeComponent();
            FotoLoeschen fotoLoeschen = new FotoLoeschen();
            fotoLoeschen.OrdnerClear();
          
            
            Startseite startseite = new Startseite();
            Kontakte.Content = startseite;

            AlterBerechnen alter = new AlterBerechnen();
            alter.Berechnen();

             ListeGeburstage();
           
            list.Clear();
            foreach (string str in lboxKontakte.Items)
            {
                list.Add(str);
                
            }
        }

        private void ListeGeburstage()
        {
                lboxKontakte.Items.Clear();

                List<GeburtstagsListe> liste2 = GeburtstagsListe.CreateList();

                liste2 = liste2.OrderBy(x => x.Alter).ToList();
               
                int counter = 0;
                int count2 = 2;
                foreach (GeburtstagsListe i2 in liste2)
                {
                    if ( i2.Alter == -1)
                    {
                       
                    }
                    else
                    {
                        lboxKontakte.Items.Add(i2.Vorname + " " + i2.Name + " " + i2.date);
                        counter++;
                    }
                    count2++;
                    if (counter > 4)
                    { break; }

                }
                gebnext.Content = lboxKontakte.Items[0];
               
        }
         
        

        
        private void ListeNamen()
        {
            lboxKontakte.Items.Clear();

            List<GeburtstagsListe> listen = GeburtstagsListe.CreateList();

            listen = listen.OrderBy(x => x.Vorname).ToList();

            
           
            foreach (GeburtstagsListe i2 in listen)
            {
                
                    lboxKontakte.Items.Add(i2.Vorname + " " + i2.Name);
               
            }
        }
       
        private void KontaktLoeschen()
        {
            if (lboxKontakte.SelectedItem != null)
            {
                MessageBoxResult result = MessageBox.Show("Wollen Sie den Kontakt wirklich löschen", "Kontakt löschen", MessageBoxButton.YesNo);
                switch (result)
                {
                    case MessageBoxResult.Yes:
                        int counter = 2;
                        using (ExcelEngine excelEngine = new ExcelEngine())
                        {
                            IApplication application = excelEngine.Excel;
                            application.DefaultVersion = ExcelVersion.Excel2016;

                            //open a dokument

                            IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                            IWorksheet worksheet = workbook.Worksheets[0];

                            
                            string namen;

                            namen = (string)lboxKontakte.SelectedItem;

                            // in subs wird vor und nachname gespeichert des wegen der spilt

                            string[] subs = namen.Split(' ');
                          

                            while (true)
                            {
                                if (worksheet.Range["A" + counter].Text == subs[0] && worksheet.Range["B" + counter].Text == subs[1])
                                {
                                    worksheet.Range["A" + counter + ":" + "L" + counter].Text = null;

                                    
                                    

                                    lboxKontakte.Items.Clear();
                                   
                                    
                                    break;

                                        
                                }
                                else if (worksheet.Range["A" + counter].Text == null)
                                {
                                    MessageBox.Show("Kontakt nicht gefunden");
                                    break;
                                }
                                counter++;


                            }



                            //ende
                            
                            
                            MessageBox.Show("Kontakt ist gelöscht");
                            
                            workbook.SaveAs(Paths.GetFilePath("Output.xlsx"));
                            workbook.Close();
                            lboxKontakte.Items.Clear();
                            ListeNamen();



                        }

                        break;

                    case MessageBoxResult.No:

                        break;
                }

            }


        }

        public BitmapSource kontaktBildImage { get; private set; }

        private void lboxKontakte_SelectionChanged(object sender, SelectionChangedEventArgs e)
            {
                UserKontakte userKontakte = new UserKontakte();
                Kontakte.Content = userKontakte;


                using (ExcelEngine excelEngine = new ExcelEngine())
                {
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //open a dokument

                    IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                    IWorksheet worksheet = workbook.Worksheets[0];

                    int counter = 2;
                    string KontaktKlick;
                    KontaktKlick = (string)lboxKontakte.SelectedItem;


                    if (lboxKontakte.SelectedItem != null)
                    {     //teilt vornamen und nachnamen auf und durchsucht die tabelle
                        string[] teilen = KontaktKlick.Split();
                        while (true)
                        {
                            if (worksheet.Range["A" + counter].Text == teilen[0] && worksheet.Range["B" + counter].Text == teilen[1])
                            {
                                userKontakte.vorname.Content = worksheet.Range["A" + counter].Text;
                                userKontakte.name.Content = worksheet.Range["B" + counter].Text;
                                userKontakte.geburtstag.Content = worksheet.Range["K" + counter].Text;
                                userKontakte.strasse.Content = worksheet.Range["D" + counter].Text;
                                userKontakte.hausnummer.Content = worksheet.Range["E" + counter].Text;
                                userKontakte.postleizahl.Content = worksheet.Range["F" + counter].Text;
                                userKontakte.ort.Content = worksheet.Range["G" + counter].Text;
                                userKontakte.telefon.Content = worksheet.Range["H" + counter].Text;
                                userKontakte.email.Content = worksheet.Range["I" + counter].Text;


                                string VN = worksheet["A" + counter].Text + worksheet["B" + counter].Text + ".png";
                                if (worksheet["J" + counter].Text == "1")
                                {

                                Bitmap bitmap = new Bitmap(Paths.GetFilePath("Bilder\\") + VN);
                                kontaktBildImage = Imaging.CreateBitmapSourceFromHBitmap(bitmap.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

                                userKontakte.bildkontakt.Source = kontaktBildImage;


                               // userKontakte.bildkontakt.Source = new BitmapImage(new Uri(AppDomain.CurrentDomain.BaseDirectory+@"Bilder\" + VN, UriKind.RelativeOrAbsolute));

                               


                                }


                                break;
                            }
                            else if (worksheet.Range["A" + counter].Text == null)
                            {
                                MessageBox.Show("Kontakt nicht gefunden");
                                break;
                            }
                            counter++;



                            // lboxKontakte.Items.Remove(lboxKontakte.Items[lboxKontakte.SelectedIndex]);


                        }

                    }





                    //ende

                    workbook.Close();
                    lboxKontakte.Items.Refresh();

                }

            } 

        private void gebvBild_MouseDown(object sender, MouseButtonEventArgs e)
        {

            Startseite startseite = new Startseite();
            Kontakte.Content = startseite;
            ListeGeburstage();
            bearbeiten.Height = 0;
            bearbeiten.Width = 0;
        }
       
        private void kontakteBild_MouseDown(object sender, MouseButtonEventArgs e)
        {
            ListeNamen();
            list.Clear();
            foreach (string str in lboxKontakte.Items)
            {
                list.Add(str);

            }
            UserKontakte userKontakte = new UserKontakte();
            Kontakte.Content = userKontakte;
            bearbeiten.Height = 50;
            bearbeiten.Width = 50;
        }

        private void deleteBild_MouseDown(object sender, MouseButtonEventArgs e)
        {
            KontaktLoeschen();
            bearbeiten.Height = 0;
            bearbeiten.Width = 0;
        }

        private void addBild_MouseDown(object sender, MouseButtonEventArgs e)
        {

            UserHinzu userHinzu = new UserHinzu();
            Kontakte.Content = userHinzu;
            bearbeiten.Height = 0;
            bearbeiten.Width = 0;
        }

        
        List<string> list = new List<string>();
        private void suche_TextChanged(object sender, TextChangedEventArgs e)
        {
           
            if (string.IsNullOrEmpty(suche.Text) == false)
            {
                lboxKontakte.Items.Clear();
                foreach (string str in list)
                {
                    if (str.Contains(suche.Text))
                    {
                        lboxKontakte.Items.Add(str);
                    }
                }
            }
            else if (suche.Text == "")
            {
                lboxKontakte.Items.Clear();
                foreach (string str in list)
                {
                    lboxKontakte.Items.Add(str);

                }
            }
        }

        private void gebBild_MouseEnter(object sender, MouseEventArgs e)
        {

            gebBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\geb2.png"), UriKind.RelativeOrAbsolute));
            

        }

        private void gebBild_MouseLeave(object sender, MouseEventArgs e)
        {
            gebBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\geb.png"), UriKind.RelativeOrAbsolute));


        }

        private void kontakteBild_MouseEnter(object sender, MouseEventArgs e)
        {
            kontakteBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\kontakte2.png"), UriKind.RelativeOrAbsolute));

        }

        private void kontakteBild_MouseLeave(object sender, MouseEventArgs e)
        {
           kontakteBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\kontakte.png"), UriKind.RelativeOrAbsolute));

        }

        private void deleteBild_MouseEnter(object sender, MouseEventArgs e)
        {
           deleteBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\delete2.png"), UriKind.RelativeOrAbsolute));

        }

        private void deleteBild_MouseLeave(object sender, MouseEventArgs e)
        {
           deleteBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\delete.png"), UriKind.RelativeOrAbsolute));

        }

        private void addBild_MouseEnter(object sender, MouseEventArgs e)
        {
            addBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\add2.png"), UriKind.RelativeOrAbsolute));

        }

        private void addBild_MouseLeave(object sender, MouseEventArgs e)
        {
            addBild.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\add.png"), UriKind.RelativeOrAbsolute));

        }

        private void bebBild2_MouseDown(object sender, MouseButtonEventArgs e)
        {
            list.Clear();
            foreach (string str in lboxKontakte.Items)
            {
                list.Add(str);

            }

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //open a dokument

                IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                int counter = 2;
                string KontaktKlick;
                KontaktKlick = (string)lboxKontakte.SelectedItem;


                if (lboxKontakte.SelectedItem != null)
       {

                    Kontaktebearbeiten kontaktebearbeiten = new Kontaktebearbeiten();
                    Kontakte.Content = kontaktebearbeiten;
                    //teilt vornamen und nachnamen auf und durchsucht die tabelle
                    string[] teilen = KontaktKlick.Split();
                    while (true)
                    {
                        if (worksheet.Range["A" + counter].Text == teilen[0] && worksheet.Range["B" + counter].Text == teilen[1])
                        {
                            kontaktebearbeiten.vorname.Text = worksheet.Range["A" + counter].Text;
                            kontaktebearbeiten.name.Text = worksheet.Range["B" + counter].Text;
                            kontaktebearbeiten.datePicker.Text = worksheet.Range["K" + counter].Text;
                            kontaktebearbeiten.HinzufuegenStrasse.Text = worksheet.Range["D" + counter].Text;
                            kontaktebearbeiten.HinzufuegenHausnummer.Text = worksheet.Range["E" + counter].Text;
                            kontaktebearbeiten.HinzufuegenPostleizahl.Text = worksheet.Range["F" + counter].Text;
                            kontaktebearbeiten.HinzufuegenOrt.Text= worksheet.Range["G" + counter].Text;
                            kontaktebearbeiten.HinzufuegenTelefon.Text = worksheet.Range["H" + counter].Text;
                            kontaktebearbeiten.HinzufuegenEmail.Text = worksheet.Range["I" + counter].Text;


                            string VN = worksheet["A" + counter].Text + worksheet["B" + counter].Text + ".png";
                            if (worksheet["J" + counter].Text == "1")
                            {
                                Bitmap bitmap = new Bitmap(Paths.GetFilePath(@"Bilder\") + VN);
                                kontaktBildImage = Imaging.CreateBitmapSourceFromHBitmap(bitmap.GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());

                                kontaktebearbeiten.bildkontakt.Source = kontaktBildImage;
                               // kontaktebearbeiten.bildkontakt.Source = new BitmapImage(new Uri(AppDomain.CurrentDomain.BaseDirectory + @"Bilder\" + VN, UriKind.RelativeOrAbsolute));

                            }


                            break;
                        }
                        else if (worksheet.Range["A" + counter].Text == null)
                        {
                            MessageBox.Show("Kontakt nicht gefunden");
                            break;
                        }
                        counter++;

                    }

                }
                //ende
                workbook.Close();

                
                lboxKontakte.Items.Refresh();
            }

        }

        private void bebBild2_MouseEnter(object sender, MouseEventArgs e)
        {
           bearbeiten.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\bearbeiten2.png"), UriKind.RelativeOrAbsolute));

        }

        private void bebBild2_MouseLeave(object sender, MouseEventArgs e)
        {
          bearbeiten.Source = new BitmapImage(new Uri(Paths.GetFilePath(@"ressours\bearbeiten.png"), UriKind.RelativeOrAbsolute));

        }

        private void gebBild_MouseDown(object sender, MouseButtonEventArgs e)
        {

        }

       
    }
}
    


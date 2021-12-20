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

namespace Adressbuch
{
    internal class FotoLoeschen
    {
        List<string> namen = new List<string>();
        public void OrdnerClear()
        {

            int counter = 2;
            List<string> datei = new List<string>();

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];
               
                while (true)
                {
                    if (worksheet.Range["A" + counter].Text == null)
                    {

                        break;
                    }
                    else
                    {
                        namen.Add(worksheet.Range["A" + counter].Text + worksheet.Range["B" + counter].Text + ".png");
                    }
                    counter++;
                }
                workbook.Close();


            }

            string[] datein = Directory.GetFiles(Paths.GetFilePath("Bilder") );
            foreach (string s in datein)
            {
                datei.Add(Path.GetFileName(s));
            }



            List<string> vergleichliste = datei.Except(namen).ToList();

            for (int i = 0; i < vergleichliste.Count; i++)
            {
                System.IO.File.Delete(Paths.GetFilePath("Bilder\\" + vergleichliste[i]));
            }



        }

      
           
        
    }
}

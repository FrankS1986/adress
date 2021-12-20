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
    internal class GeburtstagsListe
    {

        public string Vorname { get; set; }
        public string Name { get; set; }
        public int Alter { get; set; }

        public string date { get; set; }
       
        public static List<GeburtstagsListe> CreateList()
        {
            int counter = 2;

            using (ExcelEngine excelEngine = new ExcelEngine())
            {


               
                    IApplication application = excelEngine.Excel;
                    application.DefaultVersion = ExcelVersion.Excel2016;

                    //create a dokument

                    IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                    IWorksheet worksheet = workbook.Worksheets[0];

                List<GeburtstagsListe> list = new List<GeburtstagsListe>();


                while (worksheet.Range["A"+counter].Text !=null)
                {
                   
                    if (worksheet.Range["A" + counter].Text != "")
                    {
                        list.Add(new GeburtstagsListe() { Vorname = worksheet.Range["A" + counter].Text, Name = worksheet.Range["B" + counter].Text, Alter = Convert.ToInt32( worksheet.Range["L" + counter].Text), date = worksheet.Range["C" + counter].Text });
                    }
                        counter++;
                }

                



                    workbook.Close();
                    return list;

                



               

            }




        }
    }
}

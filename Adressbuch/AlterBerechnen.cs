using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Syncfusion.XlsIO;
namespace Adressbuch
{
    internal class AlterBerechnen
    {


        public void Berechnen()
        {

            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

                //open a dokument

                IWorkbook workbook = application.Workbooks.Open(Paths.GetFilePath("Output.xlsx"));
                IWorksheet worksheet = workbook.Worksheets[0];

                DateTime today = DateTime.Today;
                TimeSpan time;

                int count = 2;
                while (worksheet.Range["A" + count].Text != null)
                {
                    if (worksheet.Range["K" + count].Text != "")
                    {
                        int alter = today.Year - Convert.ToDateTime(worksheet.Range["K" + count].Text).Year;
                        if (Convert.ToDateTime(worksheet.Range["K" + count].Text).AddYears(alter) >= today)

                        {
                            time = Convert.ToDateTime(worksheet.Range["K" + count].Text).AddYears(alter) - today;
                            worksheet.Range["L" + count].Text = Convert.ToString(time.Days);
                        }

                        else
                        {
                            time = Convert.ToDateTime(worksheet.Range["K" + count].Text).AddYears(alter + 1) - today;
                            worksheet.Range["L" + count].Text = Convert.ToString(time.Days);
                        }
                    }

                    else
                    {

                        worksheet.Range["L" + count].Text = "-1";
                    }
                    count++;
                }

                workbook.Save();
                workbook.Close();
            }
        }
    }
}

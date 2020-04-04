using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;

using OfficeOpenXml;
using static UrenTijd.MainWindow;

namespace UrenTijd
{
    class Utils
    {
        static public void CreateExcel(DayFieldsStruct[] days)
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("UrenTijd.assets.UrenBriefTemplate.xlsx");

            ExcelPackage excel = new OfficeOpenXml.ExcelPackage(stream);

            var sheet = excel.Workbook.Worksheets[0];

            for (int i = 0; i < days.Length; i++)
            {
                DayFieldsStruct day = days[i];
                // the base row is the first row of the day
                int baseRow = 5 + 8 * i;

                sheet.Cells["C" + baseRow.ToString()].Value = day.arriving.from?.ToString("HH:mm") ?? "";
                sheet.Cells["E" + baseRow.ToString()].Value = day.arriving.until?.ToString("HH:mm") ?? "";

                if (day.arriving.from != null && day.arriving.until != null)
                {
                    TimeSpan difference = day.arriving.until?.Subtract(day.arriving.from ?? DateTime.Now) ?? TimeSpan.Zero;

                    double hoursBetween = difference.TotalHours;

                    sheet.Cells["f" + baseRow.ToString()].Value = hoursBetween;
                }

                sheet.Cells["C" + (baseRow + 1).ToString()].Value = day.working.from?.ToString("HH:mm") ?? "";
                sheet.Cells["E" + (baseRow + 1).ToString()].Value = day.working.until?.ToString("HH:mm") ?? "";

                if (day.working.from != null && day.working.until != null)
                {
                    TimeSpan difference = day.working.until?.Subtract(day.working.from ?? DateTime.Now) ?? TimeSpan.Zero;

                    double hoursBetween = difference.TotalHours;

                    sheet.Cells["f" + (baseRow + 1).ToString()].Value = hoursBetween;
                }

                sheet.Cells["C" + (baseRow + 2).ToString()].Value = day.workDescription;

                sheet.Cells["C" + (baseRow + 6).ToString()].Value = day.leaving.from?.ToString("HH:mm") ?? "";
                sheet.Cells["E" + (baseRow + 6).ToString()].Value = day.leaving.until?.ToString("HH:mm") ?? "";

                if (day.leaving.from != null && day.leaving.until != null)
                {
                    TimeSpan difference = day.leaving.until?.Subtract(day.leaving.from ?? DateTime.Now) ?? TimeSpan.Zero;

                    double hoursBetween = difference.TotalHours;

                    sheet.Cells["f" + (baseRow + 6).ToString()].Value = hoursBetween;
                }

                sheet.Cells["J" + baseRow.ToString()].Value = day.workType;
            }

            FileInfo excelFile = new FileInfo(@"C:\Users\erhan\Desktop\test.xlsx");
            excel.SaveAs(excelFile);
        }
    }
}

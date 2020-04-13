using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection;

using OfficeOpenXml;
using static UrenTijd.MainWindow;
using System.Globalization;
using System.Net.Mail;
using System.Diagnostics;

namespace UrenTijd
{
    public static class Utils
    {
        public static string DefaultFileName = "Urenbrief digitaal chauffeur Vroom.xlsx";

        static public void CreateExcel(DayFieldsStruct[] days, DateTime startOfWeek)
        {
            Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("UrenTijd.assets.UrenBriefTemplate.xlsx");

            ExcelPackage excel = new OfficeOpenXml.ExcelPackage(stream);

            var sheet = excel.Workbook.Worksheets[0];

            sheet.Cells["I1"].Value = startOfWeek.ToString("d-M-yyyy");
            sheet.Cells["L1"].Value = startOfWeek.AddDays(6).ToString("d-M-yyyy");
            sheet.Cells["O1"].Value = startOfWeek.ToString("yy");

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

                    sheet.Cells["f" + (baseRow + 1).ToString()].Value = hoursBetween - ((day.hadBreak && hoursBetween >= 1.0d) ? 1 : 0);
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

            string tempPath = Environment.GetEnvironmentVariable("Temp") + @"\UrenTijd\";
            Directory.CreateDirectory(tempPath);
            FileInfo excelFile = new FileInfo(tempPath + DefaultFileName);
            excel.SaveAs(excelFile);
        }

        public static void SendMail()
        {
            string tempFile = Environment.GetEnvironmentVariable("Temp") + @"\UrenTijd\";
            var mailMessage = new MailMessage
            {
                From = new MailAddress("famarif@hotmail.com"),
                Subject = "urenstaat",
                IsBodyHtml = false,
                Body = "",
            };

            mailMessage.Attachments.Add(new Attachment(tempFile + DefaultFileName));
            string fileName = tempFile + "urenstaatMessage.eml";

            MailUtility.Save(mailMessage, fileName);
            Process.Start(fileName);
        }

        public static DateTime FirstDateOfWeekISO8601(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = DayOfWeek.Thursday - jan1.DayOfWeek;

            // Use first Thursday in January to get first week of the year as
            // it will never be in Week 52/53
            DateTime firstThursday = jan1.AddDays(daysOffset);
            var cal = CultureInfo.CurrentCulture.Calendar;
            int firstWeek = cal.GetWeekOfYear(firstThursday, CalendarWeekRule.FirstFourDayWeek, DayOfWeek.Monday);

            var weekNum = weekOfYear;
            // As we're adding days to a date in Week 1,
            // we need to subtract 1 in order to get the right date for week #1
            if (firstWeek == 1)
            {
                weekNum -= 1;
            }

            // Using the first Thursday as starting week ensures that we are starting in the right year
            // then we add number of weeks multiplied with days
            var result = firstThursday.AddDays(weekNum * 7);

            // Subtract 3 days from Thursday to get Monday, which is the first weekday in ISO8601
            return result.AddDays(-3);
        }
    }
}

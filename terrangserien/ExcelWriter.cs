using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace terrangserien
{
    struct Styles
    {
        public XSSFCellStyle normalCenter;
        public XSSFCellStyle normalLeft;
        public XSSFCellStyle normalHeaderCenter;
        public XSSFCellStyle normalHeaderLeft;
        public XSSFCellStyle bigHeader;
    }

    class ExcelWriter
    {
        private static void CreateCellWithStyle(ref IRow row, int column, string value, ref XSSFCellStyle style)
        {
            ICell cell = row.CreateCell(column, CellType.String);
            cell.CellStyle = style;
            cell.SetCellValue(value);
        }

        private static void CreateCellWithStyle(ref IRow row, int column, double value, ref XSSFCellStyle style)
        {
            if (value > 0.1)
            {
                ICell cell = row.CreateCell(column, CellType.Numeric);
                cell.CellStyle = style;
                cell.SetCellValue(value);
            }
        }

        public static void WriteAttendance(ref ISheet sheet, ref IList<Person> persons, ref Styles styles)
        {
            sheet.SetColumnWidth(0, 256 * 7);
            sheet.SetColumnWidth(1, 256 * 7);
            sheet.SetColumnWidth(2, 256 * 15);
            sheet.SetColumnWidth(3, 256 * 7);
            sheet.SetColumnWidth(4, 256 * 15);
            sheet.SetColumnWidth(5, 256 * 17);
            sheet.SetColumnWidth(6, 256 * 7);
            sheet.SetColumnWidth(7, 256 * 7);
            sheet.SetColumnWidth(8, 256 * 7);
            sheet.SetColumnWidth(9, 256 * 7);
            sheet.SetColumnWidth(10, 256 * 7);
            sheet.SetColumnWidth(11, 256 * 7);
            sheet.SetColumnWidth(12, 256 * 7);

            int i = 0;
            WriteSmallHeaderAthletes(ref sheet, i++, ref styles);
            foreach (Person person in persons)
            {
                IRow row = sheet.CreateRow(i);
                CreateCellWithStyle(ref row, 0, person.Distance, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 1, person.Gender, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 2, person.SocialNumber, ref styles.normalCenter);
                if (person.Number.Length == 0)
                {
                    CreateCellWithStyle(ref row, 3, person.Number, ref styles.normalCenter);
                }
                else
                {
                    CreateCellWithStyle(ref row, 3, int.Parse(person.Number), ref styles.normalCenter);
                }
                CreateCellWithStyle(ref row, 4, person.Name, ref styles.normalLeft);
                CreateCellWithStyle(ref row, 5, person.Surname, ref styles.normalLeft);
                CreateCellWithStyle(ref row, 6, person.Klass, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 7, person.Result(0).Time, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 8, person.Result(1).Time, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 9, person.Result(2).Time, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 10, person.Result(3).Time, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 11, person.Result(4).Time, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 12, person.Result(5).Time, ref styles.normalCenter);
                i++;
            }
        }
        private static IList<Person> DoQueryAndSort(ref IList<Person> persons, ResultFilter filter)
        {
            IList<Person> filtered = persons.FilterResults(filter);
            IEnumerable<Person> sortedEnum = filtered.OrderBy(f => f.Result(filter.Day));
            IList<Person> sorted = sortedEnum.ToList();
            return sorted;
        }

        private static int WriteByClass(ref ISheet sheet, int startRow, ref IList<Person> persons, ref Styles styles, int day)
        {
            int currentRow = startRow;
            int placering = 1;
            foreach (Person person in persons)
            {
                IRow row = sheet.CreateRow(currentRow);
                CreateCellWithStyle(ref row, 0, placering.ToString(), ref styles.normalCenter);
                CreateCellWithStyle(ref row, 1, person.Number, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 2, person.Name, ref styles.normalLeft);
                CreateCellWithStyle(ref row, 3, person.Surname, ref styles.normalLeft);
                CreateCellWithStyle(ref row, 4, person.Klass, ref styles.normalCenter);
                CreateCellWithStyle(ref row, 5, person.Result(day).ToString(), ref styles.normalCenter);
                currentRow++;
                placering++;
            }
            return currentRow;
        }

        private static int WriteSmallHeader(ref ISheet sheet, int startRow, ref Styles styles)
        {
            int currentRow = startRow;
            IRow row = sheet.CreateRow(currentRow);
            CreateCellWithStyle(ref row, 0, "Placering", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 1, "Nr", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 2, "Namn", ref styles.normalHeaderLeft);
            CreateCellWithStyle(ref row, 3, "Efternamn", ref styles.normalHeaderLeft);
            CreateCellWithStyle(ref row, 4, "Åldersgrupp", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 5, "Resultat", ref styles.normalHeaderCenter);
            currentRow++;
            return currentRow;
        }

        private static int WriteSmallHeaderAthletes(ref ISheet sheet, int startRow, ref Styles styles)
        {
            int currentRow = startRow;
            IRow row = sheet.CreateRow(currentRow);
            CreateCellWithStyle(ref row, 0, "Sträcka", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 1, "Kön", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 2, "Född år", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 3, "Startnr", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 4, "Förnamn", ref styles.normalHeaderLeft);
            CreateCellWithStyle(ref row, 5, "Efternamn", ref styles.normalHeaderLeft);
            CreateCellWithStyle(ref row, 6, "Klass", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 7, "1", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 8, "2", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 9, "3", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 10, "4", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 11, "5", ref styles.normalHeaderCenter);
            CreateCellWithStyle(ref row, 12, "6", ref styles.normalHeaderCenter);
            currentRow++;
            return currentRow;
        }

        private static int WriteBigHeader(ref ISheet sheet, int startRow, string header1, string header2, string header3, ref Styles styles)
        {
            int currentRow = startRow;
            IRow r = sheet.CreateRow(currentRow);
            currentRow++;
            CreateCellWithStyle(ref r, 0, header1, ref styles.bigHeader);

            r = sheet.CreateRow(currentRow);
            currentRow++;
            CreateCellWithStyle(ref r, 0, header2, ref styles.bigHeader);

            r = sheet.CreateRow(currentRow);
            currentRow++;
            CreateCellWithStyle(ref r, 0, header3, ref styles.bigHeader);
            return currentRow;
        }

        private static int WriteEmptyLine(ref ISheet sheet, int startRow)
        {
            int currentRow = startRow;
            IRow r = sheet.CreateRow(currentRow);
            currentRow++;
            return currentRow;
        }

        private static int WriteGroupResult(ref ISheet sheet, int startRow, ref IList<Person> persons, ref Styles styles, int day, ResultFilter filter)
        {
            int lastRow = startRow;
            IList<Person> filtered = DoQueryAndSort(ref persons, filter);
            if (filtered.Count > 0)
            {
                lastRow = WriteSmallHeader(ref sheet, lastRow, ref styles);
                lastRow = WriteByClass(ref sheet, lastRow, ref filtered, ref styles, day);
                lastRow = WriteEmptyLine(ref sheet, lastRow);
            }
            return lastRow;
        }

        private static void WriteResult(ref ISheet sheet, ref IList<Person> persons, ref Styles styles, int day)
        {
            sheet.SetColumnWidth(0, 256 * 15);
            sheet.SetColumnWidth(1, 256 * 15);
            sheet.SetColumnWidth(2, 256 * 15);
            sheet.SetColumnWidth(3, 256 * 15);
            sheet.SetColumnWidth(4, 256 * 15);
            sheet.SetColumnWidth(5, 256 * 15);

            int lastRow = 0;
            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Flickor", "600 m", ref styles);

            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "0-4", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "5-6", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "7-8", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "9-10", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "11-12", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "13-14", Distance = "600", Gender = "f", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Pojkar", "600 m", ref styles);
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "0-4", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "5-6", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "7-8", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "9-10", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "11-12", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "13-14", Distance = "600", Gender = "p", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Flickor", "1200 m", ref styles);
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "0-4", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "5-6", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "7-8", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "9-10", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "11-12", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "13-14", Distance = "1200", Gender = "f", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Pojkar", "1200 m", ref styles);
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "0-4", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "5-6", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "7-8", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "9-10", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "11-12", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "13-14", Distance = "1200", Gender = "p", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Flickor", "3000 m", ref styles);
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "0-4", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "5-6", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "7-8", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "9-10", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "11-12", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "13-14", Distance = "3000", Gender = "f", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Pojkar", "3000 m", ref styles);
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "0-4", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "5-6", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "7-8", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "9-10", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "11-12", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref styles, day, new ResultFilter { Klass = "13-14", Distance = "3000", Gender = "p", Day = day });
        }

        public static void Write(ref string filePath, ref IList<Person> persons)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();

            XSSFFont arialNormalFont = (XSSFFont)workbook.CreateFont();
            arialNormalFont.FontHeightInPoints = (short)10;
            arialNormalFont.FontName = "Arial";

            XSSFCellStyle centerStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            centerStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            centerStyle.SetFont(arialNormalFont);

            XSSFCellStyle leftStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            leftStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            leftStyle.SetFont(arialNormalFont);

            XSSFFont smallHeaderFont = (XSSFFont)workbook.CreateFont();
            smallHeaderFont.FontHeightInPoints = (short)10;
            smallHeaderFont.FontName = "Arial";
            smallHeaderFont.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            XSSFCellStyle normalHeaderCenterStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            normalHeaderCenterStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            normalHeaderCenterStyle.SetFont(smallHeaderFont);

            XSSFCellStyle normalHeaderLeftStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            normalHeaderLeftStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            normalHeaderLeftStyle.SetFont(smallHeaderFont);

            XSSFFont bigHeaderFont = (XSSFFont)workbook.CreateFont();
            bigHeaderFont.FontHeightInPoints = (short)12;
            bigHeaderFont.FontName = "Arial";
            bigHeaderFont.Boldweight = (short)NPOI.SS.UserModel.FontBoldWeight.Bold;

            XSSFCellStyle bigHeaderStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            bigHeaderStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            bigHeaderStyle.SetFont(bigHeaderFont);

            Styles styles = new Styles();
            styles.normalCenter = centerStyle;
            styles.normalLeft = leftStyle;
            styles.bigHeader = bigHeaderStyle;
            styles.normalHeaderCenter = normalHeaderCenterStyle;
            styles.normalHeaderLeft = normalHeaderLeftStyle;

            ISheet attendanceSheet = workbook.CreateSheet();
            WriteAttendance(ref attendanceSheet, ref persons, ref styles);

            ISheet sheet0 = workbook.CreateSheet();
            WriteResult(ref sheet0, ref persons, ref styles, 0);

            ISheet sheet1 = workbook.CreateSheet();
            WriteResult(ref sheet1, ref persons, ref styles, 1);

            ISheet sheet2 = workbook.CreateSheet();
            WriteResult(ref sheet2, ref persons, ref styles, 2);

            ISheet sheet3 = workbook.CreateSheet();
            WriteResult(ref sheet3, ref persons, ref styles, 3);

            ISheet sheet4 = workbook.CreateSheet();
            WriteResult(ref sheet4, ref persons, ref styles, 4);

            ISheet sheet5 = workbook.CreateSheet();
            WriteResult(ref sheet5, ref persons, ref styles, 5);

            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }
            workbook.Close();
        }
    }
}

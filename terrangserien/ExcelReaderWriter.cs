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
    class ExcelReaderWriter
    {
        public static IList<Person> Read(ref string filePath)
        {
            IList<Person> persons = new List<Person>();
            using (FileStream stream = new FileStream(@filePath, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(stream);

                // Sträcka | p/ f | Född år | Startnr | Förnamn | Efternamn | Klass | May 07 | May | 09 | May | 14 | May | 16 | May | 21 | May 23
                ISheet sheet = workbook.GetSheetAt(0);
                int count = sheet.LastRowNum;
                for (int i = 0; i < sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i + 1);
                    if (row == null)
                    {
                        continue;
                    }
                    Person person = Person.Create();

                    person.Distance = row.GetCell(0).ToString().Trim();
                    person.Gender = row.GetCell(1).StringCellValue.Trim();
                    person.SocialNumber = row.GetCell(2).ToString().Trim();
                    person.Number = row.GetCell(3).ToString().Trim();
                    person.Name = row.GetCell(4).StringCellValue.Trim();
                    person.Surname = row.GetCell(5).StringCellValue.Trim();
                    person.Klass = row.GetCell(6).StringCellValue.Trim();
                    string r0 = row.GetCell(7).NumericCellValue.ToString().Trim();
                    person.Result(0, Result.Create(ref r0));

                    string r1 = row.GetCell(8).NumericCellValue.ToString().Trim();
                    person.Result(1, Result.Create(ref r1));

                    string r2 = row.GetCell(9).NumericCellValue.ToString().Trim();
                    person.Result(2, Result.Create(ref r2));

                    string r3 = row.GetCell(10).NumericCellValue.ToString().Trim();
                    person.Result(3, Result.Create(ref r3));

                    string r4 = row.GetCell(11).NumericCellValue.ToString().Trim();
                    person.Result(4, Result.Create(ref r4));

                    string r5 = row.GetCell(12).NumericCellValue.ToString().Trim();
                    person.Result(5, Result.Create(ref r5));
                    persons.Add(person);
                }
                workbook.Close();
            }
            return persons;
        }

        public static void WriteAttendance(ref ISheet sheet, ref IList<Person> persons, ref XSSFCellStyle centerStyle, ref XSSFCellStyle leftStyle)
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

            void CreateCellWithStyle(ref IRow row, int column, string value, ref XSSFCellStyle style)
            {
                ICell cell = row.CreateCell(column, CellType.String);
                cell.CellStyle = style;
                cell.SetCellValue(value);
            }

            int i = 0;
            foreach (Person person in persons)
            {
                IRow row = sheet.CreateRow(i);
                CreateCellWithStyle(ref row, 0, person.Distance, ref centerStyle);
                CreateCellWithStyle(ref row, 1, person.Gender, ref centerStyle);
                CreateCellWithStyle(ref row, 2, person.SocialNumber, ref centerStyle);
                CreateCellWithStyle(ref row, 3, person.Number, ref centerStyle);
                CreateCellWithStyle(ref row, 4, person.Name, ref leftStyle);
                CreateCellWithStyle(ref row, 5, person.Surname, ref leftStyle);
                CreateCellWithStyle(ref row, 6, person.Klass, ref centerStyle);
                CreateCellWithStyle(ref row, 7, person.Result(0).ToString(), ref leftStyle);
                CreateCellWithStyle(ref row, 8, person.Result(1).ToString(), ref leftStyle);
                CreateCellWithStyle(ref row, 9, person.Result(2).ToString(), ref leftStyle);
                CreateCellWithStyle(ref row, 10, person.Result(3).ToString(), ref leftStyle);
                CreateCellWithStyle(ref row, 11, person.Result(4).ToString(), ref leftStyle);
                CreateCellWithStyle(ref row, 12, person.Result(5).ToString(), ref leftStyle);
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

        private static int WriteByClass(ref ISheet sheet, int startRow, ref IList<Person> persons, ref XSSFCellStyle centerStyle, ref XSSFCellStyle leftStyle, int day)
        {
            void CreateCellWithStyle(ref IRow row, int column, string value, ref XSSFCellStyle style)
            {
                ICell cell = row.CreateCell(column, CellType.String);
                cell.CellStyle = style;
                cell.SetCellValue(value);
            }

            int currentRow = startRow;
            int placering = 1;
            foreach (Person person in persons)
            {
                IRow row = sheet.CreateRow(currentRow);
                CreateCellWithStyle(ref row, 0, placering.ToString(), ref centerStyle);
                CreateCellWithStyle(ref row, 1, person.Number, ref centerStyle);
                CreateCellWithStyle(ref row, 2, person.Name, ref leftStyle);
                CreateCellWithStyle(ref row, 3, person.Surname, ref leftStyle);
                CreateCellWithStyle(ref row, 4, person.Klass, ref centerStyle);
                CreateCellWithStyle(ref row, 5, person.Result(day).ToString(), ref centerStyle);
                currentRow++;
                placering++;
            }
            return currentRow;
        }

        private static int WriteSmallHeader(ref ISheet sheet, int startRow, ref XSSFCellStyle centerStyle, ref XSSFCellStyle leftStyle)
        {
            void CreateCellWithStyle(ref IRow row, int column, string value, ref XSSFCellStyle style)
            {
                ICell cell = row.CreateCell(column, CellType.String);
                cell.CellStyle = style;
                cell.SetCellValue(value);
            }

            int currentRow = startRow;
            IRow r = sheet.CreateRow(currentRow);
            CreateCellWithStyle(ref r, 0, "Placering", ref centerStyle);
            CreateCellWithStyle(ref r, 1, "Nr", ref centerStyle);
            CreateCellWithStyle(ref r, 2, "Namn", ref leftStyle);
            CreateCellWithStyle(ref r, 3, "Efternamn", ref leftStyle);
            CreateCellWithStyle(ref r, 4, "Åldersgrupp", ref centerStyle);
            CreateCellWithStyle(ref r, 5, "Resultat", ref centerStyle);
            currentRow++;
            return currentRow;
        }

        private static int WriteBigHeader(ref ISheet sheet, int startRow, string header1, string header2, string header3)
        {
            void CreateCell(ref IRow row, int column, string value)
            {
                ICell cell = row.CreateCell(column, CellType.String);
                cell.SetCellValue(value);
            }

            int currentRow = startRow;
            IRow r = sheet.CreateRow(currentRow);
            currentRow++;
            CreateCell(ref r, 0, header1);

            r = sheet.CreateRow(currentRow);
            currentRow++;
            CreateCell(ref r, 0, header2);

            r = sheet.CreateRow(currentRow);
            currentRow++;
            CreateCell(ref r, 0, header3);
            return currentRow;
        }

        private static int WriteEmptyLine(ref ISheet sheet, int startRow)
        {
            int currentRow = startRow;
            IRow r = sheet.CreateRow(currentRow);
            currentRow++;
            return currentRow;
        }

        private static int WriteGroupResult(ref ISheet sheet, int startRow, ref IList<Person> persons, ref XSSFCellStyle centerStyle, ref XSSFCellStyle leftStyle, int day, ResultFilter filter)
        {
            int lastRow = startRow;
            IList<Person> filtered = DoQueryAndSort(ref persons, filter);
            if (filtered.Count > 0)
            {
                lastRow = WriteSmallHeader(ref sheet, lastRow, ref centerStyle, ref leftStyle);
                lastRow = WriteByClass(ref sheet, lastRow, ref filtered, ref centerStyle, ref leftStyle, day);
                lastRow = WriteEmptyLine(ref sheet, lastRow);
            }
            return lastRow;
        }

        private static void WriteResult(ref ISheet sheet, ref IList<Person> persons, ref XSSFCellStyle centerStyle, ref XSSFCellStyle leftStyle, int day)
        {
            sheet.SetColumnWidth(0, 256 * 15);
            sheet.SetColumnWidth(1, 256 * 15);
            sheet.SetColumnWidth(2, 256 * 15);
            sheet.SetColumnWidth(3, 256 * 15);
            sheet.SetColumnWidth(4, 256 * 15);
            sheet.SetColumnWidth(5, 256 * 15);

            int lastRow = 0;
            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Flickor", "600 m");

            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "0-4", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "5-6", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "7-8", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "9-10", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "11-12", Distance = "600", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "13-14", Distance = "600", Gender = "f", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Pojkar", "600 m");
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "0-4", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "5-6", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "7-8", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "9-10", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "11-12", Distance = "600", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "13-14", Distance = "600", Gender = "p", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Flickor", "1200 m");
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "0-4", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "5-6", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "7-8", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "9-10", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "11-12", Distance = "1200", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "13-14", Distance = "1200", Gender = "f", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Pojkar", "1200 m");
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "0-4", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "5-6", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "7-8", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "9-10", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "11-12", Distance = "1200", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "13-14", Distance = "1200", Gender = "p", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Flickor", "3000 m");
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "0-4", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "5-6", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "7-8", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "9-10", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "11-12", Distance = "3000", Gender = "f", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "13-14", Distance = "3000", Gender = "f", Day = day });

            lastRow = WriteBigHeader(ref sheet, lastRow, "Terrängserien 2019, xx maj", "Pojkar", "3000 m");
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "0-4", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "5-6", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "7-8", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "9-10", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "11-12", Distance = "3000", Gender = "p", Day = day });
            lastRow = WriteGroupResult(ref sheet, lastRow, ref persons, ref centerStyle, ref leftStyle, day, new ResultFilter { Klass = "13-14", Distance = "3000", Gender = "p", Day = day });
        }

        public static void Write(ref string filePath, ref IList<Person> persons)
        {
            XSSFWorkbook workbook = new XSSFWorkbook();

            XSSFFont arialFont = (XSSFFont)workbook.CreateFont();
            arialFont.FontHeightInPoints = (short)10;
            arialFont.FontName = "Arial";

            XSSFCellStyle centerStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            centerStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            centerStyle.SetFont(arialFont);

            XSSFCellStyle leftStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            leftStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            leftStyle.SetFont(arialFont);

            ISheet attendanceSheet = workbook.CreateSheet();
            WriteAttendance(ref attendanceSheet, ref persons, ref centerStyle, ref leftStyle);

            ISheet sheet0 = workbook.CreateSheet();
            WriteResult(ref sheet0, ref persons, ref centerStyle, ref leftStyle, 0);

            ISheet sheet1 = workbook.CreateSheet();
            WriteResult(ref sheet1, ref persons, ref centerStyle, ref leftStyle, 1);

            ISheet sheet2 = workbook.CreateSheet();
            WriteResult(ref sheet2, ref persons, ref centerStyle, ref leftStyle, 2);

            ISheet sheet3 = workbook.CreateSheet();
            WriteResult(ref sheet3, ref persons, ref centerStyle, ref leftStyle, 3);

            ISheet sheet4 = workbook.CreateSheet();
            WriteResult(ref sheet4, ref persons, ref centerStyle, ref leftStyle, 4);

            ISheet sheet5 = workbook.CreateSheet();
            WriteResult(ref sheet5, ref persons, ref centerStyle, ref leftStyle, 5);

            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }
            workbook.Close();
        }
    }
}

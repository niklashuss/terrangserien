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
                    Person person = new Person();

                    person.Distance = row.GetCell(0).ToString().Trim();
                    person.Gender = row.GetCell(1).StringCellValue.Trim();
                    person.SocialNumber = row.GetCell(2).ToString().Trim();
                    person.Number = row.GetCell(3).ToString().Trim();
                    person.Name = row.GetCell(4).StringCellValue.Trim();
                    person.Surname = row.GetCell(5).StringCellValue.Trim();
                    person.Klass = row.GetCell(6).StringCellValue.Trim();
                    person.Result0 = row.GetCell(7).NumericCellValue.ToString().Trim();
                    person.Result1 = row.GetCell(8).NumericCellValue.ToString().Trim();
                    person.Result2 = row.GetCell(9).NumericCellValue.ToString().Trim();
                    person.Result3 = row.GetCell(10).NumericCellValue.ToString().Trim();
                    person.Result4 = row.GetCell(11).NumericCellValue.ToString().Trim();
                    person.Result5 = row.GetCell(12).NumericCellValue.ToString().Trim();
                    persons.Add(person);
                }
                workbook.Close();
            }
            return persons;
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

            ISheet sheet = workbook.CreateSheet();
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
                CreateCellWithStyle(ref row, 7, person.Result0, ref leftStyle);
                CreateCellWithStyle(ref row, 8, person.Result1, ref leftStyle);
                CreateCellWithStyle(ref row, 9, person.Result2, ref leftStyle);
                CreateCellWithStyle(ref row, 10, person.Result3, ref leftStyle);
                CreateCellWithStyle(ref row, 11, person.Result4, ref leftStyle);
                CreateCellWithStyle(ref row, 12, person.Result5, ref leftStyle);
                i++;
            }
            using (FileStream file = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(file);
            }
            workbook.Close();
        }

    }
}

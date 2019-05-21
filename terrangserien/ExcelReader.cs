using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.IO;

namespace terrangserien
{
    class ExcelReader
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
    }
}

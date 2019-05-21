﻿using NPOI.SS.UserModel;
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
                    person.Result(0, Result.Create(row.GetCell(7).NumericCellValue));
                    person.Result(1, Result.Create(row.GetCell(8).NumericCellValue));
                    person.Result(2, Result.Create(row.GetCell(9).NumericCellValue));
                    person.Result(3, Result.Create(row.GetCell(10).NumericCellValue));
                    person.Result(4, Result.Create(row.GetCell(11).NumericCellValue));
                    person.Result(5, Result.Create(row.GetCell(12).NumericCellValue));
                    persons.Add(person);
                }
                workbook.Close();
            }
            return persons;
        }
    }
}

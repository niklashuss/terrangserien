using System.IO;
using System.Collections.Generic;
using Serilog;

namespace terrangserien
{
    class IntermediateReaderWriter
    {
        static public IList<Person> Read(string filePath)
        {
            Log.Logger.Information("Reading {filePath}", filePath);
            IList<Person> persons = new List<Person>();
            using (StreamReader file = new StreamReader(filePath))
            {
                string line;
                while ((line = file.ReadLine()) != null)
                {
                    int i = 0;
                    string[] entries = line.Split(';');
                    Person person = Person.Create();
                    person.Name = entries[i++];
                    person.Surname = entries[i++];
                    person.Distance = entries[i++];
                    person.Gender = entries[i++];
                    person.SocialNumber = entries[i++];
                    person.Number = entries[i++];
                    person.Klass = entries[i++];
                    person.Result(0, Result.Create(ref entries[i++]));
                    person.Result(1, Result.Create(ref entries[i++]));
                    person.Result(2, Result.Create(ref entries[i++]));
                    person.Result(3, Result.Create(ref entries[i++]));
                    person.Result(4, Result.Create(ref entries[i++]));
                    person.Result(5, Result.Create(ref entries[i++]));
                    persons.Add(person);
                }
            }
            Log.Logger.Information("Read {0} entries", persons.Count);
            return persons;
        }

        static public void Write(string filePath, ref IList<Person> persons)
        {
            Log.Logger.Information("Writing ${filePath}", filePath);
            using (StreamWriter file = new StreamWriter(filePath))
            {
                foreach (Person person in persons)
                {
                    file.WriteLine("{0};{1};{2};{3};{4};{5};{6};{7};{8};{9};{10};{11};{12}",
                        person.Name,
                        person.Surname,
                        person.Distance,
                        person.Gender,
                        person.SocialNumber,
                        person.Number,
                        person.Klass,
                        person.Result(0),
                        person.Result(1),
                        person.Result(2),
                        person.Result(3),
                        person.Result(4),
                        person.Result(5)
                        );
                }
            }
            Log.Logger.Information("Wrote {0} entries", persons.Count);
        }
    }
}

using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace terrangserien
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IList<Person> persons;
        IList<Person> filteredPersons;

        const string INTERMEDIATE_FILE_NAME = "terrangserien.csv";

        public MainWindow()
        {
            InitializeComponent();

            Log.Logger = new LoggerConfiguration()
                .WriteTo.Console()
                .WriteTo.Debug()
                .WriteTo.File("log-.txt", rollingInterval: RollingInterval.Day)
                .CreateLogger();

            Log.Logger.Information("Starting application");
//            string inFilePath = @"Terrängserie 2019_14_maj.xlsx";
//            persons = ExcelReader.Read(ref inFilePath);

//            IntermediateReaderWriter.Write("terrangserien.csv", ref persons);
            persons = IntermediateReaderWriter.Read(INTERMEDIATE_FILE_NAME);
            string outFilePath = @"Terrängserien_new.xlsx";
            ExcelWriter.Write(ref outFilePath, ref persons);

            DataGridPersons.ItemsSource = persons;
            filteredPersons = persons;
        }

        private IList<Person> CreateFilter()
        {
            string name = TextBox_Name.Text.ToLower().Trim();
            string surname = TextBox_Surname.Text.ToLower().Trim();
            string gender = TextBox_Gender.Text.ToLower().Trim();
            string socialNumber = TextBox_SocialNumber.Text.ToLower().Trim();
            string number = TextBox_Number.Text.ToLower().Trim();
            string klass = TextBox_Klass.Text.ToLower().Trim();
            string distance = TextBox_Distance.Text.ToLower().Trim();

            PersonFilter filter = new PersonFilter {
                Name = name,
                Surname = surname,
                Gender = gender,
                SocialNumber = socialNumber,
                Number = number,
                Klass = klass,
                Distance = distance,
            };
            filteredPersons = persons.FilterPersons(filter);
            return filteredPersons;
        }

        private void TextBox_Name_TextChanged(object sender, TextChangedEventArgs e)
        {
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Surname_TextChanged(object sender, TextChangedEventArgs e)
        {
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Gender_TextChanged(object sender, TextChangedEventArgs e)
        {
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_SocialNumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Number_TextChanged(object sender, TextChangedEventArgs e)
        {
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Klass_TextChanged(object sender, TextChangedEventArgs e)
        {
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Distance_TextChanged(object sender, TextChangedEventArgs e)
        {
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void Button_Add_Click(object sender, RoutedEventArgs e)
        {
            string name = TextBox_Name.Text;
            if (name.Length == 0)
            {
                return;
            }
            string surname = TextBox_Surname.Text;
            if (surname.Length == 0)
            {
                return;
            }
            string gender = TextBox_Gender.Text;
            if (gender.Length == 0)
            {
                return;
            }
            string socialNumber = TextBox_SocialNumber.Text;
            if (socialNumber.Length == 0)
            {
                return;
            }
            string number = TextBox_Number.Text;
            if (number.Length == 0)
            {
                return;
            }
//            DataGridPersons.ItemsSource = items;
            Console.WriteLine("NIKLAS. button pressed");
        }

        private void UpdatePersonFromRowColumn(int row, int col, string value)
        {
            Person tempPerson = filteredPersons.ElementAt(row);
            int realRow = tempPerson.Id;
            Person person = persons.ElementAt(realRow);
            Log.Logger.Information("Före uppdateringen {Namn}, {Efternamn}, {Sträcka}, {Kön}, {Personnummer}, {Nummer}, {Klass}, {Result0}, {Result1}, {Result2}, {Result3}, {Result4}, {Result5}", 
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

            // stupid...
            if (col == 0)
            {
                person.Name = value;
            }
            else if (col == 1)
            {
                person.Surname = value;
            }
            else if (col == 2)
            {
                person.Gender = value;
            }
            else if (col == 3)
            {
                person.SocialNumber = value;
            }
            else if (col == 4)
            {
                person.Distance = value;
            }
            else if (col == 5)
            {
                person.Number = value;
            }
            else if (col == 6)
            {
                person.Klass = value;
            }
            else if (col == 7)
            {
                person.Result(0, Result.Create(ref value));
            }
            else if (col == 8)
            {
                person.Result(1, Result.Create(ref value));
            }
            else if (col == 9)
            {
                person.Result(2, Result.Create(ref value));
            }
            else if (col == 10)
            {
                person.Result(3, Result.Create(ref value));
            }
            else if (col == 11)
            {
                person.Result(4, Result.Create(ref value));
            }
            else if (col == 12)
            {
                person.Result(5, Result.Create(ref value));
            }

            persons[realRow] = person;

            Log.Logger.Information("Efter uppdateringen {Namn}, {Efternamn}, {Sträcka}, {Kön}, {Personnummer}, {Nummer}, {Klass}, {Result0}, {Result1}, {Result2}, {Result3}, {Result4}, {Result5}",
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

            IntermediateReaderWriter.Write(INTERMEDIATE_FILE_NAME, ref persons);
        }

        private void DataGridPersons_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int row = e.Row.GetIndex();
            int col = e.Column.DisplayIndex;

            var firstRow = (e.Row.Item as DataGridRow);
            var editedTextbox = e.EditingElement as TextBox;
            UpdatePersonFromRowColumn(row, col, editedTextbox.Text);
        }
    }
}

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
//            persons = ExcelReaderWriter.Read(ref inFilePath);

//            IntermediateReaderWriter.Write("terrangserien.csv", ref persons);
            persons = IntermediateReaderWriter.Read(INTERMEDIATE_FILE_NAME);
            /*
                        string outFilePath = @"Terrängserien_new.xlsx";
                        ExcelReaderWriter.Write(ref outFilePath, ref persons);
                        */
            DataGridPersons.ItemsSource = persons;
        }

        private IList<Person> CreateFilter()
        {
            string name = TextBox_Name.Text.ToLower().Trim();
            string surname = TextBox_Surname.Text.ToLower().Trim();
            string gender = TextBox_Gender.Text.ToLower().Trim();
            string socialNumber = TextBox_SocialNumber.Text.ToLower().Trim();
            string number = TextBox_Number.Text.ToLower().Trim();

            PersonFilter filter = new PersonFilter {
                Name = name,
                Surname = surname,
                Gender = gender,
                SocialNumber = socialNumber,
                Number = number,
            };
            return persons.FilterPersons(filter);
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
            Person person = persons.ElementAt(row);
            Log.Logger.Information("Före uppdateringen {Namn}, {Efternamn}, {Sträcka}, {Kön}, {Personnummer}, {Nummer}, {Klass}, {Result0}, {Result1}, {Result2}, {Result3}, {Result4}, {Result5}", 
                person.Name, 
                person.Surname, 
                person.Distance, 
                person.Gender, 
                person.SocialNumber,
                person.Number,
                person.Klass,
                person.Result0,
                person.Result1,
                person.Result2,
                person.Result3,
                person.Result4,
                person.Result5
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
                person.Result0 = value;
            }
            else if (col == 8)
            {
                person.Result1 = value;
            }
            else if (col == 9)
            {
                person.Result2 = value;
            }
            else if (col == 10)
            {
                person.Result3 = value;
            }
            else if (col == 11)
            {
                person.Result4 = value;
            }
            else if (col == 12)
            {
                person.Result5 = value;
            }

            persons[row] = person;

            Log.Logger.Information("Efter uppdateringen {Namn}, {Efternamn}, {Sträcka}, {Kön}, {Personnummer}, {Nummer}, {Klass}, {Result0}, {Result1}, {Result2}, {Result3}, {Result4}, {Result5}",
                person.Name,
                person.Surname,
                person.Distance,
                person.Gender,
                person.SocialNumber,
                person.Number,
                person.Klass,
                person.Result0,
                person.Result1,
                person.Result2,
                person.Result3,
                person.Result4,
                person.Result5
                );

            IntermediateReaderWriter.Write(INTERMEDIATE_FILE_NAME, ref persons);
        }

        private void DataGridPersons_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            int row = e.Row.GetIndex();
            int col = e.Column.DisplayIndex;
            var editedTextbox = e.EditingElement as TextBox;
            UpdatePersonFromRowColumn(row, col, editedTextbox.Text);
        }
    }
}

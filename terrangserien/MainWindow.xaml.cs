using Microsoft.Win32;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

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
            persons = IntermediateReaderWriter.Read(INTERMEDIATE_FILE_NAME);

            DataGridPersons.ItemsSource = persons;
            filteredPersons = persons;
            foreach (Person person in persons)
            {
                string klass = LooseFunctions.GetClassFromSocialNumber(person.SocialNumber);
                if (klass != "")
                {
                    person.Klass = klass;
                }
            }
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


        private void SetFieldAsError(ref TextBox textBox)
        {
            textBox.Background = Brushes.Red;
        }

        private void SetFieldAsOk(ref TextBox textBox)
        {
            textBox.Background = Brushes.White;
        }

        private void SetAllFieldsAsOk()
        {
            SetFieldAsOk(ref TextBox_Name);
            SetFieldAsOk(ref TextBox_Surname);
            SetFieldAsOk(ref TextBox_Gender);
            SetFieldAsOk(ref TextBox_SocialNumber);
            SetFieldAsOk(ref TextBox_Number);
        }

        private void ClearAllFields()
        {
            TextBox_Name.Text = "";
            TextBox_Surname.Text = "";
            TextBox_Gender.Text = "";
            TextBox_SocialNumber.Text = "";
            TextBox_Number.Text = "";
            TextBox_Klass.Text = "";
            TextBox_Distance.Text = "";
        }

        private void TextBox_Name_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetFieldAsOk(ref TextBox_Name);
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Surname_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetFieldAsOk(ref TextBox_Surname);
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Gender_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetFieldAsOk(ref TextBox_Gender);
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_SocialNumber_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetFieldAsOk(ref TextBox_SocialNumber);
            DataGridPersons.ItemsSource = CreateFilter();
        }

        private void TextBox_Number_TextChanged(object sender, TextChangedEventArgs e)
        {
            SetFieldAsOk(ref TextBox_Number);
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
            bool hasError = false;
            if (name.Length == 0)
            {
                SetFieldAsError(ref TextBox_Name);
                hasError = true;
            }

            string surname = TextBox_Surname.Text;
            if (surname.Length == 0)
            {
                SetFieldAsError(ref TextBox_Surname);
                hasError = true;
            }

            string gender = TextBox_Gender.Text;
            if (gender.Length == 0)
            {
                SetFieldAsError(ref TextBox_Gender);
                hasError = true;
            }

            string socialNumber = TextBox_SocialNumber.Text;
            if (socialNumber.Length == 0)
            {
                SetFieldAsError(ref TextBox_SocialNumber);
                hasError = true;
            }

            string number = TextBox_Number.Text;
            if (number.Length == 0)
            {
                SetFieldAsError(ref TextBox_Number);
                hasError = true;
            }

            if (hasError)
            {
                return;
            }

            Person person = Person.Create();
            person.Name = name;
            person.Surname = surname;
            person.SocialNumber = socialNumber;
            person.Gender = gender;
            person.Number = number;
            person.Klass = LooseFunctions.GetClassFromSocialNumber(socialNumber);
            persons.Add(person);

            TextBox_Name.Text = "";
            TextBox_Surname.Text = "";
            TextBox_Gender.Text = "";
            TextBox_SocialNumber.Text = "";
            TextBox_Number.Text = "";
            DataGridPersons.ItemsSource = persons;
            DataGridPersons.ScrollIntoView(person);
        }

        private void Button_Clear_Click(object sender, RoutedEventArgs e)
        {
            SetAllFieldsAsOk();
            ClearAllFields();
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
                person.Klass = LooseFunctions.GetClassFromSocialNumber(person.SocialNumber);
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
                person.Result(0, Result.Create(value));
            }
            else if (col == 8)
            {
                person.Result(1, Result.Create(value));
            }
            else if (col == 9)
            {
                person.Result(2, Result.Create(value));
            }
            else if (col == 10)
            {
                person.Result(3, Result.Create(value));
            }
            else if (col == 11)
            {
                person.Result(4, Result.Create(value));
            }
            else if (col == 12)
            {
                person.Result(5, Result.Create(value));
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

        private void MenuItem_Open_Excel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel (*.xlsx)|*.xlsx";
            openFileDialog.InitialDirectory = System.AppDomain.CurrentDomain.BaseDirectory;
            if (openFileDialog.ShowDialog() == true)
            {
                Log.Logger.Information("File selected: {0}", openFileDialog.FileName);
                string filePath = openFileDialog.FileName;
                persons = ExcelReader.Read(ref filePath);
                filteredPersons = persons;
                DataGridPersons.ItemsSource = persons;
                IntermediateReaderWriter.Write("terrangserien.csv", ref persons);
            }
        }

        private void MenuItem_Save_Excel_Click(object sender, RoutedEventArgs e)
        {
            string outFilePath = @"Terrängserien_new.xlsx";
            ExcelWriter.Write(ref outFilePath, ref persons);
            MessageBox.Show("Sparade filen till en ny Excel-fil", "Sparad");
        }

        private void MenuItem_Quit_Click(object sender, RoutedEventArgs e)
        {
            Log.Logger.Information("Closing program.");
            Close();
        }
    }
}

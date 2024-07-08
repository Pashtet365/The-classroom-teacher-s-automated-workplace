using Microsoft.Office.Interop.Word;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Documents;
using System.Windows.Media;
using DataTable = System.Data.DataTable;
using Window = System.Windows.Window;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;
using Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Word.Range;
using System.Text;
using Application = Microsoft.Office.Interop.Word.Application;

namespace ClassroomTeachersWorkplace
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        public int tableIndex;

        //глобальные переменные
        public static int changing = 0;
        public static object element1;
        public static object element2;
        public static object element3;
        public static object element4;
        public static object element5;
        public static object element6;
        public static object element7;
        public static object element8;
        public static object element9;
        public static object element10;
        public static object element11;
        public static object element12;
        public static object element13;
        public static object element14;

        public System.Windows.Controls.DataGrid pavlyuchkov;

        private List<UIElement> comboBoxElements;

        private List<UIElement> documentElements;

        public MainWindow()
        {
            InitializeComponent();

            //collection for filtration
            comboBoxElements = new List<UIElement>();
            foreach (UIElement element in menuFilter.Items)
            {
                comboBoxElements.Add(element);
            }
            menuFilter.Items.Clear();

            documentElements = new List<UIElement>();
            foreach (UIElement element in documentsMenu.Items)
            {
                documentElements.Add(element);
            }
            documentsMenu.Items.Clear();

            documentsMenu.Visibility = Visibility.Collapsed;
        }


        //-----------------ВЫВОД ТАБЛИЦ-----------------

        private void LoadData(DatabaseConnection dbConnection, string query)
        {
            try
            {
                SqlConnection connection = dbConnection.GetConnection();
                if (dbConnection.OpenConnection())
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    adapter.Fill(dataTable);
                    dataGridForm.ItemsSource = dataTable.DefaultView;
                    dbConnection.CloseConnection();
                }
                else
                {
                    MessageBox.Show("Failed to open connection.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        //предметы - 1
        private void menuTableItems_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[3]);

            nameTable.Header = "Предметы";

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();
            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT id_предмета AS ID, название AS Название FROM Предметы";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 1;
        }

        //отметки - 2
        private void menuTableMarks_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            nameTable.Header = "Отметки";

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[1]);
            documentsMenu.Items.Add(documentElements[3]);
            documentsMenu.Items.Add(documentElements[4]);

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT o.id_отметки AS ID, " +
                           "u.фио AS 'ФИО Учащегося', " +
                           "p.название AS 'Название предмета', " +
                           "o.отметка AS Отметка, " +
                           "FORMAT(o.дата, 'dd.MM.yyyy') AS Дата " +
                           "FROM Отметки o " +
                           "INNER JOIN Учащийся u ON o.id_учащегося = u.id_учащегося " +
                           "INNER JOIN Предметы p ON o.id_предмета = p.id_предмета";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 2;
        }

        // учащиеся - 3
        private void menuTableStudent_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            nameTable.Header = "Учащиеся";

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[6]);
            documentsMenu.Items.Add(documentElements[7]);
            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT id_учащегося AS ID, " +
                           "фио AS 'ФИО учащегося', " +
                           "телефон AS Телефон, " +
                           "FORMAT(дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', " +
                           "соц_бытовые_условия AS 'Социальные и бытовые условия', " +
                           "материальное_состояние AS 'Материальное состояние', " +
                           "трудности_в_учёбе AS 'Трудности в учебе', " +
                           "состояние_на_учёте AS 'Состояние на учете', " +
                           "нарушение_общения AS 'Нарушение общения', " +
                           "инвалидность AS 'Инвалидность', " +
                           "группа_здоровья AS 'Группа здоровья', " +
                           "чаэс AS 'ЧАЭС', " +
                           "кружки AS 'Кружки', " +
                           "сирота AS 'Сирота' " +
                           "FROM Учащийся";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 3;
        }

        //пропуски - 4
        private void menuTableSkip_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            nameTable.Header = "Пропуски";

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[2]);
            documentsMenu.Items.Add(documentElements[3]);
            documentsMenu.Items.Add(documentElements[4]);

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT Пропуски.id_пропуска AS ID, " +
                           "Учащийся.фио AS 'ФИО учащегося', " +
                           "Пропуски.количество_часов AS 'Количество часов', " +
                           "FORMAT(Пропуски.дата, 'dd.MM.yyyy') AS 'Дата' " +
                           "FROM Пропуски " +
                           "INNER JOIN Учащийся ON Пропуски.id_учащегося = Учащийся.id_учащегося";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 4;
        }

        //справки - 5
        private void menuTableLisainces_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            nameTable.Header = "Справки";

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[5]);
            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT Справка.id_справки AS ID, " +
                           "Учащийся.фио AS 'ФИО учащегося', " +
                           "FORMAT(Справка.дата_начала, 'dd.MM.yyyy') AS 'Дата начала', " +
                           "FORMAT(Справка.дата_конца, 'dd.MM.yyyy') AS 'Дата конца', " +
                           "Справка.вид_справки AS 'Вид справки' " +
                           "FROM Справка " +
                           "INNER JOIN Учащийся ON Справка.id_учащегося = Учащийся.id_учащегося";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 5;
        }

        // родители - 6
        private void menuTableParents_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            nameTable.Header = "Родители";

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();
            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT Родители.id_родителя AS ID, " +
                           "Учащийся.фио AS 'ФИО учащегося', " +
                           "Родители.фио AS 'ФИО родителя', " +
                           "Родители.пол AS 'Семейное положение', " +
                           "FORMAT(Родители.дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', " +
                           "Родители.место_работы AS 'Место работы', " +
                           "Родители.должность AS Должность, " +
                           "Родители.адресс AS Адрес " +
                           "FROM Родители " +
                           "INNER JOIN Учащийся ON Родители.id_учащегося = Учащийся.id_учащегося";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 6;
        }

        //собрания - 7
        private void menuTableEvents_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Visible;
            menuFilter.Items.Clear();
            menuFilter.Items.Add(comboBoxElements[0]);
            menuFilter.Items.Add(comboBoxElements[1]);
            menuFilter.Items.Add(comboBoxElements[2]);
            menuFilter.Items.Add(comboBoxElements[3]);

            nameTable.Header = "Собрания";

            documentsMenu.Visibility = Visibility.Visible;
            documentsMenu.Items.Clear();
            documentsMenu.Items.Add(documentElements[0]);
            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT Собрание.id_собрания AS ID, " +
                           "Родители.фио AS 'ФИО родителя', " +
                           "FORMAT(Собрание.дата, 'dd.MM.yyyy') AS 'Дата', " +
                           "Собрание.тема AS 'Тема' " +
                           "FROM Собрание " +
                           "INNER JOIN Родители ON Собрание.id_родителя = Родители.id_родителя";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 7;
        }

        // класс - 8
        private void menuTableClass_Click(object sender, RoutedEventArgs e)
        {
            menuFilter.Visibility = Visibility.Collapsed;
            menuFilter.Items.Clear();

            nameTable.Header = "Класс";

            documentsMenu.Visibility = Visibility.Collapsed;
            documentsMenu.Items.Clear();

            // Проверка текущей даты
            DateTime currentDate = DateTime.Now;
            if (currentDate.Month == 9 && currentDate.Day == 1)
            {
                UpdateClassNumbers();
            }

            DatabaseConnection dbConnection = new DatabaseConnection();
            string query = "SELECT id_класса AS ID, наименование AS Название FROM Класс";
            LoadData(dbConnection, query);
            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
            tableIndex = 8;
        }

        // Метод для обновления номеров классов
        private void UpdateClassNumbers()
        {
            using (DatabaseConnection dbConnection = new DatabaseConnection())
            {
                if (dbConnection.OpenConnection())
                {
                    // Получение всех классов
                    string selectQuery = "SELECT id_класса, наименование, дата_последнего_обновления FROM Класс";
                    SqlCommand selectCommand = new SqlCommand(selectQuery, dbConnection.GetConnection());

                    DataTable classTable = new DataTable();
                    SqlDataAdapter adapter = new SqlDataAdapter(selectCommand);
                    adapter.Fill(classTable);

                    DateTime currentDate = DateTime.Now;
                    DateTime newSchoolYearStartDate = new DateTime(currentDate.Year, 5, 23);

                    foreach (DataRow row in classTable.Rows)
                    {
                        string className = row["наименование"].ToString();
                        int classId = (int)row["id_класса"];
                        DateTime? lastUpdatedDate = row.IsNull("дата_последнего_обновления") ? (DateTime?)null : (DateTime)row["дата_последнего_обновления"];

                        // Проверка, если обновление уже произошло в этом году
                        if (lastUpdatedDate.HasValue && lastUpdatedDate.Value >= newSchoolYearStartDate)
                        {
                            continue; // Пропустить обновление для этого класса
                        }

                        // Разделение номера класса и буквы
                        string numberPart = new string(className.TakeWhile(char.IsDigit).ToArray());
                        string letterPart = new string(className.SkipWhile(char.IsDigit).ToArray());

                        if (int.TryParse(numberPart, out int classNumber))
                        {
                            // Проверка, что класс не 11-й
                            if (classNumber < 11)
                            {
                                classNumber++; // Увеличение номера класса на 1

                                string newClassName = classNumber.ToString() + letterPart;

                                // Обновление класса в базе данных
                                string updateQuery = "UPDATE Класс SET наименование = @newName, дата_последнего_обновления = @updateDate WHERE id_класса = @classId";
                                SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                                updateCommand.Parameters.AddWithValue("@newName", newClassName);
                                updateCommand.Parameters.AddWithValue("@updateDate", currentDate);
                                updateCommand.Parameters.AddWithValue("@classId", classId);
                                updateCommand.ExecuteNonQuery();
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Ошибка подключения к базе данных.");
                }
            }
        }


        //-----------------КОНЕЦ ВЫВОД ТАБЛИЦ-----------------



        //-----------------ДОБАВЛЕНИЕ-----------------
        private void menuTableAddedRow_Click(object sender, RoutedEventArgs e)
        {
            switch (tableIndex)
            {
                case 1:
                    // Открыть новую форму Организации
                    Предметы itemsForm = new Предметы(this);
                    itemsForm.Owner = this;
                    itemsForm.ShowDialog();
                    break;
                case 2:
                    // Открыть новую форму Педагоги
                    Отметки marksForm = new Отметки(this);
                    marksForm.Owner = this;
                    marksForm.ShowDialog();
                    break;
                case 3:
                    // Открыть новую форму Участники
                    Учащийся studentForm = new Учащийся(this);
                    studentForm.Owner = this;
                    studentForm.ShowDialog();
                    break;
                case 4:
                    // Открыть новую форму Предметы
                    Пропуски skipForm = new Пропуски(this);
                    skipForm.Owner = this;
                    skipForm.ShowDialog();
                    break;
                case 5:
                    // Открыть новую форму Этапы
                    Справки spravkyForm = new Справки(this);
                    spravkyForm.Owner = this;
                    spravkyForm.ShowDialog();
                    break;
                case 6:
                    // Открыть новую форму Мероприятия
                    Родители parentsForm = new Родители(this);
                    parentsForm.Owner = this;
                    parentsForm.ShowDialog();
                    break;
                case 7:
                    // Открыть новую форму Заявки
                    Собрание applicationForm = new Собрание(this);
                    applicationForm.Owner = this;
                    applicationForm.ShowDialog();
                    break;
                case 8:
                    // Открыть новую форму Класса
                    Класс classForm = new Класс(this);
                    classForm.Owner = this;
                    classForm.ShowDialog();
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }
        //-----------------КОНЕЦ ДОБАВЛЕНИЕ-----------------



        //-----------------ИЗМЕНЕНИЕ-----------------
        private void menuTableChanging_Click(object sender, RoutedEventArgs e)
        {
            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;

            if (selectedRow == null)
            {
                MessageBox.Show("Выберите строку!");
                return;
            }

            switch (tableIndex)
            {
                case 1:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["Название"].ToString();
                    changing = 1;
                    Предметы itemForm = new Предметы(this);
                    itemForm.Owner = this;
                    itemForm.ShowDialog();
                    break;
                case 2:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО Учащегося"].ToString();
                    element3 = selectedRow["Название предмета"].ToString();
                    element4 = selectedRow["Отметка"].ToString();
                    element5 = selectedRow["Дата"].ToString();
                    changing = 1;
                    Отметки marksForm = new Отметки(this);
                    marksForm.Owner = this;
                    marksForm.ShowDialog();
                    break;
                case 3:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО учащегося"].ToString();
                    element3 = selectedRow["Телефон"].ToString();
                    element4 = selectedRow["Дата рождения"].ToString();
                    element5 = selectedRow["Социальные и бытовые условия"].ToString();
                    element6 = selectedRow["Материальное состояние"].ToString();
                    element7 = selectedRow["Трудности в учебе"].ToString();
                    element8 = selectedRow["Состояние на учете"].ToString();
                    element9 = selectedRow["Нарушение общения"].ToString();
                    element10 = selectedRow["Инвалидность"].ToString();
                    element11 = selectedRow["Группа здоровья"].ToString();
                    element12 = selectedRow["ЧАЭС"].ToString();
                    element13 = selectedRow["Кружки"].ToString();
                    element14 = selectedRow["Сирота"].ToString();
                    changing = 1;
                    Учащийся studentForm = new Учащийся(this);
                    studentForm.Owner = this;
                    studentForm.ShowDialog();
                    break;
                case 4:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО учащегося"].ToString();
                    element3 = selectedRow["Количество часов"].ToString();
                    element4 = selectedRow["Дата"].ToString();
                    changing = 1;
                    Пропуски subjectForm = new Пропуски(this);
                    subjectForm.Owner = this;
                    subjectForm.ShowDialog();
                    break;
                case 5:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО учащегося"].ToString();
                    element3 = selectedRow["Дата начала"].ToString();
                    element4 = selectedRow["Дата конца"].ToString();
                    element5 = selectedRow["Вид справки"].ToString();
                    changing = 1;
                    Справки spravkyForm = new Справки(this);
                    spravkyForm.Owner = this;
                    spravkyForm.ShowDialog();
                    break;
                case 6:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО учащегося"].ToString();
                    element3 = selectedRow["ФИО родителя"].ToString();
                    element4 = selectedRow["Семейное положение"].ToString();
                    element5 = selectedRow["Дата рождения"].ToString();
                    element6 = selectedRow["Место работы"].ToString();
                    element7 = selectedRow["Должность"].ToString();
                    element8 = selectedRow["Адрес"].ToString();
                    changing = 1;
                    // Открыть новую форму Родители
                    Родители parentsForm = new Родители(this);
                    parentsForm.Owner = this;
                    parentsForm.ShowDialog();
                    break;
                case 7:
                    // Передача данных из выбранной строки таблицы "Заявки" в переменные класса MainWindow
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["ФИО родителя"].ToString();
                    element3 = selectedRow["Дата"].ToString();
                    element4 = selectedRow["Тема"].ToString();
                    // Установка значения changing для указания на редактирование
                    changing = 1;
                    // Открытие новой формы Заявки
                    Собрание applicationForm = new Собрание(this);
                    applicationForm.Owner = this;
                    applicationForm.ShowDialog();
                    break;
                case 8:
                    element1 = selectedRow["ID"].ToString();
                    element2 = selectedRow["Название"].ToString();
                    changing = 1;
                    Класс classForm = new Класс(this);
                    classForm.Owner = this;
                    classForm.ShowDialog();
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }
        //-----------------КОНЕЦ ИЗМЕНЕНИЕ-----------------


        //-----------------ОБНОВЛЕНИЕ-----------------
        public void menuTableRefresh_Click(object sender, RoutedEventArgs e)
        {
            Refresh(sender, e);
        }

        public void Refresh(object sender, RoutedEventArgs e)
        {
            switch (tableIndex)
            {
                case 1:
                    menuTableItems_Click(sender, e);
                    break;
                case 2:
                    menuTableMarks_Click(sender, e);
                    break;
                case 3:
                    menuTableStudent_Click(sender, e);
                    break;
                case 4:
                    menuTableSkip_Click(sender, e);
                    break;
                case 5:
                    menuTableLisainces_Click(sender, e);
                    break;
                case 6:
                    menuTableParents_Click(sender, e);
                    break;
                case 7:
                    menuTableEvents_Click(sender, e);
                    break;
                case 8:
                    menuTableClass_Click(sender, e);
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }
        //-----------------КОНЕЦ ОБНОВЛЕНИЕ-----------------

        //-----------------УДАЛЕНИЕ-----------------
        private void menuTableDelete_Click(object sender, RoutedEventArgs e)
        {
            // Получение выбранной строки из DataGrid
            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;

            // Проверка, что строка действительно выбрана
            if (selectedRow != null)
            {
                try
                {
                    // Получение идентификатора из первой колонки (предполагается, что идентификатор находится в первой колонке)
                    int idToDelete = Convert.ToInt32(selectedRow[0]);

                    // Выполнение операции удаления в базе данных в зависимости от выбранной таблицы
                    DatabaseConnection dbConnection = new DatabaseConnection();
                    SqlConnection connection = dbConnection.GetConnection();
                    if (dbConnection.OpenConnection())
                    {
                        SqlCommand command = null;
                        switch (tableIndex)
                        {
                            case 1:
                                command = new SqlCommand("DELETE FROM Предметы WHERE id_предмета = @ID", connection);
                                break;
                            case 2:
                                command = new SqlCommand("DELETE FROM Отметки WHERE id_отметки = @ID", connection);
                                break;
                            case 3:
                                command = new SqlCommand("DELETE FROM Учащийся WHERE id_учащегося = @ID", connection);
                                break;
                            case 4:
                                command = new SqlCommand("DELETE FROM Пропуски WHERE id_пропуска = @ID", connection);
                                break;
                            case 5:
                                command = new SqlCommand("DELETE FROM Справка WHERE id_справки = @ID", connection);
                                break;
                            case 6:
                                command = new SqlCommand("DELETE FROM Родители WHERE id_родителя = @ID", connection);
                                break;
                            case 7:
                                command = new SqlCommand("DELETE FROM Собрание WHERE id_собрания = @ID", connection);
                                break;
                            case 8:
                                command = new SqlCommand("DELETE FROM Класс WHERE id_класса = @ID", connection);
                                break;
                            default:
                                MessageBox.Show("Выберите таблицу!");
                                return; // Прекращаем выполнение метода, так как нет команды для удаления
                        }

                        // Установка параметра и выполнение команды удаления
                        if (command != null)
                        {
                            command.Parameters.AddWithValue("@ID", idToDelete);
                            command.ExecuteNonQuery();
                        }

                        // Обновление DataGrid после удаления
                        Refresh(sender, e);

                        dbConnection.CloseConnection();
                    }
                    else
                    {
                        MessageBox.Show("Failed to open connection.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Данная запись связана с другой таблицей!");
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Выберите строку для удаления!");
            }
        }
        //-----------------КОНЕЦ УДАЛЕНИЕ-----------------


        //-----------------ВЫВОД В EXCEL-----------------
        private void printExsel_Click(object sender, RoutedEventArgs e)
        {
            if (tableIndex == 0)
            {
                MessageBox.Show("Выберите таблицу!");
                return;
            }

            string nameTable = string.Empty;

            // Устанавливаем имя таблицы в зависимости от выбора пользователя
            switch (tableIndex)
            {
                case 1:
                    nameTable = "Предметы";
                    break;
                case 2:
                    nameTable = "Отметки";
                    break;
                case 3:
                    nameTable = "Учащиеся";
                    break;
                case 4:
                    nameTable = "Пропуски";
                    break;
                case 5:
                    nameTable = "Справки";
                    break;
                case 6:
                    nameTable = "Родители";
                    break;
                case 7:
                    nameTable = "Собрания";
                    break;
                case 8:
                    nameTable = "Класс";
                    break;
            }

            // Создание объекта SaveFileDialog
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
            saveFileDialog.Title = "Сохранить как Excel";
            saveFileDialog.DefaultExt = "xlsx";

            if (saveFileDialog.ShowDialog() == true)
            {
                // Получение выбранного пользователем пути и имени файла
                string filePath = saveFileDialog.FileName;

                // Создание нового объекта приложения Excel
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Visible = false; // Скрываем Excel

                // Создание новой книги Excel
                Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Add(Type.Missing);

                // Создание нового листа Excel
                Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Sheets[1];

                // Заполнение листа данными из вашего DataGrid
                for (int i = 0; i < dataGridForm.Items.Count; i++)
                {
                    var dataGridRow = (DataGridRow)dataGridForm.ItemContainerGenerator.ContainerFromIndex(i);
                    if (dataGridRow != null)
                    {
                        for (int j = 0; j < dataGridForm.Columns.Count; j++)
                        {
                            var content = dataGridForm.Columns[j].GetCellContent(dataGridRow);
                            if (content is TextBlock)
                            {
                                var text = (content as TextBlock).Text;
                                excelSheet.Cells[i + 2, j + 1] = text; // Начинаем с второй строки

                                // Если это столбец с телефонным номером, установить формат ячейки в текстовый
                                if (dataGridForm.Columns[j].Header.ToString() == "Телефон")
                                {
                                    excelSheet.Cells[i + 2, j + 1].NumberFormat = "@";
                                }
                            }
                        }
                    }
                }


                // Удаление столбца A
                Microsoft.Office.Interop.Excel.Range columnA = (Microsoft.Office.Interop.Excel.Range)excelSheet.Columns["A"];
                columnA.Delete();

                // Объединение ячеек в первой строке
                Microsoft.Office.Interop.Excel.Range headerRange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, dataGridForm.Columns.Count - 1]];
                headerRange.Merge();

                // Установка текста в объединенной ячейке
                excelSheet.Cells[1, 1] = nameTable;

                // Выравнивание текста по центру и установка жирного шрифта для первой строки
                headerRange.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                headerRange.Font.Bold = true;

                // Добавление обводки для всей таблицы Excel
                Microsoft.Office.Interop.Excel.Range tableRange = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[dataGridForm.Items.Count + 1, dataGridForm.Columns.Count - 1]];
                tableRange.Borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                tableRange.Borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;

                // Выравнивание ширины столбцов
                excelSheet.UsedRange.Columns.AutoFit();

                // Сохранение книги Excel по выбранному пути
                excelBook.SaveAs(filePath);

                // Закрытие книги и приложения Excel
                excelBook.Close();
                excelApp.Quit();

                // Освобождение ресурсов COM
                Marshal.ReleaseComObject(excelSheet);
                Marshal.ReleaseComObject(excelBook);
                Marshal.ReleaseComObject(excelApp);
            }
        }
        //-----------------КОНЕЦ ВЫВОД В EXCEL-----------------


        //-----------------ФИЛЬТРАЦИЯ-----------------
        private void buttonFilter_Click(object sender, RoutedEventArgs e)
        {
            switch (tableIndex)
            {
                case 1: // Предметы
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = "SELECT id_предмета AS ID, название AS Название FROM Предметы WHERE название LIKE @FilterText";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                            // Здесь вы можете использовать dataTable для отображения результатов
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 2: //Отметки
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = "SELECT o.id_отметки AS ID, " +
                                          "u.фио AS 'ФИО Учащегося', " +
                                          "p.название AS 'Название предмета', " +
                                          "o.отметка AS Отметка, " +
                                          "FORMAT(o.дата, 'dd.MM.yyyy') AS Дата " +
                                          "FROM Отметки o " +
                                          "INNER JOIN Учащийся u ON o.id_учащегося = u.id_учащегося " +
                                          "INNER JOIN Предметы p ON o.id_предмета = p.id_предмета " +
                                          "WHERE (u.фио LIKE @FilterText OR p.название LIKE @FilterText) " +
                                          "AND (@StartDate IS NULL OR o.дата >= @StartDate) " +
                                          "AND (@EndDate IS NULL OR o.дата <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 3: // Учащиеся
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = "SELECT id_учащегося AS ID, " +
                                          "фио AS 'ФИО учащегося', " +
                                          "телефон AS Телефон, " +
                                          "FORMAT(дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', " +
                                          "соц_бытовые_условия AS 'Социальные и бытовые условия', " +
                                          "материальное_состояние AS 'Материальное состояние', " +
                                          "трудности_в_учёбе AS 'Трудности в учебе', " +
                                          "состояние_на_учёте AS 'Состояние на учете', " +
                                          "нарушение_общения AS 'Нарушение общения', " +
                                          "инвалидность AS 'Инвалидность', " +
                                          "группа_здоровья AS 'Группа здоровья', " +
                                          "чаэс AS 'ЧАЭС', " +
                                          "кружки AS 'Кружки', " +
                                          "сирота AS 'Сирота' " +
                                          "FROM Учащийся " +
                                          "WHERE (фио LIKE @FilterText OR " +
                                          "телефон LIKE @FilterText OR " +
                                          "FORMAT(дата_рождения, 'dd.MM.yyyy') LIKE @FilterText OR " +
                                          "соц_бытовые_условия LIKE @FilterText OR " +
                                          "материальное_состояние LIKE @FilterText OR " +
                                          "трудности_в_учёбе LIKE @FilterText OR " +
                                          "состояние_на_учёте LIKE @FilterText OR " +
                                          "нарушение_общения LIKE @FilterText OR " +
                                          "инвалидность LIKE @FilterText OR " +
                                          "группа_здоровья LIKE @FilterText OR " +
                                          "чаэс LIKE @FilterText OR " +
                                          "кружки LIKE @FilterText OR " +
                                          "сирота LIKE @FilterText) " +
                                          "AND (@StartDate IS NULL OR дата_рождения >= @StartDate) " +
                                          "AND (@EndDate IS NULL OR дата_рождения <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 4: //Пропуски
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = "SELECT Пропуски.id_пропуска AS ID, " +
                                          "Учащийся.фио AS 'ФИО учащегося', " +
                                          "Пропуски.количество_часов AS 'Количество часов', " +
                                          "FORMAT(Пропуски.дата, 'dd.MM.yyyy') AS 'Дата' " +
                                          "FROM Пропуски " +
                                          "INNER JOIN Учащийся ON Пропуски.id_учащегося = Учащийся.id_учащегося " +
                                          "WHERE (Учащийся.фио LIKE @FilterText OR " +
                                          "Пропуски.количество_часов LIKE @FilterText OR " +
                                          "FORMAT(Пропуски.дата, 'dd.MM.yyyy') LIKE @FilterText) " +
                                          "AND (@StartDate IS NULL OR Пропуски.дата >= @StartDate) " +
                                          "AND (@EndDate IS NULL OR Пропуски.дата <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 5: //Справки
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = "SELECT Справка.id_справки AS ID, " +
                                          "Учащийся.фио AS 'ФИО учащегося', " +
                                          "FORMAT(Справка.дата_начала, 'dd.MM.yyyy') AS 'Дата начала', " +
                                          "FORMAT(Справка.дата_конца, 'dd.MM.yyyy') AS 'Дата конца', " +
                                          "Справка.вид_справки AS 'Вид справки' " +
                                          "FROM Справка " +
                                          "INNER JOIN Учащийся ON Справка.id_учащегося = Учащийся.id_учащегося " +
                                          "WHERE (Учащийся.фио LIKE @FilterText OR " +
                                          "FORMAT(Справка.дата_начала, 'dd.MM.yyyy') LIKE @FilterText OR " +
                                          "FORMAT(Справка.дата_конца, 'dd.MM.yyyy') LIKE @FilterText OR " +
                                          "Справка.вид_справки LIKE @FilterText) " +
                                          "AND (@StartDate IS NULL OR Справка.дата_начала >= @StartDate) " +
                                          "AND (@EndDate IS NULL OR Справка.дата_конца <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 6: // Родители
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = "SELECT Родители.id_родителя AS ID, " +
                                          "Учащийся.фио AS 'ФИО учащегося', " +
                                          "Родители.фио AS 'ФИО родителя', " +
                                          "Родители.пол AS 'Семейное положение', " +
                                          "FORMAT(Родители.дата_рождения, 'dd.MM.yyyy') AS 'Дата рождения', " +
                                          "Родители.место_работы AS 'Место работы', " +
                                          "Родители.должность AS Должность, " +
                                          "Родители.адресс AS Адрес " +
                                          "FROM Родители " +
                                          "INNER JOIN Учащийся ON Родители.id_учащегося = Учащийся.id_учащегося " +
                                          "WHERE (Учащийся.фио LIKE @FilterText OR " +
                                          "Родители.фио LIKE @FilterText OR " +
                                          "Родители.пол LIKE @FilterText OR " +
                                          "FORMAT(Родители.дата_рождения, 'dd.MM.yyyy') LIKE @FilterText OR " +
                                          "Родители.место_работы LIKE @FilterText OR " +
                                          "Родители.должность LIKE @FilterText OR " +
                                          "Родители.адресс LIKE @FilterText) " +
                                          "AND (@StartDate IS NULL OR Родители.дата_рождения >= @StartDate) " +
                                          "AND (@EndDate IS NULL OR Родители.дата_рождения <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                case 7: // Собрание
                    try
                    {
                        string filterText = textBoxFilter.Text.Trim(); // Получаем текст фильтрации
                        DateTime? startDate = datePickerFilterFirstDate.SelectedDate;
                        DateTime? endDate = datePickerFilterLastDate.SelectedDate;

                        // Формируем SQL-запрос для фильтрации и выборки данных
                        string sqlQuery = "SELECT Собрание.id_собрания AS ID, " +
                                          "Родители.фио AS 'ФИО родителя', " +
                                          "FORMAT(Собрание.дата, 'dd.MM.yyyy') AS 'Дата', " +
                                          "Собрание.тема AS 'Тема' " +
                                          "FROM Собрание " +
                                          "INNER JOIN Родители ON Собрание.id_родителя = Родители.id_родителя " +
                                          "WHERE (Родители.фио LIKE @FilterText OR " +
                                          "FORMAT(Собрание.дата, 'dd.MM.yyyy') LIKE @FilterText OR " +
                                          "Собрание.тема LIKE @FilterText) " +
                                          "AND (@StartDate IS NULL OR Собрание.дата >= @StartDate) " +
                                          "AND (@EndDate IS NULL OR Собрание.дата <= @EndDate)";

                        // Создаем подключение к базе данных
                        using (DatabaseConnection dbConnection = new DatabaseConnection())
                        {
                            if (dbConnection.OpenConnection())
                            {
                                using (SqlCommand command = new SqlCommand(sqlQuery, dbConnection.GetConnection()))
                                {
                                    command.Parameters.AddWithValue("@FilterText", "%" + filterText + "%");
                                    command.Parameters.AddWithValue("@StartDate", startDate ?? (object)DBNull.Value);
                                    command.Parameters.AddWithValue("@EndDate", endDate ?? (object)DBNull.Value);

                                    try
                                    {
                                        using (SqlDataReader reader = command.ExecuteReader())
                                        {
                                            DataTable dataTable = new DataTable();
                                            dataTable.Load(reader);
                                            dataGridForm.ItemsSource = dataTable.DefaultView;
                                            // Скрываем колонку ID, если это нужно
                                            dataGridForm.Columns[0].Visibility = Visibility.Hidden;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        MessageBox.Show("Ошибка при выполнении запроса: " + ex.Message);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("Ошибка подключения к базе данных.");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Ошибка: " + ex.Message);
                    }
                    break;
                default:
                    MessageBox.Show("Выберите таблицу!");
                    break;
            }
        }
        //-----------------КОНЕЦ ФИЛЬТРАЦИЯ-----------------


        //-----------------ПОИСК-----------------
        private void buttonSearch_Click(object sender, RoutedEventArgs e)
        {
            string searchText = txtSearch.Text;

            if (searchText == "")
            {
                return;
            }
            if (string.IsNullOrWhiteSpace(searchText))
            {
                // Если поле поиска пустое, очистите подсветку и выходите
                ClearSearchHighlighting();
                return;
            }

            // Пройдитесь по всем строкам и ячейкам в DataGrid
            foreach (DataGridRow row in GetDataGridRows(dataGridForm))
            {
                foreach (DataGridColumn column in dataGridForm.Columns)
                {
                    if (column is DataGridTextColumn)
                    {
                        var cell = GetCell(row, column);
                        if (cell != null)
                        {
                            TextBlock textBlock = cell.Content as TextBlock;
                            if (textBlock != null)
                            {
                                string cellContent = textBlock.Text;
                                if (!string.IsNullOrEmpty(cellContent))
                                {
                                    if (cellContent.IndexOf(searchText, StringComparison.OrdinalIgnoreCase) >= 0)
                                    {
                                        // Если найдено совпадение, подсветите текст
                                        int index = cellContent.IndexOf(searchText, StringComparison.OrdinalIgnoreCase);
                                        string preMatch = cellContent.Substring(0, index);
                                        string match = cellContent.Substring(index, searchText.Length);
                                        string postMatch = cellContent.Substring(index + searchText.Length);

                                        textBlock.Inlines.Clear();
                                        textBlock.Inlines.Add(new Run(preMatch));
                                        Run matchRun = new Run(match);
                                        matchRun.Background = Brushes.Yellow; // Задайте цвет подсветки
                                        textBlock.Inlines.Add(matchRun);
                                        textBlock.Inlines.Add(new Run(postMatch));
                                    }
                                    else
                                    {
                                        // Если совпадение не найдено, очистите подсветку
                                        textBlock.Inlines.Clear();
                                        textBlock.Inlines.Add(new Run(cellContent));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }

        private void ClearSearchHighlighting()
        {
            // Очистите подсветку во всех ячейках DataGrid
            foreach (DataGridRow row in GetDataGridRows(dataGridForm))
            {
                foreach (DataGridColumn column in dataGridForm.Columns)
                {
                    if (column is DataGridTextColumn)
                    {
                        var cell = GetCell(row, column);
                        if (cell != null)
                        {
                            TextBlock textBlock = cell.Content as TextBlock;
                            if (textBlock != null)
                            {
                                textBlock.Inlines.Clear();
                                textBlock.Inlines.Add(new Run(textBlock.Text));
                            }
                        }
                    }
                }
            }
        }

        private System.Windows.Controls.DataGridCell GetCell(DataGridRow row, DataGridColumn column)
        {
            if (column != null)
            {
                DataGridCellsPresenter presenter = GetVisualChild<DataGridCellsPresenter>(row);
                if (presenter == null)
                    return null;

                int columnIndex = dataGridForm.Columns.IndexOf(column);
                if (columnIndex > -1)
                {
                    System.Windows.Controls.DataGridCell cell = (System.Windows.Controls.DataGridCell)presenter.ItemContainerGenerator.ContainerFromIndex(columnIndex);
                    return cell;
                }
            }
            return null;
        }

        private List<DataGridRow> GetDataGridRows(System.Windows.Controls.DataGrid grid)
        {
            List<DataGridRow> rows = new List<DataGridRow>();
            for (int i = 0; i < dataGridForm.Items.Count; i++)
            {
                DataGridRow row = (DataGridRow)dataGridForm.ItemContainerGenerator.ContainerFromIndex(i);
                if (row != null)
                {
                    rows.Add(row);
                }
            }
            return rows;
        }

        private childItem GetVisualChild<childItem>(DependencyObject obj) where childItem : DependencyObject
        {
            for (int i = 0; i < VisualTreeHelper.GetChildrenCount(obj); i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(obj, i);
                if (child != null && child is childItem)
                    return (childItem)child;
                else
                {
                    childItem childOfChild = GetVisualChild<childItem>(child);
                    if (childOfChild != null)
                        return childOfChild;
                }
            }
            return null;
        }
        //-----------------КОНЕЦ ПОИСК-----------------


        //печать протокола собрания
        private void protocol_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridForm.SelectedItem == null)
            {
                MessageBox.Show("Выберите собрание!");
                return;
            }

            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;
            string id = selectedRow["ID"].ToString();
            string date = selectedRow["Дата"].ToString();
            string nameEvent = selectedRow["Тема"].ToString();

            // Преобразуем формат даты из дд.мм.гггг в гггг.мм.дд
            DateTime parsedDate = DateTime.ParseExact(date, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            string formattedDate = parsedDate.ToString("yyyy-MM-dd");

            // Создаем объект подключения к базе данных
            using (DatabaseConnection dbConnection = new DatabaseConnection())
            {
                // Открываем соединение
                if (dbConnection.OpenConnection())
                {
                    // Создаем SQL-запрос для получения всех родителей
                    string query = @"SELECT фио FROM Родители";

                    // Создаем список для хранения имен родителей
                    List<string> parentsList = new List<string>();

                    // Создаем объект команды SQL
                    using (SqlCommand cmd = new SqlCommand(query, dbConnection.GetConnection()))
                    {
                        // Выполняем запрос и добавляем каждое имя родителя в список
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            // Добавляем каждое имя родителя в список, удаляя начальные и конечные пробелы и символы перевода строки
                            while (reader.Read())
                            {
                                string parentName = reader["фио"].ToString().Trim();
                                parentsList.Add(parentName);
                            }
                        }
                    }

                    // Создаем строку для хранения списка родителей с номерами
                    StringBuilder parentsWithNumbers = new StringBuilder();

                    // Добавляем номера перед каждым именем родителя
                    for (int i = 0; i < parentsList.Count; i++)
                    {
                        parentsWithNumbers.Append($"{i + 1}. {parentsList[i]} присутствовал(а) отсутствовал(а){Environment.NewLine}");
                    }

                    // laptop
                    // string templateFilePath = @"C:\GitHub\OpenAccess\GGAEK\4course\БД\КурсоваяБД\ИсходныеДокументы\Zayavka.doc";

                    // computer
                    string templateFilePath = @"G:\Files\ЗаконченыеПроекты\3КУРС\4course\Дипломы\Дубровский\Программа Дубровского\ClassroomTeachersWorkplace\ClassroomTeachersWorkplace\Documents\protocolParentsEvent.docx";

                    Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
                    Document doc = wordApp.Documents.Open(templateFilePath);
                    doc.Activate();

                    // Замена данных в документе
                    FindAndReplace(wordApp, "[дата]", date);
                    FindAndReplace(wordApp, "[тема]", nameEvent);
                    FindAndReplace(wordApp, "[люди]", parentsWithNumbers.ToString()); // Замена [люди] на список родителей с номерами

                    // Создание диалогового окна "Сохранить файл"
                    Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                    saveFileDialog.Filter = "Файлы Word (*.docx)|*.docx";
                    saveFileDialog.Title = "Сохранить документ Word как";
                    saveFileDialog.DefaultExt = "docx";

                    // Сохранение документа
                    if (saveFileDialog.ShowDialog() == true)
                    {
                        string saveFilePath = saveFileDialog.FileName;
                        doc.SaveAs(saveFilePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                        MessageBox.Show("Документ успешно сохранен как Word.");
                    }

                    // Закрытие документа и Word приложения
                    doc.Close();
                    wordApp.Quit();
                }
                else
                {
                    // Обработка ошибки открытия соединения
                    MessageBox.Show("Ошибка подключения к базе данных.");
                }
            }
        }


        //печать ведомости успеваемости
        private void vedomostUspev_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем выбранные даты
                DateTime? startDate = datePickerFilterFirstDateVedom.SelectedDate;
                DateTime? endDate = datePickerFilterLastDateVedom.SelectedDate;

                // Проверяем, что даты были выбраны и что начальная дата не больше конечной
                if (startDate.HasValue && endDate.HasValue && startDate <= endDate)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                    // Устанавливаем значение в ячейку A1
                    worksheet.Cells[1, 1] = "Ведомость успеваемости";

                    // Устанавливаем значение в ячейку A1
                    worksheet.Cells[2, 1] = "ФИО учащегося";

                    // Подключаемся к базе данных
                    using (DatabaseConnection dbConnection = new DatabaseConnection())
                    {
                        if (dbConnection.OpenConnection())
                        {
                            // Создаем команду для получения списка учащихся и их средних баллов по предметам в заданном диапазоне дат
                            SqlCommand command = new SqlCommand("SELECT Учащийся.фио, Предметы.название, AVG(ISNULL(Отметки.отметка, 0)) AS средний_балл " +
                                                                "FROM Учащийся " +
                                                                "CROSS JOIN Предметы " +
                                                                "LEFT JOIN Отметки ON Учащийся.id_учащегося = Отметки.id_учащегося AND Предметы.id_предмета = Отметки.id_предмета " +
                                                                "WHERE Отметки.дата BETWEEN @StartDate AND @EndDate " +
                                                                "GROUP BY Учащийся.фио, Предметы.название", dbConnection.GetConnection());
                            command.Parameters.AddWithValue("@StartDate", startDate);
                            command.Parameters.AddWithValue("@EndDate", endDate);
                            SqlDataReader reader = command.ExecuteReader();

                            // Словарь для хранения средних баллов по предметам для каждого учащегося
                            Dictionary<string, Dictionary<string, double>> studentMarks = new Dictionary<string, Dictionary<string, double>>();

                            // Заполняем словарь данными из базы
                            while (reader.Read())
                            {
                                string studentName = reader["фио"].ToString();
                                string subjectName = reader["название"].ToString();
                                double averageMark = Convert.ToDouble(reader["средний_балл"]);

                                if (!studentMarks.ContainsKey(studentName))
                                {
                                    studentMarks.Add(studentName, new Dictionary<string, double>());
                                }

                                studentMarks[studentName].Add(subjectName, averageMark);
                            }

                            reader.Close(); // Закрываем Reader после использования

                            // Настройка заголовков столбцов
                            int column = 2;
                            foreach (var subject in studentMarks.Values.SelectMany(dict => dict.Keys).Distinct())
                            {
                                worksheet.Cells[2, column] = subject;
                                column++;
                            }

                            // Заполнение таблицы данными из словаря
                            int row = 3;
                            foreach (var student in studentMarks)
                            {
                                worksheet.Cells[row, 1] = student.Key; // Имя учащегося в первом столбце
                                column = 2;
                                foreach (var subject in studentMarks.Values.SelectMany(dict => dict.Keys).Distinct())
                                {
                                    double mark = student.Value.ContainsKey(subject) ? student.Value[subject] : 0;
                                    worksheet.Cells[row, column] = mark; // Средний балл по предмету
                                    column++;
                                }
                                row++;
                            }

                            // Применяем автоподбор ширины столбцов
                            Excel.Range usedRange = worksheet.UsedRange;
                            usedRange.Columns.AutoFit();

                            // Применяем обводку к данным
                            ApplyBordersToAllWorksheets(workbook);

                            // Сохраняем книгу Excel
                            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                            saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
                            saveFileDialog.Title = "Сохранить ведомость успеваемости как";

                            if (saveFileDialog.ShowDialog() == true)
                            {
                                string fileName = saveFileDialog.FileName;
                                workbook.SaveAs(fileName);
                                excelApp.Quit();
                                MessageBox.Show("Ведомость успеваемости успешно сохранена.");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Ошибка подключения к базе данных.");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Выберите корректный диапазон дат.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }








        //печать ведомости посещяемости
        private void vedomostPosesh_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Получаем выбранные даты
                DateTime? startDate = datePickerFilterFirstDateVedom.SelectedDate;
                DateTime? endDate = datePickerFilterLastDateVedom.SelectedDate;

                // Проверяем, что даты были выбраны и что дата начала не больше даты окончания
                if (startDate.HasValue && endDate.HasValue && startDate <= endDate)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = (Excel.Worksheet)workbook.ActiveSheet;

                    // Устанавливаем значение в ячейку A1
                    worksheet.Cells[1, 1] = "Ведомость посещаемости";

                    // Устанавливаем заголовок столбца A
                    worksheet.Cells[2, 1] = "ФИО учащегося";

                    // Устанавливаем заголовок столбца B
                    worksheet.Cells[2, 2] = "Сумма часов пропуска";

                    // Форматирование дат
                    string formattedStartDate = startDate.Value.ToString("dd.MM.yyyy");
                    string formattedEndDate = endDate.Value.ToString("dd.MM.yyyy");

                    // Вывод диапазона дат в ячейку B1
                    worksheet.Cells[1, 2] = $"{formattedStartDate} по {formattedEndDate}";


                    // Подключаемся к базе данных
                    using (DatabaseConnection dbConnection = new DatabaseConnection())
                    {
                        if (dbConnection.OpenConnection())
                        {
                            // Создаем команду для получения данных о суммарном количестве часов пропуска для каждого учащегося в заданном диапазоне дат
                            SqlCommand command = new SqlCommand("SELECT Учащийся.фио, ISNULL(SUM(Пропуски.количество_часов), 0) AS Сумма_часов_пропуска " +
                                                                  "FROM Учащийся " +
                                                                  "LEFT JOIN Пропуски ON Учащийся.id_учащегося = Пропуски.id_учащегося " +
                                                                  "WHERE Пропуски.дата BETWEEN @StartDate AND @EndDate " +
                                                                  "GROUP BY Учащийся.фио", dbConnection.GetConnection());
                            command.Parameters.AddWithValue("@StartDate", startDate);
                            command.Parameters.AddWithValue("@EndDate", endDate);
                            SqlDataReader reader = command.ExecuteReader();

                            int row = 3; // начиная с третьей строки

                            // Заполняем таблицу Excel данными из результата запроса
                            while (reader.Read())
                            {
                                string studentName = reader["фио"].ToString();
                                worksheet.Cells[row, 1] = studentName;

                                int hoursMissed = (int)reader["Сумма_часов_пропуска"];
                                worksheet.Cells[row, 2] = hoursMissed;

                                // Переходим к следующей строке
                                row++;
                            }

                            reader.Close();

                            // Применяем автоподбор ширины столбцов
                            Excel.Range usedRange = worksheet.UsedRange;
                            usedRange.Columns.AutoFit();

                            // Применяем обводку к данным
                            ApplyBordersToAllWorksheets(workbook);

                            // Сохраняем книгу Excel
                            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
                            saveFileDialog.Filter = "Файлы Excel (*.xlsx)|*.xlsx";
                            saveFileDialog.Title = "Сохранить ведомость посещаемости как";

                            if (saveFileDialog.ShowDialog() == true)
                            {
                                string fileName = saveFileDialog.FileName;
                                workbook.SaveAs(fileName);
                                excelApp.Quit();
                                MessageBox.Show("Ведомость посещаемости успешно сохранена.");
                            }
                        }
                        else
                        {
                            MessageBox.Show("Ошибка подключения к базе данных.");
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Выберите корректный диапазон дат.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }






        //печать справки
        private void spravka_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridForm.SelectedItem == null)
            {
                MessageBox.Show("Выберите учащегося!");
                return;
            }

            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;
            string fio = selectedRow["ФИО учащегося"].ToString();
            string startDate = selectedRow["Дата начала"].ToString();
            string lastDate = selectedRow["Дата конца"].ToString();
            string vidSpravky = selectedRow["Вид справки"].ToString();
            string todayDate = DateTime.Today.ToString("yyyy-MM-dd");

            // laptop
            // string templateFilePath = @"C:\GitHub\OpenAccess\GGAEK\4course\БД\КурсоваяБД\ИсходныеДокументы\Zayavka.doc";

            // computer
            string templateFilePath = @"G:\Files\ЗаконченыеПроекты\3КУРС\4course\Дипломы\Дубровский\Программа Дубровского\ClassroomTeachersWorkplace\ClassroomTeachersWorkplace\Documents\spravka.docx";

            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Open(templateFilePath);
            doc.Activate();

            // Замена данных в документе
            FindAndReplace(wordApp, "[дата_сегодня]", todayDate);
            FindAndReplace(wordApp, "[дата_сегодня]", todayDate);
            FindAndReplace(wordApp, "[фио]", fio);
            FindAndReplace(wordApp, "[вид]", vidSpravky);
            FindAndReplace(wordApp, "[начальная_дата]", startDate);
            FindAndReplace(wordApp, "[конечная_дата]", lastDate);

            // Создание диалогового окна "Сохранить файл"
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Файлы Word (*.doc)|*.doc";
            saveFileDialog.Title = "Сохранить документ Word как";
            saveFileDialog.DefaultExt = "doc";

            // Сохранение документа
            if (saveFileDialog.ShowDialog() == true)
            {
                string saveFilePath = saveFileDialog.FileName;
                doc.SaveAs(saveFilePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument);
                MessageBox.Show("Документ успешно сохранен как Word.");
            }

            // Закрытие документа и Word приложения
            doc.Close();
            wordApp.Quit();
        }

        //обводка данных
        private void ApplyBordersToAllWorksheets(Microsoft.Office.Interop.Excel.Workbook workbook)
        {
            foreach (Microsoft.Office.Interop.Excel.Worksheet worksheet in workbook.Sheets)
            {
                Microsoft.Office.Interop.Excel.Range usedRange = worksheet.UsedRange;

                foreach (Microsoft.Office.Interop.Excel.Range cell in usedRange)
                {
                    Microsoft.Office.Interop.Excel.Borders borders = cell.Borders;
                    borders.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                    borders.Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
                }
            }
        }

        //метод для замены в WORD
        private void FindAndReplace(Microsoft.Office.Interop.Word.Application wordApp, object findText, string replaceText)
        {
            foreach (Range range in wordApp.ActiveDocument.StoryRanges)
            {
                range.Find.ClearFormatting();

                if (replaceText.Length <= 255)
                {
                    // Если текст для замены короче или равен 255 символам, используем стандартный метод замены
                    range.Find.Execute(FindText: findText, ReplaceWith: replaceText);
                }
                else
                {
                    // Разбиваем текст для замены на части
                    List<string> parts = SplitByLength(replaceText, 255);

                    // Вставляем первую часть и заменяем оставшийся текст
                    if (range.Find.Execute(FindText: findText, ReplaceWith: parts[0]))
                    {
                        Range endRange = range.Duplicate;
                        endRange.SetRange(range.End, range.End);

                        for (int i = 1; i < parts.Count; i++)
                        {
                            endRange.InsertAfter(parts[i]);
                            endRange.SetRange(endRange.End, endRange.End);
                        }
                    }
                }
            }
        }

        private List<string> SplitByLength(string text, int length)
        {
            List<string> parts = new List<string>();

            for (int i = 0; i < text.Length; i += length)
            {
                if (i + length > text.Length)
                {
                    parts.Add(text.Substring(i));
                }
                else
                {
                    parts.Add(text.Substring(i, length));
                }
            }

            return parts;
        }


        private void karta_Click(object sender, RoutedEventArgs e)
        {
            if (dataGridForm.SelectedItem == null)
            {
                MessageBox.Show("Выберите учащегося!");
                return;
            }

            DataRowView selectedRow = (DataRowView)dataGridForm.SelectedItem;
            string id = selectedRow["ID"].ToString();
            string fio = selectedRow["ФИО учащегося"].ToString();
            string phone = selectedRow["Телефон"].ToString();
            string dateOfBirth = selectedRow["Дата рождения"].ToString();
            string socBitovyeUsloviya = selectedRow["Социальные и бытовые условия"].ToString();
            string materialnoeSostoyanie = selectedRow["Материальное состояние"].ToString();
            string trudnostiVUchebe = selectedRow["Трудности в учебе"].ToString();
            string sostoyanieNaUchte = selectedRow["Состояние на учете"].ToString();
            string narushenieObshcheniya = selectedRow["Нарушение общения"].ToString();

            // Загрузка шаблона документа Word
            string templateFilePath = @"G:\Files\ЗаконченыеПроекты\3КУРС\4course\Дипломы\Дубровский\Программа Дубровского\ClassroomTeachersWorkplace\ClassroomTeachersWorkplace\Documents\lichCartUchash.docx";
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Document doc = wordApp.Documents.Open(templateFilePath);
            doc.Activate();

            // Замена данных в документе
            FindAndReplace(wordApp, "[фио]", fio);
            FindAndReplace(wordApp, "[телефон]", phone);
            FindAndReplace(wordApp, "[дата_рождения]", dateOfBirth);
            FindAndReplace(wordApp, "[соц_быт]", socBitovyeUsloviya);
            FindAndReplace(wordApp, "[мат_сост]", materialnoeSostoyanie);
            FindAndReplace(wordApp, "[труд_учеб]", trudnostiVUchebe);
            FindAndReplace(wordApp, "[сост_учет]", sostoyanieNaUchte);
            FindAndReplace(wordApp, "[общение]", narushenieObshcheniya);

            // Найти родителей учащегося по id
            List<string> parents = GetParents(id);

            // Сформировать строку с ФИО родителей
            string parentsStr = string.Join(", ", parents);

            // Заменить [родители] на список родителей
            FindAndReplace(wordApp, "[родители]", parentsStr);

            // Создание диалогового окна "Сохранить файл"
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Файлы Word (*.docx)|*.docx";
            saveFileDialog.Title = "Сохранить личную карточку учащегося как";
            saveFileDialog.DefaultExt = "docx";

            // Сохранение документа
            if (saveFileDialog.ShowDialog() == true)
            {
                string saveFilePath = saveFileDialog.FileName;
                doc.SaveAs(saveFilePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                MessageBox.Show("Личная карточка учащегося успешно сохранена.");
            }

            // Закрытие документа и Word приложения
            doc.Close();
            wordApp.Quit();
        }

        private List<string> GetParents(string studentId)
        {
            List<string> parentsInfo = new List<string>();

            // Создаем объект подключения к базе данных
            using (DatabaseConnection dbConnection = new DatabaseConnection())
            {
                // Открываем соединение
                if (dbConnection.OpenConnection())
                {
                    // Создаем SQL-запрос для получения данных о родителях учащегося по его id
                    string query = @"SELECT Родители.фио, Родители.пол, CONVERT(NVARCHAR, Родители.дата_рождения, 104) AS дата_рождения
                             FROM Родители 
                             WHERE Родители.id_учащегося = @StudentId";

                    // Создаем объект команды SQL
                    using (SqlCommand cmd = new SqlCommand(query, dbConnection.GetConnection()))
                    {
                        // Добавляем параметр для id учащегося
                        cmd.Parameters.AddWithValue("@StudentId", studentId);

                        // Выполняем запрос и получаем результирующий набор данных
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            // Читаем каждую запись результата запроса
                            while (reader.Read())
                            {
                                // Получаем данные о родителе из текущей записи
                                string parentName = reader["фио"].ToString();
                                string gender = reader["пол"].ToString();
                                string birthDate = reader["дата_рождения"].ToString();

                                // Создаем строку с информацией о родителе
                                string parentInfo = $"{parentName}, семейное положение {gender}, дата рождения {birthDate}";

                                // Добавляем информацию о родителе в список
                                parentsInfo.Add(parentInfo);
                            }
                        }
                    }
                }
                else
                {
                    // Обработка ошибки открытия соединения
                    MessageBox.Show("Ошибка подключения к базе данных.");
                }
            }

            return parentsInfo;
        }

        // Печать социального паспорта
        private void pasport_Click(object sender, RoutedEventArgs e)
        {
            // Переменные для хранения данных из базы данных
            int количествоУчащихся = 0;
            string наименованиеКласса = string.Empty;
            int наУчете = 0;
            int асоциальные = 0;
            int трудностиВУчёбе = 0;
            int чаэс = 0;
            int сироты = 0;

            // Социально-бытовые условия
            int хорошие = 0;
            int средние = 0;
            int плохие = 0;
            int критические = 0;

            // Материальное состояние
            int вышеСреднего = 0;
            int среднеМат = 0;
            int нижеСреднего = 0;
            int нижеПрожиточного = 0;

            // Группы инвалидности
            int перваяГруппа = 0;
            int втораяГруппа = 0;
            int третьяГруппа = 0;

            // Группы здоровья
            int перваяГруппаЗдоровья = 0;
            int втораяГруппаЗдоровья = 0;
            int третьяГруппаЗдоровья = 0;
            int четвертаяГруппаЗдоровья = 0;
            int пятаяГруппаЗдоровья = 0;

            // Получение данных из базы данных
            using (DatabaseConnection dbConnection = new DatabaseConnection())
            {
                if (dbConnection.OpenConnection())
                {
                    SqlConnection connection = dbConnection.GetConnection();

                    // Выполнение SQL запросов для получения данных
                    количествоУчащихся = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся");
                    наименованиеКласса = GetSingleValueFromDatabase(connection, "SELECT наименование FROM Класс");
                    наУчете = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE состояние_на_учёте = N'Да'");
                    асоциальные = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE нарушение_общения = N'Асоциальный'");
                    трудностиВУчёбе = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE трудности_в_учёбе = N'Да'");
                    чаэс = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE чаэс = N'Да'");
                    сироты = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE сирота = N'Да'");
                    хорошие = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE соц_бытовые_условия = N'Хорошие'");
                    средние = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE соц_бытовые_условия = N'Средние'");
                    плохие = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE соц_бытовые_условия = N'Плохие'");
                    критические = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE соц_бытовые_условия = N'Критические'");
                    вышеСреднего = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE материальное_состояние = N'Выше среднего'");
                    среднеМат = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE материальное_состояние = N'Средние'");
                    нижеСреднего = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE материальное_состояние = N'Ниже среднего'");
                    нижеПрожиточного = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE материальное_состояние = N'Ниже прожиточного минимума'");
                    перваяГруппа = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE инвалидность = N'1 группа'");
                    втораяГруппа = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE инвалидность = N'2 группа'");
                    третьяГруппа = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE инвалидность = N'3 группа'");
                    перваяГруппаЗдоровья = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE группа_здоровья = N'1 группа'");
                    втораяГруппаЗдоровья = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE группа_здоровья = N'2 группа'");
                    третьяГруппаЗдоровья = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE группа_здоровья = N'3 группа'");
                    четвертаяГруппаЗдоровья = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE группа_здоровья = N'4 группа'");
                    пятаяГруппаЗдоровья = GetCountFromDatabase(connection, "SELECT COUNT(*) FROM Учащийся WHERE группа_здоровья = N'5 группа'");
                }
            }

            // Загрузка шаблона документа Word
            string templateFilePath = @"G:\Files\ЗаконченыеПроекты\3КУРС\4course\Дипломы\Дубровский\Программа Дубровского\ClassroomTeachersWorkplace\ClassroomTeachersWorkplace\Documents\SocPasport.docx";
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = wordApp.Documents.Open(templateFilePath);
            doc.Activate();

            // Замена данных в документе
            FindAndReplace(wordApp, "[колво]", количествоУчащихся.ToString());
            FindAndReplace(wordApp, "[класс]", наименованиеКласса);
            FindAndReplace(wordApp, "[учёт]", наУчете.ToString());
            FindAndReplace(wordApp, "[общение]", асоциальные.ToString());
            FindAndReplace(wordApp, "[трудности]", трудностиВУчёбе.ToString());
            FindAndReplace(wordApp, "[чаэс]", чаэс.ToString());
            FindAndReplace(wordApp, "[сироты]", сироты.ToString());
            FindAndReplace(wordApp, "[хорошие]", хорошие.ToString());
            FindAndReplace(wordApp, "[средние]", средние.ToString());
            FindAndReplace(wordApp, "[плохие]", плохие.ToString());
            FindAndReplace(wordApp, "[критические]", критические.ToString());
            FindAndReplace(wordApp, "[выше_среднего]", вышеСреднего.ToString());
            FindAndReplace(wordApp, "[средние_мат]", среднеМат.ToString());
            FindAndReplace(wordApp, "[ниже_среднего]", нижеСреднего.ToString());
            FindAndReplace(wordApp, "[ниже_прожиточного]", нижеПрожиточного.ToString());
            FindAndReplace(wordApp, "[1групп]", перваяГруппа.ToString());
            FindAndReplace(wordApp, "[2групп]", втораяГруппа.ToString());
            FindAndReplace(wordApp, "[3групп]", третьяГруппа.ToString());
            FindAndReplace(wordApp, "[1_групп]", перваяГруппаЗдоровья.ToString());
            FindAndReplace(wordApp, "[2_групп]", втораяГруппаЗдоровья.ToString());
            FindAndReplace(wordApp, "[3_групп]", третьяГруппаЗдоровья.ToString());
            FindAndReplace(wordApp, "[4_групп]", четвертаяГруппаЗдоровья.ToString());
            FindAndReplace(wordApp, "[5_групп]", пятаяГруппаЗдоровья.ToString());

            // Создание диалогового окна "Сохранить файл"
            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Файлы Word (*.docx)|*.docx",
                Title = "Сохранить социальный паспорт как",
                DefaultExt = "docx"
            };

            // Сохранение документа
            if (saveFileDialog.ShowDialog() == true)
            {
                string saveFilePath = saveFileDialog.FileName;
                doc.SaveAs(saveFilePath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocumentDefault);
                MessageBox.Show("Социальный паспорт успешно сохранен.");
            }

            // Закрытие документа и Word приложения
            doc.Close();
            wordApp.Quit();
        }

        // Метод для получения количества записей из базы данных
        private int GetCountFromDatabase(SqlConnection connection, string query)
        {
            using (SqlCommand cmd = new SqlCommand(query, connection))
            {
                return (int)cmd.ExecuteScalar();
            }
        }

        // Метод для получения одного значения из базы данных
        private string GetSingleValueFromDatabase(SqlConnection connection, string query)
        {
            using (SqlCommand cmd = new SqlCommand(query, connection))
            {
                return (string)cmd.ExecuteScalar();
            }
        }
    }
}

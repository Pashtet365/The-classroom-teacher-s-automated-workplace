using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml.Linq;
using static ClassroomTeachersWorkplace.Отметки;
using Window = System.Windows.Window;

namespace ClassroomTeachersWorkplace
{
    /// <summary>
    /// Логика взаимодействия для Отметки.xaml
    /// </summary>
    public partial class Отметки : Window
    {
        private MainWindow _main;

        public Отметки()
        {
            InitializeComponent();
            LoadStudents();
            LoadItems();
            LoadMarks();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных оценки";
                comboStudent.Text = MainWindow.element2.ToString();
                comboItem.Text = MainWindow.element3.ToString();
                comboMarks.Text = MainWindow.element4.ToString();
                pickerStartDate.Text = MainWindow.element5.ToString();
            }
        }

        public Отметки(MainWindow main) : this()
        {
            _main = main;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Получаем выбранные значения из ComboBox и DatePicker
            Student selectedStudent = (Student)comboStudent.SelectedItem;
            Items selectedItem = (Items)comboItem.SelectedItem;
            string selectedMark = comboMarks.SelectedItem?.ToString();
            DateTime? selectedDate = pickerStartDate.SelectedDate;

            // Проверяем, что все значения были выбраны
            if (string.IsNullOrEmpty(selectedStudent.ToString()) || string.IsNullOrEmpty(selectedItem.ToString()) || string.IsNullOrEmpty(selectedMark) || selectedDate == null)
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            // Проверяем, что выбранная дата не в будущем
            if (selectedDate > DateTime.Today)
            {
                MessageBox.Show("Дата не может быть в будущем.");
                return;
            }

            // Получение id организации из выбранного объекта Organization
            int idItem = selectedItem.Id;
            int idStep = selectedStudent.Id;

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            try
                            {

                                    if (dbConnection.OpenConnection())
                                    {
                                        // Проверяем, удалось ли получить ID предмета и учащегося
                                        if (idItem == -1 || idStep == -1)
                                        {
                                            MessageBox.Show("Не удалось получить ID предмета или учащегося.");
                                            return;
                                        }

                                        // Получаем отметку и преобразуем ее в int
                                        int mark;
                                        if (!int.TryParse(selectedMark, out mark))
                                        {
                                            MessageBox.Show("Неверный формат отметки.");
                                            return;
                                        }

                                        // Выполнение запроса на добавление новой записи
                                        string insertQuery = "INSERT INTO Отметки (id_учащегося, id_предмета, отметка, дата) VALUES (@idStep, @idItem, @mark, @selectedDate)";
                                        SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                                        insertCommand.Parameters.AddWithValue("@idStep", idStep);
                                        insertCommand.Parameters.AddWithValue("@idItem", idItem);
                                        insertCommand.Parameters.AddWithValue("@mark", mark);
                                        insertCommand.Parameters.AddWithValue("@selectedDate", selectedDate);
                                        insertCommand.ExecuteNonQuery();

                                        MessageBox.Show("Новая отметка успешно добавлена.");
                                        comboStudent.Text = "";
                                        comboItem.Text = "";
                                        comboMarks.Text = "";
                                        pickerStartDate.Text = "";
                                }
                                    else
                                    {
                                        MessageBox.Show("Ошибка подключения к базе данных.");
                                    }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Произошла ошибка при добавлении новой отметки: " + ex.Message);
                            }
                        }
                        else // Изменение записи
                        {
                            try
                            {
                                if (dbConnection.OpenConnection())
                                {
                                    // Получаем отметку и преобразуем ее в int
                                    int mark;
                                    if (!int.TryParse(selectedMark, out mark))
                                    {
                                        MessageBox.Show("Неверный формат отметки.");
                                        return;
                                    }

                                    // Выполнение запроса на изменение записи
                                    string updateQuery = "UPDATE Отметки SET id_учащегося = @idStep, id_предмета = @idItem, отметка = @mark, дата = @selectedDate WHERE id_отметки = @markId";
                                    SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                                    updateCommand.Parameters.AddWithValue("@idStep", idStep);
                                    updateCommand.Parameters.AddWithValue("@idItem", idItem);
                                    updateCommand.Parameters.AddWithValue("@mark", mark);
                                    updateCommand.Parameters.AddWithValue("@selectedDate", selectedDate);
                                    updateCommand.Parameters.AddWithValue("@markId", MainWindow.element1);
                                    updateCommand.ExecuteNonQuery();

                                    MessageBox.Show("Запись успешно изменена.");
                                }
                                else
                                {
                                    MessageBox.Show("Ошибка подключения к базе данных.");
                                }
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Произошла ошибка при изменении записи: " + ex.Message);
                            }
                        }


                        // Обновление данных в таблице (если необходимо)
                        _main.Refresh(sender, e);
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Произошла ошибка при добавлении данных: " + ex.Message);
            }
        }


        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        public class Student
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private void LoadStudents()
        {
            try
            {
                List<Student> Students = new List<Student>();
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        string selectQuery = "SELECT id_учащегося, фио FROM Учащийся";
                        SqlCommand selectCmd = new SqlCommand(selectQuery, dbConnection.GetConnection());
                        SqlDataReader reader = selectCmd.ExecuteReader();

                        while (reader.Read())
                        {
                            int id = reader.GetInt32(0);
                            string name = reader.GetString(1);
                            Students.Add(new Student { Id = id, Name = name });
                        }

                        reader.Close();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }

                // Привязать список предметов к ComboBox
                comboStudent.ItemsSource = Students;
                comboStudent.DisplayMemberPath = "Name"; // Указать, какое свойство использовать для отображения в ComboBox
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }

        private void LoadMarks()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboMarks != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboMarks.Items.Clear();

                // Добавляем данные в комбобокс
                comboMarks.Items.Add("1");
                comboMarks.Items.Add("2");
                comboMarks.Items.Add("3");
                comboMarks.Items.Add("4");
                comboMarks.Items.Add("5");
                comboMarks.Items.Add("6");
                comboMarks.Items.Add("7");
                comboMarks.Items.Add("8");
                comboMarks.Items.Add("9");
                comboMarks.Items.Add("10");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        public class Items
        {
            public int Id { get; set; }
            public string Name { get; set; }
        }

        private void LoadItems()
        {
            try
            {
                List<Items> Students = new List<Items>();
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        string selectQuery = "SELECT id_предмета, название FROM Предметы";
                        SqlCommand selectCmd = new SqlCommand(selectQuery, dbConnection.GetConnection());
                        SqlDataReader reader = selectCmd.ExecuteReader();

                        while (reader.Read())
                        {
                            int id = reader.GetInt32(0);
                            string name = reader.GetString(1);
                            Students.Add(new Items { Id = id, Name = name });
                        }

                        reader.Close();
                    }
                    else
                    {
                        MessageBox.Show("Ошибка подключения к базе данных.");
                    }
                }

                // Привязать список предметов к ComboBox
                comboItem.ItemsSource = Students;
                comboItem.DisplayMemberPath = "Name"; // Указать, какое свойство использовать для отображения в ComboBox
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: " + ex.Message);
            }
        }
    }
}

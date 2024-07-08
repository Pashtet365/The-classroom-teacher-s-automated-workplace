using Microsoft.Office.Interop.Excel;
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
    /// Логика взаимодействия для Справки.xaml
    /// </summary>
    public partial class Справки : Window
    {
        private MainWindow _main;

        public Справки()
        {
            InitializeComponent();
            LoadStudents();
            LoadVidSpravky();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных справки";
                comboStudent.Text = MainWindow.element2.ToString();
                pickerStartDate.Text = MainWindow.element3.ToString();
                pickerLastDate.Text = MainWindow.element4.ToString();
                comboVidSpravky.Text = MainWindow.element5.ToString();
            }
        }

        public Справки(MainWindow main) : this()
        {
            _main = main;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что все поля заполнены
            if (comboStudent.SelectedItem == null ||
                pickerStartDate.SelectedDate == null ||
                pickerLastDate.SelectedDate == null ||
                comboVidSpravky.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }
            // Проверка, что дата начала не больше даты окончания
            if (pickerStartDate.SelectedDate > pickerLastDate.SelectedDate)
            {
                MessageBox.Show("Дата начала не может быть позже даты окончания.");
                return;
            }

            // Получаем выбранные значения из ComboBox и DatePicker
            Student selectedStudent = (Student)comboStudent.SelectedItem;
            string selectedVidSpravky = comboVidSpravky.SelectedItem?.ToString();
            DateTime? selectedStartDate = pickerStartDate.SelectedDate;
            DateTime? selectedLastDate = pickerLastDate.SelectedDate;

            // Получение id организации из выбранного объекта Organization
            int idStep = selectedStudent.Id;

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            string insertQuery = @"
                    INSERT INTO Справка (id_учащегося, дата_начала, дата_конца, вид_справки)
                    VALUES (@StudentId, @StartDate, @LastDate, @VidSpravky)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                            insertCommand.Parameters.AddWithValue("@StudentId", idStep);
                            insertCommand.Parameters.AddWithValue("@StartDate", selectedStartDate);
                            insertCommand.Parameters.AddWithValue("@LastDate", selectedLastDate);
                            insertCommand.Parameters.AddWithValue("@VidSpravky", selectedVidSpravky);

                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                // Очистка полей
                                comboStudent.SelectedItem = null;
                                pickerStartDate.SelectedDate = null;
                                pickerLastDate.SelectedDate = null;
                                comboVidSpravky.SelectedItem = null;
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить запись.");
                            }
                        }
                        else // Изменение записи
                        {
                            string updateQuery = @"
                    UPDATE Справка
                    SET id_учащегося = @StudentId, дата_начала = @StartDate, дата_конца = @LastDate, вид_справки = @VidSpravky
                    WHERE id_справки = @SpravkaId";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                            updateCommand.Parameters.AddWithValue("@SpravkaId", MainWindow.element1); // ID записи
                            updateCommand.Parameters.AddWithValue("@StudentId", idStep);
                            updateCommand.Parameters.AddWithValue("@StartDate", selectedStartDate);
                            updateCommand.Parameters.AddWithValue("@LastDate", selectedLastDate);
                            updateCommand.Parameters.AddWithValue("@VidSpravky", selectedVidSpravky);

                            int rowsAffected = updateCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно изменена.");
                            }
                            else
                            {
                                MessageBox.Show("Не удалось изменить запись.");
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
                MessageBox.Show("Произошла ошибка при выполнении запроса: " + ex.Message);
            }


        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void LoadVidSpravky()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboVidSpravky != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboVidSpravky.Items.Clear();

                // Добавляем данные в комбобокс
                comboVidSpravky.Items.Add("I Группа");
                comboVidSpravky.Items.Add("II группа");
                comboVidSpravky.Items.Add("III группа");
                comboVidSpravky.Items.Add("IV группа ");
                comboVidSpravky.Items.Add("V группа");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
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
    }
}

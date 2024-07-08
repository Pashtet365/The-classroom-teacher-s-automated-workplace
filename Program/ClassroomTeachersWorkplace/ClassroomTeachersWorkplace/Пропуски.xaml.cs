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
using Window = System.Windows.Window;

namespace ClassroomTeachersWorkplace
{
    /// <summary>
    /// Логика взаимодействия для Пропуски.xaml
    /// </summary>
    public partial class Пропуски : Window
    {
        private MainWindow _main;

        public Пропуски()
        {
            InitializeComponent();
            LoadStudents();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных пропуска";
                comboStudent.Text = MainWindow.element2.ToString();
                txtTime.Text = MainWindow.element3.ToString();
                pickerDate.Text = MainWindow.element4.ToString();
            }
        }

        public Пропуски(MainWindow main) : this()
        {
            _main = main;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что все поля заполнены
            if (comboStudent.SelectedItem == null ||
                string.IsNullOrEmpty(txtTime.Text) ||
                pickerDate.SelectedDate == null)
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            // Получение выбранной даты
            DateTime selectedDate = pickerDate.SelectedDate.Value;

            // Проверка, что выбранная дата не в будущем
            if (selectedDate > DateTime.Today)
            {
                MessageBox.Show("Дата не может быть в будущем.");
                return;
            }

            
            // Получение ID выбранного учащегося
            int studentId = ((Student)comboStudent.SelectedItem).Id;
            // Получение количества часов
            if (!int.TryParse(txtTime.Text, out int hours))
            {
                MessageBox.Show("Неверный формат количества часов.");
                return;
            }

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            string insertQuery = @"
                    INSERT INTO Пропуски (id_учащегося, количество_часов, дата)
                    VALUES (@StudentId, @Hours, @Date)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                            insertCommand.Parameters.AddWithValue("@StudentId", studentId);
                            insertCommand.Parameters.AddWithValue("@Hours", hours);
                            insertCommand.Parameters.AddWithValue("@Date", selectedDate);

                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                // Очистка полей
                                comboStudent.SelectedItem = null;
                                txtTime.Text = "";
                                pickerDate.SelectedDate = null;
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить запись.");
                            }
                        }
                        else // Изменение записи
                        {
                            string updateQuery = @"
                    UPDATE Пропуски
                    SET id_учащегося = @StudentId, количество_часов = @Hours, дата = @Date
                    WHERE id_пропуска = @AbsenceId";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                            updateCommand.Parameters.AddWithValue("@AbsenceId", MainWindow.element1); // ID записи
                            updateCommand.Parameters.AddWithValue("@StudentId", studentId);
                            updateCommand.Parameters.AddWithValue("@Hours", hours);
                            updateCommand.Parameters.AddWithValue("@Date", selectedDate);

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


        private void txtTime_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            foreach (char c in e.Text)
            {
                // Проверяем, является ли символ цифрой
                if (!char.IsDigit(c))
                {
                    // Если символ не является цифрой, отменяем ввод
                    e.Handled = true;
                    break;
                }
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
    }
}

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
using static ClassroomTeachersWorkplace.Отметки;
using Window = System.Windows.Window;

namespace ClassroomTeachersWorkplace
{
    /// <summary>
    /// Логика взаимодействия для Собрание.xaml
    /// </summary>
    public partial class Собрание : Window
    {
        private MainWindow _main;

        public Собрание()
        {
            InitializeComponent();
            LoadStudents();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных собрания";
                comboStudent.Text = MainWindow.element2.ToString();
                pickerStartDate.Text = MainWindow.element3.ToString();
                txtName.Text = MainWindow.element4.ToString();
            }
        }

        public Собрание(MainWindow main) : this()
        {
            _main = main;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что все поля заполнены
            if (comboStudent.SelectedItem == null ||
                string.IsNullOrEmpty(txtName.Text) ||
                pickerStartDate.SelectedDate == null)
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            // Проверка, что дата собрания не в прошлом времени
            DateTime selectedDate = pickerStartDate.SelectedDate.Value;
            if (selectedDate < DateTime.Today)
            {
                MessageBox.Show("Дата собрания не может быть в прошлом времени.");
                return;
            }

            // Получаем выбранные значения из ComboBox и DatePicker
            Student selectedStudent = (Student)comboStudent.SelectedItem;
            string themeEvent = txtName.Text;
            DateTime? selectedDateEvent = pickerStartDate.SelectedDate;

            // Получение id родителя
            int idParent = selectedStudent.Id;

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            string insertQuery = @"
                        INSERT INTO Собрание (id_родителя, дата, тема)
                        VALUES (@ParentId, @EventDate, @Theme)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                            insertCommand.Parameters.AddWithValue("@ParentId", idParent);
                            insertCommand.Parameters.AddWithValue("@EventDate", selectedDateEvent);
                            insertCommand.Parameters.AddWithValue("@Theme", themeEvent);

                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                // Очистка полей
                                comboStudent.SelectedItem = null;
                                txtName.Text = "";
                                pickerStartDate.SelectedDate = null;
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить запись.");
                            }
                        }
                        else // Изменение записи
                        {
                            

                            string updateQuery = @"
        UPDATE Собрание
        SET id_родителя = @ParentId, дата = @EventDate, тема = @Theme
        WHERE id_собрания = @EventId";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                            updateCommand.Parameters.AddWithValue("@ParentId", idParent);
                            updateCommand.Parameters.AddWithValue("@EventDate", selectedDateEvent);
                            updateCommand.Parameters.AddWithValue("@Theme", themeEvent);
                            updateCommand.Parameters.AddWithValue("@EventId", MainWindow.element1);

                            int rowsAffected = updateCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно изменена.");
                                // Очистка полей
                                comboStudent.SelectedItem = null;
                                txtName.Text = "";
                                pickerStartDate.SelectedDate = null;
                                // Сброс режима изменения
                                MainWindow.changing = 0;
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
                        string selectQuery = "SELECT id_родителя, фио FROM Родители";
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

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
using static ClassroomTeachersWorkplace.Отметки;
using Window = System.Windows.Window;

namespace ClassroomTeachersWorkplace
{
    /// <summary>
    /// Логика взаимодействия для Родители.xaml
    /// </summary>
    public partial class Родители : Window
    {
        private MainWindow _main;

        public Родители()
        {
            InitializeComponent();
            LoadStudents();
            LoadGender();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных  Родителя";
                comboStudent.Text = MainWindow.element2.ToString();
                txtName.Text = MainWindow.element3.ToString();
                comboGender.Text = MainWindow.element4.ToString();
                pickerStartDate.Text = MainWindow.element5.ToString();
                txtWorck.Text = MainWindow.element6.ToString();
                txtPost.Text = MainWindow.element7.ToString();
                txtAdress.Text = MainWindow.element8.ToString();
            }
        }

        public Родители(MainWindow main) : this()
        {
            _main = main;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что все поля заполнены
            if (comboStudent.SelectedItem == null ||
                string.IsNullOrEmpty(txtName.Text) ||
                comboGender.SelectedItem == null ||
                pickerStartDate.SelectedDate == null ||
                string.IsNullOrEmpty(txtWorck.Text) ||
                string.IsNullOrEmpty(txtPost.Text) ||
                string.IsNullOrEmpty(txtAdress.Text))
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            // Проверка, что дата рождения не больше 150 лет
            DateTime selectedDate = pickerStartDate.SelectedDate.Value;
            DateTime minDate = DateTime.Today.AddYears(-150);
            if (selectedDate > DateTime.Today || selectedDate < minDate)
            {
                MessageBox.Show("Дата рождения должна быть в пределах последних 150 лет.");
                return;
            }

            // Проверка возраста родителя (не менее 18 лет)
            DateTime today = DateTime.Today;
            DateTime eighteenYearsAgo = today.AddYears(-18);
            if (selectedDate > eighteenYearsAgo)
            {
                MessageBox.Show("Родитель должен быть старше 18 лет.");
                return;
            }

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0)
                        {
                            // Реализация добавления записи
                            string insertQuery = @"
                        INSERT INTO Родители (id_учащегося, фио, пол, дата_рождения, место_работы, должность, адресс)
                        VALUES (@StudentId, @ParentName, @Gender, @DateOfBirth, @WorkPlace, @Position, @Address)";

                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                            insertCommand.Parameters.AddWithValue("@StudentId", ((Student)comboStudent.SelectedItem).Id);
                            insertCommand.Parameters.AddWithValue("@ParentName", txtName.Text);
                            insertCommand.Parameters.AddWithValue("@Gender", comboGender.SelectedItem.ToString());
                            insertCommand.Parameters.AddWithValue("@DateOfBirth", selectedDate);
                            insertCommand.Parameters.AddWithValue("@WorkPlace", txtWorck.Text);
                            insertCommand.Parameters.AddWithValue("@Position", txtPost.Text);
                            insertCommand.Parameters.AddWithValue("@Address", txtAdress.Text);

                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись успешно добавлена.");
                                // Очистка полей
                                comboStudent.SelectedItem = null;
                                txtName.Text = "";
                                comboGender.SelectedItem = null;
                                pickerStartDate.SelectedDate = null;
                                txtWorck.Text = "";
                                txtPost.Text = "";
                                txtAdress.Text = "";
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить запись.");
                            }
                        }
                        else
                        {
                            // Реализация изменения записи
                            string updateQuery = @"
                        UPDATE Родители
                        SET id_учащегося = @StudentId, фио = @ParentName, пол = @Gender, дата_рождения = @DateOfBirth,
                            место_работы = @WorkPlace, должность = @Position, адресс = @Address
                        WHERE id_родителя = @ParentId";

                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                            updateCommand.Parameters.AddWithValue("@ParentId", MainWindow.element1); // ID записи
                            updateCommand.Parameters.AddWithValue("@StudentId", ((Student)comboStudent.SelectedItem).Id);
                            updateCommand.Parameters.AddWithValue("@ParentName", txtName.Text);
                            updateCommand.Parameters.AddWithValue("@Gender", comboGender.SelectedItem.ToString());
                            updateCommand.Parameters.AddWithValue("@DateOfBirth", selectedDate);
                            updateCommand.Parameters.AddWithValue("@WorkPlace", txtWorck.Text);
                            updateCommand.Parameters.AddWithValue("@Position", txtPost.Text);
                            updateCommand.Parameters.AddWithValue("@Address", txtAdress.Text);

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

        private void LoadGender()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboGender != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboGender.Items.Clear();

                // Добавляем данные в комбобокс
                comboGender.Items.Add("Мама");
                comboGender.Items.Add("Папа");
                comboGender.Items.Add("Опекун");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        private void txtName_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Проверка, что введенный символ является буквой или символом пробела
            if (!char.IsLetter(e.Text, 0) && !char.IsWhiteSpace(e.Text, 0))
            {
                e.Handled = true; // Отмена ввода символа, если он не является буквой или пробелом
            }
        }
    }
}

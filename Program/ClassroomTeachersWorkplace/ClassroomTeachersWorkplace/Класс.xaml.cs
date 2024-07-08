using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ClassroomTeachersWorkplace
{
    /// <summary>
    /// Логика взаимодействия для Класс.xaml
    /// </summary>
    public partial class Класс : Window
    {
        private MainWindow _main;

        public Класс()
        {
            InitializeComponent();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных класса";
                txtName.Text = MainWindow.element2.ToString();
            }
        }

        public Класс(MainWindow main) : this()
        {
            _main = main;
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            MainWindow.changing = 0;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            string name = txtName.Text.Trim();

            if (string.IsNullOrEmpty(name))
            {
                MessageBox.Show("Введите класс!");
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
                            // Проверяем, есть ли уже записи в таблице Класс
                            string checkQuery = "SELECT COUNT(*) FROM Класс";
                            SqlCommand checkCommand = new SqlCommand(checkQuery, dbConnection.GetConnection());
                            int count = (int)checkCommand.ExecuteScalar();

                            if (count > 0)
                            {
                                // Если записи существуют, выводим сообщение и отменяем добавление
                                MessageBox.Show("Класс уже существует. Нельзя добавить новый класс.");
                            }
                            else
                            {
                                // Выполнение запроса на добавление новой записи
                                string insertQuery = "INSERT INTO Класс (наименование) VALUES (@name)";
                                SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());
                                insertCommand.Parameters.AddWithValue("@name", name);
                                insertCommand.ExecuteNonQuery();

                                MessageBox.Show("Новый класс успешно добавлен.");
                            }
                        }
                        else // Изменение записи
                        {
                            // Выполнение запроса на изменение записи
                            string updateQuery = "UPDATE Класс SET наименование = @name WHERE id_класса = @subjectId";
                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());
                            updateCommand.Parameters.AddWithValue("@name", name);
                            updateCommand.Parameters.AddWithValue("@subjectId", MainWindow.element1);
                            updateCommand.ExecuteNonQuery();

                            MessageBox.Show("Название класса успешно изменено.");
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

        private void txtName_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            Regex regex = new Regex("^[0-9]{1,2}[А-Я]{0,1}$");
            if (!regex.IsMatch((sender as TextBox).Text + e.Text))
                e.Handled = true;
        }

        private void txtName_TextChanged(object sender, TextChangedEventArgs e)
        {
            Regex regex = new Regex("^[0-9]{1,2}[А-Я]{0,1}$");
            if (!regex.IsMatch((sender as TextBox).Text))
            {
                // Если текст не соответствует формату, очистите TextBox
                (sender as TextBox).Text = "";
            }
        }
    }
}

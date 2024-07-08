using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using Window = System.Windows.Window;

namespace ClassroomTeachersWorkplace
{
    /// <summary>
    /// Логика взаимодействия для Учащийся.xaml
    /// </summary>
    public partial class Учащийся : Window
    {
        private MainWindow _main;

        public Учащийся()
        {
            InitializeComponent();
            LoadcomboSocBit();
            LoadcomboMatSot();
            LoadcomboCommunication();
            LoadcomboTrudStady();
            LoadcomboUchet();
            LoadcomboInvalid();
            LoadcomboGrupZdorov();
            LoadcomboChernobol();
            LoadcomboSirota();
            if (MainWindow.changing == 1)
            {
                mainLabel.Content = "Изменение";
                mainButton.Content = "Сохранить";
                this.Title = "Изменение данных учащегося";
                txtName.Text = MainWindow.element2.ToString();
                txtPhone.Text = MainWindow.element3.ToString();
                pickerDateOfBirtzd.Text = MainWindow.element4.ToString();
                comboSocBit.Text = MainWindow.element5.ToString();
                comboMatSot.Text = MainWindow.element6.ToString();
                comboTrudStady.Text = MainWindow.element7.ToString();
                comboUchet.Text = MainWindow.element8.ToString();
                comboCommunication.Text = MainWindow.element9.ToString();
                comboInvalid.Text = MainWindow.element10.ToString();
                comboGrupZdorov.Text = MainWindow.element11.ToString();
                comboChernobol.Text = MainWindow.element12.ToString();
                txtHobbi.Text = MainWindow.element13.ToString();
                comboSirota.Text = MainWindow.element14.ToString();
            }
        }

        public Учащийся(MainWindow main) : this()
        {
            _main = main;
        }

        private void mainButton_Click(object sender, RoutedEventArgs e)
        {
            // Проверка, что все поля заполнены
            if (string.IsNullOrEmpty(txtName.Text) ||
                string.IsNullOrEmpty(txtPhone.Text) ||
                comboSocBit.SelectedItem == null ||
                comboMatSot.SelectedItem == null ||
                comboTrudStady.SelectedItem == null ||
                comboUchet.SelectedItem == null ||
                comboCommunication.SelectedItem == null ||
                pickerDateOfBirtzd.SelectedDate == null ||
                comboInvalid.SelectedItem == null ||
                comboGrupZdorov.SelectedItem == null ||
                comboChernobol.SelectedItem == null ||
                string.IsNullOrEmpty(txtHobbi.Text) ||
                comboSirota.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, заполните все поля.");
                return;
            }

            // Проверка правильности ввода номера телефона
            if (txtPhone.Text.Length < 13)
            {
                MessageBox.Show("Введите правильно номер!");
                return;
            }

            // Получение даты рождения из DateTimePicker
            DateTime? birthDate = pickerDateOfBirtzd.SelectedDate;

            // Проверка, что дата рождения выбрана
            if (birthDate == null)
            {
                MessageBox.Show("Пожалуйста, выберите дату рождения.");
                return;
            }

            // Вычисление возраста на основе выбранной даты рождения
            int age = DateTime.Today.Year - birthDate.Value.Year;

            // Проверка, что возраст больше или равен 6 годам
            if (age < 6)
            {
                MessageBox.Show("Возраст учащегося не может быть меньше 6 лет.");
                return;
            }

            // Получение данных из полей
            string name = txtName.Text;
            string phone = txtPhone.Text;
            string socBit = (string)comboSocBit.SelectedItem;
            string matSot = (string)comboMatSot.SelectedItem;
            string trudStady = (string)comboTrudStady.SelectedItem;
            string uchet = (string)comboUchet.SelectedItem;
            string communication = (string)comboCommunication.SelectedItem;
            string invalid = (string)comboInvalid.SelectedItem;
            string grupZdorov = (string)comboGrupZdorov.SelectedItem;
            string chernobol = (string)comboChernobol.SelectedItem;
            string hobbi = txtHobbi.Text;
            string sirota = (string)comboSirota.SelectedItem;

            try
            {
                using (DatabaseConnection dbConnection = new DatabaseConnection())
                {
                    if (dbConnection.OpenConnection())
                    {
                        if (MainWindow.changing == 0) // Добавление записи
                        {
                            // Создание SQL запроса для добавления записи учащегося
                            string insertQuery = @"
                        INSERT INTO Учащийся (фио, телефон, дата_рождения, соц_бытовые_условия, материальное_состояние, трудности_в_учёбе, состояние_на_учёте, 
                        нарушение_общения, инвалидность, группа_здоровья, чаэс, кружки, сирота)
                        VALUES (@Name, @Phone, @BirthDate, @SocBit, @MatSot, @TrudStady, @Uchet, @Communication, @Invalid, @GrupZdorov, @Chernobol, @Hobbi, @Sirota)";

                            // Создание команды для выполнения SQL запроса
                            SqlCommand insertCommand = new SqlCommand(insertQuery, dbConnection.GetConnection());

                            // Добавление параметров к запросу
                            insertCommand.Parameters.AddWithValue("@Name", name);
                            insertCommand.Parameters.AddWithValue("@Phone", phone);
                            insertCommand.Parameters.AddWithValue("@BirthDate", birthDate);
                            insertCommand.Parameters.AddWithValue("@SocBit", socBit);
                            insertCommand.Parameters.AddWithValue("@MatSot", matSot);
                            insertCommand.Parameters.AddWithValue("@TrudStady", trudStady);
                            insertCommand.Parameters.AddWithValue("@Uchet", uchet);
                            insertCommand.Parameters.AddWithValue("@Communication", communication);
                            insertCommand.Parameters.AddWithValue("@Invalid", invalid);
                            insertCommand.Parameters.AddWithValue("@GrupZdorov", grupZdorov);
                            insertCommand.Parameters.AddWithValue("@Chernobol", chernobol);
                            insertCommand.Parameters.AddWithValue("@Hobbi", hobbi);
                            insertCommand.Parameters.AddWithValue("@Sirota", sirota);

                            // Выполнение запроса на добавление новой записи
                            int rowsAffected = insertCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Новая запись учащегося успешно добавлена.");
                                txtName.Text = "";
                                txtPhone.Text = "";
                                comboSocBit.Text = "";
                                comboMatSot.Text = "";
                                comboTrudStady.Text = "";
                                comboUchet.Text = "";
                                comboCommunication.Text = "";
                                pickerDateOfBirtzd.SelectedDate = null;
                                comboInvalid.Text = "";
                                comboGrupZdorov.Text = "";
                                comboChernobol.Text = "";
                                txtHobbi.Text = "";
                                comboSirota.Text = "";
                            }
                            else
                            {
                                MessageBox.Show("Не удалось добавить новую запись учащегося.");
                            }
                        }
                        else // Изменение записи
                        {
                            // Создание SQL запроса для обновления записи учащегося
                            string updateQuery = @"
        UPDATE Учащийся 
        SET фио = @Name, 
            телефон = @Phone, 
            дата_рождения = @BirthDate, 
            соц_бытовые_условия = @SocBit, 
            материальное_состояние = @MatSot, 
            трудности_в_учёбе = @TrudStady, 
            состояние_на_учёте = @Uchet, 
            нарушение_общения = @Communication,
            инвалидность = @Invalid,
            группа_здоровья = @GrupZdorov,
            чаэс = @Chernobol,
            кружки = @Hobbi,
            сирота = @Sirota
        WHERE id_учащегося = @StudentId";

                            // Создание команды для выполнения SQL запроса
                            SqlCommand updateCommand = new SqlCommand(updateQuery, dbConnection.GetConnection());

                            // Добавление параметров к запросу
                            updateCommand.Parameters.AddWithValue("@Name", name);
                            updateCommand.Parameters.AddWithValue("@Phone", phone);
                            updateCommand.Parameters.AddWithValue("@BirthDate", birthDate);
                            updateCommand.Parameters.AddWithValue("@SocBit", socBit);
                            updateCommand.Parameters.AddWithValue("@MatSot", matSot);
                            updateCommand.Parameters.AddWithValue("@TrudStady", trudStady);
                            updateCommand.Parameters.AddWithValue("@Uchet", uchet);
                            updateCommand.Parameters.AddWithValue("@Communication", communication);
                            updateCommand.Parameters.AddWithValue("@Invalid", invalid);
                            updateCommand.Parameters.AddWithValue("@GrupZdorov", grupZdorov);
                            updateCommand.Parameters.AddWithValue("@Chernobol", chernobol);
                            updateCommand.Parameters.AddWithValue("@Hobbi", hobbi);
                            updateCommand.Parameters.AddWithValue("@Sirota", sirota);
                            updateCommand.Parameters.AddWithValue("@StudentId", MainWindow.element1);

                            // Выполнение запроса на обновление записи
                            int rowsAffected = updateCommand.ExecuteNonQuery();

                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Запись учащегося успешно изменена.");
                            }
                            else
                            {
                                MessageBox.Show("Не удалось изменить запись учащегося.");
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

        //Заполнение данными ComboBox

        //социально бытовые условия
        private void LoadcomboSocBit()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboSocBit != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboSocBit.Items.Clear();

                // Добавляем данные в комбобокс
                comboSocBit.Items.Add("Хорошие");
                comboSocBit.Items.Add("Средние");
                comboSocBit.Items.Add("Плохие");
                comboSocBit.Items.Add("Критические");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //материальное состояние
        private void LoadcomboMatSot()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboMatSot != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboMatSot.Items.Clear();

                // Добавляем данные в комбобокс
                comboMatSot.Items.Add("Выше среднего");
                comboMatSot.Items.Add("Средние");
                comboMatSot.Items.Add("Ниже среднего");
                comboMatSot.Items.Add("Ниже прожиточного минимума");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //трудности в учёбе
        private void LoadcomboTrudStady()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboTrudStady != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboTrudStady.Items.Clear();

                // Добавляем данные в комбобокс
                comboTrudStady.Items.Add("Да");
                comboTrudStady.Items.Add("Средне");
                comboTrudStady.Items.Add("Нет");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //состояние на учёте
        private void LoadcomboUchet()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboUchet != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboUchet.Items.Clear();

                // Добавляем данные в комбобокс
                comboUchet.Items.Add("Да");
                comboUchet.Items.Add("Нет");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //нарушение общения
        private void LoadcomboCommunication()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboCommunication != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboCommunication.Items.Clear();

                // Добавляем данные в комбобокс
                comboCommunication.Items.Add("Асоциальный");
                comboCommunication.Items.Add("Социолизированный");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //инвалидность
        private void LoadcomboInvalid()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboInvalid != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboInvalid.Items.Clear();

                // Добавляем данные в комбобокс
                comboInvalid.Items.Add("1 группа");
                comboInvalid.Items.Add("2 группа");
                comboInvalid.Items.Add("3 группа");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //группа здоровья
        private void LoadcomboGrupZdorov()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboGrupZdorov != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboGrupZdorov.Items.Clear();

                // Добавляем данные в комбобокс
                comboGrupZdorov.Items.Add("1 группа");
                comboGrupZdorov.Items.Add("2 группа");
                comboGrupZdorov.Items.Add("3 группа");
                comboGrupZdorov.Items.Add("4 группа");
                comboGrupZdorov.Items.Add("5 группа");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //чаэс
        private void LoadcomboChernobol()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboChernobol != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboChernobol.Items.Clear();

                // Добавляем данные в комбобокс
                comboChernobol.Items.Add("Да");
                comboChernobol.Items.Add("Нет");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        //чаэс
        private void LoadcomboSirota()
        {

            // Проверяем, что комбобокс существует и не равен null
            if (comboSirota != null)
            {
                // Очищаем комбобокс перед добавлением новых данных
                comboSirota.Items.Clear();

                // Добавляем данные в комбобокс
                comboSirota.Items.Add("Да");
                comboSirota.Items.Add("Нет");
            }
            else
            {
                MessageBox.Show("Комбобокс не найден.");
            }
        }

        private void txtName_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Получаем вводимый символ
            char inputChar = e.Text[0];

            // Проверяем, является ли символ буквой или пробелом
            if (!char.IsLetter(inputChar) && !char.IsWhiteSpace(inputChar))
            {
                // Отменяем ввод, если символ не является буквой или пробелом
                e.Handled = true;
            }
        }

        private void txtPhone_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {
            // Получаем текущий текст в текстовом поле
            string currentText = txtPhone.Text;

            // Получаем вводимый символ
            char inputChar = e.Text[0];

            // Проверяем, является ли вводимый символ цифрой или плюсом
            if (!char.IsDigit(inputChar) && inputChar != '+')
            {
                // Отменяем ввод, если символ не является цифрой или плюсом
                e.Handled = true;
            }

            // Проверяем, чтобы символ "+" вводился только в начале строки
            if (inputChar == '+' && currentText.Length > 0)
            {
                e.Handled = true;
            }

            // Проверяем, чтобы первые три символа были "+375"
            if (currentText.Length == 0 && inputChar != '+')
            {
                e.Handled = true;
            }
            else if (currentText.Length == 1 && inputChar != '3')
            {
                e.Handled = true;
            }
            else if (currentText.Length == 2 && inputChar != '7')
            {
                e.Handled = true;
            }

            // Проверяем, чтобы количество цифр не превышало 13 (включая "+")
            if (currentText.Length >= 13)
            {
                e.Handled = true;
            }
        }

    }
}

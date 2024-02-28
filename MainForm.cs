using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;

namespace drivingschool
{
    public partial class MainForm : Form
    {
        public SqlConnection sqlConnection = null;

        public MainForm()
        {
            InitializeComponent();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["drivingschooldb"].ConnectionString);

            sqlConnection.Open();

            Refreshdball();

        }

        private void Refreshdball()
        {
            Refreshdbstudents();
            Refreshdblocations();
            Refreshdblessontypes();
            RefreshdbScheduleFromSelectDate();

            RefreshTimeComboBox();
            RefreshStudentComboBox();
            RefreshLocationComboBox();
            RefreshLessonTypeComboBox();
        }

        private void Refreshdbstudents()
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT * FROM Students", sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            students_dataGridView.DataSource = dataSet.Tables[0];

            students_dataGridView.Columns["StudentID"].HeaderText = "ID студента";
            students_dataGridView.Columns["FirstName"].HeaderText = "Имя";
            students_dataGridView.Columns["LastName"].HeaderText = "Фамилия";
            students_dataGridView.Columns["BirthDate"].HeaderText = "День рождения";
            students_dataGridView.Columns["Phone"].HeaderText = "Телефон";
        }

        private void Refreshdblocations()
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT * FROM Locations", sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            locations_dataGridView.DataSource = dataSet.Tables[0];

            locations_dataGridView.Columns["LocationID"].HeaderText = "ID локации";
            locations_dataGridView.Columns["Address"].HeaderText = "Адрес";
            locations_dataGridView.Columns["Description"].HeaderText = "Дополнительная информация";
        }

        private void Refreshdblessontypes()
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT * FROM LessonTypes", sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            lessontypes_dataGridView.DataSource = dataSet.Tables[0];

            lessontypes_dataGridView.Columns["LessonTypeID"].HeaderText = "ID типа занятия";
            lessontypes_dataGridView.Columns["Name"].HeaderText = "Именование";
            lessontypes_dataGridView.Columns["InstructorNotes"].HeaderText = "Заметка инструктора";
        }


        private void RefreshdbSchedule()
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT s.ScheduleID, s.LessonDate, s.StartTime, s.EndTime, st.FirstName + ' ' + st.LastName AS StudentName, " +
                "l.Name AS LocationName, lt.Name AS LessonTypeName " +
                "FROM Schedule s " +
                "JOIN Students st ON s.StudentID = st.StudentID " +
                "JOIN Locations l ON s.LocationID = l.LocationID " +
                "JOIN LessonTypes lt ON s.LessonTypeID = lt.LessonTypeID", sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            schedule_dataGridView.DataSource = dataSet.Tables[0];

            schedule_dataGridView.Columns["ScheduleID"].HeaderText = "ID занятия";
            schedule_dataGridView.Columns["LessonDate"].HeaderText = "Дата занятия";
            schedule_dataGridView.Columns["StartTime"].HeaderText = "Время начала";
            schedule_dataGridView.Columns["EndTime"].HeaderText = "Время конца";
            schedule_dataGridView.Columns["StudentName"].HeaderText = "Курсант";
            schedule_dataGridView.Columns["LocationName"].HeaderText = "Локация";
            schedule_dataGridView.Columns["LessonTypeName"].HeaderText = "Тип занятия";
        }

        private void add_students_button_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(firstname_students_textBox.Text) ||
                string.IsNullOrWhiteSpace(lastname_students_textBox.Text) ||
                string.IsNullOrWhiteSpace(birthdate_students_textBox.Text) ||
                string.IsNullOrWhiteSpace(phone_students_textBox.Text))
            {
                MessageBox.Show("Введите данные во все поля", "Ошибка");
                return;
            }

            if (!DateTime.TryParseExact(birthdate_students_textBox.Text, "dd.MM.yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime birthDate))
            {
                MessageBox.Show("Неверный формат даты рождения. Используйте формат ДД.ММ.ГГГГ", "Ошибка");
                return;
            }

            SqlCommand command = new SqlCommand(
                $"INSERT INTO [Students] (FirstName, LastName, BirthDate, Phone) VALUES (@FirstName, @LastName, @BirthDate, @Phone)",
                sqlConnection);

            DateTime date = DateTime.Parse(birthdate_students_textBox.Text);

            command.Parameters.AddWithValue("FirstName", firstname_students_textBox.Text);
            command.Parameters.AddWithValue("LastName", lastname_students_textBox.Text);
            command.Parameters.AddWithValue("BirthDate", birthDate.ToString("yyyy-MM-dd"));
            command.Parameters.AddWithValue("Phone", phone_students_textBox.Text);

            int rowsAffected = command.ExecuteNonQuery();
            if (rowsAffected == 1)
            {
                MessageBox.Show($"Данные успешно добавлены:\n\nИмя: {firstname_students_textBox.Text}\nФамилия: {lastname_students_textBox.Text}\nДата рождения: {birthdate_students_textBox.Text}\nТелефон: {phone_students_textBox.Text}", "Успех");
                firstname_students_textBox.Clear();
                lastname_students_textBox.Clear();
                birthdate_students_textBox.Clear();
                phone_students_textBox.Clear();
            }
            else
            {
                MessageBox.Show("Не удалось добавить данные.", "Ошибка");
            }
            Refreshdbstudents();
        }

        private void add_locations_button_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(name_locations_textBox.Text) ||
                string.IsNullOrWhiteSpace(address_locations_textBox.Text) ||
                string.IsNullOrWhiteSpace(description_locations_textBox.Text))
            {
                MessageBox.Show("Введите данные во все поля", "Ошибка");
                return;
            }

            SqlCommand command = new SqlCommand(
                $"INSERT INTO [Locations] (Name, Address, Description) VALUES (@Name, @Address, @Description)",
                sqlConnection);

            command.Parameters.AddWithValue("Name", name_locations_textBox.Text);
            command.Parameters.AddWithValue("Address", address_locations_textBox.Text);
            command.Parameters.AddWithValue("Description", description_locations_textBox.Text);

            int rowsAffected = command.ExecuteNonQuery();
            if (rowsAffected == 1)
            {
                MessageBox.Show($"Данные успешно добавлены:\n\nНазвание: {name_locations_textBox.Text}\nАдрес: {address_locations_textBox.Text}\nОписание: {description_locations_textBox.Text}", "Успех");
                name_locations_textBox.Clear();
                address_locations_textBox.Clear();
                description_locations_textBox.Clear();
            }
            else
            {
                MessageBox.Show("Не удалось добавить данные.", "Ошибка");
            }
            Refreshdblocations();
        }

        private void add_lessontypes_button_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(name_lessontypes_textBox.Text) ||
                string.IsNullOrWhiteSpace(description_lessontypes_textBox.Text))
            {
                MessageBox.Show("Введите данные во все поля", "Ошибка");
                return;
            }

            SqlCommand command = new SqlCommand(
                $"INSERT INTO [LessonTypes] (Name, Description, InstructorNotes) VALUES (@Name, @Description, @InstructorNotes)",
                sqlConnection);

            command.Parameters.AddWithValue("Name", name_lessontypes_textBox.Text);
            command.Parameters.AddWithValue("Description", description_lessontypes_textBox.Text);
            command.Parameters.AddWithValue("InstructorNotes", instructornotes_lessontypes_textBox.Text);

            int rowsAffected = command.ExecuteNonQuery();
            if (rowsAffected == 1)
            {
                MessageBox.Show($"Данные успешно добавлены:\n\nНазвание: {name_lessontypes_textBox.Text}\nОписание: {description_lessontypes_textBox.Text}\nЗаметки инструктора: {instructornotes_lessontypes_textBox.Text}", "Успех");
                name_lessontypes_textBox.Clear();
                description_lessontypes_textBox.Clear();
                instructornotes_lessontypes_textBox.Clear();
            }
            else
            {
                MessageBox.Show("Не удалось добавить данные.", "Ошибка");
            }
            Refreshdblessontypes();
        }


        private void RefreshStudentComboBox()
        {
            studentID_schedule_comboBox.Items.Clear();
            SqlCommand command = new SqlCommand("SELECT StudentID, FirstName, LastName FROM Students", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                Student student = new Student
                {
                    StudentID = Convert.ToInt32(reader["StudentID"]),
                    FirstName = reader["FirstName"].ToString(),
                    LastName = reader["LastName"].ToString()
                };
                studentID_schedule_comboBox.Items.Add(student);
            }
            reader.Close();
        }

        private void RefreshLocationComboBox()
        {
            locationID_schedule_comboBox.Items.Clear();
            SqlCommand command = new SqlCommand("SELECT LocationID, Name FROM Locations", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                Location location = new Location
                {
                    LocationID = Convert.ToInt32(reader["LocationID"]),
                    Name = reader["Name"].ToString()
                };
                locationID_schedule_comboBox.Items.Add(location);
            }
            reader.Close();
        }

        private void RefreshLessonTypeComboBox()
        {
            lessontypeID_schedule_comboBox.Items.Clear();
            SqlCommand command = new SqlCommand("SELECT LessonTypeID, Name FROM LessonTypes", sqlConnection);
            SqlDataReader reader = command.ExecuteReader();
            while (reader.Read())
            {
                LessonType lessonType = new LessonType
                {
                    LessonTypeID = Convert.ToInt32(reader["LessonTypeID"]),
                    Name = reader["Name"].ToString()
                };
                lessontypeID_schedule_comboBox.Items.Add(lessonType);
            }
            reader.Close();
        }

        private void RefreshTimeComboBox()
        {
            for (int hour = 6; hour <= 20; hour++)
            {
                for (int minute = 0; minute <= 30; minute += 30)
                {
                    string time = string.Format("{0:D2}:{1:D2}", hour, minute);

                    // Добавляем время в комбо-боксы
                    starttime_schedule_comboBox.Items.Add(time);
                    endtime_schedule_comboBox.Items.Add(time);
                }
            }
        }

        private void add_schedule_button_Click(object sender, EventArgs e)
        {
            if (studentID_schedule_comboBox.SelectedItem == null ||
                locationID_schedule_comboBox.SelectedItem == null ||
                lessontypeID_schedule_comboBox.SelectedItem == null ||
                string.IsNullOrWhiteSpace(lessondate_schedule_dateTimePicker.Text) ||
                string.IsNullOrWhiteSpace(starttime_schedule_comboBox.Text) ||
                string.IsNullOrWhiteSpace(endtime_schedule_comboBox.Text))
            {
                MessageBox.Show("Выберите курсанта, локацию, тип занятия и укажите дату и время занятия из выпадающих списков", "Ошибка");
                return;
            }

            int studentID = ((Student)studentID_schedule_comboBox.SelectedItem).StudentID;
            int locationID = ((Location)locationID_schedule_comboBox.SelectedItem).LocationID;
            int lessontypeID = ((LessonType)lessontypeID_schedule_comboBox.SelectedItem).LessonTypeID;
            DateTime lessonDate = lessondate_schedule_dateTimePicker.Value.Date;
            TimeSpan startTime = TimeSpan.Parse(starttime_schedule_comboBox.Text);
            TimeSpan endTime = TimeSpan.Parse(endtime_schedule_comboBox.Text);

            if (IsLessonOverlap(lessonDate, startTime, endTime))
            {
                MessageBox.Show($"Невозможно добавить занятие. Промежуток времени {lessonDate.ToShortDateString()} c {startTime} по {endTime} занят.", "Ошибка");
                return;
            }

            SqlCommand command = new SqlCommand(
                @"INSERT INTO Schedule (LessonDate, StartTime, EndTime, StudentID, LocationID, LessonTypeID)
            VALUES (@LessonDate, @StartTime, @EndTime, @StudentID, @LocationID, @LessonTypeID)", sqlConnection);

            command.Parameters.AddWithValue("@LessonDate", lessonDate);
            command.Parameters.AddWithValue("@StartTime", startTime);
            command.Parameters.AddWithValue("@EndTime", endTime);
            command.Parameters.AddWithValue("@StudentID", studentID);
            command.Parameters.AddWithValue("@LocationID", locationID);
            command.Parameters.AddWithValue("@LessonTypeID", lessontypeID);
            
            int rowsAffected = command.ExecuteNonQuery();
            if (rowsAffected == 1)
            {
                MessageBox.Show("Занятие успешно добавлено.", "Успех");
                RefreshdbScheduleFromSelectDate();
            }
            else
            {
                MessageBox.Show("Не удалось добавить занятие.", "Ошибка");
            }
        }

        // Метод для проверки пересечения занятий
        private bool IsLessonOverlap(DateTime lessonDate, TimeSpan startTime, TimeSpan endTime)
        {
            SqlCommand command = new SqlCommand(
                @"SELECT COUNT(*) 
                FROM Schedule 
                WHERE LessonDate = @LessonDate 
                AND (
                ((@StartTime >= StartTime AND @StartTime < EndTime) OR (@EndTime > StartTime AND @EndTime <= EndTime))
                OR (StartTime >= @StartTime AND EndTime <= @EndTime)
                )", sqlConnection);

            command.Parameters.AddWithValue("@LessonDate", lessonDate);
            command.Parameters.AddWithValue("@StartTime", startTime);
            command.Parameters.AddWithValue("@EndTime", endTime);

            int count = (int)command.ExecuteScalar();
            return count > 0;
        }

        private void change_lessondate_scheduledateTimePicker_ValueChanged(object sender, EventArgs e)
        {
            RefreshdbScheduleFromSelectDate();
        }

        private void allschedule_button_Click(object sender, EventArgs e)
        {
            RefreshdbSchedule();
        }

        private void RefreshdbScheduleFromSelectDate()
        {
            DateTime selectedDate = change_lessondate_scheduledateTimePicker.Value.Date;

            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT s.ScheduleID, s.LessonDate, s.StartTime, s.EndTime, st.FirstName + ' ' + st.LastName AS StudentName, " +
                "l.Name AS LocationName, lt.Name AS LessonTypeName " +
                "FROM Schedule s " +
                "JOIN Students st ON s.StudentID = st.StudentID " +
                "JOIN Locations l ON s.LocationID = l.LocationID " +
                "JOIN LessonTypes lt ON s.LessonTypeID = lt.LessonTypeID " +
                "WHERE s.LessonDate = @LessonDate", sqlConnection);

            dataAdapter.SelectCommand.Parameters.AddWithValue("@LessonDate", selectedDate);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            schedule_dataGridView.DataSource = dataSet.Tables[0];

            schedule_dataGridView.Columns["ScheduleID"].HeaderText = "ID занятия";
            schedule_dataGridView.Columns["LessonDate"].HeaderText = "Дата занятия";
            schedule_dataGridView.Columns["StartTime"].HeaderText = "Время начала";
            schedule_dataGridView.Columns["EndTime"].HeaderText = "Время конца";
            schedule_dataGridView.Columns["StudentName"].HeaderText = "Курсант";
            schedule_dataGridView.Columns["LocationName"].HeaderText = "Локация";
            schedule_dataGridView.Columns["LessonTypeName"].HeaderText = "Тип занятия";
        }

        private void students_dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && students_dataGridView.Rows[e.RowIndex].Selected)
            {
                DataGridViewRow selectedRow = students_dataGridView.Rows[e.RowIndex];

                string firstName = selectedRow.Cells["FirstName"].Value.ToString();
                string lastName = selectedRow.Cells["LastName"].Value.ToString();
                DateTime birthDate = (DateTime)selectedRow.Cells["BirthDate"].Value;
                string phone = selectedRow.Cells["Phone"].Value.ToString();

                firstname_students_textBox.Text = firstName;
                lastname_students_textBox.Text = lastName;
                birthdate_students_textBox.Text = birthDate.ToString("dd.MM.yyyy");
                phone_students_textBox.Text = phone;

                add_students_button.Enabled = false;
            }
            else
            {
                add_students_button.Enabled = true;

                firstname_students_textBox.Clear();
                lastname_students_textBox.Clear();
                birthdate_students_textBox.Clear();
                phone_students_textBox.Clear();
            }
        }

        private void locations_dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && locations_dataGridView.Rows[e.RowIndex].Selected)
            {
                DataGridViewRow selectedRow = locations_dataGridView.Rows[e.RowIndex];

                string name = selectedRow.Cells["Name"].Value.ToString();
                string address = selectedRow.Cells["Address"].Value.ToString();
                string description = selectedRow.Cells["Description"].Value.ToString();

                name_locations_textBox.Text = name;
                address_locations_textBox.Text = address;
                description_locations_textBox.Text = description;

                add_locations_button.Enabled = false;
            }
            else
            {
                add_locations_button.Enabled = true;

                name_locations_textBox.Clear();
                address_locations_textBox.Clear();
                description_locations_textBox.Clear();
            }
        }

        private void lessontypes_dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && lessontypes_dataGridView.Rows[e.RowIndex].Selected)
            {
                DataGridViewRow selectedRow = lessontypes_dataGridView.Rows[e.RowIndex];

                string name = selectedRow.Cells["Name"].Value.ToString();
                string description = selectedRow.Cells["Description"].Value.ToString();
                string instructorNotes = selectedRow.Cells["InstructorNotes"].Value.ToString();

                name_lessontypes_textBox.Text = name;
                description_lessontypes_textBox.Text = description;
                instructornotes_lessontypes_textBox.Text = instructorNotes;

                add_lessontypes_button.Enabled = false;
            }
            else
            {
                add_lessontypes_button.Enabled = true;

                name_lessontypes_textBox.Clear();
                description_lessontypes_textBox.Clear();
                instructornotes_lessontypes_textBox.Clear();
            }
        }

        private void schedule_dataGridView_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && schedule_dataGridView.Rows[e.RowIndex].Selected)
            {
                DataGridViewRow selectedRow = schedule_dataGridView.Rows[e.RowIndex];

                DateTime lessonDate = Convert.ToDateTime(selectedRow.Cells["LessonDate"].Value);
                TimeSpan startTime = TimeSpan.Parse(selectedRow.Cells["StartTime"].Value.ToString());
                TimeSpan endTime = TimeSpan.Parse(selectedRow.Cells["EndTime"].Value.ToString());
                string studentName = selectedRow.Cells["StudentName"].Value.ToString();
                string locationName = selectedRow.Cells["LocationName"].Value.ToString();
                string lessonTypeName = selectedRow.Cells["LessonTypeName"].Value.ToString();

                lessondate_schedule_dateTimePicker.Value = lessonDate;
                starttime_schedule_comboBox.SelectedItem = startTime.ToString("hh\\:mm");
                endtime_schedule_comboBox.SelectedItem = endTime.ToString("hh\\:mm");
                studentID_schedule_comboBox.Text = studentName;
                locationID_schedule_comboBox.Text = locationName;
                lessontypeID_schedule_comboBox.Text = lessonTypeName;

                add_schedule_button.Enabled = false;
            }
            else
            {
                add_schedule_button.Enabled = true;

                // Очистите значения элементов управления при снятии выделения
                lessondate_schedule_dateTimePicker.Value = DateTime.Today;
                starttime_schedule_comboBox.SelectedIndex = -1;
                endtime_schedule_comboBox.SelectedIndex = -1;
                studentID_schedule_comboBox.SelectedIndex = -1;
                locationID_schedule_comboBox.SelectedIndex = -1;
                lessontypeID_schedule_comboBox.SelectedIndex = -1;
            }
        }
    }
}

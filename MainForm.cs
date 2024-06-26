﻿using System;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Reporting.WinForms;

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

            change_lessondate_scheduledateTimePicker.Value = DateTime.Now.Date;
            Refreshdball();
            RefreshdbScheduleFromSelectDate();
            this.students_reportViewer.RefreshReport();
            this.schedule_reportViewer.RefreshReport();
        }


        private void Refreshdball()
        {
            Refreshdbstudents();
            Refreshdblocations();
            Refreshdblessontypes();


            RefreshTimeComboBox();
            RefreshStudentComboBox();
            RefreshLocationComboBox();
            RefreshLessonTypeComboBox();


            FillColumnsStudentsComboBox();
            FillColumnsLocationsComboBox();
            FillColumnsLessonTypesComboBox();
            FillColumnsScheduleComboBox();
        }


        private void FillColumnsStudentsComboBox()
        {
            changecolums_students_comboBox.Items.Clear();
            changecolums_students_comboBox.Items.Add("Все столбцы");

            foreach (DataGridViewColumn column in students_dataGridView.Columns)
            {
                changecolums_students_comboBox.Items.Add(column.HeaderText);
            }

            changecolums_students_comboBox.SelectedIndex = 0;

            changecolums_students_report_comboBox.Items.Clear();
            changecolums_students_report_comboBox.Items.Add("Все столбцы");

            foreach (DataGridViewColumn column in students_dataGridView.Columns)
            {
                changecolums_students_report_comboBox.Items.Add(column.HeaderText);
            }

            changecolums_students_report_comboBox.SelectedIndex = 0;
        }

        private void FillColumnsLocationsComboBox()
        {
            changecolums_locations_comboBox.Items.Clear();
            changecolums_locations_comboBox.Items.Add("Все столбцы");

            foreach (DataGridViewColumn column in locations_dataGridView.Columns)
            {
                changecolums_locations_comboBox.Items.Add(column.HeaderText);
            }

            changecolums_locations_comboBox.SelectedIndex = 0;
        }


        private void FillColumnsLessonTypesComboBox()
        {
            changecolums_lessontypes_comboBox.Items.Clear();
            changecolums_lessontypes_comboBox.Items.Add("Все столбцы");

            foreach (DataGridViewColumn column in lessontypes_dataGridView.Columns)
            {
                changecolums_lessontypes_comboBox.Items.Add(column.HeaderText);
            }

            changecolums_lessontypes_comboBox.SelectedIndex = 0;
        }

        private void FillColumnsScheduleComboBox()
        {
            changecolums_schedule_comboBox.Items.Clear();
            changecolums_schedule_comboBox.Items.Add("Все столбцы");

            foreach (DataGridViewColumn column in schedule_dataGridView.Columns)
            {
                changecolums_schedule_comboBox.Items.Add(column.HeaderText);
            }

            changecolums_schedule_comboBox.SelectedIndex = 0;

            changecolums_schedule_report_comboBox.Items.Clear();
            changecolums_schedule_report_comboBox.Items.Add("Все столбцы");

            foreach (DataGridViewColumn column in schedule_dataGridView.Columns)
            {
                changecolums_schedule_report_comboBox.Items.Add(column.HeaderText);
            }

            changecolums_schedule_report_comboBox.SelectedIndex = 0;
        }

        private void Refreshdbstudents()
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT * FROM Students", sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            students_dataGridView.DataSource = dataSet.Tables[0];
            students_dataGridView.Columns["BirthDate"].HeaderText = "День рождения";
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
            locations_dataGridView.Columns["Name"].HeaderText = "Название";
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
            lessontypes_dataGridView.Columns["Description"].HeaderText = "Описание";
            lessontypes_dataGridView.Columns["InstructorNotes"].HeaderText = "Заметка инструктора";
        }


        private void RefreshdbSchedule()
        {
            SqlDataAdapter dataAdapter = new SqlDataAdapter(
                "SELECT s.ScheduleID, s.LessonDate, s.StartTime, s.EndTime, st.FirstName + ' ' + st.LastName AS StudentName, " +
                "l.Name AS LocationName, lt.Name AS LessonTypeName, s.Mark " +
                "FROM Schedule s " +
                "JOIN Students st ON s.StudentID = st.StudentID " +
                "JOIN Locations l ON s.LocationID = l.LocationID " +
                "JOIN LessonTypes lt ON s.LessonTypeID = lt.LessonTypeID", sqlConnection);

            DataSet dataSet = new DataSet();
            dataAdapter.Fill(dataSet);

            schedule_dataGridView.DataSource = dataSet.Tables[0];

            // Установка заголовков столбцов
            schedule_dataGridView.Columns["ScheduleID"].HeaderText = "ID занятия";
            schedule_dataGridView.Columns["LessonDate"].HeaderText = "Дата занятия";
            schedule_dataGridView.Columns["StartTime"].HeaderText = "Время начала";
            schedule_dataGridView.Columns["EndTime"].HeaderText = "Время конца";
            schedule_dataGridView.Columns["StudentName"].HeaderText = "Курсант";
            schedule_dataGridView.Columns["LocationName"].HeaderText = "Локация";
            schedule_dataGridView.Columns["LessonTypeName"].HeaderText = "Тип занятия";
            schedule_dataGridView.Columns["Mark"].HeaderText = "Оценка"; 
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
            RefreshStudentComboBox();
            UpdateReportStudents();
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
            RefreshLocationComboBox();
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
            RefreshLessonTypeComboBox();
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

        private void RefreshdbScheduleFromSelectDate()
        {
            DateTime selectedDate = change_lessondate_scheduledateTimePicker.Value.Date;

            SqlDataAdapter dataAdapter;

            if (allschedule_checkbox.Checked)
            {
                dataAdapter = new SqlDataAdapter(
                    "SELECT s.ScheduleID, s.LessonDate, s.StartTime, s.EndTime, st.FirstName + ' ' + st.LastName AS StudentName, " +
                    "l.Name AS LocationName, lt.Name AS LessonTypeName, s.Mark " +
                    "FROM Schedule s " +
                    "JOIN Students st ON s.StudentID = st.StudentID " +
                    "JOIN Locations l ON s.LocationID = l.LocationID " +
                    "JOIN LessonTypes lt ON s.LessonTypeID = lt.LessonTypeID", sqlConnection);
            }
            else
            {
                dataAdapter = new SqlDataAdapter(
                    "SELECT s.ScheduleID, s.LessonDate, s.StartTime, s.EndTime, st.FirstName + ' ' + st.LastName AS StudentName, " +
                    "l.Name AS LocationName, lt.Name AS LessonTypeName, s.Mark " +
                    "FROM Schedule s " +
                    "JOIN Students st ON s.StudentID = st.StudentID " +
                    "JOIN Locations l ON s.LocationID = l.LocationID " +
                    "JOIN LessonTypes lt ON s.LessonTypeID = lt.LessonTypeID " +
                    "WHERE s.LessonDate = @LessonDate", sqlConnection);

                dataAdapter.SelectCommand.Parameters.AddWithValue("@LessonDate", selectedDate);
            }

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
            schedule_dataGridView.Columns["Mark"].HeaderText = "Оценка";
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

                int studentID = (int)selectedRow.Cells["StudentID"].Value;

                string queryStudentLessons = "SELECT COUNT(*) AS StudentLessons FROM Schedule WHERE StudentID = @StudentID";
                string queryFirstLesson = "SELECT MIN(LessonDate) AS FirstLesson FROM Schedule WHERE StudentID = @StudentID";
                string queryLastLesson = "SELECT MAX(LessonDate) AS LastLesson FROM Schedule WHERE StudentID = @StudentID";
                string queryTotalMarks = "SELECT COUNT(*) AS TotalMarks FROM Schedule WHERE StudentID = @StudentID AND Mark IS NOT NULL";
                string queryAverageMark = "SELECT AVG(CAST(Mark AS FLOAT)) AS AverageMark FROM Schedule WHERE StudentID = @StudentID";
                string queryZeroMarks = "SELECT COUNT(*) AS ZeroMarks FROM Schedule WHERE StudentID = @StudentID AND Mark = 0";

                SqlCommand commandStudentLessons = new SqlCommand(queryStudentLessons, sqlConnection);
                commandStudentLessons.Parameters.AddWithValue("@StudentID", studentID);

                SqlCommand commandFirstLesson = new SqlCommand(queryFirstLesson, sqlConnection);
                commandFirstLesson.Parameters.AddWithValue("@StudentID", studentID);

                SqlCommand commandLastLesson = new SqlCommand(queryLastLesson, sqlConnection);
                commandLastLesson.Parameters.AddWithValue("@StudentID", studentID);

                SqlCommand commandTotalMarks = new SqlCommand(queryTotalMarks, sqlConnection);
                commandTotalMarks.Parameters.AddWithValue("@StudentID", studentID);

                SqlCommand commandAverageMark = new SqlCommand(queryAverageMark, sqlConnection);
                commandAverageMark.Parameters.AddWithValue("@StudentID", studentID);

                SqlCommand commandZeroMarks = new SqlCommand(queryZeroMarks, sqlConnection);
                commandZeroMarks.Parameters.AddWithValue("@StudentID", studentID);

                int studentLessons = 0;
                DateTime? firstLessonDate = null;
                DateTime? lastLessonDate = null;
                int totalMarks = 0;
                double averageMark = 0;
                int zeroMarks = 0;

                SqlDataReader reader = null;
                try
                {
                    reader = commandStudentLessons.ExecuteReader();
                    if (reader.Read() && !reader.IsDBNull(reader.GetOrdinal("StudentLessons")))
                    {
                        studentLessons = Convert.ToInt32(reader["StudentLessons"]);
                    }
                    reader.Close();

                    reader = commandFirstLesson.ExecuteReader();
                    if (reader.Read() && !reader.IsDBNull(reader.GetOrdinal("FirstLesson")))
                    {
                        firstLessonDate = Convert.ToDateTime(reader["FirstLesson"]);
                    }
                    reader.Close();

                    reader = commandLastLesson.ExecuteReader();
                    if (reader.Read() && !reader.IsDBNull(reader.GetOrdinal("LastLesson")))
                    {
                        lastLessonDate = Convert.ToDateTime(reader["LastLesson"]);
                    }
                    reader.Close();

                    reader = commandTotalMarks.ExecuteReader();
                    if (reader.Read() && !reader.IsDBNull(reader.GetOrdinal("TotalMarks")))
                    {
                        totalMarks = Convert.ToInt32(reader["TotalMarks"]);
                    }
                    reader.Close();

                    reader = commandAverageMark.ExecuteReader();
                    if (reader.Read() && !reader.IsDBNull(reader.GetOrdinal("AverageMark")))
                    {
                        averageMark = Convert.ToDouble(reader["AverageMark"]);
                    }
                    reader.Close();

                    reader = commandZeroMarks.ExecuteReader();
                    if (reader.Read() && !reader.IsDBNull(reader.GetOrdinal("ZeroMarks")))
                    {
                        zeroMarks = Convert.ToInt32(reader["ZeroMarks"]);
                    }
                }
                finally
                {
                    reader?.Close();
                }

                double percent = 25 == 0 ? 0 : ((double)studentLessons / 25) * 100;

                firstname_students_textBox.Text = firstName;
                lastname_students_textBox.Text = lastName;
                birthdate_students_textBox.Text = birthDate.ToString("dd.MM.yyyy");
                phone_students_textBox.Text = phone;
                info_students_textBox.Text = $"Пройдено занятий: {studentLessons} из 25 ({percent:F2}%)." +
                                                $"{Environment.NewLine}" +
                                                $"{Environment.NewLine}Первое занятие в расписании: {firstLessonDate?.ToString("dd.MM.yyyy") ?? "нет данных в расписании"}" +
                                                $"{Environment.NewLine}Последнее занятие в расписании: {lastLessonDate?.ToString("dd.MM.yyyy") ?? "нет данных в расписании"}" +
                                                $"{Environment.NewLine}" +
                                                $"{Environment.NewLine}Общее количество оценок: {totalMarks}" +
                                                $"{Environment.NewLine}Средний балл: {averageMark:F2}" +
                                                $"{Environment.NewLine}" +
                                                $"{Environment.NewLine}Количество несданных занятий: {zeroMarks}";

                add_students_button.Enabled = false;
            }
            else
            {
                firstname_students_textBox.Clear();
                lastname_students_textBox.Clear();
                birthdate_students_textBox.Clear();
                phone_students_textBox.Clear();
                info_students_textBox.Clear();

                add_students_button.Enabled = true;
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
                string mark = selectedRow.Cells["Mark"].Value.ToString();

                lessondate_schedule_dateTimePicker.Value = lessonDate;
                starttime_schedule_comboBox.SelectedItem = startTime.ToString("hh\\:mm");
                endtime_schedule_comboBox.SelectedItem = endTime.ToString("hh\\:mm");
                studentID_schedule_comboBox.Text = studentName;
                locationID_schedule_comboBox.Text = locationName;
                lessontypeID_schedule_comboBox.Text = lessonTypeName;
                mark_schedule_textBox.Text = mark;

                add_schedule_button.Enabled = false;
                mark_schedule_label.Visible = true;
                mark_schedule_textBox.Visible = true;
            }
            else
            {
                mark_schedule_label.Visible = false;
                mark_schedule_textBox.Visible = false;

                add_schedule_button.Enabled = true;

                lessondate_schedule_dateTimePicker.Value = DateTime.Today;
                starttime_schedule_comboBox.SelectedIndex = -1;
                endtime_schedule_comboBox.SelectedIndex = -1;
                studentID_schedule_comboBox.SelectedIndex = -1;
                locationID_schedule_comboBox.SelectedIndex = -1;
                lessontypeID_schedule_comboBox.SelectedIndex = -1;
                mark_schedule_textBox.Clear();
            }
        }

        private void delete_students_button_Click(object sender, EventArgs e)
        {
            if (students_dataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите студентов для удаления", "Ошибка");
                return;
            }

            StringBuilder deletedStudents = new StringBuilder();

            foreach (DataGridViewRow selectedRow in students_dataGridView.SelectedRows)
            {
                string firstName = selectedRow.Cells["FirstName"].Value.ToString();
                string lastName = selectedRow.Cells["LastName"].Value.ToString();
                int studentID = Convert.ToInt32(selectedRow.Cells["StudentID"].Value);

                SqlCommand checkScheduleCommand = new SqlCommand(
                    "SELECT COUNT(*) FROM [Schedule] WHERE StudentID = @StudentID",
                    sqlConnection);

                checkScheduleCommand.Parameters.AddWithValue("StudentID", studentID);
                int scheduleCount = (int)checkScheduleCommand.ExecuteScalar();

                string message = $"Вы уверены, что хотите удалить студента {firstName} {lastName} из базы данных?";
                if (scheduleCount > 0)
                {
                    message += $"\nУ этого студента есть записи в расписании ({scheduleCount} записей). Хотите удалить их вместе с курсантом?";
                }

                DialogResult dialogResult = MessageBox.Show(message, "Подтверждение удаления", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (scheduleCount > 0)
                    {
                        SqlCommand deleteScheduleCommand = new SqlCommand(
                            "DELETE FROM [Schedule] WHERE StudentID = @StudentID",
                            sqlConnection);

                        deleteScheduleCommand.Parameters.AddWithValue("StudentID", studentID);
                        deleteScheduleCommand.ExecuteNonQuery();
                    }

                    SqlCommand deleteStudentCommand = new SqlCommand(
                        "DELETE FROM [Students] WHERE StudentID = @StudentID",
                        sqlConnection);
                    deleteStudentCommand.Parameters.AddWithValue("StudentID", studentID);

                    int rowsAffected = deleteStudentCommand.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        deletedStudents.AppendLine($"{firstName} {lastName}");
                    }
                }
            }

            string deletedStudentsMessage = "Следующие студенты были успешно удалены из базы данных:\n\n" + deletedStudents.ToString();
            if (!string.IsNullOrWhiteSpace(deletedStudents.ToString()))
            {
                MessageBox.Show(deletedStudentsMessage, "Успех");
                RefreshStudentComboBox();
                Refreshdbstudents();
                UpdateReportStudents();
            }
            else
            {
                MessageBox.Show("Не удалось удалить выбранных студентов из базы данных.", "Ошибка");
            }
        }

        private void delete_locations_button_Click(object sender, EventArgs e)
        {
            if (locations_dataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите местоположения для удаления", "Ошибка");
                return;
            }

            StringBuilder deletedLocations = new StringBuilder();

            foreach (DataGridViewRow selectedRow in locations_dataGridView.SelectedRows)
            {
                string locationName = selectedRow.Cells["Name"].Value.ToString();
                int locationID = Convert.ToInt32(selectedRow.Cells["LocationID"].Value);

                SqlCommand checkScheduleCommand = new SqlCommand(
                    "SELECT COUNT(*) FROM [Schedule] WHERE LocationID = @LocationID",
                    sqlConnection);

                checkScheduleCommand.Parameters.AddWithValue("LocationID", locationID);
                int scheduleCount = (int)checkScheduleCommand.ExecuteScalar();

                string message = $"Вы уверены, что хотите удалить местоположение '{locationName}' из базы данных?";
                if (scheduleCount > 0)
                {
                    message += $"\nУ этого местоположения есть записи в расписании ({scheduleCount} записей). Хотите удалить их вместе с местоположением?";
                }

                DialogResult dialogResult = MessageBox.Show(message, "Подтверждение удаления", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (scheduleCount > 0)
                    {
                        SqlCommand deleteScheduleCommand = new SqlCommand(
                            "DELETE FROM [Schedule] WHERE LocationID = @LocationID",
                            sqlConnection);

                        deleteScheduleCommand.Parameters.AddWithValue("LocationID", locationID);
                        deleteScheduleCommand.ExecuteNonQuery();
                    }

                    SqlCommand deleteLocationCommand = new SqlCommand(
                        "DELETE FROM [Locations] WHERE LocationID = @LocationID",
                        sqlConnection);
                    deleteLocationCommand.Parameters.AddWithValue("LocationID", locationID);

                    int rowsAffected = deleteLocationCommand.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        deletedLocations.AppendLine(locationName);
                    }
                }
            }

            string deletedLocationsMessage = "Следующие местоположения были успешно удалены из базы данных:\n\n" + deletedLocations.ToString();
            if (!string.IsNullOrWhiteSpace(deletedLocations.ToString()))
            {
                MessageBox.Show(deletedLocationsMessage, "Успех");
                Refreshdblocations();
                RefreshLocationComboBox();
            }
            else
            {
                MessageBox.Show("Не удалось удалить выбранные местоположения из базы данных.", "Ошибка");
            }
        }

        private void delete_lessontypes_button_Click(object sender, EventArgs e)
        {
            if (lessontypes_dataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите типы занятий для удаления", "Ошибка");
                return;
            }

            StringBuilder deletedLessonTypes = new StringBuilder();

            foreach (DataGridViewRow selectedRow in lessontypes_dataGridView.SelectedRows)
            {
                string lessonTypeName = selectedRow.Cells["Name"].Value.ToString();
                int lessonTypeID = Convert.ToInt32(selectedRow.Cells["LessonTypeID"].Value);

                SqlCommand checkScheduleCommand = new SqlCommand(
                    "SELECT COUNT(*) FROM [Schedule] WHERE LessonTypeID = @LessonTypeID",
                    sqlConnection);

                checkScheduleCommand.Parameters.AddWithValue("LessonTypeID", lessonTypeID);
                int scheduleCount = (int)checkScheduleCommand.ExecuteScalar();

                string message = $"Вы уверены, что хотите удалить тип занятия '{lessonTypeName}' из базы данных?";
                if (scheduleCount > 0)
                {
                    message += $"\nУ этого типа занятия есть записи в расписании ({scheduleCount} записей). Хотите удалить их вместе с типом занятия?";
                }

                DialogResult dialogResult = MessageBox.Show(message, "Подтверждение удаления", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    if (scheduleCount > 0)
                    {
                        SqlCommand deleteScheduleCommand = new SqlCommand(
                            "DELETE FROM [Schedule] WHERE LessonTypeID = @LessonTypeID",
                            sqlConnection);

                        deleteScheduleCommand.Parameters.AddWithValue("LessonTypeID", lessonTypeID);
                        deleteScheduleCommand.ExecuteNonQuery();
                    }

                    SqlCommand deleteLessonTypeCommand = new SqlCommand(
                        "DELETE FROM [LessonTypes] WHERE LessonTypeID = @LessonTypeID",
                        sqlConnection);
                    deleteLessonTypeCommand.Parameters.AddWithValue("LessonTypeID", lessonTypeID);

                    int rowsAffected = deleteLessonTypeCommand.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        deletedLessonTypes.AppendLine(lessonTypeName);
                    }
                }
            }

            string deletedLessonTypesMessage = "Следующие типы занятий были успешно удалены из базы данных:\n\n" + deletedLessonTypes.ToString();
            if (!string.IsNullOrWhiteSpace(deletedLessonTypes.ToString()))
            {
                MessageBox.Show(deletedLessonTypesMessage, "Успех");
                Refreshdblessontypes();
                RefreshLessonTypeComboBox();
            }
            else
            {
                MessageBox.Show("Не удалось удалить выбранные типы занятий из базы данных.", "Ошибка");
            }
        }

        private void delete_schedule_button_Click(object sender, EventArgs e)
        {
            if (schedule_dataGridView.SelectedRows.Count == 0)
            {
                MessageBox.Show("Выберите записи о расписании для удаления", "Ошибка");
                return;
            }

            StringBuilder deletedSchedules = new StringBuilder();

            foreach (DataGridViewRow selectedRow in schedule_dataGridView.SelectedRows)
            {
                int scheduleID = Convert.ToInt32(selectedRow.Cells["ScheduleID"].Value);
                DateTime lessonDate = Convert.ToDateTime(selectedRow.Cells["LessonDate"].Value);
                TimeSpan startTime = TimeSpan.Parse(selectedRow.Cells["StartTime"].Value.ToString());
                TimeSpan endTime = TimeSpan.Parse(selectedRow.Cells["EndTime"].Value.ToString());
                string studentName = selectedRow.Cells["StudentName"].Value.ToString();
                int mark = Convert.ToInt32(selectedRow.Cells["Mark"].Value);

                string message = $"Вы уверены, что хотите удалить запись о расписании для курсанта {studentName} на {lessonDate.ToShortDateString()} c {startTime} по {endTime} c отметкой {mark}?";

                DialogResult dialogResult = MessageBox.Show(message, "Подтверждение удаления", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    SqlCommand deleteScheduleCommand = new SqlCommand(
                        "DELETE FROM [Schedule] WHERE ScheduleID = @ScheduleID",
                        sqlConnection);
                    deleteScheduleCommand.Parameters.AddWithValue("ScheduleID", scheduleID);

                    int rowsAffected = deleteScheduleCommand.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        deletedSchedules.AppendLine($"Расписание для курсанта {studentName} на {lessonDate.ToShortDateString()} c {startTime} по {endTime} c отметкой {mark}");
                    }
                }
            }

            string deletedSchedulesMessage = "Следующие записи о расписании были успешно удалены из базы данных:\n\n" + deletedSchedules.ToString();
            if (!string.IsNullOrWhiteSpace(deletedSchedules.ToString()))
            {
                MessageBox.Show(deletedSchedulesMessage, "Успех");
                RefreshdbScheduleFromSelectDate();
            }
            else
            {
                MessageBox.Show("Не удалось удалить выбранные записи о расписании из базы данных.", "Ошибка");
            }
        }

        private void change_students_button_Click(object sender, EventArgs e)
        {
            if (students_dataGridView.SelectedRows.Count > 0)
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

                int studentID = Convert.ToInt32(students_dataGridView.SelectedRows[0].Cells["StudentID"].Value);

                SqlCommand updateStudentCommand = new SqlCommand(
                    "UPDATE [Students] SET FirstName = @FirstName, LastName = @LastName, BirthDate = @BirthDate, Phone = @Phone WHERE StudentID = @StudentID",
                    sqlConnection);

                updateStudentCommand.Parameters.AddWithValue("@FirstName", firstname_students_textBox.Text);
                updateStudentCommand.Parameters.AddWithValue("@LastName", lastname_students_textBox.Text);
                updateStudentCommand.Parameters.AddWithValue("@BirthDate", birthDate.ToString("yyyy-MM-dd"));
                updateStudentCommand.Parameters.AddWithValue("@Phone", phone_students_textBox.Text);
                updateStudentCommand.Parameters.AddWithValue("@StudentID", studentID);

                int rowsAffected = updateStudentCommand.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show("Изменения успешно сохранены в базе данных", "Успех");
                    Refreshdbstudents();
                    RefreshStudentComboBox();
                    UpdateReportStudents();
                }
                else
                {
                    MessageBox.Show("Не удалось сохранить изменения в базе данных", "Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите запись, которую хотите изменить", "Ошибка");
            }
        }

        private void change_locations_button_Click(object sender, EventArgs e)
        {
            if (locations_dataGridView.SelectedRows.Count > 0)
            {
                if (string.IsNullOrWhiteSpace(name_locations_textBox.Text) ||
                    string.IsNullOrWhiteSpace(address_locations_textBox.Text) ||
                    string.IsNullOrWhiteSpace(description_locations_textBox.Text))
                {
                    MessageBox.Show("Введите данные во все поля", "Ошибка");
                    return;
                }

                int locationID = Convert.ToInt32(locations_dataGridView.SelectedRows[0].Cells["LocationID"].Value);

                SqlCommand updateLocationCommand = new SqlCommand(
                    "UPDATE [Locations] SET Name = @Name, Address = @Address, Description = @Description WHERE LocationID = @LocationID",
                    sqlConnection);

                updateLocationCommand.Parameters.AddWithValue("@Name", name_locations_textBox.Text);
                updateLocationCommand.Parameters.AddWithValue("@Address", address_locations_textBox.Text);
                updateLocationCommand.Parameters.AddWithValue("@Description", description_locations_textBox.Text);
                updateLocationCommand.Parameters.AddWithValue("@LocationID", locationID);

                int rowsAffected = updateLocationCommand.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show("Изменения успешно сохранены в базе данных", "Успех");
                    Refreshdblocations();
                    RefreshLocationComboBox();
                }
                else
                {
                    MessageBox.Show("Не удалось сохранить изменения в базе данных", "Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите запись, которую хотите изменить", "Ошибка");
            }
        }

        private void change_lessontypes_button_Click(object sender, EventArgs e)
        {
            if (lessontypes_dataGridView.SelectedRows.Count > 0)
            {
                if (string.IsNullOrWhiteSpace(name_lessontypes_textBox.Text) ||
                    string.IsNullOrWhiteSpace(description_lessontypes_textBox.Text) ||
                    string.IsNullOrWhiteSpace(instructornotes_lessontypes_textBox.Text))
                {
                    MessageBox.Show("Введите данные во все поля", "Ошибка");
                    return;
                }

                int lessonTypeID = Convert.ToInt32(lessontypes_dataGridView.SelectedRows[0].Cells["LessonTypeID"].Value);

                SqlCommand updateLessonTypeCommand = new SqlCommand(
                    "UPDATE [LessonTypes] SET Name = @Name, Description = @Description, InstructorNotes = @InstructorNotes WHERE LessonTypeID = @LessonTypeID",
                    sqlConnection);

                updateLessonTypeCommand.Parameters.AddWithValue("@Name", name_lessontypes_textBox.Text);
                updateLessonTypeCommand.Parameters.AddWithValue("@Description", description_lessontypes_textBox.Text);
                updateLessonTypeCommand.Parameters.AddWithValue("@InstructorNotes", instructornotes_lessontypes_textBox.Text);
                updateLessonTypeCommand.Parameters.AddWithValue("@LessonTypeID", lessonTypeID);

                int rowsAffected = updateLessonTypeCommand.ExecuteNonQuery();

                if (rowsAffected > 0)
                {
                    MessageBox.Show("Изменения успешно сохранены в базе данных", "Успех");
                    RefreshLessonTypeComboBox();
                    Refreshdblessontypes();
                }
                else
                {
                    MessageBox.Show("Не удалось сохранить изменения в базе данных", "Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите запись, которую хотите изменить", "Ошибка");
            }
        }

        private void change_schedule_button_Click(object sender, EventArgs e)
        {
            if (schedule_dataGridView.SelectedRows.Count > 0)
            {
                if (lessondate_schedule_dateTimePicker.Value == null ||
                    string.IsNullOrWhiteSpace(starttime_schedule_comboBox.Text) ||
                    string.IsNullOrWhiteSpace(endtime_schedule_comboBox.Text) ||
                    studentID_schedule_comboBox.SelectedItem == null ||
                    locationID_schedule_comboBox.SelectedItem == null ||
                    lessontypeID_schedule_comboBox.SelectedItem == null)
                {
                    MessageBox.Show("Выберите данные из выпадающих списков и заполните все поля", "Ошибка");
                    return;
                }

                int scheduleID = Convert.ToInt32(schedule_dataGridView.SelectedRows[0].Cells["ScheduleID"].Value);
                DateTime lessonDate = lessondate_schedule_dateTimePicker.Value.Date;
                TimeSpan startTime;
                TimeSpan endTime;
                if (TimeSpan.TryParse(starttime_schedule_comboBox.Text, out startTime) &&
                    TimeSpan.TryParse(endtime_schedule_comboBox.Text, out endTime))
                {
                    int studentID = ((Student)studentID_schedule_comboBox.SelectedItem).StudentID;
                    int locationID = ((Location)locationID_schedule_comboBox.SelectedItem).LocationID;
                    int lessontypeID = ((LessonType)lessontypeID_schedule_comboBox.SelectedItem).LessonTypeID;

                    int mark = -1;
                    if (!string.IsNullOrWhiteSpace(mark_schedule_textBox.Text))
                    {
                        string input = mark_schedule_textBox.Text.Trim().ToLower();

                        if (int.TryParse(input, out int parsedMark))
                        {
                            if (parsedMark < 0 || parsedMark > 10)
                            {
                                MessageBox.Show("Введите корректное значение для оценки от 0 до 10", "Ошибка");
                                return;
                            }
                            mark = parsedMark;
                        }
                        else if (input == "сдан")
                        {
                            mark = 10;
                        }
                        else if (input == "не сдан")
                        {
                            mark = 0;
                        }
                        else
                        {
                            MessageBox.Show("Введите корректное значение для оценки (число от 0 до 10, \"Сдан\" или \"Не сдан\")", "Ошибка");
                            return;
                        }
                    }

                    if (IsLessonOverlap_change(lessonDate, startTime, endTime, scheduleID))
                    {
                        MessageBox.Show($"Невозможно сохранить изменения. Промежуток времени {lessonDate.ToShortDateString()} с {startTime} по {endTime} занят.", "Ошибка");
                        return;
                    }

                    SqlCommand updateScheduleCommand = new SqlCommand(
                        "UPDATE [Schedule] SET LessonDate = @LessonDate, StartTime = @StartTime, EndTime = @EndTime, StudentID = @StudentID, LocationID = @LocationID, LessonTypeID = @LessonTypeID, Mark = @Mark WHERE ScheduleID = @ScheduleID",
                        sqlConnection);

                    updateScheduleCommand.Parameters.AddWithValue("@LessonDate", lessonDate);
                    updateScheduleCommand.Parameters.AddWithValue("@StartTime", startTime);
                    updateScheduleCommand.Parameters.AddWithValue("@EndTime", endTime);
                    updateScheduleCommand.Parameters.AddWithValue("@StudentID", studentID);
                    updateScheduleCommand.Parameters.AddWithValue("@LocationID", locationID);
                    updateScheduleCommand.Parameters.AddWithValue("@LessonTypeID", lessontypeID);
                    updateScheduleCommand.Parameters.AddWithValue("@Mark", mark);
                    updateScheduleCommand.Parameters.AddWithValue("@ScheduleID", scheduleID);

                    int rowsAffected = updateScheduleCommand.ExecuteNonQuery();

                    if (rowsAffected > 0)
                    {
                        MessageBox.Show("Изменения успешно сохранены в базе данных", "Успех");
                        RefreshdbScheduleFromSelectDate();
                    }
                    else
                    {
                        MessageBox.Show("Не удалось сохранить изменения в базе данных", "Ошибка");
                    }
                }
                else
                {
                    MessageBox.Show("Введите корректное время в формате ЧЧ:ММ", "Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Выберите запись, которую хотите изменить", "Ошибка");
            }
        }

        private bool IsLessonOverlap_change(DateTime lessonDate, TimeSpan startTime, TimeSpan endTime, int currentScheduleID)
        {
            SqlCommand command = new SqlCommand(
                @"SELECT COUNT(*) 
                FROM Schedule 
                WHERE LessonDate = @LessonDate 
                AND (
                ((@StartTime >= StartTime AND @StartTime < EndTime) OR (@EndTime > StartTime AND @EndTime <= EndTime))
                OR (StartTime >= @StartTime AND EndTime <= @EndTime)
                )
                AND ScheduleID != @CurrentScheduleID", sqlConnection);

            command.Parameters.AddWithValue("@LessonDate", lessonDate);
            command.Parameters.AddWithValue("@StartTime", startTime);
            command.Parameters.AddWithValue("@EndTime", endTime);
            command.Parameters.AddWithValue("@CurrentScheduleID", currentScheduleID);

            int count = (int)command.ExecuteScalar();
            return count > 0;
        }

        Dictionary<string, string> StudentscolumnTranslations = new Dictionary<string, string>
{
            { "ID студента", "StudentID" },
            { "Имя", "FirstName" },
            { "Фамилия", "LastName" },
            { "День рождения", "BirthDate" },
            { "Телефон", "Phone" }
           };

        private void search_students_textBox_TextChanged(object sender, EventArgs e)
        {
            Searchstudents();
        }

        private void Searchstudents()
        {
            string selectedColumn = changecolums_students_comboBox.SelectedItem?.ToString();
            string searchTerm = search_students_textBox.Text.Trim();

            if (!string.IsNullOrEmpty(selectedColumn) && selectedColumn != "Все столбцы")
            {
                string englishColumnName = StudentscolumnTranslations[selectedColumn];

                string searchQuery = $"SELECT * FROM [Students] WHERE [{englishColumnName}] LIKE @SearchTerm";

                if (students_dataGridView.Columns[englishColumnName].ValueType == typeof(DateTime))
                {
                    searchQuery = $"SELECT * FROM [Students] WHERE CONVERT(varchar, [{englishColumnName}], 104) LIKE @SearchTerm";
                }

                SqlCommand searchStudentsCommand = new SqlCommand(searchQuery, sqlConnection);
                searchStudentsCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

                SqlDataAdapter adapter = new SqlDataAdapter(searchStudentsCommand);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                students_dataGridView.DataSource = dt;
            }
            else
            {
                string searchQuery = "SELECT * FROM [Students] WHERE ";
                foreach (DataGridViewColumn column in students_dataGridView.Columns)
                {
                    string englishColumnName = StudentscolumnTranslations[column.HeaderText];

                    if (column.ValueType == typeof(DateTime))
                    {
                        searchQuery += $"CONVERT(varchar, [{englishColumnName}], 104) LIKE @SearchTerm OR ";
                    }
                    else
                    {
                        searchQuery += $"[{englishColumnName}] LIKE @SearchTerm OR ";
                    }
                }
                searchQuery = searchQuery.TrimEnd("OR ".ToCharArray());

                SqlCommand searchStudentsCommand = new SqlCommand(searchQuery, sqlConnection);
                searchStudentsCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

                SqlDataAdapter adapter = new SqlDataAdapter(searchStudentsCommand);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                students_dataGridView.DataSource = dt;

                HighlightSearchResults(students_dataGridView, searchTerm);
            }
        }

        private void changecolums_students_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            Searchstudents();
        }

        Dictionary<string, string> LocationsColumnTranslations = new Dictionary<string, string>
        {
            { "ID локации", "LocationID" },
            { "Название", "Name" },
            { "Адрес", "Address" },
            { "Дополнительная информация", "Description" },
        };


        private void SearchLocations()
        {
            string selectedColumn = changecolums_locations_comboBox.SelectedItem?.ToString();
            string searchTerm = search_locations_textBox.Text.Trim();

            if (!string.IsNullOrEmpty(selectedColumn) && selectedColumn != "Все столбцы")
            {
                string englishColumnName = LocationsColumnTranslations[selectedColumn];

                string searchQuery = $"SELECT * FROM [Locations] WHERE [{englishColumnName}] LIKE @SearchTerm";

                if (locations_dataGridView.Columns[englishColumnName].ValueType == typeof(DateTime))
                {
                    searchQuery = $"SELECT * FROM [Locations] WHERE CONVERT(varchar, [{englishColumnName}], 104) LIKE @SearchTerm";
                }

                SqlCommand searchLocationsCommand = new SqlCommand(searchQuery, sqlConnection);
                searchLocationsCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

                SqlDataAdapter adapter = new SqlDataAdapter(searchLocationsCommand);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                locations_dataGridView.DataSource = dt;
            }
            else
            {
                string searchQuery = "SELECT * FROM [Locations] WHERE ";
                foreach (DataGridViewColumn column in locations_dataGridView.Columns)
                {
                    string englishColumnName = LocationsColumnTranslations[column.HeaderText];

                    if (column.ValueType == typeof(DateTime))
                    {
                        searchQuery += $"CONVERT(varchar, [{englishColumnName}], 104) LIKE @SearchTerm OR ";
                    }
                    else
                    {
                        searchQuery += $"[{englishColumnName}] LIKE @SearchTerm OR ";
                    }
                }
                searchQuery = searchQuery.TrimEnd("OR ".ToCharArray());

                SqlCommand searchLocationsCommand = new SqlCommand(searchQuery, sqlConnection);
                searchLocationsCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

                SqlDataAdapter adapter = new SqlDataAdapter(searchLocationsCommand);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                locations_dataGridView.DataSource = dt;

                HighlightSearchResults(locations_dataGridView, searchTerm);
            }
        }

        private void search_locations_textBox_TextChanged(object sender, EventArgs e)
        {
            SearchLocations();
        }

        private void changecolums_locations_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SearchLocations();
        }

        Dictionary<string, string> LessonTypesColumnTranslations = new Dictionary<string, string>
        {
        { "ID типа занятия", "LessonTypeID" },
        { "Именование", "Name" },
        { "Описание", "Description" },
        { "Заметка инструктора", "InstructorNotes" }
        };


        private void SearchLessonTypes()
        {
            string selectedColumn = changecolums_lessontypes_comboBox.SelectedItem?.ToString();
            string searchTerm = search_lessontypes_textBox.Text.Trim();

            if (!string.IsNullOrEmpty(selectedColumn) && selectedColumn != "Все столбцы")
            {
                string englishColumnName = LessonTypesColumnTranslations[selectedColumn];

                string searchQuery = $"SELECT * FROM [LessonTypes] WHERE [{englishColumnName}] LIKE @SearchTerm";

                if (lessontypes_dataGridView.Columns[englishColumnName].ValueType == typeof(DateTime))
                {
                    searchQuery = $"SELECT * FROM [LessonTypes] WHERE CONVERT(varchar, [{englishColumnName}], 104) LIKE @SearchTerm";
                }

                SqlCommand searchLessonTypesCommand = new SqlCommand(searchQuery, sqlConnection);
                searchLessonTypesCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

                SqlDataAdapter adapter = new SqlDataAdapter(searchLessonTypesCommand);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                lessontypes_dataGridView.DataSource = dt;
            }
            else
            {
                string searchQuery = "SELECT * FROM [LessonTypes] WHERE ";
                foreach (DataGridViewColumn column in lessontypes_dataGridView.Columns)
                {
                    string englishColumnName = LessonTypesColumnTranslations[column.HeaderText];

                    if (column.ValueType == typeof(DateTime))
                    {
                        searchQuery += $"CONVERT(varchar, [{englishColumnName}], 104) LIKE @SearchTerm OR ";
                    }
                    else
                    {
                        searchQuery += $"[{englishColumnName}] LIKE @SearchTerm OR ";
                    }
                }
                searchQuery = searchQuery.TrimEnd("OR ".ToCharArray());

                SqlCommand searchLessonTypesCommand = new SqlCommand(searchQuery, sqlConnection);
                searchLessonTypesCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

                SqlDataAdapter adapter = new SqlDataAdapter(searchLessonTypesCommand);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                lessontypes_dataGridView.DataSource = dt;

                HighlightSearchResults(lessontypes_dataGridView, searchTerm);

            }
        }

        private void HighlightSearchResults(DataGridView dataGridView, string searchTerm)
        {
            if (string.IsNullOrEmpty(searchTerm))
            {
                foreach (DataGridViewRow row in dataGridView.Rows)
                {
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = dataGridView.DefaultCellStyle.BackColor;
                    }
                }
                return;
            }

            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value.ToString().Contains(searchTerm))
                    {
                        cell.Style.BackColor = Color.Yellow; 
                    }
                    else
                    {
                        cell.Style.BackColor = dataGridView.DefaultCellStyle.BackColor; 
                    }
                }
            }
        }


        private void changecolums_lessontypes_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SearchLessonTypes();
        }

        private void search_lessontypes_textBox_TextChanged(object sender, EventArgs e)
        {
            SearchLessonTypes();
        }

        Dictionary<string, string> ScheduleColumnTranslations = new Dictionary<string, string>
        {
            { "ID занятия", "ScheduleID" },
            { "Дата занятия", "LessonDate" },
            { "Время начала", "StartTime" },
            { "Время конца", "EndTime" },
            { "Курсант", "StudentID" },
            { "Локация", "LocationID" },
            { "Тип занятия", "LessonTypeID" },
            { "Оценка", "Mark" }
        };

        private void SearchSchedule()
        {
            string selectedColumn = changecolums_schedule_comboBox.SelectedItem?.ToString();
            string searchTerm = search_schedule_textBox.Text.Trim();

            string searchQuery = "SELECT s.ScheduleID, s.LessonDate, s.StartTime, s.EndTime, st.FirstName + ' ' + st.LastName AS StudentName, " +
                                 "l.Name AS LocationName, lt.Name AS LessonTypeName, s.Mark " +
                                 "FROM Schedule s " +
                                 "JOIN Students st ON s.StudentID = st.StudentID " +
                                 "JOIN Locations l ON s.LocationID = l.LocationID " +
                                 "JOIN LessonTypes lt ON s.LessonTypeID = lt.LessonTypeID ";

            if (!string.IsNullOrEmpty(selectedColumn) && selectedColumn != "Все столбцы")
            {
                string englishColumnName = ScheduleColumnTranslations[selectedColumn];

                if (englishColumnName == "StudentID")
                {
                    searchQuery += $"WHERE st.FirstName + ' ' + st.LastName LIKE @SearchTerm";
                }
                else if (englishColumnName == "LocationID")
                {
                    searchQuery += $"WHERE l.Name LIKE @SearchTerm";
                }
                else if (englishColumnName == "LessonTypeID")
                {
                    searchQuery += $"WHERE lt.Name LIKE @SearchTerm";
                }
                else if (englishColumnName == "LessonDate")
                {
                    // Предположим, что формат даты в вашей базе данных - "dd.MM.yyyy"
                    searchQuery += $"WHERE CONVERT(varchar, s.LessonDate, 104) LIKE @SearchTerm";
                }
                else
                {
                    searchQuery += $"WHERE s.[{englishColumnName}] LIKE @SearchTerm";
                }
            }
            else
            {
                searchQuery += "WHERE s.ScheduleID LIKE @SearchTerm OR " +
                               "CONVERT(varchar, s.LessonDate, 104) LIKE @SearchTerm OR " +
                               "s.StartTime LIKE @SearchTerm OR " +
                               "s.EndTime LIKE @SearchTerm OR " +
                               "st.FirstName + ' ' + st.LastName LIKE @SearchTerm OR " +
                               "l.Name LIKE @SearchTerm OR " +
                               "lt.Name LIKE @SearchTerm OR " +
                               "s.Mark LIKE @SearchTerm";
            }

            SqlCommand searchScheduleCommand = new SqlCommand(searchQuery, sqlConnection);
            searchScheduleCommand.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

            SqlDataAdapter adapter = new SqlDataAdapter(searchScheduleCommand);
            DataTable dt = new DataTable();
            adapter.Fill(dt);

            schedule_dataGridView.DataSource = dt;

            if (!string.IsNullOrEmpty(selectedColumn) && selectedColumn == "Все столбцы")
            {
                HighlightSearchResults(schedule_dataGridView, searchTerm);
            }
        }


        private void search_schedule_textBox_TextChanged(object sender, EventArgs e)
        {
            SearchSchedule();
        }

        private void changecolums_schedule_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            SearchSchedule();
        }

        private void allschedule_checkbox_CheckedChanged(object sender, EventArgs e)
        {
            RefreshdbScheduleFromSelectDate();
        }

        private void students_reportViewer_Load(object sender, EventArgs e)
        {
            UpdateReportStudents();
        }

        private void UpdateReportStudents()
        {
            string selectedColumn = changecolums_students_report_comboBox.SelectedItem?.ToString();
            string searchTerm = search_students_report_textBox.Text.Trim();

            string searchQuery = $"SELECT * FROM Students";

            if (!string.IsNullOrEmpty(selectedColumn) && selectedColumn != "Все столбцы")
            {
                string englishColumnName = StudentscolumnTranslations[selectedColumn];
                searchQuery += $" WHERE [{englishColumnName}] LIKE @SearchTerm";

                if (students_dataGridView.Columns[englishColumnName].ValueType == typeof(DateTime))
                {
                    searchQuery = $"SELECT * FROM Students WHERE CONVERT(varchar, [{englishColumnName}], 104) LIKE @SearchTerm";
                }
            }
            else if (!string.IsNullOrEmpty(searchTerm))
            {
                searchQuery += " WHERE ";
                foreach (string columnName in StudentscolumnTranslations.Values)
                {
                    if (students_dataGridView.Columns[columnName].ValueType == typeof(DateTime))
                    {
                        searchQuery += $"CONVERT(varchar, [{columnName}], 104) LIKE @SearchTerm OR ";
                    }
                    else
                    {
                        searchQuery += $"[{columnName}] LIKE @SearchTerm OR ";
                    }
                }
                searchQuery = searchQuery.TrimEnd(" OR ".ToCharArray());
            }

            SqlCommand command = new SqlCommand(searchQuery, sqlConnection);
            command.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

            SqlDataAdapter d = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            d.Fill(dt);

            students_reportViewer.LocalReport.DataSources.Clear();
            ReportDataSource sourse = new ReportDataSource("DataSet_students", dt);
            students_reportViewer.LocalReport.ReportPath = "Report_students.rdlc";
            students_reportViewer.LocalReport.DataSources.Add(sourse);
            students_reportViewer.RefreshReport();
        }



        private void UpdateReportSchedule()
        {
            string selectedColumn = changecolums_schedule_comboBox.SelectedItem?.ToString();
            string searchTerm = search_schedule_report_textBox.Text.Trim();

            string searchQuery = "SELECT s.ScheduleID, s.LessonDate, s.StartTime, s.EndTime, " +
                                 "st.FirstName + ' ' + st.LastName AS StudentName, " +
                                 "st.StudentID, " +
                                 "l.Name AS LocationName, " +
                                 "l.LocationID, " +
                                 "lt.Name AS LessonTypeName, " +
                                 "lt.LessonTypeID, " +
                                 "s.Mark " +
                                 "FROM Schedule s " +
                                 "JOIN Students st ON s.StudentID = st.StudentID " +
                                 "JOIN Locations l ON s.LocationID = l.LocationID " +
                                 "JOIN LessonTypes lt ON s.LessonTypeID = lt.LessonTypeID ";

            if (!string.IsNullOrEmpty(selectedColumn) && selectedColumn != "Все столбцы")
            {
                string englishColumnName = ScheduleColumnTranslations[selectedColumn];

                if (englishColumnName == "StudentID")
                {
                    searchQuery += $"WHERE st.FirstName + ' ' + st.LastName LIKE @SearchTerm";
                }
                else if (englishColumnName == "LocationID")
                {
                    searchQuery += $"WHERE l.Name LIKE @SearchTerm";
                }
                else if (englishColumnName == "LessonTypeID")
                {
                    searchQuery += $"WHERE lt.Name LIKE @SearchTerm";
                }
                else if (englishColumnName == "LessonDate")
                {
                    // Предположим, что формат даты в вашей базе данных - "dd.MM.yyyy"
                    searchQuery += $"WHERE CONVERT(varchar, s.LessonDate, 104) LIKE @SearchTerm";
                }
                else
                {
                    searchQuery += $"WHERE s.[{englishColumnName}] LIKE @SearchTerm";
                }
            }
            else
            {
                searchQuery += "WHERE s.ScheduleID LIKE @SearchTerm OR " +
                               "CONVERT(varchar, s.LessonDate, 104) LIKE @SearchTerm OR " +
                               "s.StartTime LIKE @SearchTerm OR " +
                               "s.EndTime LIKE @SearchTerm OR " +
                               "st.FirstName + ' ' + st.LastName LIKE @SearchTerm OR " +
                               "l.Name LIKE @SearchTerm OR " +
                               "lt.Name LIKE @SearchTerm OR " +
                               "s.Mark LIKE @SearchTerm";
            }

            SqlCommand command = new SqlCommand(searchQuery, sqlConnection);
            command.Parameters.AddWithValue("@SearchTerm", $"%{searchTerm}%");

            SqlDataAdapter d = new SqlDataAdapter(command);
            DataTable dt = new DataTable();
            d.Fill(dt);

            schedule_reportViewer.LocalReport.DataSources.Clear();
            ReportDataSource sourse = new ReportDataSource("DataSet_schedule", dt);
            schedule_reportViewer.LocalReport.ReportPath = "Report_schedule.rdlc";
            schedule_reportViewer.LocalReport.DataSources.Add(sourse);
            schedule_reportViewer.RefreshReport();
        }


        private void search_students_report_textBox_TextChanged(object sender, EventArgs e)
        {
            UpdateReportStudents();
        }

        private void changecolums_students_report_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateReportStudents();
        }

        private void schedule_reportViewer_Load(object sender, EventArgs e)
        {
            UpdateReportSchedule();
        }

        private void changecolums_schedule_report_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateReportSchedule();
        }

        private void search_schedule_report_textBox_TextChanged(object sender, EventArgs e)
        {
            UpdateReportSchedule();
        }
    }
}

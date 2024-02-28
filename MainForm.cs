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

            Refreshdbstudents();

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
        
    }
}

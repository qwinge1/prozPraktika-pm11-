using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;

namespace SchoolJournalApp
{
    public partial class AddGradeWindow : Window
    {
        private string connectionString;
        private int studentId;
        private DataTable subjectsTable;

        public AddGradeWindow(string connString, int studentId)
        {
            InitializeComponent();
            this.connectionString = connString;
            this.studentId = studentId;
            dpDate.SelectedDate = DateTime.Today;
            LoadSubjects();
        }

        private void LoadSubjects()
        {
            string query = "SELECT ID_предмета, Название FROM Предметы ORDER BY Название";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                subjectsTable = new DataTable();
                da.Fill(subjectsTable);
                cmbSubject.ItemsSource = subjectsTable.DefaultView;
                cmbSubject.DisplayMemberPath = "Название";
                cmbSubject.SelectedValuePath = "ID_предмета";
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            if (cmbSubject.SelectedValue == null || cmbGrade.SelectedItem == null)
            {
                MessageBox.Show("Выберите предмет и оценку.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            int subjectId = Convert.ToInt32(cmbSubject.SelectedValue);
            int gradeValue = Convert.ToInt32(((ComboBoxItem)cmbGrade.SelectedItem).Content);
            string controlType = ((ComboBoxItem)cmbControlType.SelectedItem)?.Content.ToString() ?? "текущая";
            DateTime date = dpDate.SelectedDate ?? DateTime.Today;

            string insertQuery = @"
                INSERT INTO Оценки (ID_ученика, ID_предмета, Дата_получения, Оценка, Тип_контроля)
                VALUES (@sid, @subj, @date, @grade, @type)";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(insertQuery, conn);
                cmd.Parameters.AddWithValue("@sid", studentId);
                cmd.Parameters.AddWithValue("@subj", subjectId);
                cmd.Parameters.AddWithValue("@date", date);
                cmd.Parameters.AddWithValue("@grade", gradeValue);
                cmd.Parameters.AddWithValue("@type", controlType);
                conn.Open();
                cmd.ExecuteNonQuery();
            }
            DialogResult = true;
            Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
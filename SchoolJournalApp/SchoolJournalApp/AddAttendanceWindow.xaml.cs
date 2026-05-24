using System;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;

namespace SchoolJournalApp
{
    public partial class AddAttendanceWindow : Window
    {
        private string connectionString;
        private int studentId;

        public AddAttendanceWindow(string connString, int studentId)
        {
            InitializeComponent();
            this.connectionString = connString;
            this.studentId = studentId;
            dpDate.SelectedDate = DateTime.Today;
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            DateTime date = dpDate.SelectedDate ?? DateTime.Today;
            string status = ((ComboBoxItem)cmbStatus.SelectedItem).Content.ToString();
            string note = txtNote.Text.Trim();

            string insertQuery = @"
                INSERT INTO Посещаемость (ID_ученика, Дата, Статус, Примечание)
                VALUES (@sid, @date, @status, @note)";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(insertQuery, conn);
                cmd.Parameters.AddWithValue("@sid", studentId);
                cmd.Parameters.AddWithValue("@date", date);
                cmd.Parameters.AddWithValue("@status", status);
                cmd.Parameters.AddWithValue("@note", string.IsNullOrEmpty(note) ? (object)DBNull.Value : note);
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
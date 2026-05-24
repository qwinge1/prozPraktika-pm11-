using System;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;

namespace SchoolJournalApp
{
    public partial class EditAttendanceWindow : Window
    {
        private string connectionString;
        private int attendanceId;
        private int studentId;

        public EditAttendanceWindow(string connString, int attendanceId, int studentId)
        {
            InitializeComponent();
            this.connectionString = connString;
            this.attendanceId = attendanceId;
            this.studentId = studentId;
            LoadAttendanceData();
        }

        private void LoadAttendanceData()
        {
            string query = "SELECT Дата, Статус, Примечание FROM Посещаемость WHERE ID_посещаемости = @id";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", attendanceId);
                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    dpDate.SelectedDate = reader.GetDateTime(0);
                    string status = reader.GetString(1);
                    foreach (ComboBoxItem item in cmbStatus.Items)
                    {
                        if (item.Content.ToString() == status)
                        {
                            cmbStatus.SelectedItem = item;
                            break;
                        }
                    }
                    txtNote.Text = reader.IsDBNull(2) ? "" : reader.GetString(2);
                }
                reader.Close();
            }
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            DateTime date = dpDate.SelectedDate ?? DateTime.Today;
            string status = ((ComboBoxItem)cmbStatus.SelectedItem).Content.ToString();
            string note = txtNote.Text.Trim();

            string updateQuery = @"
                UPDATE Посещаемость
                SET Дата = @date, Статус = @status, Примечание = @note
                WHERE ID_посещаемости = @id";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(updateQuery, conn);
                cmd.Parameters.AddWithValue("@date", date);
                cmd.Parameters.AddWithValue("@status", status);
                cmd.Parameters.AddWithValue("@note", string.IsNullOrEmpty(note) ? (object)DBNull.Value : note);
                cmd.Parameters.AddWithValue("@id", attendanceId);
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
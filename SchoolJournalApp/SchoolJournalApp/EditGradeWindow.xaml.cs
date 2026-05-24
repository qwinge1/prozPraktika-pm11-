using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows;
using System.Windows.Controls;

namespace SchoolJournalApp
{
    public partial class EditGradeWindow : Window
    {
        private string connectionString;
        private int gradeId;
        private int studentId;
        private DataTable subjectsTable;

        public EditGradeWindow(string connString, int gradeId, int studentId)
        {
            InitializeComponent();
            this.connectionString = connString;
            this.gradeId = gradeId;
            this.studentId = studentId;
            LoadSubjects();
            LoadGradeData();
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

        private void LoadGradeData()
        {
            string query = "SELECT ID_предмета, Оценка, Тип_контроля, Дата_получения FROM Оценки WHERE ID_оценки = @id";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", gradeId);
                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    // Предмет
                    cmbSubject.SelectedValue = reader.GetInt32(0);

                    // Оценка
                    int gradeValue = reader.GetInt32(1);
                    foreach (ComboBoxItem item in cmbGrade.Items)
                    {
                        if (Convert.ToInt32(item.Content) == gradeValue)
                        {
                            cmbGrade.SelectedItem = item;
                            break;
                        }
                    }

                    // Тип контроля (может быть NULL)
                    string controlType = reader.IsDBNull(2) ? "текущая" : reader.GetString(2);
                    foreach (ComboBoxItem item in cmbControlType.Items)
                    {
                        if (item.Content.ToString() == controlType)
                        {
                            cmbControlType.SelectedItem = item;
                            break;
                        }
                    }

                    // Дата (вряд ли NULL, но на всякий случай)
                    DateTime date = reader.IsDBNull(3) ? DateTime.Today : reader.GetDateTime(3);
                    dpDate.SelectedDate = date;
                }
                reader.Close();
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

            string updateQuery = @"
                UPDATE Оценки
                SET ID_предмета = @subj, Оценка = @grade, Тип_контроля = @type, Дата_получения = @date
                WHERE ID_оценки = @id";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(updateQuery, conn);
                cmd.Parameters.AddWithValue("@subj", subjectId);
                cmd.Parameters.AddWithValue("@grade", gradeValue);
                cmd.Parameters.AddWithValue("@type", controlType);
                cmd.Parameters.AddWithValue("@date", date);
                cmd.Parameters.AddWithValue("@id", gradeId);
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
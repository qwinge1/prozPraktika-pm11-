using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace SchoolJournalApp
{
    public partial class MainWindow : Window
    {
        private string connectionString = "Server=DESKTOP-7CTIM0M\\SQLEXPRESS;Database=Школьный_Журнал;Trusted_Connection=True;";

        private bool isEditMode = false;
        private int editStudentId = 0;

        // Переменные для хранения текущих фильтров
        private int? currentClassFilter = null;
        private string currentNameFilter = "";
        private int? currentGradeSubjectFilter = null;
        private string currentAttendanceStatusFilter = "Все";
        private int? currentTopClassFilter = null;
        private string currentYearFilter = "2025-2026";
        private int? currentSubjectClassFilter = null;
        private DateTime? currentAttendanceDateFrom = null;
        private DateTime? currentAttendanceDateTo = null;

        public MainWindow()
        {
            InitializeComponent();
            this.WindowState = WindowState.Maximized;
            LoadFilterClasses();
            LoadStudents();
            LoadClasses();
            LoadStudentsForAttendance();
            LoadClassesForAttendanceReport();
            LoadSubjectsForFilter();
            LoadTopClassFilter();
            LoadSubjectClassFilter();
        }

        // ---- Загрузка вспомогательных списков для фильтров ----
        private void LoadFilterClasses()
        {
            string query = "SELECT ID_класса, CAST(Номер_класса AS VARCHAR) + ' ' + Буква_класса + ' (' + CAST(Учебный_год AS VARCHAR) + ')' AS Класс FROM Классы ORDER BY Учебный_год DESC, Номер_класса, Буква_класса";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                DataRow rowAll = dt.NewRow();
                rowAll["ID_класса"] = DBNull.Value;
                rowAll["Класс"] = "Все классы";
                dt.Rows.InsertAt(rowAll, 0);
                cmbFilterClass.ItemsSource = dt.DefaultView;
                cmbFilterClass.SelectedIndex = 0;
            }
        }

        private void LoadAttendanceFilterClasses()
        {
            string query = "SELECT ID_класса, CAST(Номер_класса AS VARCHAR) + ' ' + Буква_класса AS Класс FROM Классы ORDER BY Номер_класса, Буква_класса";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                cmbClassForAttendanceReport.ItemsSource = dt.DefaultView;
            }
        }

        private void LoadTopClassFilter()
        {
            string query = "SELECT ID_класса, CAST(Номер_класса AS VARCHAR) + ' ' + Буква_класса + ' (' + CAST(Учебный_год AS VARCHAR) + ')' AS Класс FROM Классы ORDER BY Учебный_год DESC, Номер_класса, Буква_класса";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                DataRow rowAll = dt.NewRow();
                rowAll["ID_класса"] = DBNull.Value;
                rowAll["Класс"] = "Все классы";
                dt.Rows.InsertAt(rowAll, 0);
                cmbTopClassFilter.ItemsSource = dt.DefaultView;
                cmbTopClassFilter.SelectedIndex = 0;
            }
        }

        private void LoadSubjectClassFilter()
        {
            string query = "SELECT ID_класса, CAST(Номер_класса AS VARCHAR) + ' ' + Буква_класса AS Класс FROM Классы ORDER BY Номер_класса, Буква_класса";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                DataRow rowAll = dt.NewRow();
                rowAll["ID_класса"] = DBNull.Value;
                rowAll["Класс"] = "Все классы";
                dt.Rows.InsertAt(rowAll, 0);
                cmbSubjectClassFilter.ItemsSource = dt.DefaultView;
                cmbSubjectClassFilter.SelectedIndex = 0;
            }
        }

        private void LoadSubjectsForFilter()
        {
            string query = "SELECT ID_предмета, Название FROM Предметы ORDER BY Название";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                DataRow rowAll = dt.NewRow();
                rowAll["ID_предмета"] = DBNull.Value;
                rowAll["Название"] = "Все предметы";
                dt.Rows.InsertAt(rowAll, 0);
                cmbGradeSubjectFilter.ItemsSource = dt.DefaultView;
                cmbGradeSubjectFilter.SelectedIndex = 0;
            }
        }

        // ---- Загрузка учеников с фильтрацией ----
        private void LoadStudents()
        {
            try
            {
                string query = @"
                SELECT 
                    у.ID_ученика, у.ФИО, у.Дата_рождения, у.Логин,
                    CAST(к.Номер_класса AS VARCHAR) + ' ' + к.Буква_класса AS Класс,
                    у.ID_класса
                FROM Учащиеся у
                JOIN Классы к ON у.ID_класса = к.ID_класса
                WHERE 1=1";

                if (currentClassFilter.HasValue)
                    query += " AND у.ID_класса = @classId";
                if (!string.IsNullOrWhiteSpace(currentNameFilter))
                    query += " AND у.ФИО LIKE @nameFilter";

                query += " ORDER BY Класс, у.ФИО";

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand(query, conn);
                    if (currentClassFilter.HasValue)
                        cmd.Parameters.AddWithValue("@classId", currentClassFilter.Value);
                    if (!string.IsNullOrWhiteSpace(currentNameFilter))
                        cmd.Parameters.AddWithValue("@nameFilter", "%" + currentNameFilter + "%");

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    da.Fill(dt);
                    lvStudents.ItemsSource = dt.DefaultView;
                    txtStudentCount.Text = dt.Rows.Count.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка загрузки учеников: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void FilterStudents_Click(object sender, RoutedEventArgs e)
        {
            if (cmbFilterClass.SelectedValue != null && cmbFilterClass.SelectedValue != DBNull.Value)
                currentClassFilter = Convert.ToInt32(cmbFilterClass.SelectedValue);
            else
                currentClassFilter = null;

            currentNameFilter = txtFilterName.Text.Trim();
            LoadStudents();
        }

        private void ResetStudentsFilter_Click(object sender, RoutedEventArgs e)
        {
            cmbFilterClass.SelectedIndex = 0;
            txtFilterName.Text = "";
            currentClassFilter = null;
            currentNameFilter = "";
            LoadStudents();
        }

        // Загрузка классов для формы редактирования
        private void LoadClasses()
        {
            string query = "SELECT ID_класса, CAST(Номер_класса AS VARCHAR) + ' ' + Буква_класса + ' (' + CAST(Учебный_год AS VARCHAR) + ')' AS Класс FROM Классы ORDER BY Учебный_год DESC, Номер_класса, Буква_класса";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                cmbClass.ItemsSource = dt.DefaultView;
            }
        }

        private void LoadClassesForAttendanceReport()
        {
            string query = "SELECT ID_класса, CAST(Номер_класса AS VARCHAR) + ' ' + Буква_класса + ' (' + CAST(Учебный_год AS VARCHAR) + ')' AS Класс FROM Классы ORDER BY Учебный_год DESC, Номер_класса, Буква_класса";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                cmbClassForAttendanceReport.ItemsSource = dt.DefaultView;
            }
        }

        private void LoadStudentsForAttendance()
        {
            string query = "SELECT ID_ученика, ФИО FROM Учащиеся ORDER BY ФИО";
            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                cmbStudentsForAttendance.ItemsSource = dt.DefaultView;
            }
        }

        // ---- CRUD для учеников ----
        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            isEditMode = false;
            editStudentId = 0;
            ClearForm();
            borderForm.Visibility = Visibility.Visible;
        }

        private void btnEdit_Click(object sender, RoutedEventArgs e)
        {
            if (lvStudents.SelectedItem == null)
            {
                MessageBox.Show("Выберите ученика для редактирования.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            DataRowView row = (DataRowView)lvStudents.SelectedItem;
            editStudentId = Convert.ToInt32(row["ID_ученика"]);
            txtFIO.Text = row["ФИО"].ToString();
            txtBirth.Text = Convert.ToDateTime(row["Дата_рождения"]).ToString("yyyy-MM-dd");
            txtLogin.Text = row["Логин"].ToString();

            int classId = Convert.ToInt32(row["ID_класса"]);
            foreach (DataRowView item in cmbClass.Items)
            {
                if (Convert.ToInt32(item["ID_класса"]) == classId)
                {
                    cmbClass.SelectedItem = item;
                    break;
                }
            }

            isEditMode = true;
            borderForm.Visibility = Visibility.Visible;
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            if (lvStudents.SelectedItem == null)
            {
                MessageBox.Show("Выберите ученика для удаления.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (MessageBox.Show("Удалить выбранного ученика? Все его оценки также будут удалены.", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                DataRowView row = (DataRowView)lvStudents.SelectedItem;
                int studentId = Convert.ToInt32(row["ID_ученика"]);

                string deleteQuery = "DELETE FROM Учащиеся WHERE ID_ученика = @id";
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand(deleteQuery, conn);
                    cmd.Parameters.AddWithValue("@id", studentId);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                LoadStudents();
            }
        }

        private void btnRefresh_Click(object sender, RoutedEventArgs e)
        {
            LoadStudents();
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!DateTime.TryParseExact(txtBirth.Text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _))
            {
                MessageBox.Show("Дата должна быть в формате ГГГГ-ММ-ДД (например, 2010-05-15)", "Ошибка ввода", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string fio = txtFIO.Text.Trim();
            string birth = txtBirth.Text.Trim();
            int classId = Convert.ToInt32(cmbClass.SelectedValue);
            string login = txtLogin.Text.Trim();

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                if (!isEditMode)
                {
                    string insertQuery = @"
                        INSERT INTO Учащиеся (ФИО, Дата_рождения, ID_класса, Логин)
                        VALUES (@fio, @birth, @classId, @login)";
                    SqlCommand cmd = new SqlCommand(insertQuery, conn);
                    cmd.Parameters.AddWithValue("@fio", fio);
                    cmd.Parameters.AddWithValue("@birth", birth);
                    cmd.Parameters.AddWithValue("@classId", classId);
                    cmd.Parameters.AddWithValue("@login", login);
                    cmd.ExecuteNonQuery();
                }
                else
                {
                    string updateQuery = @"
                        UPDATE Учащиеся
                        SET ФИО = @fio, Дата_рождения = @birth, ID_класса = @classId, Логин = @login
                        WHERE ID_ученика = @id";
                    SqlCommand cmd = new SqlCommand(updateQuery, conn);
                    cmd.Parameters.AddWithValue("@fio", fio);
                    cmd.Parameters.AddWithValue("@birth", birth);
                    cmd.Parameters.AddWithValue("@classId", classId);
                    cmd.Parameters.AddWithValue("@login", login);
                    cmd.Parameters.AddWithValue("@id", editStudentId);
                    cmd.ExecuteNonQuery();
                }
            }
            LoadStudents();
            borderForm.Visibility = Visibility.Collapsed;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            borderForm.Visibility = Visibility.Collapsed;
        }

        private void ClearForm()
        {
            txtFIO.Text = "";
            txtBirth.Text = "";
            txtLogin.Text = "";
            cmbClass.SelectedIndex = -1;
        }

        // ---- Вкладка "Оценки и посещаемость" с фильтрацией ----
        private void cmbStudentsForAttendance_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbStudentsForAttendance.SelectedValue != null)
            {
                int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
                LoadStudentGrades(studentId);
                LoadStudentAttendance(studentId);
            }
        }

        private void LoadStudentGrades(int studentId)
        {
            string query = @"
                SELECT о.ID_оценки, п.Название AS Предмет, о.Дата_получения, о.Оценка, о.Тип_контроля
                FROM Оценки о
                JOIN Предметы п ON о.ID_предмета = п.ID_предмета
                WHERE о.ID_ученика = @sid";
            if (currentGradeSubjectFilter.HasValue)
                query += " AND о.ID_предмета = @subjectId";
            query += " ORDER BY о.Дата_получения DESC";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@sid", studentId);
                if (currentGradeSubjectFilter.HasValue)
                    cmd.Parameters.AddWithValue("@subjectId", currentGradeSubjectFilter.Value);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                lvStudentGrades.ItemsSource = dt.DefaultView;
            }
        }

        private void LoadStudentAttendance(int studentId)
        {
            string query = "SELECT ID_посещаемости, Дата, Статус, Примечание FROM Посещаемость WHERE ID_ученика = @sid";
            if (currentAttendanceStatusFilter != "Все")
                query += " AND Статус = @status";
            query += " ORDER BY Дата DESC";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@sid", studentId);
                if (currentAttendanceStatusFilter != "Все")
                    cmd.Parameters.AddWithValue("@status", currentAttendanceStatusFilter);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                lvAttendance.ItemsSource = dt.DefaultView;
            }
        }

        private void FilterGrades_Click(object sender, SelectionChangedEventArgs e)
        {
            if (cmbGradeSubjectFilter.SelectedValue != null && cmbGradeSubjectFilter.SelectedValue != DBNull.Value)
                currentGradeSubjectFilter = Convert.ToInt32(cmbGradeSubjectFilter.SelectedValue);
            else
                currentGradeSubjectFilter = null;

            if (cmbStudentsForAttendance.SelectedValue != null)
            {
                int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
                LoadStudentGrades(studentId);
            }
        }

        private void FilterAttendance_Click(object sender, SelectionChangedEventArgs e)
        {
            if (cmbAttendanceStatusFilter.SelectedItem is ComboBoxItem item)
                currentAttendanceStatusFilter = item.Content.ToString();
            else
                currentAttendanceStatusFilter = "Все";

            if (cmbStudentsForAttendance.SelectedValue != null)
            {
                int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
                LoadStudentAttendance(studentId);
            }
        }

        private void btnRefreshAttendance_Click(object sender, RoutedEventArgs e)
        {
            if (cmbStudentsForAttendance.SelectedValue != null)
            {
                int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
                LoadStudentGrades(studentId);
                LoadStudentAttendance(studentId);
            }
        }

        private void btnAddGrade_Click(object sender, RoutedEventArgs e)
        {
            if (cmbStudentsForAttendance.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите ученика.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
            var addGradeWindow = new AddGradeWindow(connectionString, studentId);
            addGradeWindow.Owner = this;
            if (addGradeWindow.ShowDialog() == true)
            {
                LoadStudentGrades(studentId);
            }
        }

        private void btnAddAttendance_Click(object sender, RoutedEventArgs e)
        {
            if (cmbStudentsForAttendance.SelectedValue == null)
            {
                MessageBox.Show("Сначала выберите ученика.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
            var addAttendanceWindow = new AddAttendanceWindow(connectionString, studentId);
            addAttendanceWindow.Owner = this;
            if (addAttendanceWindow.ShowDialog() == true)
            {
                LoadStudentAttendance(studentId);
            }
        }

        // ---- Отчёты с фильтрацией ----
        private void btnTopStudents_Click(object sender, RoutedEventArgs e)
        {
            if (cmbTopClassFilter.SelectedValue != null && cmbTopClassFilter.SelectedValue != DBNull.Value)
                currentTopClassFilter = Convert.ToInt32(cmbTopClassFilter.SelectedValue);
            else
                currentTopClassFilter = null;

            string query = @"
                SELECT TOP 5
                    у.ФИО,
                    CAST(к.Номер_класса AS VARCHAR) + ' ' + к.Буква_класса AS Класс,
                    ROUND(AVG(CAST(о.Оценка AS FLOAT)), 2) AS Средний_балл
                FROM Учащиеся у
                JOIN Классы к ON у.ID_класса = к.ID_класса
                JOIN Оценки о ON у.ID_ученика = о.ID_ученика
                WHERE 1=1";
            if (currentTopClassFilter.HasValue)
                query += " AND у.ID_класса = @classId";
            query += @" GROUP BY у.ФИО, к.Номер_класса, к.Буква_класса
                        HAVING MIN(о.Оценка) >= 4 AND AVG(CAST(о.Оценка AS FLOAT)) >= 4.5
                        ORDER BY Средний_балл DESC";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                if (currentTopClassFilter.HasValue)
                    cmd.Parameters.AddWithValue("@classId", currentTopClassFilter.Value);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                lvReports.ItemsSource = dt.DefaultView;
                GenerateColumnsForListView(lvReports, dt);
            }
        }

        private void btnClassAvg_Click(object sender, RoutedEventArgs e)
        {
            if (cmbYearFilter.SelectedItem is ComboBoxItem item)
                currentYearFilter = item.Content.ToString();
            else
                currentYearFilter = "2025-2026";

            string query = @"
                SELECT 
                    CAST(к.Номер_класса AS VARCHAR) + ' ' + к.Буква_класса AS Класс,
                    ROUND(AVG(CAST(о.Оценка AS FLOAT)), 2) AS Средний_балл_класса
                FROM Классы к
                JOIN Учащиеся у ON к.ID_класса = у.ID_класса
                JOIN Оценки о ON у.ID_ученика = о.ID_ученика
                WHERE к.Учебный_год = @year
                GROUP BY к.Номер_класса, к.Буква_класса
                ORDER BY Класс";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@year", currentYearFilter.Split('-')[0]);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                lvReports.ItemsSource = dt.DefaultView;
                GenerateColumnsForListView(lvReports, dt);
            }
        }

        private void btnSubjectAvg_Click(object sender, RoutedEventArgs e)
        {
            if (cmbSubjectClassFilter.SelectedValue != null && cmbSubjectClassFilter.SelectedValue != DBNull.Value)
                currentSubjectClassFilter = Convert.ToInt32(cmbSubjectClassFilter.SelectedValue);
            else
                currentSubjectClassFilter = null;

            string query = @"
                SELECT 
                    п.Название AS Предмет,
                    ROUND(AVG(CAST(о.Оценка AS FLOAT)), 2) AS Средняя_оценка
                FROM Предметы п
                JOIN Оценки о ON п.ID_предмета = о.ID_предмета
                JOIN Учащиеся у ON о.ID_ученика = у.ID_ученика
                WHERE 1=1";
            if (currentSubjectClassFilter.HasValue)
                query += " AND у.ID_класса = @classId";
            query += " GROUP BY п.Название ORDER BY Средняя_оценка DESC";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                if (currentSubjectClassFilter.HasValue)
                    cmd.Parameters.AddWithValue("@classId", currentSubjectClassFilter.Value);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                lvReports.ItemsSource = dt.DefaultView;
                GenerateColumnsForListView(lvReports, dt);
            }
        }

        private void GenerateColumnsForListView(ListView listView, DataTable dataTable)
        {
            if (dataTable == null || dataTable.Columns.Count == 0) return;
            var gridView = new GridView();
            foreach (DataColumn col in dataTable.Columns)
            {
                var binding = new Binding(col.ColumnName);
                var column = new GridViewColumn
                {
                    Header = col.ColumnName,
                    DisplayMemberBinding = binding,
                    Width = Double.NaN
                };
                gridView.Columns.Add(column);
            }
            listView.View = gridView;
        }

        // ---- Отчёт по посещаемости с фильтром по дате ----
        private void cmbClassForAttendanceReport_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            LoadAttendanceReport();
        }

        private void btnShowAttendanceReport_Click(object sender, RoutedEventArgs e)
        {
            LoadAttendanceReport();
        }

        private void LoadAttendanceReport()
        {
            if (cmbClassForAttendanceReport.SelectedValue == null) return;

            int classId = Convert.ToInt32(cmbClassForAttendanceReport.SelectedValue);
            currentAttendanceDateFrom = dpAttendanceFrom.SelectedDate;
            currentAttendanceDateTo = dpAttendanceTo.SelectedDate;

            string query = @"
                SELECT 
                    у.ФИО,
                    SUM(CASE WHEN п.Статус = 'Присутствовал' THEN 1 ELSE 0 END) AS Присутствий,
                    SUM(CASE WHEN п.Статус = 'Прогул неуважительный' THEN 1 ELSE 0 END) AS Прогулы_неуваж,
                    SUM(CASE WHEN п.Статус = 'Прогул уважительный' THEN 1 ELSE 0 END) AS Прогулы_уваж,
                    SUM(CASE WHEN п.Статус = 'Больничный' THEN 1 ELSE 0 END) AS Больничных,
                    COUNT(п.ID_посещаемости) AS Всего_записей
                FROM Учащиеся у
                LEFT JOIN Посещаемость п ON у.ID_ученика = п.ID_ученика
                WHERE у.ID_класса = @classId";
            if (currentAttendanceDateFrom.HasValue)
                query += " AND п.Дата >= @dateFrom";
            if (currentAttendanceDateTo.HasValue)
                query += " AND п.Дата <= @dateTo";

            query += " GROUP BY у.ФИО ORDER BY у.ФИО";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@classId", classId);
                if (currentAttendanceDateFrom.HasValue)
                    cmd.Parameters.AddWithValue("@dateFrom", currentAttendanceDateFrom.Value);
                if (currentAttendanceDateTo.HasValue)
                    cmd.Parameters.AddWithValue("@dateTo", currentAttendanceDateTo.Value);
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);
                lvAttendanceReport.ItemsSource = dt.DefaultView;
            }
        }

        // ---- Экспорт (без изменений) ----
        private void btnExportCsv_Click(object sender, RoutedEventArgs e)
        {
            if (lvReports.ItemsSource == null || lvReports.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта. Сначала сформируйте отчёт.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "CSV файлы (*.csv)|*.csv";
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.FileName = $"Отчёт_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    DataTable dt = ((DataView)lvReports.ItemsSource).Table;
                    using (var sw = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sw.Write($"\"{dt.Columns[i].ColumnName}\"");
                            if (i < dt.Columns.Count - 1) sw.Write(";");
                        }
                        sw.WriteLine();

                        foreach (DataRow row in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                string value = row[i]?.ToString() ?? "";
                                value = value.Replace("\"", "\"\"");
                                sw.Write($"\"{value}\"");
                                if (i < dt.Columns.Count - 1) sw.Write(";");
                            }
                            sw.WriteLine();
                        }
                    }
                    MessageBox.Show($"Отчёт сохранён в файл:\n{saveFileDialog.FileName}", "Экспорт завершён", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте в CSV: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void btnExportXlsx_Click(object sender, RoutedEventArgs e)
        {
            if (lvReports.ItemsSource == null || lvReports.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта. Сначала сформируйте отчёт.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.FileName = $"Отчёт_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    DataTable dt = ((DataView)lvReports.ItemsSource).Table;

                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Отчёт");
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = dt.Columns[i].ColumnName;
                            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                        }
                        for (int row = 0; row < dt.Rows.Count; row++)
                        {
                            for (int col = 0; col < dt.Columns.Count; col++)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = dt.Rows[row][col]?.ToString();
                            }
                        }
                        worksheet.Cells[1, 1, dt.Rows.Count + 1, dt.Columns.Count].AutoFitColumns();
                        package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    }
                    MessageBox.Show($"Отчёт сохранён в файл:\n{saveFileDialog.FileName}", "Экспорт завершён", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void btnExportAttendanceReport_Click(object sender, RoutedEventArgs e)
        {
            if (lvAttendanceReport.ItemsSource == null || lvAttendanceReport.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта. Сначала сформируйте отчёт по посещаемости.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.FileName = $"Отчёт_посещаемости_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    DataTable dt = ((DataView)lvAttendanceReport.ItemsSource).Table;
                    using (var package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("Посещаемость");
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            worksheet.Cells[1, i + 1].Value = dt.Columns[i].ColumnName;
                            worksheet.Cells[1, i + 1].Style.Font.Bold = true;
                        }
                        for (int row = 0; row < dt.Rows.Count; row++)
                        {
                            for (int col = 0; col < dt.Columns.Count; col++)
                            {
                                worksheet.Cells[row + 2, col + 1].Value = dt.Rows[row][col]?.ToString();
                            }
                        }
                        worksheet.Cells[1, 1, dt.Rows.Count + 1, dt.Columns.Count].AutoFitColumns();
                        package.SaveAs(new FileInfo(saveFileDialog.FileName));
                    }
                    MessageBox.Show($"Отчёт сохранён в файл:\n{saveFileDialog.FileName}", "Экспорт завершён", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
        // Редактирование оценки
        private void btnEditGrade_Click(object sender, RoutedEventArgs e)
        {
            if (lvStudentGrades.SelectedItem == null)
            {
                MessageBox.Show("Выберите оценку для редактирования.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DataRowView row = (DataRowView)lvStudentGrades.SelectedItem;
            int gradeId = Convert.ToInt32(row["ID_оценки"]);
            int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
            var editWindow = new EditGradeWindow(connectionString, gradeId, studentId);
            editWindow.Owner = this;
            if (editWindow.ShowDialog() == true)
            {
                LoadStudentGrades(studentId);
            }
        }

        // Удаление оценки
        private void btnDeleteGrade_Click(object sender, RoutedEventArgs e)
        {
            if (lvStudentGrades.SelectedItem == null)
            {
                MessageBox.Show("Выберите оценку для удаления.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DataRowView row = (DataRowView)lvStudentGrades.SelectedItem;
            int gradeId = Convert.ToInt32(row["ID_оценки"]);
            if (MessageBox.Show("Удалить выбранную оценку?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                string query = "DELETE FROM Оценки WHERE ID_оценки = @id";
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@id", gradeId);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
                LoadStudentGrades(studentId);
            }
        }

        // Редактирование посещаемости
        private void btnEditAttendance_Click(object sender, RoutedEventArgs e)
        {
            if (lvAttendance.SelectedItem == null)
            {
                MessageBox.Show("Выберите запись посещаемости для редактирования.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DataRowView row = (DataRowView)lvAttendance.SelectedItem;
            int attendanceId = Convert.ToInt32(row["ID_посещаемости"]);
            int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
            var editWindow = new EditAttendanceWindow(connectionString, attendanceId, studentId);
            editWindow.Owner = this;
            if (editWindow.ShowDialog() == true)
            {
                LoadStudentAttendance(studentId);
            }
        }

        // Удаление посещаемости
        private void btnDeleteAttendance_Click(object sender, RoutedEventArgs e)
        {
            if (lvAttendance.SelectedItem == null)
            {
                MessageBox.Show("Выберите запись посещаемости для удаления.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DataRowView row = (DataRowView)lvAttendance.SelectedItem;
            int attendanceId = Convert.ToInt32(row["ID_посещаемости"]);
            if (MessageBox.Show("Удалить выбранную запись посещаемости?", "Подтверждение", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                string query = "DELETE FROM Посещаемость WHERE ID_посещаемости = @id";
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@id", attendanceId);
                    conn.Open();
                    cmd.ExecuteNonQuery();
                }
                int studentId = Convert.ToInt32(cmbStudentsForAttendance.SelectedValue);
                LoadStudentAttendance(studentId);
            }
        }
        private void btnExportAttendanceCsv_Click(object sender, RoutedEventArgs e)
        {
            if (lvAttendanceReport.ItemsSource == null || lvAttendanceReport.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для экспорта. Сначала сформируйте отчёт по посещаемости.", "Информация", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "CSV файлы (*.csv)|*.csv";
            saveFileDialog.DefaultExt = ".csv";
            saveFileDialog.FileName = $"Посещаемость_{DateTime.Now:yyyyMMdd_HHmmss}";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    DataTable dt = ((DataView)lvAttendanceReport.ItemsSource).Table;
                    using (var sw = new StreamWriter(saveFileDialog.FileName, false, Encoding.UTF8))
                    {
                        for (int i = 0; i < dt.Columns.Count; i++)
                        {
                            sw.Write($"\"{dt.Columns[i].ColumnName}\"");
                            if (i < dt.Columns.Count - 1) sw.Write(";");
                        }
                        sw.WriteLine();

                        foreach (DataRow row in dt.Rows)
                        {
                            for (int i = 0; i < dt.Columns.Count; i++)
                            {
                                string value = row[i]?.ToString() ?? "";
                                value = value.Replace("\"", "\"\"");
                                sw.Write($"\"{value}\"");
                                if (i < dt.Columns.Count - 1) sw.Write(";");
                            }
                            sw.WriteLine();
                        }
                    }
                    MessageBox.Show($"CSV сохранён:\n{saveFileDialog.FileName}", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка: {ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void btnAbout_Click(object sender, RoutedEventArgs e)
        {
            AboutWindow about = new AboutWindow();
            about.Owner = this;
            about.ShowDialog();
        }
    }
}
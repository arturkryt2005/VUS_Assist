using Microsoft.Data.Sqlite;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;

namespace VUS_Assist_v._1._0
{
    public partial class MainPage : Page
    {
        #region НАСТРОЙКИ (ШАБЛОН) - МЕНЯЙТЕ ЗДЕСЬ
        // recruitment  — карточка призывника (1 запись на призывника)
        // recruit_training — ВЭК/обучение/права/протокол (0 или 1 запись на призывника,
        //                     столбец recruit_code в ней уникален)
        private const string TABLE_NAME = "recruitment";
        private const string PRIMARY_KEY = "code";
        private const string DISPLAY_ONLY_COLUMN = "pers_second_name";

        private const string TRAINING_TABLE = "recruit_training";
        private const string TRAINING_RECRUIT_FK = "recruit_code";

        // Технический столбец (не показывается в гриде, нужен только для сохранения):
        // есть ли уже строка в recruit_training у этого призывника.
        private const string HAS_TRAINING_ROW_COL = "has_training_row";

        private readonly string[] EDITABLE_COLUMNS = {
            "training_card_number",
            "training_graduation_date",
            "driver_license_number",
            "driver_license_date",
            "training_protocol_number",
            "training_protocol_date"
        };

        private readonly string[] VEC_COLUMNS = { "training_card_number", "training_graduation_date" };

        private readonly string[] PREPARED_COLUMNS_DRIVER = {
            "driver_license_number",
            "driver_license_date"
        };

        private readonly string[] PREPARED_COLUMNS_TRAINING = {
            "training_protocol_number",
            "training_protocol_date"
        };
        #endregion

        private string _dbPath;
        private DataTable _dataTable;

        public MainPage()
        {
            InitializeComponent();
        }

        // Кнопка: Открыть базу данных
        private void OpenDb_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "SQLite Database (*.db)|*.db|All files (*.*)|*.*",
                Title = "Выберите файл базы данных"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                _dbPath = openFileDialog.FileName;

                // СОХРАНЯЕМ ПУТЬ В ОБЩЕЕ ХРАНИЛИЩЕ
                AppData.DbPath = _dbPath;

                TxtFilePath.Text = _dbPath;

                try
                {
                    LoadData();
                    BtnSave.IsEnabled = true;
                    TxtStatus.Text = "Данные успешно загружены.";
                    TxtStatus.Foreground = System.Windows.Media.Brushes.Green;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка загрузки БД:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    TxtStatus.Text = "Ошибка загрузки.";
                    TxtStatus.Foreground = System.Windows.Media.Brushes.Red;
                }
            }
        }

        // Загрузка данных из БД (JOIN recruitment + recruit_training) в DataTable
        private void LoadData()
        {
            _dataTable = new DataTable();
            DataGridMain.Columns.Clear();

            using (var connection = new SqliteConnection($"Data Source={_dbPath}"))
            {
                connection.Open();

                string query = $@"
                    SELECT
                        r.{PRIMARY_KEY} AS {PRIMARY_KEY},
                        r.{DISPLAY_ONLY_COLUMN} AS {DISPLAY_ONLY_COLUMN},
                        rt.card_number AS training_card_number,
                        rt.graduation_date AS training_graduation_date,
                        rt.license_number AS driver_license_number,
                        rt.license_date AS driver_license_date,
                        rt.protocol_number AS training_protocol_number,
                        rt.protocol_date AS training_protocol_date,
                        CASE WHEN rt.{TRAINING_RECRUIT_FK} IS NULL THEN 0 ELSE 1 END AS {HAS_TRAINING_ROW_COL}
                    FROM {TABLE_NAME} r
                    LEFT JOIN {TRAINING_TABLE} rt 
                        ON rt.{TRAINING_RECRUIT_FK} = r.{PRIMARY_KEY} AND rt.status_code = 1";

                using (var command = new SqliteCommand(query, connection))
                using (var reader = command.ExecuteReader())
                {
                    _dataTable.Load(reader);
                }
            }

            SetupDataGridColumns();
            DataGridMain.ItemsSource = _dataTable.DefaultView;
            _dataTable.AcceptChanges();
            UpdateStatistics();
        }

        // Динамическое создание столбцов в DataGrid (технический столбец не показываем)
        private void SetupDataGridColumns()
        {
            DataGridMain.Columns.Add(new DataGridTextColumn
            {
                Header = PRIMARY_KEY,
                Binding = new Binding(PRIMARY_KEY),
                IsReadOnly = true,
                FontWeight = FontWeights.Bold
            });

            DataGridMain.Columns.Add(new DataGridTextColumn
            {
                Header = DISPLAY_ONLY_COLUMN,
                Binding = new Binding(DISPLAY_ONLY_COLUMN),
                IsReadOnly = true,
                Foreground = System.Windows.Media.Brushes.DarkSlateGray
            });

            foreach (var colName in EDITABLE_COLUMNS)
            {
                DataGridMain.Columns.Add(new DataGridTextColumn
                {
                    Header = colName,
                    Binding = new Binding(colName),
                    IsReadOnly = false
                });
            }
        }

        // Подсчёт и отображение статистики
        private void UpdateStatistics()
        {
            if (string.IsNullOrEmpty(_dbPath)) return;

            try
            {
                using (var connection = new SqliteConnection($"Data Source={_dbPath}"))
                {
                    connection.Open();

                    string query = $@"
                        SELECT
                            COUNT(*) AS Total,
                            SUM(CASE WHEN rt.card_number IS NOT NULL AND rt.card_number != ''
                                      AND rt.graduation_date IS NOT NULL AND rt.graduation_date != ''
                                     THEN 1 ELSE 0 END) AS VecPassed,
                            SUM(CASE WHEN
                                    (rt.license_number IS NOT NULL AND rt.license_number != ''
                                     AND rt.license_date IS NOT NULL AND rt.license_date != '')
                                    OR
                                    (rt.protocol_number IS NOT NULL AND rt.protocol_number != ''
                                     AND rt.protocol_date IS NOT NULL AND rt.protocol_date != '')
                                THEN 1 ELSE 0 END) AS Prepared
                        FROM {TABLE_NAME} r
                        LEFT JOIN {TRAINING_TABLE} rt 
                            ON rt.{TRAINING_RECRUIT_FK} = r.{PRIMARY_KEY} AND rt.status_code = 1";

                    using (var command = new SqliteCommand(query, connection))
                    using (var reader = command.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            int total = reader.GetInt32(0);
                            int vecPassed = reader.IsDBNull(1) ? 0 : reader.GetInt32(1);
                            int prepared = reader.IsDBNull(2) ? 0 : reader.GetInt32(2);

                            TxtTotal.Text = total.ToString();
                            TxtVec.Text = $"{vecPassed} / {total}";
                            TxtPrepared.Text = $"{prepared} / {total}";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TxtTotal.Text = "—";
                TxtVec.Text = "—";
                TxtPrepared.Text = "—";
                TxtStatus.Text = $"Ошибка подсчёта статистики: {ex.Message}";
                TxtStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        // Кнопка: Сохранить изменения в БД
        private void SaveDb_Click(object sender, RoutedEventArgs e)
        {
            if (_dataTable == null) return;

            var modifiedRows = _dataTable.GetChanges(DataRowState.Modified);

            if (modifiedRows == null || modifiedRows.Rows.Count == 0)
            {
                TxtStatus.Text = "Нет изменений для сохранения.";
                TxtStatus.Foreground = System.Windows.Media.Brushes.Orange;
                return;
            }

            try
            {
                using (var connection = new SqliteConnection($"Data Source={_dbPath}"))
                {
                    connection.Open();
                    using (var transaction = connection.BeginTransaction())
                    {
                        foreach (DataRow row in modifiedRows.Rows)
                        {
                            string recruitCode = row[PRIMARY_KEY].ToString();

                            bool hasTrainingRow = row[HAS_TRAINING_ROW_COL] != DBNull.Value
                                && Convert.ToInt32(row[HAS_TRAINING_ROW_COL]) == 1;

                            object cardNumber = NormalizeValue(row["training_card_number"]);
                            object graduationDate = NormalizeValue(row["training_graduation_date"]);
                            object licenseNumber = NormalizeValue(row["driver_license_number"]);
                            object licenseDate = NormalizeValue(row["driver_license_date"]);
                            object protocolNumber = NormalizeValue(row["training_protocol_number"]);
                            object protocolDate = NormalizeValue(row["training_protocol_date"]);

                            if (hasTrainingRow)
                            {
                                using (var cmd = new SqliteCommand($@"
                                    UPDATE {TRAINING_TABLE}
                                    SET card_number = @cardNumber,
                                        graduation_date = @graduationDate,
                                        license_number = @licenseNumber,
                                        license_date = @licenseDate,
                                        protocol_number = @protocolNumber,
                                        protocol_date = @protocolDate
                                    WHERE {TRAINING_RECRUIT_FK} = @recruitCode AND status_code = 1", connection, transaction))
                                {
                                    cmd.Parameters.AddWithValue("@cardNumber", cardNumber);
                                    cmd.Parameters.AddWithValue("@graduationDate", graduationDate);
                                    cmd.Parameters.AddWithValue("@licenseNumber", licenseNumber);
                                    cmd.Parameters.AddWithValue("@licenseDate", licenseDate);
                                    cmd.Parameters.AddWithValue("@protocolNumber", protocolNumber);
                                    cmd.Parameters.AddWithValue("@protocolDate", protocolDate);
                                    cmd.Parameters.AddWithValue("@recruitCode", recruitCode);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                            else
                            {
                                bool anyValue = cardNumber != DBNull.Value || graduationDate != DBNull.Value
                                    || licenseNumber != DBNull.Value || licenseDate != DBNull.Value
                                    || protocolNumber != DBNull.Value || protocolDate != DBNull.Value;

                                if (!anyValue)
                                    continue;

                                // У призывника ещё нет записи в recruit_training — создаём новую.
                                // code не автоинкремент (составной ключ code+recruit_code), 
                                // поэтому вычисляем следующий свободный номер для этого recruit_code.
                                using (var cmd = new SqliteCommand($@"
                                    INSERT INTO {TRAINING_TABLE}
                                        (code, status_code, recruit_code, card_number, graduation_date,
                                         license_number, license_date, protocol_number, protocol_date)
                                    VALUES (
                                        (SELECT COALESCE(MAX(code), 0) + 1 FROM {TRAINING_TABLE} WHERE {TRAINING_RECRUIT_FK} = @recruitCode),
                                        1, @recruitCode, @cardNumber, @graduationDate,
                                        @licenseNumber, @licenseDate, @protocolNumber, @protocolDate)", connection, transaction))
                                {
                                    cmd.Parameters.AddWithValue("@recruitCode", recruitCode);
                                    cmd.Parameters.AddWithValue("@cardNumber", cardNumber);
                                    cmd.Parameters.AddWithValue("@graduationDate", graduationDate);
                                    cmd.Parameters.AddWithValue("@licenseNumber", licenseNumber);
                                    cmd.Parameters.AddWithValue("@licenseDate", licenseDate);
                                    cmd.Parameters.AddWithValue("@protocolNumber", protocolNumber);
                                    cmd.Parameters.AddWithValue("@protocolDate", protocolDate);
                                    cmd.ExecuteNonQuery();
                                }
                            }
                        }

                        transaction.Commit();
                    }
                }

                TxtStatus.Text = $"Успешно сохранено строк: {modifiedRows.Rows.Count}";
                TxtStatus.Foreground = System.Windows.Media.Brushes.Green;

                // Перезагружаем данные, чтобы has_training_row и новые записи были актуальны
                LoadData();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка сохранения:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtStatus.Text = "Ошибка сохранения!";
                TxtStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        // Пустую строку/DBNull приводим единообразно к DBNull.Value для параметров SQL
        private static object NormalizeValue(object value)
        {
            if (value == null || value == DBNull.Value)
                return DBNull.Value;

            if (value is string s && string.IsNullOrWhiteSpace(s))
                return DBNull.Value;

            return value;
        }
    }
}
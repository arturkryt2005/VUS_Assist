using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace VUS_Assist_v._1._0
{
    public partial class AddingPage : Page
    {
        // Все нужные поля (ВЭК, права, протокол) лежат в recruit_training —
        // одна запись на призывника (связь через recruit_code).
        // ВНИМАНИЕ: у recruit_training составной первичный ключ (code, recruit_code),
        // столбец code НЕ уникален сам по себе и не является автоинкрементом — 
        // поэтому все операции идут по recruit_code, который в этой таблице уникален.
        private const string TRAINING_TABLE = "recruit_training";
        private const string RECRUIT_FK = "recruit_code";

        private Random _random = new Random();
        private const string DATE_FORMAT = "dd.MM.yyyy";

        // Путь к звуковому файлу
        private readonly string _soundFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "notification.mp3");

        public AddingPage()
        {
            InitializeComponent();
        }

        private void AddVec_Click(object sender, RoutedEventArgs e)
        {
            if (!AppData.IsDbLoaded)
            {
                TxtVecStatus.Text = "Ошибка: сначала загрузите базу данных на странице 'Редактор БД'!";
                TxtVecStatus.Foreground = System.Windows.Media.Brushes.Red;
                return;
            }

            if (!int.TryParse(TxtVecCount.Text, out int count) || count <= 0)
            {
                TxtVecStatus.Text = "Ошибка: введите корректное количество!";
                TxtVecStatus.Foreground = System.Windows.Media.Brushes.Red;
                return;
            }

            try
            {
                int addedCount = GenerateAndInsertVecRecords(count);
                TxtVecStatus.Text = $"Запрошено: {count}, успешно обновлено: {addedCount}";
                TxtVecStatus.Foreground = System.Windows.Media.Brushes.Green;

                ShowNotification($"Успешно обновлено записей ВЭК: {addedCount}");
            }
            catch (Exception ex)
            {
                TxtVecStatus.Text = $"Ошибка: {ex.Message}";
                TxtVecStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        private void AddPrepared_Click(object sender, RoutedEventArgs e)
        {
            if (!AppData.IsDbLoaded)
            {
                TxtPreparedStatus.Text = "Ошибка: сначала загрузите базу данных на странице 'Редактор БД'!";
                TxtPreparedStatus.Foreground = System.Windows.Media.Brushes.Red;
                return;
            }

            if (!int.TryParse(TxtPreparedCount.Text, out int count) || count <= 0)
            {
                TxtPreparedStatus.Text = "Ошибка: введите корректное количество!";
                TxtPreparedStatus.Foreground = System.Windows.Media.Brushes.Red;
                return;
            }

            try
            {
                int addedCount = GenerateAndInsertPreparedRecords(count);
                TxtPreparedStatus.Text = $"Запрошено: {count}, успешно обновлено: {addedCount}";
                TxtPreparedStatus.Foreground = System.Windows.Media.Brushes.Green;

                // Показываем уведомление и воспроизводим звук
                ShowNotification($"Успешно обновлено подготовленных записей: {addedCount}");
            }
            catch (Exception ex)
            {
                TxtPreparedStatus.Text = $"Ошибка: {ex.Message}";
                TxtPreparedStatus.Foreground = System.Windows.Media.Brushes.Red;
            }
        }

        // Метод для показа уведомления со звуком
        private void ShowNotification(string message)
        {
            // Воспроизводим звук
            PlayNotificationSound();

            // Показываем MessageBox
            MessageBox.Show(message, "Уведомление", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        // Метод для воспроизведения звука
        private void PlayNotificationSound()
        {
            try
            {
                if (File.Exists(_soundFilePath))
                {
                    var mediaPlayer = new MediaPlayer();
                    mediaPlayer.Open(new Uri(_soundFilePath, UriKind.Absolute));
                    mediaPlayer.Play();
                }
            }
            catch { }
        }

        private class VecRecordInfo
        {
            public string RecruitCode { get; set; }
            public bool NeedCardNumber { get; set; }
            public bool NeedGraduationDate { get; set; }
        }

        // Заполняем недостающие card_number / graduation_date (ВЭК) в recruit_training.
        // Если у призывника вообще нет записи в recruit_training (не зачислен на
        // обучение), он в выборку не попадает — это не ошибка, а норма.
        private int GenerateAndInsertVecRecords(int requestedCount)
        {
            using (var connection = new SqliteConnection($"Data Source={AppData.DbPath}"))
            {
                connection.Open();

                var existingNumbers = GetExistingNumbers(connection, "card_number");
                var availableRecords = GetAvailableRecordsForVec(connection);

                if (availableRecords.Count == 0)
                {
                    throw new Exception("Нет записей для заполнения ВЭК! Все записи уже заполнены или таблица пуста.");
                }

                int updatedCount = 0;
                int maxAttempts = requestedCount * 50;
                int attempts = 0;

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        while (updatedCount < requestedCount && attempts < maxAttempts && availableRecords.Count > 0)
                        {
                            attempts++;

                            int randomIndex = _random.Next(availableRecords.Count);
                            var record = availableRecords[randomIndex];
                            availableRecords.RemoveAt(randomIndex);

                            // Формируем запрос динамически в зависимости от того, какие столбцы нужно заполнить
                            var updateParts = new List<string>();
                            var parameters = new List<(string name, object value)>();

                            if (record.NeedCardNumber)
                            {
                                string number = GenerateUniqueVecNumber(existingNumbers);
                                existingNumbers.Add(number);
                                updateParts.Add("card_number = @number");
                                parameters.Add(("@number", number));
                            }

                            if (record.NeedGraduationDate)
                            {
                                string vecDate = GenerateRandomDate(14, 0);
                                updateParts.Add("graduation_date = @vecDate");
                                parameters.Add(("@vecDate", vecDate));
                            }

                            if (updateParts.Count == 0)
                                continue;

                            string query = $@"
                        UPDATE {TRAINING_TABLE} 
                        SET {string.Join(", ", updateParts)}
                        WHERE {RECRUIT_FK} = @recruitCode AND status_code = 1";

                            using (var command = new SqliteCommand(query, connection, transaction))
                            {
                                foreach (var param in parameters)
                                {
                                    command.Parameters.AddWithValue(param.name, param.value);
                                }
                                command.Parameters.AddWithValue("@recruitCode", record.RecruitCode);

                                int rowsAffected = command.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    updatedCount++;
                                }
                            }
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }

                return updatedCount;
            }
        }

        private List<VecRecordInfo> GetAvailableRecordsForVec(SqliteConnection connection)
        {
            var records = new List<VecRecordInfo>();

            // Ищем записи в recruit_training, где хотя бы один из столбцов пустой
            string query = $@"
        SELECT {RECRUIT_FK}, card_number, graduation_date
        FROM {TRAINING_TABLE}
        WHERE status_code = 1
        AND ((card_number IS NULL OR card_number = '')
             OR (graduation_date IS NULL OR graduation_date = ''))";

            using (var command = new SqliteCommand(query, connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (reader.IsDBNull(0))
                        continue;

                    var record = new VecRecordInfo
                    {
                        RecruitCode = reader.GetString(0)
                    };

                    // Проверяем, какие столбцы пустые
                    string cardNumber = reader.IsDBNull(1) ? "" : reader.GetString(1);
                    string gradDate = reader.IsDBNull(2) ? "" : reader.GetString(2);

                    record.NeedCardNumber = string.IsNullOrWhiteSpace(cardNumber);
                    record.NeedGraduationDate = string.IsNullOrWhiteSpace(gradDate);

                    // Добавляем только если хотя бы один столбец пустой
                    if (record.NeedCardNumber || record.NeedGraduationDate)
                    {
                        records.Add(record);
                    }
                }
            }

            return records;
        }

        private class PreparedRecordInfo
        {
            public string RecruitCode { get; set; }
            public string VecDate { get; set; }
        }

        // Заполняем недостающие license_number / license_date (водительские права)
        // в recruit_training — только для тех, у кого уже проставлена дата
        // окончания обучения (graduation_date).
        private int GenerateAndInsertPreparedRecords(int requestedCount)
        {
            using (var connection = new SqliteConnection($"Data Source={AppData.DbPath}"))
            {
                connection.Open();

                var existingNumbers = GetExistingNumbers(connection, "license_number");
                var availableRecords = GetAvailableRecordsForPrepared(connection);

                if (availableRecords.Count == 0)
                {
                    throw new Exception("Нет записей для заполнения данных о водительских правах! Все записи уже заполнены или таблица пуста.");
                }

                int updatedCount = 0;
                int maxAttempts = requestedCount * 50;
                int attempts = 0;

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        while (updatedCount < requestedCount && attempts < maxAttempts && availableRecords.Count > 0)
                        {
                            attempts++;

                            string number = GenerateUniquePreparedNumber(existingNumbers);
                            existingNumbers.Add(number);

                            int randomIndex = _random.Next(availableRecords.Count);
                            var selectedRecord = availableRecords[randomIndex];
                            availableRecords.RemoveAt(randomIndex);

                            string preparedDate = GeneratePreparedDate(selectedRecord.VecDate);

                            string query = $@"
                                UPDATE {TRAINING_TABLE} 
                                SET license_number = @number, 
                                    license_date = @date
                                WHERE {RECRUIT_FK} = @recruitCode AND status_code = 1";

                            using (var command = new SqliteCommand(query, connection, transaction))
                            {
                                command.Parameters.AddWithValue("@number", number);
                                command.Parameters.AddWithValue("@date", preparedDate);
                                command.Parameters.AddWithValue("@recruitCode", selectedRecord.RecruitCode);
                                int rowsAffected = command.ExecuteNonQuery();

                                if (rowsAffected > 0)
                                {
                                    updatedCount++;
                                }
                            }
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }

                return updatedCount;
            }
        }

        private HashSet<string> GetExistingNumbers(SqliteConnection connection, string columnName)
        {
            var numbers = new HashSet<string>();
            string query = $"SELECT {columnName} FROM {TRAINING_TABLE} WHERE {columnName} IS NOT NULL AND {columnName} != ''";

            using (var command = new SqliteCommand(query, connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (!reader.IsDBNull(0))
                    {
                        numbers.Add(reader.GetString(0));
                    }
                }
            }

            return numbers;
        }

        private List<PreparedRecordInfo> GetAvailableRecordsForPrepared(SqliteConnection connection)
        {
            var records = new List<PreparedRecordInfo>();
            string query = $@"
                SELECT {RECRUIT_FK}, graduation_date
                FROM {TRAINING_TABLE} 
                WHERE status_code = 1
                AND (license_number IS NULL OR license_number = '')
                AND (license_date IS NULL OR license_date = '')
                AND graduation_date IS NOT NULL 
                AND graduation_date != ''";

            using (var command = new SqliteCommand(query, connection))
            using (var reader = command.ExecuteReader())
            {
                while (reader.Read())
                {
                    if (!reader.IsDBNull(0) && !reader.IsDBNull(1))
                    {
                        records.Add(new PreparedRecordInfo
                        {
                            RecruitCode = reader.GetString(0),
                            VecDate = reader.GetString(1)
                        });
                    }
                }
            }

            return records;
        }

        private string GenerateUniqueVecNumber(HashSet<string> existingNumbers)
        {
            string number;
            do
            {
                if (_random.Next(2) == 0)
                {
                    number = $"АВ{_random.Next(1000, 10000)}";
                }
                else
                {
                    number = $"1600{_random.Next(1000000, 10000000)}";
                }
            }
            while (existingNumbers.Contains(number));

            return number;
        }

        private string GenerateUniquePreparedNumber(HashSet<string> existingNumbers)
        {
            string number;
            do
            {
                number = _random.Next(1000000000, int.MaxValue).ToString();
            }
            while (existingNumbers.Contains(number));

            return number;
        }

        private string GenerateRandomDate(int daysAgoMax, int daysAgoMin)
        {
            DateTime today = DateTime.Now.Date;
            int daysBack = _random.Next(daysAgoMin, daysAgoMax + 1);
            DateTime randomDate = today.AddDays(-daysBack);
            return randomDate.ToString(DATE_FORMAT);
        }

        private string GeneratePreparedDate(string vecDateString)
        {
            if (!DateTime.TryParseExact(vecDateString, DATE_FORMAT,
                CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime vecDate))
            {
                if (!DateTime.TryParse(vecDateString, out vecDate))
                {
                    throw new Exception($"Не удалось распознать дату: {vecDateString}");
                }
            }

            DateTime today = DateTime.Now.Date;
            DateTime minDate = vecDate.AddDays(3);

            if (minDate > today)
            {
                return today.ToString(DATE_FORMAT);
            }

            int daysRange = (today - minDate).Days;
            if (daysRange <= 0)
            {
                return minDate.ToString(DATE_FORMAT);
            }

            DateTime preparedDate = minDate.AddDays(_random.Next(0, daysRange + 1));
            return preparedDate.ToString(DATE_FORMAT);
        }
    }
}
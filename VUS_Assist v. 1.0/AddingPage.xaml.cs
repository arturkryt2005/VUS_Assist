using Microsoft.Data.Sqlite;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using Microsoft.Win32;
using System.Linq;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
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

        private HashSet<string> GetExistingNumbersInTransaction(
    SqliteConnection connection,
    SqliteTransaction transaction,
    string columnName)
        {
            var numbers = new HashSet<string>();
            string query = $@"
        SELECT {columnName} 
        FROM {TRAINING_TABLE} 
        WHERE {columnName} IS NOT NULL 
          AND {columnName} != ''";

            using var command = new SqliteCommand(query, connection, transaction);
            using var reader = command.ExecuteReader();

            while (reader.Read())
            {
                if (!reader.IsDBNull(0))
                {
                    numbers.Add(reader.GetString(0));
                }
            }
            return numbers;
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


        private class SentImportRow
        {
            public int RowNumber { get; set; }
            public string LastName { get; set; } = "";
            public string FirstName { get; set; } = "";
            public string Patronymic { get; set; } = "";
            public string BirthDate { get; set; } = "";
            public string SentDate { get; set; } = "";
            public string TrainingType { get; set; } = "";      // Вид подготовки (ДОСААФ, СПО...)
            public string CategoryText { get; set; } = "";      // Категория/спецтехника (столбец L: С, Д, Е, МТ-ЛБ, Парашютист)
            public string SpecialtyText { get; set; } = "";     // Текст специальности
            public string SpecialtyCode { get; set; } = "";     // Код специальности
            public string EduOrganization { get; set; } = "";   // Учебная организация
            public string VkText { get; set; } = "";            // Текст военкомата (столбец J)
            public string DriverLicenseSeries { get; set; } = ""; // Серия прав (столбец O)
            public string DriverLicenseNumber { get; set; } = "";
        }

        private class ImportReportRow
        {
            public int RowNumber { get; set; }
            public string FullName { get; set; } = "";
            public string BirthDate { get; set; } = "";
            public string SentDate { get; set; } = "";
            public string Training { get; set; } = "";
            public string Status { get; set; } = "";
            public string RecruitCode { get; set; } = "";
            public string Comment { get; set; } = "";
        }

        private void ImportSent_Click(object sender, RoutedEventArgs e)
        {
            if (!AppData.IsDbLoaded)
            {
                TxtSentImportStatus.Text = "Ошибка: сначала загрузите базу данных на странице 'Редактор БД'!";
                TxtSentImportStatus.Foreground = Brushes.Red;
                return;
            }

            var openDialog = new OpenFileDialog
            {
                Filter = "Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*",
                Title = "Выберите Excel со списком отправленных"
            };

            if (openDialog.ShowDialog() != true)
                return;

            try
            {
                var rows = ReadSentRowsFromXlsx(openDialog.FileName);
                var report = ImportSentRows(rows);
                var reportPath = SaveImportReport(report, openDialog.FileName);

                int updated = report.Count(r => r.Status == "обновлен");
                int inserted = report.Count(r => r.Status == "добавлен");
                int skipped = report.Count(r => r.Status == "уже отправлен");
                int errors = report.Count(r => r.Status == "ошибка");

                TxtSentImportStatus.Text = $"Готово. Обновлено: {updated}, добавлено: {inserted}, уже отправлены: {skipped}, ошибок: {errors}. Отчет: {reportPath}";
                TxtSentImportStatus.Foreground = errors == 0 ? Brushes.Green : Brushes.DarkOrange;
                ShowNotification($"Импорт отправленных завершён. Обновлено: {updated}, добавлено: {inserted}, уже отправлены: {skipped}, ошибок: {errors}");
            }
            catch (Exception ex)
            {
                TxtSentImportStatus.Text = $"Ошибка импорта: {ex.Message}";
                TxtSentImportStatus.Foreground = Brushes.Red;
            }
        }

        private List<ImportReportRow> ImportSentRows(List<SentImportRow> rows)
        {
            var report = new List<ImportReportRow>();

            using (var connection = new SqliteConnection($"Data Source={AppData.DbPath}"))
            {
                connection.Open();
                string defaultMunicipalCode = GetDefaultMunicipalCode(connection);

                using (var transaction = connection.BeginTransaction())
                {
                    try
                    {
                        foreach (var row in rows)
                        {
                            var reportRow = new ImportReportRow
                            {
                                RowNumber = row.RowNumber,
                                FullName = $"{row.LastName} {row.FirstName} {row.Patronymic}".Trim(),
                                BirthDate = row.BirthDate,
                                SentDate = row.SentDate,
                                Training = $"{row.TrainingType} {row.SpecialtyText} {row.SpecialtyCode}".Trim()
                            };

                            if (string.IsNullOrWhiteSpace(row.LastName) || string.IsNullOrWhiteSpace(row.FirstName) ||
                                string.IsNullOrWhiteSpace(row.BirthDate) || string.IsNullOrWhiteSpace(row.SentDate))
                            {
                                reportRow.Status = "ошибка";
                                reportRow.Comment = "не заполнены обязательные поля Excel: фамилия, имя, дата рождения или отправка";
                                report.Add(reportRow);
                                continue;
                            }

                            string? recruitCode = FindRecruitCode(connection, transaction, row);
                            bool inserted = false;
                            if (recruitCode == null)
                            {
                                // Пытаемся определить настоящий муниципальный код по тексту
                                // военкомата из файла (столбец J) — иначе всем новым записям
                                // ставился один и тот же муниципальный код (первый попавшийся
                                // в базе), что неверно, если в файле люди из разных районов.
                                string municipalCode = ResolveMunicipalCode(connection, transaction, row.VkText)
                                                        ?? defaultMunicipalCode;
                                recruitCode = InsertRecruit(connection, transaction, row, municipalCode);
                                inserted = true;
                            }

                            reportRow.RecruitCode = recruitCode;

                            // ВАЖНО: "дата отправки" хранится в recruitment.ap_departure_date
                            // (эта дата совпадает с ap_order_date — датой приказа об отправке —
                            // во всех существующих записях базы, это и есть искомое поле).
                            // ap_arrival_date (дата прибытия на сборный пункт) мы теперь тоже
                            // проставляем (пара дней до отправки), но проверка дубликатов идёт
                            // строго по ap_departure_date — у части призывников в ap_arrival_date
                            // уже стояли посторонние значения до этого импорта, и завязывать на
                            // него проверку "уже отправлен" по-прежнему нельзя.
                            if (IsAlreadySent(connection, transaction, recruitCode, row.SentDate))
                            {
                                reportRow.Status = "уже отправлен";
                                reportRow.Comment = "запись не изменялась (дата отправки в базе уже совпадает с датой из файла)";
                                report.Add(reportRow);
                                continue;
                            }

                            DateTime sentDate = ParseDate(row.SentDate);
                            string callDate = sentDate.AddDays(-_random.Next(14, 22)).ToString(DATE_FORMAT);
                            // Прибытие на сборный пункт — за 3-4 дня до отправки.
                            string arrivalDate = sentDate.AddDays(-_random.Next(3, 5)).ToString(DATE_FORMAT);

                            UpsertSummons(connection, transaction, recruitCode, callDate);
                            UpsertCommission(connection, transaction, recruitCode, callDate);
                            UpsertTraining(connection, transaction, recruitCode, row, sentDate);
                            UpdateRecruitDispatchInfo(connection, transaction, recruitCode, row.SentDate, arrivalDate, row.EduOrganization);

                            reportRow.Status = inserted ? "добавлен" : "обновлен";
                            reportRow.Comment = $"дата вызова/комиссии: {callDate}";
                            report.Add(reportRow);
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }
            }

            Logger.Success($"Импорт отправленных: обработано {report.Count} строк");
            return report;
        }

        private string GetDefaultMunicipalCode(SqliteConnection connection)
        {
            using var command = new SqliteCommand("SELECT address_municipal_code FROM recruitment WHERE address_municipal_code IS NOT NULL AND address_municipal_code != '' LIMIT 1", connection);
            return command.ExecuteScalar()?.ToString() ?? "00000000";
        }

        // Пытаемся определить реальный муниципальный код (recruitment.address_municipal_code)
        // по тексту военкомата из Excel (столбец J, напр. "ВК города Альметьевск и
        // Альметьевского района Республики Татарстан, г. Альметьевск").
        // r_dep_military_commissariats.name — это укороченная версия того же текста
        // (без "Республики Татарстан, г. ..." в конце), поэтому ищем её как подстроку.
        // Дальше через r_dep_municipalities.mc_municipal_code (код группы военкомата)
        // находим любой муниципальный код этой группы — он нужен только чтобы потом
        // (в ResolveDosaafOrgCode) найти, в какую автошколу обычно направляли людей
        // из этого же военкомата.
        // Кэш справочника военкоматов, чтобы не читать его заново на каждую строку файла
        private List<(string Code, string NormalizedName)>? _commissariatsCache;

        // ВАЖНО: сравнение делаем в C#, а не в SQL — встроенная SQLite-функция lower()
        // работает только с ASCII и не понимает кириллицу, поэтому LIKE lower(...) 
        // молча не находил бы часть военкоматов с иным регистром/пунктуацией.
        private static string NormalizeForMatch(string value) =>
            (value ?? "").ToLowerInvariant().Replace(",", "").Replace("-", "").Replace("  ", " ").Trim();

        private string? ResolveMunicipalCode(SqliteConnection connection, SqliteTransaction transaction, string vkText)
        {
            if (string.IsNullOrWhiteSpace(vkText))
                return null;

            if (_commissariatsCache == null)
            {
                _commissariatsCache = new List<(string, string)>();
                using var listCmd = new SqliteCommand("SELECT code, name FROM r_dep_military_commissariats", connection, transaction);
                using var reader = listCmd.ExecuteReader();
                while (reader.Read())
                {
                    _commissariatsCache.Add((reader.GetString(0), NormalizeForMatch(reader.GetString(1))));
                }
            }

            string vkNormalized = NormalizeForMatch(vkText);

            // Берём военкомат, чьё (нормализованное) название целиком входит в текст
            // из файла, предпочитая самое длинное/специфичное совпадение.
            string? commissariatCode = _commissariatsCache
                .Where(c => vkNormalized.Contains(c.NormalizedName))
                .OrderByDescending(c => c.NormalizedName.Length)
                .Select(c => (string?)c.Code)
                .FirstOrDefault();

            if (string.IsNullOrEmpty(commissariatCode))
                return null;

            using var municipalCmd = new SqliteCommand(@"
                SELECT code FROM r_dep_municipalities
                WHERE mc_municipal_code = @commissariatCode
                LIMIT 1", connection, transaction);
            municipalCmd.Parameters.AddWithValue("@commissariatCode", commissariatCode);
            return municipalCmd.ExecuteScalar()?.ToString();
        }

        private string? FindRecruitCode(SqliteConnection connection, SqliteTransaction transaction, SentImportRow row)
        {
            const string query = @"SELECT code FROM recruitment
WHERE lower(pers_second_name) = lower(@lastName)
  AND lower(pers_name) = lower(@firstName)
  AND ifnull(lower(pers_patronimic), '') = lower(@patronymic)
  AND pers_birth_date = @birthDate
LIMIT 1";
            using var command = new SqliteCommand(query, connection, transaction);
            command.Parameters.AddWithValue("@lastName", row.LastName);
            command.Parameters.AddWithValue("@firstName", row.FirstName);
            command.Parameters.AddWithValue("@patronymic", row.Patronymic);
            command.Parameters.AddWithValue("@birthDate", row.BirthDate);
            return command.ExecuteScalar()?.ToString();
        }

        private string InsertRecruit(SqliteConnection connection, SqliteTransaction transaction, SentImportRow row, string municipalCode)
        {
            string code;
            do
            {
                code = municipalCode + DateTime.Now.ToString("ddMMyyyy") + _random.Next(100000, 999999).ToString();
            }
            while (RecruitExists(connection, transaction, code));

            const string query = @"INSERT INTO recruitment
(code, card_create_date, pers_second_name, pers_name, pers_patronimic, pers_birth_date, address_municipal_code)
VALUES (@code, @cardDate, @lastName, @firstName, @patronymic, @birthDate, @municipalCode)";
            using var command = new SqliteCommand(query, connection, transaction);
            command.Parameters.AddWithValue("@code", code);
            command.Parameters.AddWithValue("@cardDate", DateTime.Now.ToString(DATE_FORMAT));
            command.Parameters.AddWithValue("@lastName", row.LastName);
            command.Parameters.AddWithValue("@firstName", row.FirstName);
            command.Parameters.AddWithValue("@patronymic", string.IsNullOrWhiteSpace(row.Patronymic) ? DBNull.Value : row.Patronymic);
            command.Parameters.AddWithValue("@birthDate", row.BirthDate);
            command.Parameters.AddWithValue("@municipalCode", municipalCode);
            command.ExecuteNonQuery();
            return code;
        }

        private bool RecruitExists(SqliteConnection connection, SqliteTransaction transaction, string code)
        {
            using var command = new SqliteCommand("SELECT 1 FROM recruitment WHERE code = @code LIMIT 1", connection, transaction);
            command.Parameters.AddWithValue("@code", code);
            return command.ExecuteScalar() != null;
        }

        // Считаем строку уже импортированной ТОЛЬКО если дата отправки в базе
        // (ap_departure_date) уже точно совпадает с датой из файла. Раньше здесь
        // проверялось "поле не пустое", причём читалось не то поле (ap_arrival_date),
        // из-за чего ~250 призывников с посторонним значением в ap_arrival_date
        // ложно считались "уже отправленными" и весь импорт для них пропускался.
        private bool IsAlreadySent(SqliteConnection connection, SqliteTransaction transaction, string recruitCode, string sentDateFromFile)
        {
            using var command = new SqliteCommand("SELECT ap_departure_date FROM recruitment WHERE code = @code", connection, transaction);
            command.Parameters.AddWithValue("@code", recruitCode);
            var value = command.ExecuteScalar()?.ToString();
            if (string.IsNullOrWhiteSpace(value))
                return false;
            return string.Equals(value.Trim(), sentDateFromFile.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        // Пишем дату отправки (ap_departure_date), дату прибытия на сборный пункт
        // (ap_arrival_date, за 1-3 дня до отправки) и учебную организацию/часть
        // (edu_org, берётся из столбца K файла). edu_org обновляем только если
        // в файле для этой строки реально есть значение — иначе не затираем то,
        // что уже было в базе.
        private void UpdateRecruitDispatchInfo(SqliteConnection connection, SqliteTransaction transaction, string recruitCode, string sentDate, string arrivalDate, string eduOrg)
        {
            using var command = new SqliteCommand(@"UPDATE recruitment SET
                ap_departure_date = @sentDate,
                ap_arrival_date = @arrivalDate,
                edu_org = CASE WHEN @eduOrg != '' THEN @eduOrg ELSE edu_org END
                WHERE code = @code", connection, transaction);
            command.Parameters.AddWithValue("@sentDate", sentDate);
            command.Parameters.AddWithValue("@arrivalDate", arrivalDate);
            command.Parameters.AddWithValue("@eduOrg", eduOrg ?? "");
            command.Parameters.AddWithValue("@code", recruitCode);
            command.ExecuteNonQuery();
        }

        private void UpsertSummons(SqliteConnection connection, SqliteTransaction transaction, string recruitCode, string callDate)
        {
            int code = GetNextChildCode(connection, transaction, "log_summonses", recruitCode);
            const string query = "INSERT INTO log_summonses (code, status_code, recruit_code, arrival_date, reason_code) VALUES (@code, 1, @recruitCode, @date, 3)";
            using var command = new SqliteCommand(query, connection, transaction);
            command.Parameters.AddWithValue("@code", code);
            command.Parameters.AddWithValue("@recruitCode", recruitCode);
            command.Parameters.AddWithValue("@date", callDate);
            command.ExecuteNonQuery();
        }

        private void UpsertCommission(SqliteConnection connection, SqliteTransaction transaction, string recruitCode, string commissionDate)
        {
            int code = GetNextChildCode(connection, transaction, "recruit_commissions", recruitCode);
            const string query = @"INSERT INTO recruit_commissions
(code, status_code, recruit_code, commission_type, arrival_date, decision_code, protocol_date, med_category_code)
VALUES (@code, 1, @recruitCode, 2, @date, 1, @date, @medCategory)";
            using var command = new SqliteCommand(query, connection, transaction);
            command.Parameters.AddWithValue("@code", code);
            command.Parameters.AddWithValue("@recruitCode", recruitCode);
            command.Parameters.AddWithValue("@date", commissionDate);
            command.Parameters.AddWithValue("@medCategory", _random.Next(1, 3));
            command.ExecuteNonQuery();
        }


        private void UpsertTraining(SqliteConnection connection, SqliteTransaction transaction, string recruitCode, SentImportRow row, DateTime sentDate)
        {
            int? typeCode = ResolveTrainingType(row.TrainingType);
            string? specialtyCode = ResolveSpecialtyCode(connection, transaction, row.SpecialtyCode, row.SpecialtyText, row.CategoryText, typeCode);

            if (typeCode == null && specialtyCode == null)
                return;

            int code = GetExistingTrainingCode(connection, transaction, recruitCode)
                       ?? GetNextChildCode(connection, transaction, "recruit_training", recruitCode);

            string graduationDate = sentDate.AddDays(-_random.Next(14, 31)).ToString(DATE_FORMAT);

            var existingCards = GetExistingNumbersInTransaction(connection, transaction, "card_number");
            var existingLicenses = GetExistingNumbersInTransaction(connection, transaction, "license_number");

            string cardNumber = GenerateUniqueVecNumber(existingCards);
            string protocolNumber = _random.Next(1000, 999999).ToString();

            bool isDriver = IsDriverSpecialty(row.SpecialtyText, specialtyCode);
            string? licenseNumber = null;
            string? licenseDate = null;

            if (isDriver)
            {
                // В базе license_number хранится как "серия номер" (напр. "9949 062171") —
                // серия в столбце O файла, номер в столбце P. Раньше сюда шёл только
                // номер (P) без серии.
                bool hasSeries = !string.IsNullOrWhiteSpace(row.DriverLicenseSeries) && row.DriverLicenseSeries != "-";
                bool hasNumber = !string.IsNullOrWhiteSpace(row.DriverLicenseNumber) && row.DriverLicenseNumber != "-";

                licenseNumber = (hasSeries && hasNumber)
                    ? $"{row.DriverLicenseSeries} {row.DriverLicenseNumber}"
                    : GenerateUniquePreparedNumber(existingLicenses);

                licenseDate = sentDate.AddDays(-_random.Next(3, 11)).ToString(DATE_FORMAT);
            }

            // Автошкола (org_code -> r_edu_training_schools) нужна только для пути
            // ДОСААФ (typeCode == 1) — для СПО/самостоятельного пути её не проставляем.
            // Жёсткого правила "военкомат -> автошкола" в данных нет (по одному и тому
            // же военкомату исторически направляли в разные школы), поэтому берём
            // школу, которая чаще всего встречалась у уже внесённых призывников из
            // того же военкомата — это наиболее вероятный, а не гарантированный вариант.
            int? orgCode = typeCode == 1
                ? ResolveDosaafOrgCode(connection, transaction, recruitCode)
                : null;

            const string query = @"
        INSERT INTO recruit_training 
        (code, status_code, recruit_code, specialty_code, type_code, org_code, graduation_date, 
         card_number, protocol_date, protocol_number, license_date, license_number)
        VALUES (@code, 1, @recruitCode, @specialtyCode, @typeCode, @orgCode, @graduationDate, 
                @cardNumber, @protocolDate, @protocolNumber, @licenseDate, @licenseNumber)
        ON CONFLICT(code, recruit_code) DO UPDATE SET
            specialty_code = coalesce(excluded.specialty_code, specialty_code),
            type_code = coalesce(excluded.type_code, type_code),
            org_code = coalesce(org_code, excluded.org_code),
            graduation_date = coalesce(nullif(graduation_date, ''), excluded.graduation_date),
            card_number = coalesce(nullif(card_number, ''), excluded.card_number),
            protocol_date = coalesce(nullif(protocol_date, ''), excluded.protocol_date),
            protocol_number = coalesce(nullif(protocol_number, ''), excluded.protocol_number),
            license_date = CASE WHEN excluded.license_date IS NOT NULL THEN coalesce(nullif(license_date, ''), excluded.license_date) ELSE license_date END,
            license_number = CASE WHEN excluded.license_number IS NOT NULL THEN coalesce(nullif(license_number, ''), excluded.license_number) ELSE license_number END";

            using var command = new SqliteCommand(query, connection, transaction);
            command.Parameters.AddWithValue("@code", code);
            command.Parameters.AddWithValue("@recruitCode", recruitCode);
            command.Parameters.AddWithValue("@specialtyCode", (object?)specialtyCode ?? DBNull.Value);
            command.Parameters.AddWithValue("@typeCode", (object?)typeCode ?? DBNull.Value);
            command.Parameters.AddWithValue("@orgCode", (object?)orgCode ?? DBNull.Value);
            command.Parameters.AddWithValue("@graduationDate", graduationDate);
            command.Parameters.AddWithValue("@cardNumber", cardNumber);
            command.Parameters.AddWithValue("@protocolDate", graduationDate);
            command.Parameters.AddWithValue("@protocolNumber", protocolNumber);
            command.Parameters.AddWithValue("@licenseDate", (object?)licenseDate ?? DBNull.Value);
            command.Parameters.AddWithValue("@licenseNumber", (object?)licenseNumber ?? DBNull.Value);

            command.ExecuteNonQuery();
        }

        // Ищет автошколу (org_code), в которую чаще всего направляли призывников из
        // того же военкомата (той же группы r_dep_municipalities.mc_municipal_code),
        // что и текущий призывник, среди уже внесённых записей recruit_training.
        // Возвращает null, если прецедентов нет — тогда org_code просто остаётся пустым.
        private int? ResolveDosaafOrgCode(SqliteConnection connection, SqliteTransaction transaction, string recruitCode)
        {
            using var groupCmd = new SqliteCommand(@"
                SELECT rdm.mc_municipal_code
                FROM recruitment r
                JOIN r_dep_municipalities rdm ON rdm.code = r.address_municipal_code
                WHERE r.code = @recruitCode", connection, transaction);
            groupCmd.Parameters.AddWithValue("@recruitCode", recruitCode);
            string? mcGroup = groupCmd.ExecuteScalar()?.ToString();

            if (string.IsNullOrEmpty(mcGroup))
                return null;

            using var orgCmd = new SqliteCommand(@"
                SELECT rt.org_code
                FROM recruit_training rt
                JOIN recruitment r2 ON r2.code = rt.recruit_code
                JOIN r_dep_municipalities rdm2 ON rdm2.code = r2.address_municipal_code
                WHERE rdm2.mc_municipal_code = @mcGroup
                  AND rt.org_code IS NOT NULL
                  AND rt.status_code = 1
                GROUP BY rt.org_code
                ORDER BY COUNT(*) DESC
                LIMIT 1", connection, transaction);
            orgCmd.Parameters.AddWithValue("@mcGroup", mcGroup);
            var result = orgCmd.ExecuteScalar();
            return result == null ? (int?)null : Convert.ToInt32(result);
        }

        private int? GetExistingTrainingCode(SqliteConnection connection, SqliteTransaction transaction, string recruitCode)
        {
            using var command = new SqliteCommand("SELECT code FROM recruit_training WHERE recruit_code = @recruitCode AND status_code = 1 LIMIT 1", connection, transaction);
            command.Parameters.AddWithValue("@recruitCode", recruitCode);
            var value = command.ExecuteScalar();
            return value == null ? null : Convert.ToInt32(value);
        }

        private int? ResolveTrainingType(string trainingType)
        {
            string normalized = trainingType.Trim().ToUpperInvariant();
            if (normalized.Contains("ДОСААФ")) return 1;
            if (normalized.Contains("СПО")) return 2;
            if (normalized.Contains("САМО")) return 3;
            return null;
        }

        // Категория (столбец L) -> код специальности в r_xrt_specialties, на случай если
        // столбцы M/N её не содержат. Используется ТОЛЬКО для ДОСААФ (typeCode == 1) —
        // именно там нужна специальность независимо от того, водитель человек или нет
        // (например МТ-ЛБ — это не про водительские права, но специальность нужна тоже).
        private static readonly Dictionary<string, string> CATEGORY_TO_SPECIALTY = new(StringComparer.OrdinalIgnoreCase)
        {
            ["С"] = "837037",       // водитель автомобилей категории «C»
            ["Д"] = "845037",       // водитель автомобилей категории «Д»
            ["Е"] = "846037",       // водитель автомобилей категории «E»
            ["МТ-ЛБ"] = "843258",   // механик-водитель плавающих гусеничных тягачей МТЛБ
            ["Парашютист"] = "837100" // водитель-парашютист
        };

        private string? ResolveSpecialtyCode(SqliteConnection connection, SqliteTransaction transaction,
            string excelCode, string specialtyText, string categoryText, int? typeCode)
        {
            // 1) Код из файла (столбец N) — но проверяем, что он реально существует
            //    в справочнике, а не просто похож на код (в файле встречается, например,
            //    несуществующий код "837038" — раньше он записался бы как есть).
            string digits = Regex.Replace(excelCode ?? "", "\\D", "");
            if (digits.Length == 6 && SpecialtyCodeExists(connection, transaction, digits))
                return digits;

            // 2) Текстовое совпадение по столбцу M (срабатывает редко, т.к. там обычно
            //    просто перечислены категории прав, а не название специальности).
            if (!string.IsNullOrWhiteSpace(specialtyText))
            {
                using var command = new SqliteCommand("SELECT code FROM r_xrt_specialties WHERE lower(name) LIKE lower(@name) AND actuality = 1 LIMIT 1", connection, transaction);
                command.Parameters.AddWithValue("@name", $"%{specialtyText.Trim()}%");
                var byText = command.ExecuteScalar()?.ToString();
                if (!string.IsNullOrEmpty(byText))
                    return byText;
            }

            // 3) Для ДОСААФ, если ни код, ни текст не дали результата — подставляем
            //    специальность по категории (столбец L), чтобы у ДОСААФ-обучавшихся
            //    специальность была в любом случае, а не только у водителей.
            if (typeCode == 1 && !string.IsNullOrWhiteSpace(categoryText)
                && CATEGORY_TO_SPECIALTY.TryGetValue(categoryText.Trim(), out var fallbackCode))
            {
                return fallbackCode;
            }

            return null;
        }

        private bool SpecialtyCodeExists(SqliteConnection connection, SqliteTransaction transaction, string code)
        {
            using var command = new SqliteCommand("SELECT 1 FROM r_xrt_specialties WHERE code = @code AND actuality = 1 LIMIT 1", connection, transaction);
            command.Parameters.AddWithValue("@code", code);
            return command.ExecuteScalar() != null;
        }

        private bool IsDriverSpecialty(string specialtyText, string? specialtyCode)
        {
            string text = (specialtyText ?? "").ToLowerInvariant();
            if (text.Contains("параш") || text.Contains("мед"))
                return false;
            if (text.Contains("водител") || text.Contains("категории c") || text.Contains("категории с"))
                return true;
            return !string.IsNullOrWhiteSpace(specialtyCode) && specialtyCode.EndsWith("037", StringComparison.Ordinal);
        }

        private int GetNextChildCode(SqliteConnection connection, SqliteTransaction transaction, string tableName, string recruitCode)
        {
            using var command = new SqliteCommand($"SELECT ifnull(max(code), 0) + 1 FROM {tableName} WHERE recruit_code = @recruitCode", connection, transaction);
            command.Parameters.AddWithValue("@recruitCode", recruitCode);
            return Convert.ToInt32(command.ExecuteScalar());
        }

        private List<SentImportRow> ReadSentRowsFromXlsx(string path)
        {
            var cells = ReadWorksheetCells(path);
            var rows = new List<SentImportRow>();
            foreach (var rowGroup in cells.GroupBy(c => c.Row).Where(g => g.Key > 1).OrderBy(g => g.Key))
            {
                string lastName = GetCell(rowGroup, "A");
                string firstName = GetCell(rowGroup, "B");
                string patronymic = GetCell(rowGroup, "C");
                string birthDate = NormalizeExcelDate(GetCell(rowGroup, "D"));
                string sentDate = NormalizeExcelDate(GetCell(rowGroup, "F"));
                string trainingType = GetCell(rowGroup, "I");
                string vkText = GetCell(rowGroup, "J");
                // Столбец L — категория/спецтехника (напр. "С", "Парашютист", "МТ-ЛБ") —
                // используется как запасной вариант для определения специальности,
                // когда столбцы M/N её не содержат.
                string categoryText = GetCell(rowGroup, "L");
                string specialtyText = GetCell(rowGroup, "M");
                string specialtyCode = GetCell(rowGroup, "N");
                // Реальный текст пункта направления/учебной части
                // (напр. "ВКС, в/ч 20925 (УЧ),Белгород") лежит в столбце K.
                string eduOrg = GetCell(rowGroup, "K");
                // Права хранятся в БД как "серия номер" (напр. "9949 062171") —
                // серия лежит в столбце O, номер в столбце P.
                string driverLicenseSeries = GetCell(rowGroup, "O");
                string driverLicense = GetCell(rowGroup, "P");
                if (string.IsNullOrWhiteSpace(lastName) && string.IsNullOrWhiteSpace(firstName) && string.IsNullOrWhiteSpace(sentDate))
                    continue;
                rows.Add(new SentImportRow
                {
                    RowNumber = rowGroup.Key,
                    LastName = lastName.Trim(),
                    FirstName = firstName.Trim(),
                    Patronymic = patronymic.Trim(),
                    BirthDate = birthDate,
                    SentDate = sentDate,
                    TrainingType = trainingType.Trim(),
                    CategoryText = categoryText.Trim(),
                    SpecialtyText = specialtyText.Trim(),
                    SpecialtyCode = specialtyCode.Trim(),
                    EduOrganization = eduOrg.Trim(),
                    VkText = vkText.Trim(),
                    DriverLicenseSeries = driverLicenseSeries.Trim(),
                    DriverLicenseNumber = driverLicense.Trim()
                });
            }
            return rows;
        }

        private static string GetCell(IEnumerable<(int Row, string Column, string Value)> row, string column) =>
            row.FirstOrDefault(c => c.Column == column).Value ?? "";

        private static List<(int Row, string Column, string Value)> ReadWorksheetCells(string path)
        {
            using var archive = ZipFile.OpenRead(path);
            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            var shared = new List<string>();
            var sharedEntry = archive.GetEntry("xl/sharedStrings.xml");
            if (sharedEntry != null)
            {
                using var stream = sharedEntry.Open();
                var doc = XDocument.Load(stream);
                shared = doc.Descendants(ns + "si").Select(si => string.Concat(si.Descendants(ns + "t").Select(t => t.Value))).ToList();
            }

            var sheetEntry = archive.GetEntry("xl/worksheets/sheet1.xml") ?? throw new Exception("В Excel не найден первый лист");
            using var sheetStream = sheetEntry.Open();
            var sheet = XDocument.Load(sheetStream);
            return sheet.Descendants(ns + "c").Select(c =>
            {
                string reference = c.Attribute("r")?.Value ?? "";
                string column = Regex.Match(reference, "[A-Z]+").Value;
                int row = int.Parse(Regex.Match(reference, "[0-9]+").Value);
                string value = c.Element(ns + "v")?.Value ?? "";
                if (c.Attribute("t")?.Value == "s" && int.TryParse(value, out int idx) && idx >= 0 && idx < shared.Count)
                    value = shared[idx];
                return (row, column, value);
            }).ToList();
        }

        private string NormalizeExcelDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "";
            value = value.Trim();
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double oa))
                return DateTime.FromOADate(oa).ToString(DATE_FORMAT);
            return ParseDate(value).ToString(DATE_FORMAT);
        }

        private DateTime ParseDate(string value)
        {
            string[] formats = { DATE_FORMAT, "d.M.yyyy", "M/d/yy", "M/d/yyyy", "yyyy-MM-dd" };
            if (DateTime.TryParseExact(value, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime date))
                return date;
            if (DateTime.TryParse(value, CultureInfo.GetCultureInfo("ru-RU"), DateTimeStyles.None, out date))
                return date;
            throw new Exception($"Не удалось распознать дату: {value}");
        }

        private string SaveImportReport(List<ImportReportRow> report, string sourcePath)
        {
            string reportPath = Path.Combine(Path.GetDirectoryName(sourcePath) ?? Environment.CurrentDirectory, $"sent_import_report_{DateTime.Now:yyyyMMdd_HHmmss}.csv");
            var sb = new StringBuilder();
            sb.AppendLine("Строка;ФИО;Дата рождения;Дата отправки;Подготовка;Статус;Код;Комментарий");
            foreach (var row in report)
                sb.AppendLine(string.Join(";", EscapeCsv(row.RowNumber.ToString()), EscapeCsv(row.FullName), EscapeCsv(row.BirthDate), EscapeCsv(row.SentDate), EscapeCsv(row.Training), EscapeCsv(row.Status), EscapeCsv(row.RecruitCode), EscapeCsv(row.Comment)));
            File.WriteAllText(reportPath, sb.ToString(), Encoding.UTF8);
            return reportPath;
        }

        private static string EscapeCsv(string value) => $"\"{(value ?? "").Replace("\"", "\"\"")}\"";

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
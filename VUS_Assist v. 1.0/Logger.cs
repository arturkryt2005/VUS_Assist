using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;

namespace VUS_Assist_v._1._0
{
    public static class Logger
    {
        // Список всех логов в памяти
        private static readonly List<LogEntry> _logs = new List<LogEntry>();

        // Путь к файлу логов
        private static readonly string _logFilePath = Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            "application.log");

        // Объект для синхронизации (многопоточность)
        private static readonly object _lock = new object();

        // Событие, которое уведомляет страницу логов о новом логе
        public static event Action<LogEntry> OnLogAdded;

        // Теги для разных типов логов
        public const string TAG_INFO = "INFO";
        public const string TAG_ERROR = "ERROR";
        public const string TAG_WARN = "WARN";
        public const string TAG_DB = "DB";
        public const string TAG_GEN = "GEN";
        public const string TAG_UI = "UI";
        public const string TAG_SUCCESS = "SUCCESS";

        static Logger()
        {
            // Очищаем файл логов при запуске (или можно оставить историю)
            try
            {
                File.WriteAllText(_logFilePath,
                    $"=== ЛОГИ ПРИЛОЖЕНИЯ VUS Assist v1.0 ===\n" +
                    $"Запуск: {DateTime.Now:dd.MM.yyyy HH:mm:ss}\n" +
                    $"=========================================\n\n");
            }
            catch { }
        }

        // Получить все логи
        public static List<LogEntry> GetAllLogs()
        {
            lock (_lock)
            {
                return new List<LogEntry>(_logs);
            }
        }

        // Очистить все логи
        public static void ClearLogs()
        {
            lock (_lock)
            {
                _logs.Clear();
                try
                {
                    File.WriteAllText(_logFilePath,
                        $"=== Логи очищены: {DateTime.Now:dd.MM.yyyy HH:mm:ss} ===\n\n");
                }
                catch { }
            }
        }

        // Базовый метод логирования
        private static void Log(LogLevel level, string tag, string message)
        {
            var entry = new LogEntry(level, tag, message);

            lock (_lock)
            {
                _logs.Add(entry);

                // Записываем в файл
                try
                {
                    File.AppendAllText(_logFilePath, entry.FormattedMessage + Environment.NewLine);
                }
                catch { }
            }

            // Уведомляем подписчиков (страницу логов)
            OnLogAdded?.Invoke(entry);

            // Также дублируем в Debug окно Visual Studio
            System.Diagnostics.Debug.WriteLine(entry.FormattedMessage);
        }

        // === Публичные методы для разных типов логов ===

        public static void Info(string message)
        {
            Log(LogLevel.Info, TAG_INFO, message);
        }

        public static void Warning(string message)
        {
            Log(LogLevel.Warning, TAG_WARN, message);
        }

        public static void Error(string message, Exception ex = null)
        {
            string fullMessage = ex != null
                ? $"{message} | Исключение: {ex.GetType().Name}: {ex.Message}"
                : message;
            Log(LogLevel.Error, TAG_ERROR, fullMessage);
        }

        public static void Success(string message)
        {
            Log(LogLevel.Success, TAG_SUCCESS, message);
        }

        public static void Database(string message)
        {
            Log(LogLevel.Database, TAG_DB, message);
        }

        public static void Generation(string message)
        {
            Log(LogLevel.Generation, TAG_GEN, message);
        }

        public static void UI(string message)
        {
            Log(LogLevel.UI, TAG_UI, message);
        }

        // Метод для экспорта логов в файл
        public static void ExportLogsToFile(string filePath)
        {
            try
            {
                lock (_lock)
                {
                    var lines = new List<string>();
                    lines.Add($"=== ЭКСПОРТ ЛОГОВ: {DateTime.Now:dd.MM.yyyy HH:mm:ss} ===");
                    lines.Add($"Всего записей: {_logs.Count}");
                    lines.Add(new string('=', 50));

                    foreach (var entry in _logs)
                    {
                        lines.Add(entry.FormattedMessage);
                    }

                    File.WriteAllLines(filePath, lines);
                }
                Success($"Логи экспортированы в файл: {filePath}");
            }
            catch (Exception ex)
            {
                Error($"Ошибка экспорта логов", ex);
                throw;
            }
        }
    }
}
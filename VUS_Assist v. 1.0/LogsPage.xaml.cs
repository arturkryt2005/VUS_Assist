using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace VUS_Assist_v._1._0
{
    public partial class LogsPage : Page
    {
        // Полный список логов (без фильтра)
        private List<LogEntry> _allLogs = new List<LogEntry>();
        // Отфильтрованный список
        private List<LogEntry> _filteredLogs = new List<LogEntry>();

        public LogsPage()
        {
            InitializeComponent();

            // Подписываемся на событие добавления нового лога
            Logger.OnLogAdded += OnNewLogAdded;

            // Загружаем существующие логи
            LoadAllLogs();
        }

        private void LoadAllLogs()
        {
            _allLogs = Logger.GetAllLogs();
            ApplyFilters();
            UpdateStatus();
        }

        // Обработчик добавления нового лога
        private void OnNewLogAdded(LogEntry entry)
        {
            // Вызываем в UI потоке
            Dispatcher.Invoke(() =>
            {
                _allLogs.Add(entry);
                ApplyFilters();
                UpdateStatus();

                // Прокручиваем к последнему логу
                if (LvLogs.Items.Count > 0)
                {
                    LvLogs.ScrollIntoView(LvLogs.Items[LvLogs.Items.Count - 1]);
                }
            });
        }

        // Применение фильтров
        private void ApplyFilters()
        {
            string selectedTag = (CmbTagFilter.SelectedItem as ComboBoxItem)?.Content?.ToString();
            string searchText = TxtSearch.Text.Trim().ToLower();

            _filteredLogs = _allLogs.Where(log =>
            {
                // Фильтр по тегу
                bool tagMatch = selectedTag == "Все" || $"[{log.Tag}]" == selectedTag;

                // Фильтр по поиску
                bool searchMatch = string.IsNullOrEmpty(searchText) ||
                                   log.Message.ToLower().Contains(searchText) ||
                                   log.Tag.ToLower().Contains(searchText);

                return tagMatch && searchMatch;
            }).ToList();

            LvLogs.ItemsSource = null;
            LvLogs.ItemsSource = _filteredLogs;

            // Применяем цвета для разных типов логов
            ApplyLogColors();
        }

        // Применение цветов для разных типов логов
        private void ApplyLogColors()
        {
            // Используем ItemContainerStyle для раскраски строк
            LvLogs.ItemContainerStyle = new Style(typeof(ListViewItem));
            LvLogs.ItemContainerStyle.Setters.Add(new Setter(ListViewItem.HorizontalContentAlignmentProperty, HorizontalAlignment.Stretch));

            // Добавляем триггеры для разных тегов
            var triggers = LvLogs.ItemContainerStyle.Triggers;

            var errorTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Tag"),
                Value = "ERROR"
            };
            errorTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.Red));
            errorTrigger.Setters.Add(new Setter(ListViewItem.FontWeightProperty, FontWeights.Bold));
            triggers.Add(errorTrigger);

            var warnTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Tag"),
                Value = "WARN"
            };
            warnTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.DarkOrange));
            triggers.Add(warnTrigger);

            var successTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Tag"),
                Value = "SUCCESS"
            };
            successTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.Green));
            triggers.Add(successTrigger);

            var dbTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Tag"),
                Value = "DB"
            };
            dbTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.DarkCyan));
            triggers.Add(dbTrigger);

            var genTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Tag"),
                Value = "GEN"
            };
            genTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.Purple));
            triggers.Add(genTrigger);

            var uiTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Tag"),
                Value = "UI"
            };
            uiTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.DarkBlue));
            triggers.Add(uiTrigger);

            var infoTrigger = new DataTrigger
            {
                Binding = new System.Windows.Data.Binding("Tag"),
                Value = "INFO"
            };
            infoTrigger.Setters.Add(new Setter(ListViewItem.ForegroundProperty, Brushes.Black));
            triggers.Add(infoTrigger);
        }

        private void CmbTagFilter_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (IsLoaded)
            {
                ApplyFilters();
                UpdateStatus();
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilters();
            UpdateStatus();
        }

        private void ClearLogs_Click(object sender, RoutedEventArgs e)
        {
            var result = MessageBox.Show(
                "Вы уверены, что хотите очистить все логи? Это действие нельзя отменить.",
                "Подтверждение",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result == MessageBoxResult.Yes)
            {
                Logger.ClearLogs();
                _allLogs.Clear();
                ApplyFilters();
                UpdateStatus();
                Logger.Success("Логи очищены пользователем");
            }
        }

        private void ExportLogs_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Текстовый файл (*.txt)|*.txt|Все файлы (*.*)|*.*",
                Title = "Экспорт логов",
                FileName = $"logs_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    Logger.ExportLogsToFile(saveDialog.FileName);
                    MessageBox.Show($"Логи успешно экспортированы в файл:\n{saveDialog.FileName}",
                        "Экспорт завершён", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка экспорта: {ex.Message}",
                        "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void UpdateStatus()
        {
            TxtStatus.Text = $"Всего записей: {_allLogs.Count} | Отображается: {_filteredLogs.Count}";
        }
    }
}
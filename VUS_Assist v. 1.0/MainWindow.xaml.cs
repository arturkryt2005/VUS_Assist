using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace VUS_Assist_v._1._0
{
    public partial class MainWindow : Window
    {
        private MainPage _mainPage;
        private AddingPage _addingPage;
        private LogsPage _logsPage;

        // Активный цвет кнопки (красный)
        private readonly SolidColorBrush _activeColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FFB10A0A"));
        // Неактивный цвет кнопки (зелёный)
        private readonly SolidColorBrush _inactiveColor = new SolidColorBrush((Color)ColorConverter.ConvertFromString("#FF2C6300"));

        public MainWindow()
        {
            InitializeComponent();

            // Инициализируем страницы
            _mainPage = new MainPage();
            _addingPage = new AddingPage();
            _logsPage = new LogsPage();

            // Открываем главную страницу по умолчанию
            MainFrame.Navigate(_mainPage);
            UpdateNavigationButtons(0);

            // Логирование запуска приложения
            Logger.Info("Приложение VUS Assist v1.0 запущено");
        }

        private void NavigateToMain_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(_mainPage);
            UpdateNavigationButtons(0);
            Logger.UI("Переход на страницу 'Просмотр БД'");
        }

        private void NavigateToAdding_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(_addingPage);
            UpdateNavigationButtons(1);
            Logger.UI("Переход на страницу 'Добавление ВЭК/ГИБДД'");
        }

        private void NavigateToLogs_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(_logsPage);
            UpdateNavigationButtons(2);
            Logger.UI("Переход на страницу 'Логи'");
        }

        // Обновление подсветки кнопок навигации
        // activeIndex: 0 - Просмотр БД, 1 - Добавление, 2 - Логи
        private void UpdateNavigationButtons(int activeIndex)
        {
            BtnMain.Background = activeIndex == 0 ? _activeColor : _inactiveColor;
            BtnAdding.Background = activeIndex == 1 ? _activeColor : _inactiveColor;
            BtnLogs.Background = activeIndex == 2 ? _activeColor : _inactiveColor;
        }
    }
}
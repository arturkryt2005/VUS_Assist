using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VUS_Assist_v._1._0
{
    public static class AppData
    {
        // Общий путь к базе данных
        public static string DbPath { get; set; }

        // Проверка, загружена ли база данных
        public static bool IsDbLoaded => !string.IsNullOrEmpty(DbPath);
    }
}

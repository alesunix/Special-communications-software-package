using System;
using System.Windows.Forms;

namespace Список_на_курьерские_отправления
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form2());
        }
    }
    static class Dostup
    {
        public static string Value { get; set; }
    }
}

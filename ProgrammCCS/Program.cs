using System;
using System.Windows.Forms;

namespace ProgramCCS
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
            Application.Run(new Login());
    }
    }
    static class Dostup
    {
        public static string Value { get; set; }
    }   
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RazvLogins
{
    static class Program
    {
        /// <summary>
        /// Главная точка входа для приложения.
        /// </summary>
        [STAThread]
        static void Main()
        {
            DateTime myNewDate = new DateTime(2017, 11, 1);
            DateTime myDate = DateTime.Parse("11.01.2017", System.Globalization.CultureInfo.InvariantCulture);
            if (DateTime.Now > myNewDate)
            {
                return;
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}

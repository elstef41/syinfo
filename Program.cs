using System;
using System.Collections.Generic;
using System.Threading;
using System.Windows.Forms;
using System.Globalization;
using System.Resources;

namespace syinfo
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        // Detección de compatiblidad
        static void Main()
        {
            ResourceManager rm = new ResourceManager(typeof(syinfo));
            string version_so = System.Environment.OSVersion.Version.Major.ToString() + "." + System.Environment.OSVersion.Version.Minor.ToString();
            if (System.Environment.OSVersion.Version.Major <= 5)
            {
                MessageBox.Show(rm.GetString("s_error_no_compatible"), Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
            }
            else
            {
                Application.EnableVisualStyles();
                Application.ThreadException += new ThreadExceptionEventHandler(UIThreadException);
                Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CurrentDomain_UnhandledException);
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new syinfo());
            }
                
        }

        // Controlador de excepciones casero
        private static void UIThreadException(object sender, ThreadExceptionEventArgs e)
        {
            try
            {
                excpt exception = new excpt(e.Exception.ToString());
                exception.TopMost = true;
                exception.ShowDialog();
                Application.Exit();
            }
            catch
            {
                try
                {
                    excpt exception = new excpt(e.Exception.ToString());
                    exception.TopMost = true;
                    exception.ShowDialog();
                    Application.Exit();
                }
                finally
                {
                    Application.Exit();
                }
            }
            Application.Exit();
        }
        static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            excpt exception = new excpt(e.ExceptionObject.ToString());
            exception.TopMost = true;
            exception.ShowDialog();
            Application.Exit();
        }
    }
}

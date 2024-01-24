using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using System.Resources;
using System.Globalization;

namespace syinfo
{
    class strings
    {
        public static string sy_repositorio = "https://github.com/elstef41/syinfo";
        public static ResourceManager rm = new ResourceManager(typeof(syinfo));
        public string obtenerVersion()
        {
            string s = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString();
            return s;
        }

        public string ver()
        {
            string osver = System.Environment.OSVersion.Version.Major.ToString() + "." + System.Environment.OSVersion.Version.Minor.ToString() + "." + System.Environment.OSVersion.Version.Build.ToString();
            switch (osver)
            {
                case "5.1":
                case "6.0":
                    MessageBox.Show(rm.GetString("s_error_no_compatible"), "elstef41 Syinfo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "";
                case "6.1":
                    return "Windows 7";
                case "6.2":
                    return "Windows 8.x/10/11";
                case "6.3":
                case "6.4":
                case "6.5":
                case "10.0":
                    return "Windows 10";
                default:
                    return osver;
            }
        }

        public int minAMs(int i)
        {
            i = i * 60000;
            return i;
        }

    }
}

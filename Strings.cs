using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text;
using System.Resources;

namespace syinfo
{
    class Strings
    {
        public string obtenerVersion()
        {
            string s = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor.ToString() + "." + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build.ToString();
            return s;
        }

        public string ver()
        {
            string osver = System.Environment.OSVersion.Version.Major.ToString() + "." + System.Environment.OSVersion.Version.Minor.ToString();
            switch (osver)
            {
                case "5.1":
                case "6.0":
                    MessageBox.Show("Syinfo requiere Windows 7 o posterior.", "elstef41 Syinfo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return "";
                case "6.1":
                    return "Windows 7";
                case "6.2":
                    return "Windows 8, 8.1, 10 u 11";
                case "6.3":
                case "6.4":
                case "6.5":
                case "6.6":
                case "6.7":
                case "6.8":
                case "10.0":
                    return "Windows 10";
                default:
                    return osver;
            }
        }

    }
}

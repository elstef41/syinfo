using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Management;
using Microsoft.VisualBasic;
using System.IO;

namespace syinfo
{
    public partial class Form1 : Form
    {
        Strings SC = new Strings();
        public Form1()
        {
            SC.ver();
            InitializeComponent();
            this.Text = "Syinfo ";
            this.Text += SC.obtenerVersion();
            this.Text += " por elstef41";
            this.MinimumSize = new Size(270, 268);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            Carga();
            label1.Visible = false;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        public bool Carga()
        {
            List<string> todoCpu = new List<string>();
            List<string> todoSys = new List<string>();
            List<string> todoPhy = new List<string>();
            List<string> todoAlm = new List<string>();
            ManagementObjectCollection cpu = new ManagementObjectSearcher("SELECT * from Win32_Processor").Get();
            ManagementObjectCollection sys = new ManagementObjectSearcher("SELECT * from Win32_ComputerSystem").Get();
            ManagementObjectCollection phy = new ManagementObjectSearcher("SELECT * from Win32_SystemSlot").Get();
            ManagementObjectCollection discoduro = new ManagementObjectSearcher("SELECT * from Win32_DiskDrive").Get();
            ManagementObjectCollection cddvd = new ManagementObjectSearcher("SELECT * from Win32_CDROMDrive").Get();
            foreach (ManagementObject m in cpu)
            {
                todoCpu.Add(m["AddressWidth"].ToString());
                todoCpu.Add(m["MaxClockSpeed"].ToString());
                todoCpu.Add(m["Name"].ToString());
                todoCpu.Add(m["DeviceID"].ToString());
                todoCpu.Add(m["Status"].ToString());
                switch (m["PowerManagementSupported"].ToString())
                {
                    case "True":
                        todoCpu.Add("Sí");
                        break;
                    default:
                        todoCpu.Add("No");
                        break;
                }
                switch (m["Architecture"].ToString())
                {
                    case "0":
                        todoCpu.Add("x86");
                        break;
                    case "1":
                        todoCpu.Add("MIPS");
                        break;
                    case "2":
                        todoCpu.Add("Alpha");
                        break;
                    case "3":
                        todoCpu.Add("PowerPC");
                        break;
                    case "5":
                        todoCpu.Add("ARM");
                        break;
                    case "6":
                        todoCpu.Add("ia64");
                        break;
                    case "9":
                        todoCpu.Add("x86_64");
                        break;
                    case "12":
                        todoCpu.Add("ARM64");
                        break;
                    default:
                        todoCpu.Add("Desconocida");
                        break;
                }
                switch (m["CurrentVoltage"].ToString())
                {
                    case "1":
                        todoCpu.Add("5");
                        break;
                    case "2":
                        todoCpu.Add("3.3");
                        break;
                    case "4":
                        todoCpu.Add("2.9");
                        break;
                    default:
                        todoCpu.Add("Desconocida");
                        break;
                }
                todoCpu.Add(m["CurrentClockSpeed"].ToString());
            }
            foreach (ManagementObject m in sys)
            {
                todoSys.Add(m["BootupState"].ToString());
                switch (m["BootROMSupported"].ToString()) {
                    case "True":
                        todoSys.Add("Sí");
                        break;
                    default:
                        todoSys.Add("No");
                        break;
                }
                todoSys.Add(m["Workgroup"].ToString());
                switch (m["ThermalState"].ToString())
                {
                    case "1":
                        todoSys.Add("Otro");
                        break;
                    case "3":
                        todoSys.Add("Seguro");
                        break;
                    case "4":
                        todoSys.Add("Advertencia");
                        break;
                    case "5":
                        todoSys.Add("Crítico");
                        break;
                    case "6":
                        todoSys.Add("Irrecuperable");
                        break;
                    case "2":
                    default:
                        todoSys.Add("Desconocido");
                        break;
                }
                switch (m["AutomaticResetBootOption"].ToString())
                {
                    case "True":
                        todoSys.Add("Sí");
                        break;
                    case "False":
                        todoSys.Add("No");
                        break;
                    default:
                        todoSys.Add("Desconocido");
                        break;
                }
            }
            foreach (ManagementObject m in phy)
            {
                todoPhy.Add(m["Status"].ToString());
                switch (m["SupportsHotPlug"].ToString())
                {
                    case "True":
                        todoPhy.Add("Sí");
                        break;
                    default:
                        todoPhy.Add("No");
                        break;
                }
                todoPhy.Add(m["CurrentUsage"].ToString());
            }
            foreach (ManagementObject m in discoduro)
            {
                todoAlm.Add(m["Size"].ToString());
                todoAlm.Add(m["Partitions"].ToString());
                todoAlm.Add(m["TotalSectors"].ToString());
                todoAlm.Add(m["TotalTracks"].ToString());
                todoAlm.Add(m["Manufacturer"].ToString());
                todoAlm.Add(m["TracksPerCylinder"].ToString());
                todoAlm.Add(m["InterfaceType"].ToString());Environment.GetLogicalDrives();
            }
            string cpy = "";
            // Básico
            dataGridView1.Rows.Add("Compilación", System.Environment.OSVersion.Version, cpy);
            dataGridView1.Rows.Add("Nombre del equipo", System.Environment.MachineName, cpy);
            dataGridView1.Rows.Add("Arquitectura", todoCpu[6], cpy);
            dataGridView1.Rows.Add("Nombre de usuario activo", System.Environment.UserName, cpy);
            dataGridView1.Rows.Add("Grupo de trabajo", todoSys[2], cpy);
            dataGridView1.Rows.Add("Directorio del sistema", System.Environment.SystemDirectory, cpy);
            dataGridView1.Rows.Add("Cantidad de núcleos", System.Environment.ProcessorCount, cpy);
            dataGridView1.Rows.Add("Tiempo trancurrido", Environment.TickCount / 3600000 + " horas, " + Environment.TickCount / 60000 + " minutos y " + Environment.TickCount / 1000 + " segundos.", cpy);

            // Sistema operativo
            dataGridView2.Rows.Add("Sistema operativo", SC.ver(), cpy);
            dataGridView2.Rows.Add("Versión completa de Windows NT", System.Environment.OSVersion.VersionString, cpy);
            dataGridView2.Rows.Add("Versión del núcleo", System.Environment.OSVersion.Version, cpy);
            dataGridView2.Rows.Add("Compilación", System.Environment.OSVersion.Version.Build, cpy);
            dataGridView2.Rows.Add("Plataforma", System.Environment.OSVersion.Platform, cpy);
            if (Environment.OSVersion.ServicePack == "")
            {
                dataGridView2.Rows.Add("Service Pack", "No aplica", cpy);
            }
            else
            {
                dataGridView2.Rows.Add("Service Pack", Environment.OSVersion.ServicePack, cpy);
            }
            dataGridView2.Rows.Add("Grupo de trabajo", todoSys[2], cpy);

            // Componentes internos
            dataGridView3.Rows.Add("Memoria RAM", System.Environment.Version.Major + "." + Environment.Version.MajorRevision + "." + Environment.Version.Minor + "." + Environment.Version.MinorRevision, cpy);
            dataGridView3.Rows.Add("Nombre del procesador", todoCpu[2], cpy);
            dataGridView3.Rows.Add("ID del procesador", todoCpu[3], cpy);
            dataGridView3.Rows.Add("Dirección del procesador", todoCpu[0], cpy);
            dataGridView3.Rows.Add("Cantidad de núcleos", System.Environment.ProcessorCount, cpy);
            dataGridView3.Rows.Add("Estado de arranque", todoSys[0], cpy);
            dataGridView3.Rows.Add("Soporte para memoria ROM", todoSys[1], cpy);
            dataGridView3.Rows.Add("Estado térmico", todoSys[3], cpy);
            dataGridView3.Rows.Add("Soporte para cambio en caliente", todoPhy[1], cpy);
            dataGridView3.Rows.Add("BUS", todoPhy[2], cpy);

            // CPU
            dataGridView4.Rows.Add("ID", todoCpu[3], cpy);
            dataGridView4.Rows.Add("Nombre", todoCpu[2], cpy);
            dataGridView4.Rows.Add("Dirección", todoCpu[0], cpy);
            dataGridView4.Rows.Add("Velocidad (MHz)", todoCpu[8], cpy);
            dataGridView4.Rows.Add("Velocidad máxima (MHz)", todoCpu[1], cpy);
            dataGridView4.Rows.Add("Arquitectura", todoCpu[6], cpy);
            dataGridView4.Rows.Add("Cantidad de núcleos", System.Environment.ProcessorCount, cpy);
            dataGridView4.Rows.Add("Estado", todoCpu[4], cpy);
            dataGridView4.Rows.Add("Soporte para administrar la energía", todoCpu[5], cpy);
            dataGridView4.Rows.Add("Capacidad de voltaje", todoCpu[7], cpy);

            // Memorias
            dataGridView5.Rows.Add("Tamaño total del disco en bytes", todoAlm[0], cpy);
            dataGridView5.Rows.Add("Particiones detectadas en disco duro principal", todoAlm[1], cpy);
            dataGridView5.Rows.Add("Cabezales en todos los discos montados", todoAlm[2], cpy);
            dataGridView5.Rows.Add("Pistas", todoAlm[3], cpy);
            dataGridView5.Rows.Add("Nùmero de serie", todoAlm[4], cpy);
            dataGridView5.Rows.Add("Pistas por cilindro", todoAlm[5], cpy);
            dataGridView5.Rows.Add("Interfaz", todoAlm[6], cpy);

            // Software
            dataGridView6.Rows.Add("Sistema operativo y compilación", SC.ver() + ", " + System.Environment.OSVersion.Version + ", " + System.Environment.OSVersion.Platform, cpy);
            dataGridView6.Rows.Add("Directorio del sistema", System.Environment.SystemDirectory, cpy);
            dataGridView6.Rows.Add("Límite de opción de arranque", todoSys[4], cpy);
            dataGridView6.Rows.Add("Arquitectura", todoCpu[6], cpy);
            dataGridView6.Rows.Add("Ubicación de la carpeta temporal", Environment.GetEnvironmentVariable("TEMP"), cpy);
            dataGridView6.Rows.Add("Tiempo trancurrido", Environment.TickCount / 3600000 + " horas, " + Environment.TickCount / 60000 + " minutos y " + Environment.TickCount / 1000 + " segundos.", cpy);

            toolStripStatusLabel1.Text = "Última actualización: " + DateTime.UtcNow.ToString("dd/MM/yy hh:mm:ss");
            return true;
        }
        private void acercaDeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new acercade().ShowDialog();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void refrescarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label1.Visible = true;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            Carga();
            label1.Visible = false;
        }

        private void licenciaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.apache.org/licenses/LICENSE-2.0.html");
        }

        private void siempreVisibleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!siempreVisibleToolStripMenuItem.Checked)
            {
                this.TopMost = true;
                siempreVisibleToolStripMenuItem.Checked = true;
            }
            else
            {
                this.TopMost = false;
                siempreVisibleToolStripMenuItem.Checked = false;
            }
        }

        private void exportarListaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog exportarTXT = new SaveFileDialog();
            exportarTXT.Title = "Elige dónde vas a guardar el archivo";
            exportarTXT.Filter = "Archivo de texto sin formato|*.txt";
            if (exportarTXT.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (exportarTXT)
                    {
                        string dataTXT = "";
                            dataTXT += " // Básico" + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView1.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView1.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                   dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // Sistema operativo" + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView2.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView2.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // Componentes en general" + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView3.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView3.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // CPU" + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView4.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView4.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // Memorias" + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView5.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView5.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // Software" + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView6.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView6.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + "Exportado de Syinfo - " + DateTime.UtcNow.ToString("dd/MM/yy hh:mm:ss");;
                        FileStream fsTXT = File.Create(exportarTXT.FileName);
                        StreamWriter guardarTXT = new StreamWriter(fsTXT, Encoding.GetEncoding("iso-8859-1"));
                        guardarTXT.Write(dataTXT);
                        guardarTXT.Close();
                        MessageBox.Show("Se ha guardado el archivo.", "Listo.", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show("Ha ocurrido un error al guardar el archivo. Es posible que el programa no tenga permisos para hacerlo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
        
        private void copiarSeleccionadoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    Clipboard.SetDataObject(this.dataGridView1.GetClipboardContent());
                    break;
                case 1:
                    Clipboard.SetDataObject(this.dataGridView2.GetClipboardContent());
                    break;
                case 2:
                    Clipboard.SetDataObject(this.dataGridView3.GetClipboardContent());
                    break;
                case 3:
                    Clipboard.SetDataObject(this.dataGridView4.GetClipboardContent());
                    break;
                case 4:
                    Clipboard.SetDataObject(this.dataGridView5.GetClipboardContent());
                    break;
                case 5:
                    Clipboard.SetDataObject(this.dataGridView6.GetClipboardContent());
                    break;
            }
        }

        private void repositorioToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://github.com/elstef41/syinfo");
        }
        private void básicoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (básicoToolStripMenuItem.Checked)
            {
                tabControl1.TabPages.Remove(tbBasic);
                básicoToolStripMenuItem.Checked = false;
            }
            else
            {
                tabControl1.TabPages.Add(tbBasic);
                básicoToolStripMenuItem.Checked = true;
            }
        }

        private void sistemaOperativoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (sistemaOperativoToolStripMenuItem.Checked)
            {
                tabControl1.TabPages.Remove(tbOS);
                sistemaOperativoToolStripMenuItem.Checked = false;
            }
            else
            {
                tabControl1.TabPages.Add(tbOS);
                sistemaOperativoToolStripMenuItem.Checked = true;
            }
        }

        private void componentesEnGeneralToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (componentesEnGeneralToolStripMenuItem.Checked)
            {
                tabControl1.TabPages.Remove(tbCompint);
                componentesEnGeneralToolStripMenuItem.Checked = false;
            }
            else
            {
                tabControl1.TabPages.Add(tbCompint);
                componentesEnGeneralToolStripMenuItem.Checked = true;
            }
        }
        private void procesadorToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (procesadorToolStripMenuItem.Checked)
            {
                tabControl1.TabPages.Remove(tbProc);
                procesadorToolStripMenuItem.Checked = false;
            }
            else
            {
                tabControl1.TabPages.Add(tbProc);
                procesadorToolStripMenuItem.Checked = true;
            }
        }

        private void memoriasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (memoriasToolStripMenuItem.Checked)
            {
                tabControl1.TabPages.Remove(tbMem);
                memoriasToolStripMenuItem.Checked = false;
            }
            else
            {
                tabControl1.TabPages.Add(tbMem);
                memoriasToolStripMenuItem.Checked = true;
            }
        }

        private void softwareToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (softwareToolStripMenuItem.Checked)
            {
                tabControl1.TabPages.Remove(tbSoftware);
                softwareToolStripMenuItem.Checked = false;
            }
            else
            {
                tabControl1.TabPages.Add(tbSoftware);
                softwareToolStripMenuItem.Checked = true;
            }
        }

        private void reorganizarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            tabControl1.TabPages.Remove(tbBasic);
            tabControl1.TabPages.Remove(tbCompint);
            tabControl1.TabPages.Remove(tbMem);
            tabControl1.TabPages.Remove(tbOS);
            tabControl1.TabPages.Remove(tbProc);
            tabControl1.TabPages.Remove(tbSoftware);
            tabControl1.TabPages.Add(tbBasic);
            básicoToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbCompint);
            componentesEnGeneralToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbMem);
            memoriasToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbOS);
            sistemaOperativoToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbProc);
            procesadorToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbSoftware);
            softwareToolStripMenuItem.Checked = true;
        }

        private void copiarTodoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StringBuilder cb = new StringBuilder();
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    foreach (DataGridViewRow Row in dataGridView1.Rows)
                    {
                        foreach (DataGridViewColumn Column in dataGridView1.Columns)
                        {
                            cb.Append(Row.Cells[Column.Index].FormattedValue.ToString() + " | ");
                        }
                        cb.AppendLine();
                    }
                    break;
                case 1:
                    foreach (DataGridViewRow Row in dataGridView2.Rows)
                    {
                        foreach (DataGridViewColumn Column in dataGridView2.Columns)
                        {
                            cb.Append(Row.Cells[Column.Index].FormattedValue.ToString() + " | ");
                        }
                        cb.AppendLine();
                    }
                    break;
                case 2:
                    foreach (DataGridViewRow Row in dataGridView3.Rows)
                    {
                        foreach (DataGridViewColumn Column in dataGridView3.Columns)
                        {
                            cb.Append(Row.Cells[Column.Index].FormattedValue.ToString() + " | ");
                        }
                        cb.AppendLine();
                    }
                    break;
                case 3:
                    foreach (DataGridViewRow Row in dataGridView4.Rows)
                    {
                        foreach (DataGridViewColumn Column in dataGridView4.Columns)
                        {
                            cb.Append(Row.Cells[Column.Index].FormattedValue.ToString() + " | ");
                        }
                        cb.AppendLine();
                    }
                    break;
                case 4:
                    foreach (DataGridViewRow Row in dataGridView5.Rows)
                    {
                        foreach (DataGridViewColumn Column in dataGridView5.Columns)
                        {
                            cb.Append(Row.Cells[Column.Index].FormattedValue.ToString() + " | ");
                        }
                        cb.AppendLine();
                    }
                    break;
                case 5:
                    foreach (DataGridViewRow Row in dataGridView6.Rows)
                    {
                        foreach (DataGridViewColumn Column in dataGridView6.Columns)
                        {
                            cb.Append(Row.Cells[Column.Index].FormattedValue.ToString() + " | ");
                        }
                        cb.AppendLine();
                    }
                    break;
            }
            Clipboard.SetText(cb.ToString());
        }

    }
}

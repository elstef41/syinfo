using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Collections;
using System.Management;
using System.Globalization;
using System.Resources;
using Microsoft.VisualBasic;
using System.IO;

namespace syinfo
{
    public partial class syinfo : Form
    {
        // Variables
        strings SC = new strings();
        public static ResourceManager rm = new ResourceManager(typeof(syinfo));
        public static int ancho_ventana = 586;
        public static int alto_ventana = 473;
        List<string> todoCpu = new List<string>();
        List<string> todoSys = new List<string>();
        List<string> todoPhy = new List<string>();
        List<string> todoAlm = new List<string>();

        public syinfo()
        {
            SC.ver();
            InitializeComponent();
            this.Text = "Syinfo ";
            this.Text += SC.obtenerVersion();
            this.Text += " por elstef41";
            this.MinimumSize = new Size(270, 268);
            DataSet ds = new DataSet();
            DataTable dt = new DataTable();
            carga();
            labelInfoLoad.Visible = false;
        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }

        public bool carga()
        {
            todoCpu.Clear();
            todoSys.Clear();
            todoPhy.Clear();
            todoAlm.Clear();
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
                        todoCpu.Add(rm.GetString("descSi"));
                        break;
                    default:
                        todoCpu.Add(rm.GetString("descNo"));
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
                        todoCpu.Add(rm.GetString("descDesconocida"));
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
                        todoCpu.Add(rm.GetString("descDesconocida"));
                        break;
                }
                todoCpu.Add(m["CurrentClockSpeed"].ToString());
                todoCpu.Add(m["Manufacturer"].ToString());
            }

            foreach (ManagementObject m in sys)
            {
                todoSys.Add(m["BootupState"].ToString());
                switch (m["BootROMSupported"].ToString()) {
                    case "True":
                        todoSys.Add(rm.GetString("descSi"));
                        break;
                    default:
                        todoSys.Add(rm.GetString("descNo"));
                        break;
                }
                todoSys.Add(m["Workgroup"].ToString());
                switch (m["ThermalState"].ToString())
                {
                    case "1":
                        todoSys.Add(rm.GetString("descOtro"));
                        break;
                    case "3":
                        todoSys.Add(rm.GetString("descSeguro"));
                        break;
                    case "4":
                        todoSys.Add(rm.GetString("descAdvertencia"));
                        break;
                    case "5":
                        todoSys.Add(rm.GetString("descCritico"));
                        break;
                    case "6":
                        todoSys.Add(rm.GetString("descIrrecuperable"));
                        break;
                    case "2":
                    default:
                        todoSys.Add(rm.GetString("descDesconocido"));
                        break;
                }
                switch (m["AutomaticResetBootOption"].ToString())
                {
                    case "True":
                        todoSys.Add(rm.GetString("descSi"));
                        break;
                    case "False":
                        todoSys.Add(rm.GetString("descNo"));
                        break;
                    default:
                        todoSys.Add(rm.GetString("descDesconocido"));
                        break;
                }
            }

            foreach (ManagementObject m in phy)
            {
                todoPhy.Add(m["Status"].ToString());
                switch (m["SupportsHotPlug"].ToString())
                {
                    case "True":
                        todoPhy.Add(rm.GetString("descSi"));
                        break;
                    default:
                        todoPhy.Add(rm.GetString("descNo"));
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


            // Básico
            dataGridView1.Rows.Add(rm.GetString("descCompilacion"), System.Environment.OSVersion.Version);
            dataGridView1.Rows.Add(rm.GetString("descNombreDelEquipo"), System.Environment.MachineName);
            dataGridView1.Rows.Add(rm.GetString("descArquitectura"), todoCpu[6]);
            dataGridView1.Rows.Add(rm.GetString("descNombreDeUsuarioActivo"), System.Environment.UserName);
            dataGridView1.Rows.Add(rm.GetString("descGrupoDeTrabajo"), todoSys[2]);
            dataGridView1.Rows.Add(rm.GetString("descDirectorioDelSistema"), System.Environment.SystemDirectory);
            dataGridView1.Rows.Add(rm.GetString("descCantidadDeNucleos"), System.Environment.ProcessorCount);
            dataGridView1.Rows.Add(rm.GetString("descTiempoTranscurrido"), Environment.TickCount / 3600000 + " horas, " + Environment.TickCount / 60000 + " minutos y " + Environment.TickCount / 1000 + " segundos.");

            // Sistema operativo
            dataGridView2.Rows.Add(rm.GetString("descSistemaOperativo"), SC.ver());
            dataGridView2.Rows.Add(rm.GetString("descVersionCompletadeWindowsNT"), System.Environment.OSVersion.VersionString);
            dataGridView2.Rows.Add(rm.GetString("descVersionDelNucleo"), System.Environment.OSVersion.Version);
            dataGridView2.Rows.Add(rm.GetString("descCompilacion"), System.Environment.OSVersion.Version.Build);
            dataGridView2.Rows.Add(rm.GetString("descPlataforma"), System.Environment.OSVersion.Platform);
            if (Environment.OSVersion.ServicePack == "")
            {
                dataGridView2.Rows.Add("Service Pack", "N/A");
            }
            else
            {
                dataGridView2.Rows.Add("Service Pack", Environment.OSVersion.ServicePack);
            }
            dataGridView2.Rows.Add(rm.GetString("descGrupoDeTrabajo"), todoSys[2]);

            // Componentes internos
            dataGridView3.Rows.Add(rm.GetString("descMemoriaRAM"), System.Environment.Version.Major + "." + Environment.Version.MajorRevision + "." + Environment.Version.Minor + "." + Environment.Version.MinorRevision);
            dataGridView3.Rows.Add(rm.GetString("descNombreDelProcesador"), todoCpu[2]);
            dataGridView3.Rows.Add(rm.GetString("descIDDelProcesador"), todoCpu[3]);
            dataGridView3.Rows.Add(rm.GetString("descDireccionDelProcesador"), todoCpu[0]);
            dataGridView3.Rows.Add(rm.GetString("descCantidadDeNucleos"), System.Environment.ProcessorCount);
            dataGridView3.Rows.Add(rm.GetString("descEstadoDeArranque"), todoSys[0]);
            dataGridView3.Rows.Add(rm.GetString("descSoporteParaMemoriaROM"), todoSys[1]);
            dataGridView3.Rows.Add(rm.GetString("descEstadoTermico"), todoSys[3]);
            dataGridView3.Rows.Add(rm.GetString("descSoporteParaCambioCaliente"), todoPhy[1]);
            dataGridView3.Rows.Add("BUS", todoPhy[2]);

            // CPU
            dataGridView4.Rows.Add("ID", todoCpu[3]);
            dataGridView4.Rows.Add(rm.GetString("descNombre"), todoCpu[2]);
            dataGridView4.Rows.Add(rm.GetString("descFabricante"), todoCpu[9]);
            dataGridView4.Rows.Add(rm.GetString("descDireccion"), todoCpu[0]);
            dataGridView4.Rows.Add(rm.GetString("descVelocidadMhz"), todoCpu[8]);
            dataGridView4.Rows.Add(rm.GetString("descVelocidadMinimaMhz"), todoCpu[1]);
            dataGridView4.Rows.Add(rm.GetString("descArquitectura"), todoCpu[6]);
            dataGridView4.Rows.Add(rm.GetString("descCantidadDeNucleos"), System.Environment.ProcessorCount);
            dataGridView4.Rows.Add(rm.GetString("descEstado"), todoCpu[4]);
            dataGridView4.Rows.Add(rm.GetString("descSoporteParaAdministrarEnergia"), todoCpu[5]);
            dataGridView4.Rows.Add(rm.GetString("descCapacidadDeVoltaje"), todoCpu[7]);

            // Memorias
            dataGridView5.Rows.Add(rm.GetString("descTamanoTotalDelDiscoEnBytes"), todoAlm[0]);
            dataGridView5.Rows.Add(rm.GetString("descParticionesDetectadasEnTodosLosDiscsDuros"), todoAlm[1]);
            dataGridView5.Rows.Add(rm.GetString("descCabezalesEnTodosLosDiscosMontados"), todoAlm[2]);
            dataGridView5.Rows.Add(rm.GetString("descPistas"), todoAlm[3]);
            dataGridView5.Rows.Add(rm.GetString("descNumeroDeSerie"), todoAlm[4]);
            dataGridView5.Rows.Add(rm.GetString("descPistasPorCilindro"), todoAlm[5]);
            dataGridView5.Rows.Add(rm.GetString("descInterfaz"), todoAlm[6]);

            // Software
            dataGridView6.Rows.Add(rm.GetString("descSistemaOperativoYCompilacion"), SC.ver() + ", " + System.Environment.OSVersion.Version + ", " + System.Environment.OSVersion.Platform);
            dataGridView6.Rows.Add(rm.GetString("descDirectorioDelSistema"), System.Environment.SystemDirectory);
            dataGridView6.Rows.Add(rm.GetString("descLimiteDeOpcionDeArranque"), todoSys[4]);
            dataGridView6.Rows.Add(rm.GetString("descArquitectura"), todoCpu[6]);
            dataGridView6.Rows.Add(rm.GetString("descUbicacionDeLaCarpetaTemporal"), Environment.GetEnvironmentVariable("TEMP"));
            dataGridView6.Rows.Add(rm.GetString("descTiempoTranscurrido"), Environment.TickCount / 3600000 + rm.GetString("s_tiempo1") + Environment.TickCount / 60000 + rm.GetString("s_tiempo2") + Environment.TickCount / 1000 + rm.GetString("s_tiempo3"));

            // Variables de entorno
            var EnV = Environment.GetEnvironmentVariables();
            foreach (DictionaryEntry i in EnV)
            {
                dataGridView7.Rows.Add(i.Key, i.Value);
            }

            toolStripStatusLabel1.Text = rm.GetString("s_ultima_actualización") + DateTime.Now.ToString("dd/MM/yy HH:MM:ss");
            return true;
        }
        private void acercaDeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new acercade().ShowDialog();
        }

        public void refrescar()
        {
            labelInfoLoad.Visible = true;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();
            dataGridView7.Rows.Clear();
            carga();
            labelInfoLoad.Visible = false;
        }

        private void refrescarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            refrescar();
        }

        private void licenciaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.apache.org/licenses/LICENSE-2.0.html");
        }


        private void exportarListaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog exportarTXT = new SaveFileDialog();
            exportarTXT.Title = rm.GetString("s_exportar");
            exportarTXT.Filter = "txt|*.txt";
            if (exportarTXT.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (exportarTXT)
                    {
                        string dataTXT = "";
                                dataTXT += " // " + rm.GetString("tbBasic.Text") + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView1.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView1.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                   dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // " + rm.GetString("tbOS.Text") + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView2.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView2.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // " + rm.GetString("tbCompint.Text") + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView3.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView3.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // " + rm.GetString("tbProc.Text") + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView4.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView4.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // " + rm.GetString("tbMem.Text") + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView5.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView5.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // " + rm.GetString("tbSoftware.Text") + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView6.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView6.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + " // " + rm.GetString("tbEV.Text") + Environment.NewLine;
                                foreach (DataGridViewRow Row in dataGridView7.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView7.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                dataTXT += Environment.NewLine + Environment.NewLine + rm.GetString("s_exportado") + DateTime.UtcNow.ToString("dd/MM/yy HH:mm:ss"); ;
                        FileStream fsTXT = File.Create(exportarTXT.FileName);
                        StreamWriter guardarTXT = new StreamWriter(fsTXT, Encoding.GetEncoding("iso-8859-1"));
                        guardarTXT.Write(dataTXT);
                        guardarTXT.Close();
                        MessageBox.Show(rm.GetString("s_guardado"), "Syinfo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show(rm.GetString("s_error_guardado"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
            tabControl1.TabPages.Remove(tbEV);
            tabControl1.TabPages.Add(tbBasic);
            básicoToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbOS);
            sistemaOperativoToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbCompint);
            componentesEnGeneralToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbProc);
            procesadorToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbMem);
            memoriasToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbSoftware);
            softwareToolStripMenuItem.Checked = true;
            tabControl1.TabPages.Add(tbEV);
            variablesDeEntornoToolStripMenuItem.Checked = true;
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
                case 6:
                    foreach (DataGridViewRow Row in dataGridView7.Rows)
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

        private void exportarListaVisibleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog exportarTXT = new SaveFileDialog();
            exportarTXT.Title = rm.GetString("s_exportar");
            exportarTXT.Filter = ".txt|*.txt";
            if (exportarTXT.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (exportarTXT)
                    {
                        string dataTXT = "";
                        switch (tabControl1.SelectedIndex)
                        {
                            case 0:
                                foreach (DataGridViewRow Row in dataGridView1.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView4.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                break;
                            case 1:
                                foreach (DataGridViewRow Row in dataGridView2.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView4.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                break;
                            case 2:
                                foreach (DataGridViewRow Row in dataGridView3.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView4.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                break;
                            case 3:
                                foreach (DataGridViewRow Row in dataGridView4.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView4.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                break;
                            case 4:
                                foreach (DataGridViewRow Row in dataGridView5.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView5.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                break;
                            case 5:
                                foreach (DataGridViewRow Row in dataGridView6.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView6.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                break;
                            case 6:
                                foreach (DataGridViewRow Row in dataGridView7.Rows)
                                {
                                    foreach (DataGridViewColumn Column in dataGridView7.Columns)
                                    {
                                        if (Row.Cells[Column.Index].FormattedValue.ToString() != "") { dataTXT += Row.Cells[Column.Index].FormattedValue.ToString() + " | "; }
                                    }
                                    dataTXT += Environment.NewLine;
                                }
                                break;
                        }
                        dataTXT += Environment.NewLine + Environment.NewLine + rm.GetString("s_exportado") + DateTime.UtcNow.ToString("dd/MM/yy HH:mm:ss"); ;
                        FileStream fsTXT = File.Create(exportarTXT.FileName);
                        StreamWriter guardarTXT = new StreamWriter(fsTXT, Encoding.GetEncoding("iso-8859-1"));
                        guardarTXT.Write(dataTXT);
                        guardarTXT.Close();
                        MessageBox.Show(rm.GetString("s_guardado"), "Syinfo", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (Exception)
                {
                    MessageBox.Show(rm.GetString("s_error_guardado"), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void variablesDeEntornoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (variablesDeEntornoToolStripMenuItem.Checked)
            {
                tabControl1.TabPages.Remove(tbEV);
                variablesDeEntornoToolStripMenuItem.Checked = false;
            }
            else
            {
                tabControl1.TabPages.Add(tbEV);
                variablesDeEntornoToolStripMenuItem.Checked = true;
            }
        }

        private void toolStripStatusLabel2_Click(object sender, EventArgs e)
        {
            labelInfoLoad.Visible = true;
            dataGridView1.Rows.Clear();
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            dataGridView4.Rows.Clear();
            dataGridView5.Rows.Clear();
            dataGridView6.Rows.Clear();
            dataGridView7.Rows.Clear();
            carga();
            labelInfoLoad.Visible = false;
        }

        private void restaurarTamañoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Size = new Size(ancho_ventana, alto_ventana);
            restaurarTamañoToolStripMenuItem.Enabled = false;
        }

        private void siempreVisibleToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (!siempreVisibleToolStripMenuItem1.Checked)
            {
                this.TopMost = true;
                siempreVisibleToolStripMenuItem1.Checked = true;
            }
            else
            {
                this.TopMost = false;
                siempreVisibleToolStripMenuItem1.Checked = false;
            }
        }

        private void syinfo_SizeChanged(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Maximized)
            {
                restaurarTamañoToolStripMenuItem.Enabled = false;

            }
            else if (this.Size.Width != ancho_ventana || this.Size.Height != alto_ventana)
            {
                restaurarTamañoToolStripMenuItem.Enabled = true;
            }
            else
            {
                restaurarTamañoToolStripMenuItem.Enabled = false;
            }
        }

        private void CincoMinTs_Click(object sender, EventArgs e)
        {
            if (!CincoMinTs.Checked)
            {
                timer.Interval = SC.minAMs(5);
                timer.Start();
                timerIconStatus.Visible = true;
                CincoMinTs.Checked = true;
                DiezMinTs.Checked = false;
                TreintaMinTs.Checked = false;
                refrescarCadaToolStripMenuItem.Checked = true;
            }
            else
            {
                timer.Stop();
                timerIconStatus.Visible = false;
                CincoMinTs.Checked = false;
                refrescarCadaToolStripMenuItem.Checked = false;
            }
        }

        private void DiezMinTs_Click(object sender, EventArgs e)
        {
            if (!DiezMinTs.Checked)
            {
                timer.Interval = SC.minAMs(10);
                timer.Start();
                timerIconStatus.Visible = true;
                CincoMinTs.Checked = false;
                DiezMinTs.Checked = true;
                TreintaMinTs.Checked = false;
                refrescarCadaToolStripMenuItem.Checked = true;
            }
            else
            {
                timer.Stop();
                timerIconStatus.Visible = false;
                DiezMinTs.Checked = false;
                refrescarCadaToolStripMenuItem.Checked = false;
            }
        }

        private void TreintaMinTs_Click(object sender, EventArgs e)
        {
            if (!TreintaMinTs.Checked)
            {
                timer.Interval = SC.minAMs(30);
                timer.Start();
                timerIconStatus.Visible = true;
                CincoMinTs.Checked = false;
                DiezMinTs.Checked = false;
                TreintaMinTs.Checked = true;
                refrescarCadaToolStripMenuItem.Checked = true;
            }
            else
            {
                timer.Stop();
                timerIconStatus.Visible = false;
                TreintaMinTs.Checked = false;
                refrescarCadaToolStripMenuItem.Checked = false;
            }
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            refrescar();
        }
    }
}

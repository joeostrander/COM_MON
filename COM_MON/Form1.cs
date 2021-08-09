using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

/*
 * TODO:  
 * 
 */


namespace COM_MON
{
    public partial class Form1 : Form
    {

        private bool boolStartup = false;

        //private static List<KeyValuePair<string, string>> listPorts = new List<KeyValuePair<string, string>>();

        private struct myPort
        {
            public DateTime DateAdded { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
        }

        private static BindingList<myPort> listPorts = new BindingList<myPort>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = Application.ProductName;

            notifyIcon1.Text = Application.ProductName;
            LoadRegistrySettings();

            dataGridView1.DataSource = listPorts;
            dataGridView1.Columns[0].DefaultCellStyle.Format = "yyyy.MM.dd HH:mm:ss";
            //dataGridView1.Columns[0].SortMode = DataGridViewColumnSortMode.Automatic;
            

            Get_Com_Ports();
        }

        private void Get_Com_Ports()
        {
            string strQuery = "Select * from Win32_PnPEntity Where Name LIKE '% (COM%)'";

            List<string> list_current_ports = new List<string>();

            try
            {

                ManagementObjectSearcher searcher = new ManagementObjectSearcher(strQuery);
                ManagementObjectCollection collection = searcher.Get();


                foreach (ManagementObject item in collection)
                {
                    string fn = item.Properties["Name"].Value.ToString();
                    string strPattern = @"^(?<porttext>.*?)\((?<portname>COM[\d]+)\)$";
                    Match m = Regex.Match(fn, strPattern);

                    if (m.Success)
                    {
                        string portname = m.Groups["portname"].Value.ToString();
                        string porttext = m.Groups["porttext"].Value.ToString();
                        // var kvp = new KeyValuePair<string, string>(portname, portname.PadRight(6, ' ') + porttext);
                        list_current_ports.Add(portname);

                        bool boolExists = false;
                        foreach (myPort port in listPorts)
                        {
                            if (port.Name == portname)
                            {
                                boolExists = true;
                                break;
                            }
                        }
                        if (!boolExists)
                        {
                            myPort newPort = new myPort();
                            newPort.DateAdded = DateTime.Now;
                            newPort.Name = portname;
                            newPort.Description = porttext;
                            //listPorts.Add(newPort);
                            listPorts.Insert(0, newPort);
                        }
                    }


                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not get COM port list!", "COM Ports", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // look to the table to see what can be deleted
            for (int i = listPorts.Count - 1; i >= 0; i--)
            {
                Console.WriteLine("Check for: {0}...{1}", listPorts[i].Name, list_current_ports.Contains(listPorts[i].Name));
                if (!list_current_ports.Contains(listPorts[i].Name))
                {
                    Console.WriteLine("REMOVED:  {0}", listPorts[i].Name);
                    listPorts.RemoveAt(i);
                }
            }



        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Get_Com_Ports();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutBox1 about = new AboutBox1();
            about.ShowDialog();
        }

        private void runAtStartupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            boolStartup = runAtStartupToolStripMenuItem.Checked;
            SaveRegistrySettings();
        }


        private bool SaveRegistrySettings()
        {
            try
            {
                RegistryKey regKey_RUN;
                regKey_RUN = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);

                RegistryKey appRegKey;
                appRegKey = Registry.CurrentUser.OpenSubKey("Software\\" + Application.ProductName, true);

                //Startup 
                appRegKey.SetValue("RunAtStartup", boolStartup, RegistryValueKind.DWord);

                //if boolStartup..., add executable path to the Run registry key
                if (boolStartup)
                {
                    regKey_RUN.SetValue(Application.ProductName, Application.ExecutablePath);
                }
                else
                {
                    if (regKey_RUN.GetValueNames().Contains(Application.ProductName))
                    {
                        regKey_RUN.DeleteValue(Application.ProductName);
                    }

                }

                regKey_RUN.Close();
                appRegKey.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exeption in SaveRegistrySettings()");
                Console.WriteLine(ex.Message);
                return false;
            }

            return true;
        }

        private void LoadRegistrySettings()
        {
            RegistryKey regKey_RUN;
            regKey_RUN = Registry.CurrentUser.OpenSubKey("Software\\Microsoft\\Windows\\CurrentVersion\\Run", true);

            RegistryKey appRegKey;
            appRegKey = Registry.CurrentUser.OpenSubKey("Software\\" + Application.ProductName, true);

            if (appRegKey == null)
            {
                // Create the key
                Registry.CurrentUser.CreateSubKey("Software\\" + Application.ProductName);
                appRegKey = Registry.CurrentUser.OpenSubKey("Software\\" + Application.ProductName, true);
                if (appRegKey == null)
                {
                    regKey_RUN.Close();
                    return;
                }

                //Ask user if they want it to launch auto...
                if (MessageBox.Show("Launch " + Application.ProductName + " at Startup?", Application.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    boolStartup = true;
                }
                else
                {
                    boolStartup = false;
                }
            }
            else
            {
                //Load current Startup Value and set it in the interface
                boolStartup = (int)appRegKey.GetValue("RunAtStartup", 0) == 1 ? true : false;

            }

            regKey_RUN.Close();
            appRegKey.Close();

            SaveRegistrySettings();

        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            this.Text = Application.ProductName;
            if (this.WindowState == FormWindowState.Minimized)
            {
                HideMe();
            }
        }

        private void HideMe()
        {
            this.Hide();
            notifyIcon1.ShowBalloonTip(2000, Application.ProductName, "Click to show", ToolTipIcon.Info);
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void notifyIcon1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ShowMe();
            }
        }

        private void ShowMe()
        {
            this.Show();
            this.WindowState = FormWindowState.Normal;
            this.Activate();
            if (this.Location.X < 0 || this.Location.Y < 0) { this.CenterToScreen(); }
        }

        private void Form1_Move(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                HideMe();
            }
            else
            {
                ShowMe();
            }
        }

        private void notifyIcon1_BalloonTipClicked(object sender, EventArgs e)
        {
            ShowMe();
        }
    }


}

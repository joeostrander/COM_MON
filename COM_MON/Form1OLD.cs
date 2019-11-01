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

namespace COM_MON
{
    public partial class Form1 : Form
    {
        //private static List<KeyValuePair<string, string>> listPorts = new List<KeyValuePair<string, string>>();

        private struct myPort
        {
            public DateTime DateAdded { get; set; }
            public string Name { get; set; }
            public string Description { get; set; }
        }

        private static List<myPort> listPorts = new List<myPort>();

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = Application.ProductName;

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
                            listPorts.Add(newPort);
                        }
                    }


                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not get COM port list!", "COM Ports", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // look to the table to see what can be deleted?
            for (int i = listPorts.Count - 1; i >= 0; i--)
            {
                if (!list_current_ports.Contains(listPorts[i].Name))
                {
                    listPorts.RemoveAt(i);
                    Console.WriteLine("REMOVED:  {0}", listPorts[i].Name);
                }
            }




            dataGridView1.DataSource = listPorts;
            
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            Get_Com_Ports();
        }
    }
}

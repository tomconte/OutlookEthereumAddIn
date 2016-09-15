using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookEthereumAddIn
{
    public partial class SettingsDialog : Form
    {
        public SettingsDialog()
        {
            InitializeComponent();
            this.textBox1.Text = Properties.Settings.Default.Host;
            this.textBox2.Text = Properties.Settings.Default.Account;
            this.textBox3.Text = Properties.Settings.Default.Password;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.Host = this.textBox1.Text;
            Properties.Settings.Default.Account = this.textBox2.Text;
            Properties.Settings.Default.Password = this.textBox3.Text;
            Properties.Settings.Default.Save();

            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

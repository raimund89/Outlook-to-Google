using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace OutlookToGoogle
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Activated(object sender, EventArgs e)
        {
            this.label2.Text = Program.ics.GetVersion();
            this.label4.Text = Program.ics.GetDefaultProfile();
            this.textBox1.Text = Properties.Settings.Default.icsPath;
            this.textBox2.Text = Properties.Settings.Default.icsName;

            this.comboBox1.Items.Clear();
            foreach (KeyValuePair<int, string> item in Program.Intervals)
            {
                this.comboBox1.Items.Add(item.Value);
            }
            this.comboBox1.SelectedIndex = Properties.Settings.Default.updateFreq;

            this.checkBox1.Checked = Properties.Settings.Default.startWithSystem;
            this.checkBox2.Checked = Properties.Settings.Default.notifyOnChange;
        }

        private void BtnFile_clicked(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = Properties.Settings.Default.icsPath;
            fbd.Description = "Select the folder where the calendar file will be stored. If the file exists, it will be overwritten.";

            DialogResult result = fbd.ShowDialog();

            if(result == DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
            {
                this.textBox1.Text = fbd.SelectedPath;
            }
        }

        private void BtnCancel_clicked(object sender, EventArgs e)
        {
            this.Close();
        }

        private void BtnOK_clicked(object sender, EventArgs e)
        {
            // First, when user doesn't want it to start and that wasn't the case earlier, ask them to be sure
            if(Properties.Settings.Default.startWithSystem != this.checkBox1.Checked && !Properties.Settings.Default.startWithSystem)
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure you don't want OutlookToICS\nto start with Windows?", "Are you sure?", MessageBoxButtons.YesNo);

                // Apparently the user made a mistake, so don't just close but let them look at it again.
                if (dialogResult == DialogResult.No)
                    return;
            }

            // Save all settings
            Properties.Settings.Default.icsPath = this.textBox1.Text;
            Properties.Settings.Default.icsName = this.textBox2.Text;

            if(this.comboBox1.SelectedIndex != Properties.Settings.Default.updateFreq)
            {
                Properties.Settings.Default.updateFreq = this.comboBox1.SelectedIndex;
                Properties.Settings.Default.Save();
                Program.InitializeTimer();
            }
            
            Properties.Settings.Default.startWithSystem = this.checkBox1.Checked;
            Program.ToggleStartup(this.checkBox1.Checked);

            Properties.Settings.Default.notifyOnChange = this.checkBox2.Checked;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}

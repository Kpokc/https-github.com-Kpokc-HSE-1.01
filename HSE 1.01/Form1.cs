using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HSE_1._01
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Open_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.Multiselect = true;

            DialogResult dr = openFileDialog1.ShowDialog();

            if (dr == System.Windows.Forms.DialogResult.OK)
            {
                string[] locationArray = openFileDialog1.FileNames;

                openFiles passFilePaths = new openFiles();
                //string value = textBox1.Text;
                passFilePaths.openFile(locationArray);
            }
        }

        private void Close_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        public void sendMessage(string msg)
        {

            MessageBox.Show(msg);
        }
    }
}

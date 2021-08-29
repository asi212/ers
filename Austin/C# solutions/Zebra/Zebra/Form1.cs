using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Zebra
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string LF_num = textBox1.Text;
            string num_copies = comboBox3.Text;
            string label_type = comboBox2.Text;

            Read_excel rdxl = new Read_excel();
            rdxl.read_excel(LF_num, num_copies, label_type);
            // Add code to execute print job here
            
            //MessageBox.Show(LF_num);
            //MessageBox.Show(num_copies);
            //MessageBox.Show(label_type);
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}


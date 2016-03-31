using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OdsConvert
{
    public partial class Form1 : Form
    {
        public string readFilePath;
        public string writeFilePath;
        public ConvertToXml convert2xml;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog choofdlog = new OpenFileDialog();
            choofdlog.Filter = "Open Document Spreadsheet (.ods)|*.ods";
            choofdlog.FilterIndex = 1;
            choofdlog.Multiselect = false;

            if (choofdlog.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = choofdlog.FileName;
                readFilePath = choofdlog.FileName;
                this.button2.Enabled = true;
                this.textBox2.Text = choofdlog.FileName.Replace(".ods",".xml");
                writeFilePath = choofdlog.FileName.Replace(".ods", ".xml");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            convert2xml = new ConvertToXml();
            convert2xml.ReadOdsFile(readFilePath);
            convert2xml.WriteXmlFile(writeFilePath);
        }
    }
}

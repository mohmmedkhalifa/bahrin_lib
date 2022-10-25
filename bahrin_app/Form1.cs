
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace bahrin_app
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            string path = openFileDialog.FileName.ToString();
            ExcelFileReader(path);



        }

        public void ExcelFileReader(string path)
        {
            var stream = File.Open(path, FileMode.Open, FileAccess.Read);
            Cursor.Current = Cursors.WaitCursor;
            var reader = ExcelReaderFactory.CreateReader(stream);
            var result = reader.AsDataSet();
            var tables = result.Tables.Cast<DataTable>();
            foreach(DataTable dt in tables)
            {
                dataGridView.DataSource = dt;
            }
            Cursor.Current = Cursors.Default;
        }

        private void searchBtn_Click(object sender, EventArgs e)
        {
            try
            {
                DataView dv= dataGridView.DataSource as DataView;
                if(dv!=null)
                    dv.RowFilter = textBox1.Text;
                label2.Text = $"النتائج {dataGridView.RowCount-1}";
            }
            catch (Exception ex)
            {
            
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)13)
                searchBtn.PerformClick();
        }

        private void dataGridView_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}

using CsvHelper;
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

namespace Barcode_Creator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        private string file;
        private string dir;

        private void button1_Click(object sender, EventArgs e)
        {
            int SER = 000001; // set serial number
            string jn = textBox1.Text;
            string numpg = textBox2.Text;
            string name = textBox3.Text;
            string pc = textBox4.Text;

            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog2.ShowDialog();

            file = openFileDialog2.FileName;
            dir = Path.GetDirectoryName(file);

            using (var reader = new StreamReader(file)) //load CSV from BBS
            using (var csv = new CsvReader(reader))



            {


                using (var dr = new CsvDataReader(csv))
                {
                    var dt = new System.Data.DataTable();
                    dt.Load(dr);

                    // Create new columns to be appended the start of the datatable
                    DataColumn newCol = new DataColumn("pbBarcode1", typeof(string));


                    // Add new columns for Barcodes
                    dt.Columns.Add(newCol);


                    //Set positon of new columns
                    newCol.SetOrdinal(0);


                    //loop through each row and add data
                    foreach (DataRow row in dt.Rows)
                    {
                        row["pbBarcode1"] = SER.ToString("000") + SER.ToString("000");


                        SER++;

                        //}

                        using (var textWriter = File.CreateText(dir + "\\" + "test.txt"))
                        using (var csv1 = new CsvWriter(textWriter))
                        {
                            // Write columns
                            foreach (DataColumn column in dt.Columns)
                            {
                                csv1.WriteField(column.ColumnName);
                            }
                            csv1.NextRecord();

                            // Write row values
                            foreach (DataRow row1 in dt.Rows)
                            {
                                for (var i = 0; i < dt.Columns.Count; i++)
                                {
                                    csv1.WriteField(row1[i]);
                                }
                                csv1.NextRecord();
                            }
                        }



                    }
                }
            }
        }
    }
}

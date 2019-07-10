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
            Application.DoEvents();
        }


        private string file;
        private string dir;

        private void button1_Click(object sender, EventArgs e)
        {
            long SER = 000001; // set serial number
            long jn = int.Parse(textBox1.Text);
            int numpg = int.Parse(textBox2.Text);
            int name = int.Parse(textBox3.Text) + numpg;
            int pc = int.Parse(textBox4.Text) + numpg;

            OpenFileDialog openFileDialog2 = new OpenFileDialog();
            openFileDialog2.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog2.ShowDialog();

            file = openFileDialog2.FileName;
            dir = Path.GetDirectoryName(file);

            string JobID = new string(' ', 10);
            string PieceID = new string(' ', 6);
            string JobType = new string(' ', 32);
            string SourceID = new string(' ', 32);
            var Name = new string(' ', 40);

            // string PagesInput02, PagesInput03, PagesInput04, SubsetInput01, SubsetInput02, SubsetInput03, SubsetInput04,
            // StitchInput01, StitchInput02, StitchInput03, StitchInput04 = new string(' ', 2);

            string InputWeight = new string(' ', 5);

            // int Sel1, Sel2, Sel3, Sel4, Sel5, Sel6, Sel7, Sel8, Sel9, Sel10, Sel11, Sel12, Sel13, Sel14, Sel15, Sel16,
            //AccountPull, QualityAudit, AlertClear, EdgeMark, VS1, VS2, VS3, VS4, VS5, VS6 = 0;

            string PagesInput02 = new string(' ', 2);
            string PagesInput03 = new string(' ', 2);
            string PagesInput04 = new string(' ', 2);
            string SubsetInput01 = new string(' ', 2);
            string SubsetInput02 = new string(' ', 2);
            string SubsetInput03 = new string(' ', 2);
            string SubsetInput04 = new string(' ', 2);
            string StitchInput01 = new string(' ', 2);
            string StitchInput02 = new string(' ', 2);
            string StitchInput03 = new string(' ', 2);
            string StitchInput04 = new string(' ', 2);
            
            int Sel1 = 0;
            int Sel2 = 0;
            int Sel3 = 0;
            int Sel4 = 0;
            int Sel5 = 0;
            int Sel6 = 0;
            int Sel7 = 0;
            int Sel8 = 0;
            int Sel9 = 0;
            int Sel10 = 0;
            int Sel11 = 0;
            int Sel12 = 0;
            int Sel13 = 0;
            int Sel14 = 0;
            int Sel15 = 0;
            int Sel16 = 0;
            int AccountPull = 0;
            int QualityAudit = 0;
            int AlertClear = 0;
            int EdgeMark = 0;
            int VS1 = 0;
            int VS2 = 0;
            int VS3 = 0;
            int VS4 = 0;
            int VS5 = 0;
            int VS6 = 0;


            string AccountIdentifier01 = new string(' ', 40);
            string AccountIdentifier02 = new string(' ', 40);
            string AccountIdentifier03 = new string(' ', 40);
            string AccountIdentifier04 = new string(' ', 40);
            string Address01 = new string(' ', 40);
            string Address02 = new string(' ', 40);
            string Address03 = new string(' ', 40);
            string Address04 = new string(' ', 40);
            string Address05 = new string(' ', 40);
            string Address06 = new string(' ', 40);
            var postcode = new string(' ', 16);
            string PostcodeShort = new string(' ', 16);
            string BRAddress01 = new string(' ', 40);
            string BRAddress02 = new string(' ', 40);
            string BRAddress03 = new string(' ', 40);
            string BRAddress04 = new string(' ', 40);
            string BRAddress05 = new string(' ', 40);
            string ReprintIndex = new string(' ', 30);
            string StartPageOffset = new string(' ', 6);
            string PageCountOffset = new string(' ', 6);
            string UserDefined01 = new string(' ', 40);
            string UserDefined02 = new string(' ', 40);
            string UserDefined03 = new string(' ', 40);
            string UserDefined04 = new string(' ', 39);
            string Filler01 = new string(' ', 2);

            using (var reader = new StreamReader(file)) //load CSV from BBS
            using (var csv = new CsvReader(reader))



            {


                using (var dr = new CsvDataReader(csv))
                {
                    var dt = new DataTable();
                    dt.Load(dr);



                    // Create new columns to be appended the start of the datatable
                    int i = 1;
                    int k = 0;
                    do
                    {
                        DataColumn newCol = new DataColumn("PBbarcode" + i, typeof(string));
                        dt.Columns.Add(newCol);
                        newCol.SetOrdinal(k);

                        i++;
                        k++;

                    } while (i != numpg + 1);

                    DataColumn newCol1 = new DataColumn("SER", typeof(string));
                    dt.Columns.Add(newCol1);
                    newCol1.SetOrdinal(0);


                    // Add new columns for Barcodes
                    // dt.Columns.Add(newCol);


                    //Set positon of new columns
                    //newCol.SetOrdinal(0);


                    //loop through each row and add data
                    string str = new string(' ', 32);
                    string OutDat;


                    using (StreamWriter sw = new StreamWriter(dir + "\\" + jn + " MRDF" + ".txt", append: true))
                    {
                        string header = jn.ToString("0000000000") + str + DateTime.Now;
                        sw.WriteLine(header);
                    }




                    i = 1;
                    int y = 0;
                


                    foreach (DataRow row in dt.Rows)
                    {
                        StringBuilder sb = new StringBuilder(40);


                        Name = dt.Rows[y][name].ToString();
                        Name = Name.PadRight(40, ' ');

                        postcode = dt.Rows[y][pc].ToString();
                        postcode = postcode.PadRight(16, ' ');

                        y++;

                        using (StreamWriter sw = new StreamWriter(dir + "\\" + jn + " MRDF" + ".txt", append: true))
                        {
                            OutDat = jn.ToString("0000000000") + SER.ToString("000000") + JobType + SourceID + numpg.ToString("00") + PagesInput02 + PagesInput03 + PagesInput04;
                            OutDat = OutDat + SubsetInput01 + SubsetInput02 + SubsetInput03 + SubsetInput04 + StitchInput01 + StitchInput02 + StitchInput03 + StitchInput04 + InputWeight;
                            OutDat = OutDat + Sel1 + Sel2 + Sel3 + Sel4 + Sel5 + Sel6 + Sel7 + Sel8 + Sel9 + Sel10 + Sel11 + Sel12 + Sel13 + Sel14 + Sel15 + Sel16;
                            OutDat = OutDat + AccountPull + QualityAudit + AlertClear + EdgeMark + VS1 + VS2 + VS3 + VS4 + VS5 + VS6;
                            OutDat = OutDat + AccountIdentifier01 + AccountIdentifier02 + AccountIdentifier03 + AccountIdentifier04;
                            OutDat = OutDat + Name + Address01 + Address02 + Address03 + Address04 + Address05 + Address06 + postcode + PostcodeShort;
                            OutDat = OutDat + BRAddress01 + BRAddress02 + BRAddress03 + BRAddress04 + BRAddress05;
                            OutDat = OutDat + ReprintIndex + StartPageOffset + PageCountOffset;
                            OutDat = OutDat + UserDefined01 + UserDefined02 + UserDefined03 + UserDefined04 + Filler01;
                            sw.WriteLine(OutDat);


                        }

                        for (i = 1; i < numpg + 1; i++)
                        {


                            row["SER"] = SER.ToString("000000");
                            row["PBbarcode" + i] = jn.ToString("0000000000") + SER.ToString("000000") + i.ToString("00") + numpg.ToString("00") + ("00000000");
                        }
                        SER++;
                    }



                    using (var textWriter = File.CreateText(dir + "\\" + jn + ".txt")) // write contents to text file
                        
                    using (var csv1 = new CsvWriter(textWriter))
                    {
                        // Write columns
                        csv1.Configuration.Delimiter = "\t";

                        foreach (DataColumn column in dt.Columns)
                        {
                            csv1.WriteField(column.ColumnName);
                        }
                        csv1.NextRecord();

                        // Write row values
                        foreach (DataRow row1 in dt.Rows)
                        {
                            for (var j = 0; j < dt.Columns.Count; j++)
                            {
                                csv1.WriteField(row1[j]);
                            }
                            csv1.NextRecord();
                        }
                    }



                }
            }
        }
    }
}







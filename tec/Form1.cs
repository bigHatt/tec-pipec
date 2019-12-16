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

namespace tec
{
    public partial class Form1 : Form
    {
        public string realData = "real.xlsx";
        public string randData = "rand.csv";

        public Form1()
        {
            InitializeComponent();
            dataGridView1.AutoGenerateColumns = true;
            dataGridView2.AutoGenerateColumns = true;
        }
        
        private void Button1_Click(object sender, EventArgs e)
        {
            var raelData = loadRealData();
            dataGridView1.DataSource = raelData;
            dataGridView1.DataMember = "realData";

            var randData = loadRandData();
            //dataGridView2.DataSource = randData;
            //dataGridView2.DataMember = "randData";
        }

        public DataSet loadRealData()
        {
            using (var stream = File.Open(realData, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();
                    result.Tables[0].TableName = "realData";
                    result.Tables[0].Rows.RemoveAt(0);
                    result.Tables[0].Columns[0].ColumnName = "material";
                    result.Tables[0].Columns[1].ColumnName = "zola";
                    result.Tables[0].Columns[2].ColumnName = "water";
                    result.Tables[0].Columns[3].ColumnName = "price";
                    result.Tables[0].Columns.RemoveAt(4);
                    return result;
                }
            }
        }

        public DataSet loadRandData()
        {
            using (var stream = File.Open(randData, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                {
                    var result = reader.AsDataSet();
                    DataSet dt = new DataSet();
                    for (int i = 1; i < 3; i++)
                    {
                        for (int j = 1; j < result.Tables[0].Rows.Count; j++)
                        {
                            dataGridView2[i,j].Value = result.Tables[0].Rows[j][i];
                        }
                    }
                    //result.Tables[0].TableName = "randData";
                    //result.Tables[0].Rows.RemoveAt(0);
                    //result.Tables[0].Columns[0].ColumnName = "material";
                    //result.Tables[0].Columns[1].ColumnName = "zola";
                    //result.Tables[0].Columns[2].ColumnName = "water";
                    //result.Tables[0].Columns[3].ColumnName = "price";
                    //result.Tables[0].Columns.RemoveAt(4);
                    return dt;
                }
            }
        }
    }
}

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
        public string randData1 = "rand.csv";
        public string randData = "randd.csv";
        /**
         * массивы с данными таблиц 
         **/
        public object[,] real;
        public object[,] rand;
        public List<object[]> realList = new List<object[]>();
        public List<object[]> randList = new List<object[]>();

        public Form1()
        {
            InitializeComponent();
            dataGridView1.AutoGenerateColumns = true;
            dataGridView2.AutoGenerateColumns = true;
        }
        
        public void fillData()
        {
            var raelData = loadRealData();
            dataGridView1.DataSource = raelData;
            dataGridView1.DataMember = "realData";
            real = new object[dataGridView1.RowCount, dataGridView1.ColumnCount];
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                object[] r = new object[dataGridView1.ColumnCount];
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {
                    real[i, j] = dataGridView1[j, i].Value;
                    r[j] = dataGridView1[j, i].Value;
                }
                realList.Add(r);
            }

            var randData = loadRandData();
            dataGridView2.DataSource = randData;
            dataGridView2.DataMember = "randData";
            rand = new object[dataGridView2.RowCount, dataGridView2.ColumnCount];
            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                object[] r = new object[dataGridView2.ColumnCount];
                for (int j = 0; j < dataGridView2.ColumnCount; j++)
                {
                    rand[i, j] = dataGridView2[j, i].Value;
                    r[j] = dataGridView2[j, i].Value;
                }
                randList.Add(r);
            }
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
                    string[] columnsName = new string[] { "m", "y", "t", "d" };
                    var result = reader.AsDataSet();
                    result.Tables[0].TableName = "randData";
                    result.Tables[0].Rows.RemoveAt(0);
                    result.Tables[0].Columns.RemoveAt(0);
                    result.Tables[0].Columns[0].ColumnName = "type";
                    result.Tables[0].Columns[2].ColumnName = "water";
                    result.Tables[0].Columns[1].ColumnName = "zola";
                    result.Tables[0].Columns[3].ColumnName = "price";
                    return result;
                }
            }
        }

        public double[] calc()
        {
            double[] res = new double[3]; // [zola, water, price]
            foreach (var row in randList)
            {
                // calc func
            }

            return res;
        }

        public void findCloseReal()
        {

            var rand = calc();
            rand = new double[] {0.054, 0.27,1450};
            double min = 10000;
            string minName = "";
            foreach (var real in realList)
            {
                double diffZola = Math.Pow(rand[0] - (double)real[1], 2);
                double diffWater = Math.Pow(rand[1] - (double)real[2], 2);
                double diffPrice = Math.Pow(rand[2] - (double)real[3], 2);
                var minn = Math.Sqrt(diffZola + diffWater + diffPrice);
                if (minn < min)
                {
                    min = minn;
                    minName = (string)real[0];
                }
            }
            textBox1.Text = minName;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            fillData();
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            findCloseReal();
        }
    }
}

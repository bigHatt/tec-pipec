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
    public struct CalcResult
    {
        public CalcResult(double Wp, double Ap)
        {
            this.Wp = Wp;
            this.Ap = Ap;
        }

        public double Wp;
        public double Ap;
    }

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
                    result.Tables[0].Columns[4].ColumnName = "heat";
                    result.Tables[0].Columns[5].ColumnName = "sera";
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
                    result.Tables[0].TableName = "randData";
                    result.Tables[0].Rows.RemoveAt(0);
                    result.Tables[0].Columns.RemoveAt(0);
                    result.Tables[0].Columns[0].ColumnName = "type";
                    result.Tables[0].Columns[2].ColumnName = "water";
                    result.Tables[0].Columns[1].ColumnName = "zola";
                    result.Tables[0].Columns[3].ColumnName = "price";
                    result.Tables[0].Columns[4].ColumnName = "heat";
                    result.Tables[0].Columns[5].ColumnName = "sera";
                    return result;
                }
            }
        }



        //nu - КПД, Q - удельная теплота сгорания
        //Sr - содержание серы
        //Wp_Ap_Sum - ограничение на сумму влажности и зольности
        public CalcResult calc_formula_1(double nu, double Q, double Sr, double Wp_Ap_Sum)
        {
            double b = Math.Pow(10, 14) / 3.52 / Sr;
            b *= nu;
            b *= Q;

            double Ap = (2168 + 10 * b) / 25.2;
            double Wp = Wp_Ap_Sum - Ap;

            CalcResult res;
            res.Wp = Wp;
            res.Ap = Ap;

            return res;
        }

        public CalcResult calc_formula_2(double nu, double Q, double Sr, double c, double Wp_Ap_Sum)
        {
            double b = Math.Pow(10, 14) / 3.52 / Sr;
            b *= nu;
            b *= Q;

            double Ap = -5 * (10 * b * c + 8143 * b + 6400) / (785 * b - 8064);
            double Wp = (50 * b * c + 119215 * b - 1260 * c - 1878918) / (785 * b - 8064);

            CalcResult res;
            res.Wp = Wp;
            res.Ap = Ap;

            return res;
        }

        double aConst(double N, double t, double k, double Q, double P0)
        {
            return N * t * k * P0 / /*Math.Pow(10, 6) /*/ Q;
        }

        double dfdWp(double N, double t, double k, double Q, double P0, double Wp, double Ap)
        {
            return aConst(N, t, k, Q, P0) * (25.2 * Ap - 2420) / Math.Pow((100 - Ap - Wp), 2);
        }

        double dfdAp(double N, double t, double k, double Q, double P0, double Wp, double Ap)
        {
            return aConst(N, t, k, Q, P0) * (100 - 25.2 * Wp) / Math.Pow((100 - Ap - Wp), 2);
        }

        double f(double N, double t, double k, double Q, double P0, double Wp, double Ap)
        {
            return aConst(N, t, k, Q, P0) * (100 - 25.2 * Wp) / (100 - Ap - Wp);
        }

        public CalcResult fast_Gradient(double N, double t, double k, double Q, double P0, double h)
        {
            double eps = 0.01;

            double x = 0.7;
            double y = 0.1;

            int counter = 0;

            double G, Y, x_p, y_p;

            do
            {
                x_p = x;
                y_p = y;
                G = f(N, t, k, Q, P0, x, y);
                x = x_p - h * dfdWp(N, t, k, Q, P0, x, y);
                y = y_p - h * dfdAp(N, t, k, Q, x, P0, y);
                Y = f(N, t, k, Q, P0, x, y);

                if (x + y < 0.15 || x + y > 0.9)
                    break;

                counter++;
            }
            while ((Math.Abs(Y - G)) > eps);

            return new CalcResult(x, y);

        }

        public double[] calc()
        {
            double[] res = new double[2]; // [zola, water, price, heat, sera]
            List<CalcResult> results = new List<CalcResult>();
            foreach (var row in randList)
            {
                var str = row[4].ToString();
                str = str.Replace(".", ",");
                double Q = Convert.ToDouble(str);
                str = row[3].ToString();
                str = str.Replace(".", ",");
                double P0 = Convert.ToDouble(str);
                results.Add(fast_Gradient(755, 550, 0.8, Q, P0, 0.0000001));
            }
            var r = results.Last();
            return new double[] { r.Ap, r.Wp };
            //return res;
        }

        public void findCloseReal()
        {

            var rand = calc();
            double min = 10000;
            string minName = "";
            foreach (var real in realList)
            {
                double diffZola = Math.Pow(rand[0] - (double)real[1], 2);
                double diffWater = Math.Pow(rand[1] - (double)real[2], 2);
                //double diffPrice = Math.Pow(rand[2] - (double)real[3], 2);
                //double diffHeat = Math.Pow(rand[3] - (double)real[4], 2);
                //double diffSera = Math.Pow(rand[4] - (double)real[5], 2);
                var minn = Math.Sqrt(diffZola + diffWater);
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

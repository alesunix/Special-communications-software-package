using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using LiveCharts;
using LiveCharts.Wpf;

namespace ProgramCCS
{
    public partial class Graph : Form
    {
        public SqlConnection con = new SqlConnection(@"Data Source=192.168.0.3;Initial Catalog=ccsbase;Persist Security Info=True;User ID=Lan;Password=Samsung0");
        public Graph()
        {
            InitializeComponent();
        }

        private void Graph_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Today.AddYears(-1);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)//График
        {
            con.Open();
            SqlCommand cmd = new SqlCommand("SELECT data_zapisi, SUM(plata_za_uslugu) FROM [Table_1] " +
                "WHERE (data_zapisi BETWEEN @StartDate AND @EndDate) AND status = @status GROUP BY data_zapisi ORDER BY data_zapisi", con);
            cmd.Parameters.AddWithValue("@status", "возврат");
            cmd.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd.ExecuteNonQuery();

            DataTable dataTable = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dataTable.Clear();//чистим DataTable, если он был не пуст
            dataAdapter.Fill(dataTable);//заполняем данными созданный DataTable
            con.Close();

            con.Open();
            SqlCommand cmd1 = new SqlCommand("SELECT data_zapisi, SUM(plata_za_uslugu) FROM [Table_1] " +
                "WHERE (data_zapisi BETWEEN @StartDate AND @EndDate) AND status = @status GROUP BY data_zapisi ORDER BY data_zapisi", con);
            cmd1.Parameters.AddWithValue("@status", "выдано");
            cmd1.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd1.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd1.ExecuteNonQuery();

            DataTable dataTable1 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter dataAdapter1 = new SqlDataAdapter(cmd1);//создаем экземпляр класса SqlDataAdapter
            dataTable1.Clear();//чистим DataTable, если он был не пуст
            dataAdapter1.Fill(dataTable1);//заполняем данными созданный DataTable
            con.Close();

            DataSet dataSet = new DataSet();
            DataSet dataSet1 = new DataSet();

            cartesianChart1.LegendLocation = LegendLocation.Bottom;

            if (dataSet.Tables["Table_1"] != null)
                dataSet.Tables["Table_1"].Clear();
            if (dataSet1.Tables["Table_1"] != null)
                dataSet1.Tables["Table_1"].Clear();
          
            dataAdapter.Fill(dataSet, "Table_1");
            dataTable = dataSet.Tables["Table_1"];
            dataAdapter1.Fill(dataSet1, "Table_1");
            dataTable1 = dataSet1.Tables["Table_1"];
            SeriesCollection series = new SeriesCollection();
            ChartValues<int> values = new ChartValues<int>();
            ChartValues<int> values1 = new ChartValues<int>();
            List<string> dates = new List<string>();
            foreach (DataRow row in dataTable.Rows)
            {
                values.Add(Convert.ToInt32(row[1]));
                dates.Add(Convert.ToDateTime(row[0]).ToShortDateString());
            }  
            foreach (DataRow row in dataTable1.Rows)
            {
                values1.Add(Convert.ToInt32(row[1]));
            }

            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisX.Add(new Axis()
            {
                Title = "Даты",
                Labels = dates
            });
            ColumnSeries line = new ColumnSeries
            {
                Title = "Возврат",
                Values = values,
                DataLabels = true,
                Fill = System.Windows.Media.Brushes.IndianRed,
                LabelPoint = point => (point.Y).ToString(),
            };
            ColumnSeries line2 = new ColumnSeries
            {
                Title = "Выдано",
                Values = values1,
                DataLabels = true,
                Fill = System.Windows.Media.Brushes.CadetBlue,
                LabelPoint = point => (point.Y).ToString(),
            };
            series.Add(line);
            series.Add(line2);
            cartesianChart1.Series = series;           
        }
    }
}

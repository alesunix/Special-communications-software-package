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

namespace ProgramCCS
{
    public partial class Search_registry : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели

        private DataGridView dgv1_TLC; // эта переменная будет содержать ссылку на грид dataGridView1 из формы Form1
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1
        public Search_registry(DataGridView dgv1, DataGridView dgv2)
        {
            dgv1_TLC = dgv1;// теперь dgv1_TLC будет ссылкой на грид dataGridView1
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
            InitializeComponent();
        }
        public void Partner_select()//Вывод Контрагентов в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT name FROM [Table_Partner] ORDER BY id", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow column in dt.Rows)
            {
                comboBox5.Items.Add(column[0].ToString());
            }
            con.Close();//Закрываем соединение          
        }
        private void button7_Click(object sender, EventArgs e)//Выборка
        {
            //9. Выборка поиск Реестра
            if (textBox14.Text != "" & comboBox5.Text != "" & comboBox1.Text != "")
            {
                dgv1_TLC.Visible = true;
                dgv2_TLC.Visible = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE nomer_reestra = @nomer_reestra AND client = @client AND status = @status ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@nomer_reestra", textBox14.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
            }
        }

        private void Search_registry_Load(object sender, EventArgs e)//Загрузка формы
        {
            Partner_select();
        }

        private void Search_registry_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }
    }
}

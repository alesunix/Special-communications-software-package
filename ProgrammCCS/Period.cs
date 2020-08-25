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
    public partial class Period : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели

        private DataGridView dgv1_TLC; // эта переменная будет содержать ссылку на грид dataGridView1 из формы Form1
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1
        public Period(DataGridView dgv1, DataGridView dgv2)
        {
            dgv1_TLC = dgv1;// теперь dgv1_TLC будет ссылкой на грид dataGridView1
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)//Выборка
        {
            dgv1_TLC.Visible = true;
            dgv2_TLC.Visible = false;
            //3. Выборка за период-1 (Дата обработки) - 'Статус + Период + Область + Клиент'.
            if (comboBox1.Text != "" & comboBox2.Text != "" & comboBox5.Text != "" & comboBox4.Text == "" & checkBox1.Checked)
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_obrabotki BETWEEN @StartDate AND @EndDate AND oblast LIKE '%" + comboitem.ToString() + "%' AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение   
            }
            //4. Выборка за период-2 (Дата записи) - 'Статус + Период + Область + Филиал + Клиент'.
            else if (comboBox1.Text != "" & comboBox2.Text != "" & comboBox4.Text != "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_zapisi BETWEEN @StartDate AND @EndDate AND oblast LIKE '%" + comboitem.ToString() + "%'" +
                    " AND filial = @filial AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@filial", comboBox4.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение   
            }
            //5. Выборка за период-3 (Дата записи) - 'Статус + Период + Область + Клиент +- Пункт'.
            else if (comboBox1.Text != "" & comboBox2.Text != "" & comboBox5.Text != "" & comboBox4.Text == "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_zapisi BETWEEN @StartDate AND @EndDate AND oblast LIKE '%" + comboitem.ToString() + "%'" +
                    " AND client = @client AND punkt LIKE '%" + Convert.ToString(textBox18.Text) + "%' ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение      
            }
            //6. Выборка за период-4 (Дата Обработки) - 'Статус + Период + Клиент'.
            else if (checkBox1.Checked && comboBox1.Text != "" & comboBox2.Text == "" & comboBox5.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_obrabotki BETWEEN @StartDate AND @EndDate AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение     
            }
            //7. Выборка за период-5 (Дата записи) - 'Статус + Период + Клиент'.
            else if (checkBox2.Checked && comboBox1.Text != "" & comboBox5.Text != "" & comboBox2.Text == "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение     
            }
            //10. Выборка за период-5 (Дата Обработки) - 'Период + Клиент'.
            else if (comboBox5.Text != "" & checkBox1.Checked)
            {
                DateTime date = new DateTime();
                date = dateTimePicker1.Value;
                DateTime date2 = dateTimePicker2.Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE (data_obrabotki BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                cmd.Parameters.AddWithValue("StartDate", date2);
                cmd.Parameters.AddWithValue("EndDate", date);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
            }
            //11. Выборка за период-6 (Дата записи) - 'Период + Клиент'.
            else if (comboBox5.Text != "")
            {
                DateTime date = new DateTime();
                date = dateTimePicker1.Value;
                DateTime date2 = dateTimePicker2.Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE (data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                cmd.Parameters.AddWithValue("StartDate", date2);
                cmd.Parameters.AddWithValue("EndDate", date);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
            }
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            F1.Podschet();//произвести подсчет из метода
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
        private void Period_Load(object sender, EventArgs e)//Загрузка формы
        {
            Partner_select();

            dateTimePicker2.Value = DateTime.Today.AddDays(0);
            dateTimePicker1.Value = DateTime.Today.AddMonths(-1);

            // инициализация         
            comboBox2.Items.Add(new ClassComboBoxOblast("Чу", "Чуйская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ош", "Ошская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Та", "Таласская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Жал", "Джалал - Абадская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Батк", "Баткенская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ис", "Иссык - Кульская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("На", "Нарынская область"));
        }

        private void Period_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }

        private void button1_Click(object sender, EventArgs e)//Печать
        {

        }
    }
}

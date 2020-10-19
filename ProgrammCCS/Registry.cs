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
    public partial class Registry : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели

        private DataGridView dgv1_TLC; // эта переменная будет содержать ссылку на грид dataGridView1 из формы Form1
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1
        
        public Registry(DataGridView dgv1, DataGridView dgv2)
        {
            dgv1_TLC = dgv1;// теперь dgv1_TLC будет ссылкой на грид dataGridView1
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)//Выборка
        {
            Table.DtRegistry = new DataTable();//инициализируем DataTable
            //1.Выборка на реестр-1 - 'Статус + Обработка + Филиал + Клиент'.
            if (comboBox1.Text != "" & comboBox5.Text != "" & comboBox4.Text != "")
            {
                dgv1_TLC.Visible = true;
                dgv2_TLC.Visible = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND obrabotka = @obrabotka AND filial = @filial AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                if (checkBox3.Checked == true) cmd.Parameters.AddWithValue("@obrabotka", "Обработано");
                else if (checkBox3.Checked == false) cmd.Parameters.AddWithValue("@obrabotka", "Не обработано");
                cmd.Parameters.AddWithValue("@filial", comboBox4.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                Table.DtRegistry.Clear();//чистим DataTable, если он был не пуст
                da.Fill(Table.DtRegistry);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = Table.DtRegistry;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
            }
            //2.Выборка на реестр-2 - 'Статус + Обработка + Клиент'.
            else if (comboBox1.Text != "" & comboBox5.Text != "" & comboBox4.Text == "")
            {
                dgv1_TLC.Visible = true;
                dgv1_TLC.Visible = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND obrabotka = @obrabotka AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                if (checkBox3.Checked == true) cmd.Parameters.AddWithValue("@obrabotka", "Обработано");
                else if (checkBox3.Checked == false) cmd.Parameters.AddWithValue("@obrabotka", "Не обработано");
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();

                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                Table.DtRegistry.Clear();//чистим DataTable, если он был не пуст
                da.Fill(Table.DtRegistry);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = Table.DtRegistry;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
            }
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            F1.Podschet();//произвести подсчет из метода
        }
        private void button1_Click(object sender, EventArgs e)//Печать
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            Partner.Name = comboBox5.Text;
            F1.Print_Registy();
            F1.Disp_data();
            Close();
        }

        public void Partner_select()//Вывод Контрагентов в Combobox
        {
            foreach (DataRow column in Table.DtPartner.Rows)
            {
                comboBox5.Items.Add(column[0].ToString());
            }
        }

        private void Registry_Load(object sender, EventArgs e)//Загрузка формы
        {
            Partner_select();
        }

        private void Registry_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }

        
    }
}

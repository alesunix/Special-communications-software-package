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
    public partial class Invoice : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели

        public Invoice()// получаем ссылку на грид в переменную dgv
        {
            InitializeComponent();
        }
        
        private void Invoice_Load(object sender, EventArgs e)//Загрузка формы
        {
            // инициализация         
            comboBox2.Items.Add(new ClassComboBoxOblast("Чу", "Чуйская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ош", "Ошская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Та", "Таласская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Жал", "Джалал - Абадская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Батк", "Баткенская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ис", "Иссык - Кульская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("На", "Нарынская область"));

            dateTimePicker1.Value = DateTime.Today.AddDays(0);
            Partner_select();
        }
        public void Partner_select()//Вывод Контрагентов в Combobox
        {
            foreach (DataRow column in Table.DtPartner.Rows)
            {
                comboBox5.Items.Add(column[0].ToString());
            }          
        }
        private void Button7_Click(object sender, EventArgs e)//Выборка
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            Dates.EndDate = DateTime.Today.AddDays(0);
            Dates.StartDate = dateTimePicker1.Value;
            //8. Выборка на накладную - 'Дата + Область + Клиент +- Пункт'.
            if (comboBox2.Text != "" & comboBox5.Text != "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                F1.dataGridView1.Visible = true;
                F1.dataGridView2.Visible = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE oblast LIKE '%" + comboitem.ToString() + "%' AND data_zapisi = @data_zapisi AND client = @client" +
                    " AND punkt LIKE '%" + Convert.ToString(textBox18.Text) + "%' ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@data_zapisi", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                Table.DtInvoice = new DataTable();//инициализируем DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                Table.DtInvoice.Clear();//чистим DataTable, если он был не пуст
                da.Fill(Table.DtInvoice);//заполняем данными созданный DataTable
                F1.dataGridView1.DataSource = Table.DtInvoice;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение  
            }           
            F1.Podschet();//произвести подсчет из метода
        }

        private void Invoice_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)//Печать
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            Partner.Name = comboBox5.Text;
            F1.Print_Invoice();//Печать Накладной и за период
            F1.Disp_data();
            Close();
        }
    }
}

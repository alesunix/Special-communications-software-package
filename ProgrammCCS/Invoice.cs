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

        private DataGridView dgv1_TLC; // эта переменная будет содержать ссылку на грид dataGridView1 из формы Form1
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1
        public Invoice(DataGridView dgv, DataGridView dgv2)// получаем ссылку на грид в переменную dgv
        {
            dgv1_TLC = dgv;// теперь dgv1_TLC будет ссылкой на грид dataGridView1
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
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
        private void Button7_Click(object sender, EventArgs e)//Выборка
        {
            //8. Выборка на накладную - 'Дата + Область + Клиент +- Пункт'.
            if (comboBox2.Text != "" & comboBox5.Text != "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                dgv1_TLC.Visible = true;
                dgv2_TLC.Visible = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE oblast LIKE '%" + comboitem.ToString() + "%' AND data_zapisi = @data_zapisi AND client = @client" +
                    " AND punkt LIKE '%" + Convert.ToString(textBox18.Text) + "%' ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@data_zapisi", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                DataTable dt_Invoice = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt_Invoice.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt_Invoice);//заполняем данными созданный DataTable
                dgv1_TLC.DataSource = dt_Invoice;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение  
            }
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            F1.Podschet();//произвести подсчет из метода
        }

        private void Invoice_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
        }

        private void button1_Click(object sender, EventArgs e)//Печать
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            if (dgv1_TLC.Rows.Count > 0 & Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value) == "Ожидание")
            {
                F1.Select_status_Nn();//(Для выдачи накладных)Выборка по статусу и сортировка по номеру накладеой от больших значений к меньшим.               
                if (MessageBox.Show("Вы хотите получить 'Накладную'? Нажмите Нет если хотите получить 'Cписок за период'!", "Внимание! Статус изменится на 'Отправлено' и присвоется номер", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    for (int i = 0; i < dgv1_TLC.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET nomer_nakladnoy = @nomer_nakladnoy, status = @status, Nn=@Nn, filial=@filial WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dgv1_TLC.Rows[i].Cells[11].Value));
                        cmd.Parameters.AddWithValue("@status", "Отправлено");
                        cmd.Parameters.AddWithValue("@nomer_nakladnoy", Number.Prefix_number);
                        cmd.Parameters.AddWithValue("@Nn", Number.Nn);
                        cmd.Parameters.AddWithValue("@filial", Person.Name);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение

                    string oblast = Convert.ToString(dgv1_TLC.Rows[0].Cells[9].Value);//Область
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Накладная № {Number.Prefix_number} - {oblast}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        F1.Export_Nakladnaya_To_Word(dgv1_TLC, sfd.FileName);
                    }
                }
            }
            else if (dgv1_TLC.Rows.Count <= 0)
            {
                MessageBox.Show("Выборка не дала результатов, невозможно сгенерировать реестр!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Эти данные нельзя обработать", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            Close();
        }
    }
}

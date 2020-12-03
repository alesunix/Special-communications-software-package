using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Linq;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProgramCCS
{
    public partial class Search : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели
        DataContext db = new DataContext(Connection.con);//Для работы LINQ to SQL

        private DataGridView dgv1_TLC; // эта переменная будет содержать ссылку на грид dataGridView1 из формы Form1
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1
        public Search(DataGridView dgv1, DataGridView dgv2)
        {
            dgv1_TLC = dgv1;// теперь dgv1_TLC будет ссылкой на грид dataGridView1
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
            InitializeComponent();
        }

        private void button6_Click(object sender, EventArgs e)//Удалить из базы данных
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            if (Person.Name == "root" & textBox3.Text != "" & dgv2_TLC.Rows.Count == 1)
            {
                if (MessageBox.Show("Вы хотите удалить эту запись?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("DELETE FROM [Table_1] WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dgv2_TLC.Rows[0].Cells[0].Value));//первая строка в гриде
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                    MessageBox.Show("Запись успешно удалена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    F1.Disp_data();
                    textBox3.Select();//Установка курсора
                }
                else
                {
                    F1.Disp_data();
                    textBox3.Select();//Установка курсора
                }
            }
            else if (dgv2_TLC.Rows.Count != 1)
            {
                MessageBox.Show("Произведите поиск по №Заказа или по Фамилии", "Внимание! Чтобы удалить запись из базы данных", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (dgv2_TLC.Rows.Count <= 0)
            {
                MessageBox.Show("В базе не найдено отправление", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else MessageBox.Show("Только администратор может удалять", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            textBox3.Text = "";//очистка текстовых полей
            F1.Disp_data();
        }

        private void button28_Click(object sender, EventArgs e)//Поиск по Ф.И.О
        {
            if (textBox2.Text != "")
            {
                dgv2_TLC.Visible = true;
                dgv1_TLC.Visible = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT id AS ID, oblast AS 'Область', punkt AS 'Населенный пункт', familia AS 'Ф.И.О'," +
                "summ AS 'Стоимость',plata_za_uslugu AS 'Услуга', tarif AS 'Тариф', doplata AS 'Доплата', ob_cennost AS 'Обьяв.ценность', plata_za_nalog AS 'Наложеный платеж'," +
                    "N_zakaza AS '№Заказа', status AS 'Статус', data_zapisi AS 'Дата записи', prichina AS 'Причина', obrabotka AS 'Обработка', data_obrabotki AS 'Дата обработки'," +
                    "filial AS 'Филиал', client AS 'Контрагент'," +
                    "nomer_spiska AS 'Список', nomer_nakladnoy AS 'Накладная', nomer_reestra AS 'Реестр', Ns AS 'NS', Nn AS 'NN', Nr AS 'NR', tarifs AS 'Тарифы'" +
                        "FROM [Table_1] WHERE familia LIKE N'%" + textBox2.Text.ToString() + "%'", con);
                //cmd.Parameters.AddWithValue("@punkt", textBox2.Text);
                //cmd.Parameters.AddWithValue("@familia", textBox2.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgv2_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение


                TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
                F1.Podschet();//произвести подсчет по методу
                //table1BindingSource.Filter = "[punkt] LIKE '%" + Convert.ToString(textBox2.Text) + "%' OR [familia] LIKE '%" + Convert.ToString(textBox2.Text) + "%'"; //Фильтр по гриду
            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)//Поиск по №Заказа
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            dgv2_TLC.Visible = true;
            dgv1_TLC.Visible = false;

            var command = from table in db.GetTable<Table_1>()
                          where table.N_Заказа == textBox3.Text.ToString()
                          orderby table.Дата_записи descending
                          select table;
            dgv2_TLC.DataSource = command;

            if (textBox3.Text == "")//если поле очищено, отобразить базу
            {
                F1.Disp_data();
            }
            F1.Podschet();//произвести подсчет по методу
        }

        private void Search_Load(object sender, EventArgs e)
        {

        }

        private void Search_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }
    }
}

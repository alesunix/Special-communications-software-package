﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Linq;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProgramCCS
{
    public partial class Editor : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели
        DataContext db = new DataContext(Connection.con);//Для работы LINQ to SQL

        private DataGridView dgv1_TLC; // эта переменная будет содержать ссылку на грид dataGridView1 из формы Form1
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1

        public Editor(DataGridView dgv1, DataGridView dgv2)
        {
            dgv1_TLC = dgv1;// теперь dgv1_TLC будет ссылкой на грид dataGridView1
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
            InitializeComponent();
        }
       
        private void button10_Click(object sender, EventArgs e)//Change
        {
            int currRowIndex = dgv2_TLC.CurrentCell.RowIndex;//  Запоминаем строку, которую выбрал пользователь.
            if (comboBox1.Text != "" & textBox8.Text == "" )
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET status = @status, prichina = @prichina, filial = @filial WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", dgv2_TLC.CurrentRow.Cells[0].Value);//выбранная строка в гриде
                if (Convert.ToString(dgv2_TLC.Rows[0].Cells[11].Value) == "Ожидание" ||
                    Convert.ToString(dgv2_TLC.Rows[0].Cells[11].Value) == "Отправлено" ||
                    Convert.ToString(dgv2_TLC.Rows[0].Cells[11].Value) == "Розыск" ||
                    Convert.ToString(dgv2_TLC.Rows[0].Cells[11].Value) == "Замена" ||
                    comboBox1.Text == "Розыск" || comboBox1.Text == "Замена")
                {
                    cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@prichina", comboBox8.Text);
                    cmd.Parameters.AddWithValue("@filial", Person.Name);
                    MessageBox.Show("Статус успешно обновлен", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    cmd.ExecuteNonQuery();
                }
                else MessageBox.Show("Изминение статуса невозможно, так как статус уже присвоен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                con.Close();//закрыть соединение             
                textBox3.Select();//Установка курсора
            }
            else if (textBox8.Text != "" & comboBox1.Text == "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET tarif = @tarif WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", dgv2_TLC.CurrentRow.Cells[0].Value);//выбранная строка в гриде
                cmd.Parameters.AddWithValue("@tarif", textBox8.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Тариф успешно обновлен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox3.Select();//Установка курсора
            }
            else if (textBox16.Text != "" & textBox8.Text == "" & comboBox1.Text == "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET doplata = @doplata WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", dgv2_TLC.CurrentRow.Cells[0].Value);//выбранная строка в гриде
                cmd.Parameters.AddWithValue("@doplata", textBox16.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Доплата успешно обновлена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox3.Select();//Установка курсора
            }
            else if (comboBox2.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET oblast = @oblast WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", dgv2_TLC.CurrentRow.Cells[0].Value);//выбранная строка в гриде
                cmd.Parameters.AddWithValue("@oblast", comboBox2.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Область успешно изменена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox3.Select();//Установка курсора
            }
            else if (textBox18.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET punkt = @punkt WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", dgv2_TLC.CurrentRow.Cells[0].Value);//выбранная строка в гриде
                cmd.Parameters.AddWithValue("@punkt", textBox18.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Населенный пункт успешно изменен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox3.Select();//Установка курсора
            }
            else if (textBox19.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET summ = @summ WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", dgv2_TLC.CurrentRow.Cells[0].Value);//выбранная строка в гриде
                cmd.Parameters.AddWithValue("@summ", textBox19.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Стоимость успешно изменена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                textBox3.Select();//Установка курсора
            }
            else if (textBox8.Text == "" & textBox16.Text == "" & comboBox1.Text == "" & comboBox2.Text == "" & dgv2_TLC.Rows.Count == 1)
            {
                MessageBox.Show("Введите (сумму тарифа или доплаты или стоимость) - или - Выбирите статус - или - Выбирите область", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (dgv2_TLC.Rows.Count != 1)
            {
                MessageBox.Show("Произведите поиск по №Заказа или по Фамилии", "Внимание! Чтобы изменить (Статус, Область, Стоимость, Тариф или Доплату)", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (dgv2_TLC.Rows.Count <= 0)
            {
                MessageBox.Show("В базе не найдено отправление", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.None);
            }
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            F1.Tarif_Update();//Заново ищет №Заказа и делает пересчет
            textBox3.Text = "";//очистка текстовых полей
            textBox8.Text = "";
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox16.Text = "";
            textBox19.Text = "";
            F1.Disp_data();
            dgv2_TLC.CurrentCell = dgv2_TLC[0, currRowIndex];//  Выбираем нашу строку (именно выбираем, не выделяем).
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

        private void button6_Click(object sender, EventArgs e)//Удалить из базы данных
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            if (textBox3.Text != "" & dgv2_TLC.Rows.Count == 1)
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
            textBox3.Text = "";//очистка текстовых полей
            textBox8.Text = "";
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox16.Text = "";
            textBox19.Text = "";
            F1.Disp_data();
        }

        private void Editor_Load(object sender, EventArgs e)//Загрузка формы
        {
            // инициализация         
            comboBox2.Items.Add(new ClassComboBoxOblast("Чу", "Чуйская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ош", "Ошская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Та", "Таласская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Жал", "Джалал - Абадская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Батк", "Баткенская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ис", "Иссык - Кульская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("На", "Нарынская область"));

            //если файл существует
            if (File.Exists("Prichina_vozvrat.txt"))
            {//создаем байтовый поток и привязываем его к файлу
             //в конструкторе указываем: путь кодировка
                using (var sr = new StreamReader("Prichina_vozvrat.txt"))
                {
                    while (!sr.EndOfStream)
                    {
                        comboBox8.Items.Add(sr.ReadLine());
                    }
                }
            }
            else//если файл не существует, создаем и заполняем
            {
                using (var sw = new StreamWriter("Prichina_vozvrat.txt", true, Encoding.UTF8))
                {
                    sw.WriteLine("Не соответствует рекламе");
                    sw.WriteLine("Не соответствует заказу");
                    sw.WriteLine("Не соответствует требованию");
                    sw.WriteLine("Не заказывал(а)");
                    sw.WriteLine("Не в городе");
                    sw.WriteLine("Не такой как в рекламе");
                    sw.WriteLine("Не полный товар");
                    sw.WriteLine("Не устраивает качество");
                    sw.WriteLine("Не оригинал");
                    sw.WriteLine("Товар бракованный");
                    sw.WriteLine("Плохие отзывы");
                    sw.WriteLine("Сделал отмену заказа");
                    sw.WriteLine("Не тот №");
                    sw.WriteLine("Дорого");
                    sw.WriteLine("Повторный заказ");
                    sw.WriteLine("Выехал на заказ");
                    sw.WriteLine("По истечении срока");
                    sw.WriteLine("Нет денег");
                    sw.WriteLine("Передумал");
                    sw.WriteLine("Недоступен");
                    sw.WriteLine("Уехал в командировку");
                    sw.WriteLine("Дорого, цвет и размер не тот");
                    sw.WriteLine("Размер и цвет не подошел");
                    sw.WriteLine("Сын без спроса заказал");
                    sw.WriteLine("Нет футляра");
                    sw.WriteLine("Нет сертификата");
                    sw.WriteLine("Поздняя доставка");
                    sw.WriteLine("Заказ ошибочный");
                }
                using (var sr = new StreamReader("Prichina_vozvrat.txt"))
                {
                    while (!sr.EndOfStream)
                    {
                        comboBox8.Items.Add(sr.ReadLine());
                    }
                }
            }
        }

        private void Editor_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }
    }
}

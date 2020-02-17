using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace ProgramCCS
{
    public partial class Inventory_control : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=192.168.0.3;Initial Catalog=ccsbase;Persist Security Info=True;User ID=Lan;Password=Samsung0");
        public Inventory_control()
        {
            InitializeComponent();
            comboBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "Добавить" с клавиатуры
            textBox4.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "Добавить" с клавиатуры
        }
        private void Sklad_Load(object sender, EventArgs e)//Загрузка формы
        {
            //-----------------Окраска Гридов-------------------//
            DataGridViewRow row1 = this.dataGridView1.RowTemplate;
            row1.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row1.Height = 5;
            row1.MinimumHeight = 17;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;//цвет заголовка
            DataGridViewRow row2 = this.dataGridView2.RowTemplate;
            row2.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row2.Height = 5;
            row2.MinimumHeight = 17;
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSlateGray;//цвет заголовка

            dateTimePicker1.Value = DateTime.Today.AddMonths(-1);
            dateTimePicker2.Value = DateTime.Today.AddDays(0);
            //итоги
            disp_data();
            disp_data2();
            pereschet();
            ITOGI();//Итоги
            disp_data();
            disp_data2();
        }
        private void Sklad_FormClosed(object sender, FormClosedEventArgs e)//Закрытие формы
        {
            if (MessageBox.Show("Вы действительно хотите выйти?!", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                Application.Exit();
            }
        }
        public void disp_data()//Table_Hoz_Osnovnye
        {
            con.Open();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [Table_Hoz_Osnovnye] ORDER BY date";
            cmd.ExecuteNonQuery();

            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridView1.DataSource = dt;
            con.Close();
        }
        public void disp_data2()//Table_Hoz_MBP
        {
            con.Open();
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [Table_Hoz_MBP] ORDER BY date";
            cmd.ExecuteNonQuery();

            DataTable dt = new DataTable();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            dataGridView2.DataSource = dt;
            con.Close();
        }
        public void pereschet()//Пересчет
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                //Через месяц (приход обнуляется, стоимость получает значение цены)
                con.Open();//открыть соединение
                for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл (основные средства)
                {   //Действие №1 (-------)
                    DateTime X = Convert.ToDateTime(dataGridView1.Rows[i].Cells[10].Value);//Переменная столбца даты
                    DateTime Y = DateTime.Today.AddMonths(-1);//Переменная текущей даты минус 1 месяц
                    if (X <= Y)// если дата меньше текущей на 1 месяц (основные средства)
                    { 
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_Osnovnye] SET stoimost = @stoimost, prihod = @prihod WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@stoimost", Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value));
                        cmd.Parameters.AddWithValue("@prihod", 0);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();                       
                        //dataGridView1.Rows[i].Cells[6].Value = dataGridView1.Rows[i].Cells[4].Value;//стоимость получает значение цены);
                        //dataGridView1.Rows[i].Cells[7].Value = 0;//приход обнуляется
                    }
                }
                con.Close();//закрыть соединение
                //основные обнуляются те что в расходе 
                con.Open();//открыть соединение         
                for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
                {
                    //Действие №2 (удаление) основные средства
                    DateTime X = Convert.ToDateTime(dataGridView1.Rows[i].Cells[10].Value);//Переменная столбца даты
                    DateTime Y = DateTime.Today.AddMonths(-1);//Переменная текущей даты минус 1 месяц
                    int S = Convert.ToInt32(dataGridView1.Rows[i].Cells[8].Value);//Расход
                    if (X <= Y && S > 0)// Если дата меньше текущей на 1 месяц и Если расход больше 0
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_Osnovnye] SET cena=@cena, stoimost=@stoimost, kol_vo=@kol_vo, prihod=@prihod, rashod=@rashod, itog=@itog WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@cena", 0);
                        cmd.Parameters.AddWithValue("@kol_vo", 0);
                        cmd.Parameters.AddWithValue("@stoimost", 0);
                        cmd.Parameters.AddWithValue("@prihod", 0);
                        cmd.Parameters.AddWithValue("@rashod", 0);
                        cmd.Parameters.AddWithValue("@itog", 0);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                        //dataGridView1.Rows[i].Cells[4].Value = 0;
                        //dataGridView1.Rows[i].Cells[5].Value = 0;
                        //dataGridView1.Rows[i].Cells[6].Value = 0;
                        //dataGridView1.Rows[i].Cells[7].Value = 0;
                        //dataGridView1.Rows[i].Cells[8].Value = 0;
                        //dataGridView1.Rows[i].Cells[9].Value = 0;
                    }
                }
                con.Close();//закрыть соединение
                //Сальдо на конец основные средства
                double saldo_konec = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[9].Value ?? "0").ToString().Replace(".", ","), out incom);//Итог
                    saldo_konec += incom;
                }
                textBox6.Visible = true;
                textBox6.Text = saldo_konec.ToString() + " сом";
                //Сальдо на начало основные средства
                double saldo_nachalo = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[6].Value ?? "0").ToString().Replace(".", ","), out incom);//Стоимость
                    saldo_nachalo += incom;
                }
                textBox7.Visible = true;
                textBox7.Text = saldo_nachalo.ToString() + " сом";
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                //Через месяц (приход обнуляется, стоимость получает значение цены)
                con.Open();//открыть соединение   
                for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл (МБП)
                {
                    DateTime W = Convert.ToDateTime(dataGridView2.Rows[i].Cells[12].Value);//Переменная столбца даты
                    DateTime Z = DateTime.Today.AddMonths(-1);//Переменная текущей даты минус 1 месяц
                    if (W <= Z)// если дата меньше текущей на 1 месяц (МБП)
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_MBP] SET stoimost = @stoimost, prihod = @prihod, kol_vo=@kol_vo, kol_vo_p=@kol_vo_p WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@stoimost", Convert.ToInt32(dataGridView1.Rows[i].Cells[7].Value));
                        cmd.Parameters.AddWithValue("@kol_vo", Convert.ToInt32(dataGridView1.Rows[i].Cells[6].Value));
                        cmd.Parameters.AddWithValue("@kol_vo_p", 0);
                        cmd.Parameters.AddWithValue("@prihod", 0);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                        //dataGridView2.Rows[i].Cells[5].Value = dataGridView2.Rows[i].Cells[7].Value;//стоимость получает значение цены
                        //dataGridView2.Rows[i].Cells[4].Value = dataGridView2.Rows[i].Cells[6].Value;//Кол-во получает у кол-во прихода
                        //dataGridView2.Rows[i].Cells[7].Value = 0;//приход обнуляется
                        //dataGridView2.Rows[i].Cells[6].Value = 0;//кол-во приход обнуляется
                    }
                }
                con.Close();//закрыть соединение
                // МБП обнуляются те что в расходе  
                con.Open();//открыть соединение  
                for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
                {
                    //Действие №2 (удаление) МБП
                    DateTime W = Convert.ToDateTime(dataGridView2.Rows[i].Cells[12].Value);//Переменная столбца даты
                    DateTime Z = DateTime.Today.AddMonths(-1);//Переменная текущей даты минус 1 месяц
                    int S = Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value);//Расход
                    int K = Convert.ToInt32(dataGridView2.Rows[i].Cells[8].Value);//Кол-во расход
                    int L = Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value);//Кол-во
                    if (W <= Z & S > 0 & K == L)// Если дата меньше текущей на 1 месяц и Если расход больше 0 и кол-во расхода равно кол-во
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_MBP] SET cena=@cena, stoimost=@stoimost, kol_vo=@kol_vo, kol_vo_p=@kol_vo_p, kol_vo_r=@kol_vo_r, kol_vo_i=@kol_vo_i, prihod=@prihod, rashod=@rashod, itog=@itog WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@cena", 0);
                        cmd.Parameters.AddWithValue("@kol_vo", 0);
                        cmd.Parameters.AddWithValue("@stoimost", 0);
                        cmd.Parameters.AddWithValue("@kol_vo_p", 0);
                        cmd.Parameters.AddWithValue("@kol_vo_r", 0);
                        cmd.Parameters.AddWithValue("@kol_vo_i", 0);
                        cmd.Parameters.AddWithValue("@prihod", 0);
                        cmd.Parameters.AddWithValue("@rashod", 0);
                        cmd.Parameters.AddWithValue("@itog", 0);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                        //dataGridView2.Rows[i].Cells[3].Value = 0;
                        //dataGridView2.Rows[i].Cells[4].Value = 0;
                        //dataGridView2.Rows[i].Cells[5].Value = 0;
                        //dataGridView2.Rows[i].Cells[6].Value = 0;
                        //dataGridView2.Rows[i].Cells[7].Value = 0;
                        //dataGridView2.Rows[i].Cells[8].Value = 0;
                        //dataGridView2.Rows[i].Cells[9].Value = 0;
                        //dataGridView2.Rows[i].Cells[10].Value = 0;
                        //dataGridView2.Rows[i].Cells[11].Value = 0;
                    }
                }
                con.Close();//закрыть соединение
                //Сальдо на конец МБП
                double saldo_konec_MBP = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[11].Value ?? "0").ToString().Replace(".", ","), out incom);//Итог МБП
                    saldo_konec_MBP += incom;
                }
                textBox9.Visible = true;
                textBox9.Text = saldo_konec_MBP.ToString() + " сом";
                //Сальдо на начало МБП
                double saldo_nachalo_MBP = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);//Стоимость МБП
                    saldo_nachalo_MBP += incom;
                }
                textBox8.Visible = true;
                textBox8.Text = saldo_nachalo_MBP.ToString() + " сом";
            }
        }
        public void ITOGI()//Итоги
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                con.Open();//открыть соединение
                for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл (основные средства)
                {
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_Osnovnye] SET itog = (stoimost + prihod - rashod)" +
                "UPDATE [Table_Hoz_MBP] SET itog = (stoimost + prihod - rashod), kol_vo_i = (kol_vo + kol_vo_p - kol_vo_r), prihod = (kol_vo_p * cena)", con);
                    cmd.ExecuteNonQuery();
                }
                con.Close();//закрыть соединение
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                con.Open();//открыть соединение   
                for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл (МБП)
                {
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_Osnovnye] SET itog = (stoimost + prihod - rashod)" +
                "UPDATE [Table_Hoz_MBP] SET itog = (stoimost + prihod - rashod), kol_vo_i = (kol_vo + kol_vo_p - kol_vo_r), prihod = (kol_vo_p * cena)", con);
                    cmd.ExecuteNonQuery();
                }
                con.Close();//закрыть соединение
            }
        }

        private void textBox10_TextChanged(object sender, EventArgs e)//Поиск
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT * FROM [Table_Hoz_Osnovnye] WHERE name LIKE '%" + Convert.ToString(textBox10.Text) + "%'", con);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
                pereschet();
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT * FROM [Table_Hoz_MBP] WHERE name LIKE '%" + Convert.ToString(textBox10.Text) + "%'", con);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
                pereschet();
            }
        }
        private void button1_Click(object sender, EventArgs e)//Добавить
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                if (textBox1.Text != "" & textBox2.Text != "" & textBox4.Text != "")
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("INSERT INTO [Table_Hoz_Osnovnye] (inv, name, ed, cena, prihod, rashod, date, kol_vo, stoimost) VALUES (@inv, @name, @ed, @cena, @prihod, @rashod, @date, @kol_vo, @stoimost)", con);
                    cmd.Parameters.AddWithValue("@inv", textBox1.Text);
                    cmd.Parameters.AddWithValue("@name", textBox2.Text);
                    cmd.Parameters.AddWithValue("@date", DateTime.Today);
                    cmd.Parameters.AddWithValue("@ed", "шт");
                    cmd.Parameters.AddWithValue("@cena", textBox4.Text);
                    cmd.Parameters.AddWithValue("@prihod", textBox4.Text);
                    cmd.Parameters.AddWithValue("@rashod", 0);
                    cmd.Parameters.AddWithValue("@kol_vo", 1);
                    cmd.Parameters.AddWithValue("@stoimost", 0);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                    MessageBox.Show("Запись в Основные успешно добавлена", "Внимание!");
                    textBox1.Select();//установить курсор
                }
                else if (textBox1.Text == "" & textBox2.Text == "" & textBox4.Text == "")
                {
                    label6.Visible = true;
                    label6.Text = "Заполните все поля";
                }
                else MessageBox.Show("Ошибка, Перейдите на вкладку МБП!", "Внимание!");
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                if (textBox2.Text != "" && textBox3.Text != "" && textBox4.Text != "" && comboBox1.Text != "")
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("INSERT INTO [Table_Hoz_MBP] (name, ed, cena, kol_vo, stoimost, kol_vo_p, kol_vo_r, rashod) VALUES (@name, @ed, @cena, @kol_vo, @stoimost, @kol_vo_p, @kol_vo_r, @rashod)", con);
                    cmd.Parameters.AddWithValue("@name", textBox2.Text);
                    cmd.Parameters.AddWithValue("@date", DateTime.Today);
                    cmd.Parameters.AddWithValue("@ed", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@cena", textBox4.Text);
                    cmd.Parameters.AddWithValue("@kol_vo", 0);
                    cmd.Parameters.AddWithValue("@stoimost", 0);
                    cmd.Parameters.AddWithValue("@kol_vo_p", textBox3.Text);
                    cmd.Parameters.AddWithValue("@kol_vo_r", 0);
                    cmd.Parameters.AddWithValue("@rashod", 0);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                    MessageBox.Show("Запись МБП успешно добавлена", "Внимание!");
                    textBox2.Select();//установить курсор
                }
                else if (textBox2.Text == "" && textBox3.Text == "" && textBox4.Text == "" && comboBox1.Text == "")
                {
                    label6.Visible = true;
                    label6.Text = "Заполните все поля";
                }
                else MessageBox.Show("Ошибка, Перейдите на вкладку Основные средства!", "Внимание!");
            }
            disp_data();
            disp_data2();
            pereschet();
            ITOGI();//Итоги
            disp_data();
            disp_data2();
            textBox1.Text = "";//очистка текстовых полей
            textBox2.Text = "";//
            textBox3.Text = "";//
            textBox4.Text = "";//
            comboBox1.Text = "";
        }
        private void button2_Click(object sender, EventArgs e)//Расход 
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                if (dataGridView1.Rows.Count == 1)
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_Osnovnye] SET rashod = (cena) WHERE id = @id", con);//получаем значение из столбца cena в столбец rashod
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[0].Cells[0].Value));//первая строка в гриде
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
                else if (dataGridView1.Rows.Count != 1)
                {
                    label6.Visible = true;
                    label6.Text = "Произведите поиск";
                    MessageBox.Show("Произведите поиск!", "Внимание!");
                }
                else if (dataGridView1.Rows.Count <= 0)
                {
                    label6.Text = "В базе не найдено!";
                    MessageBox.Show("В базе не найдено!", "Внимание!");
                }
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                if (dataGridView2.Rows.Count == 1 & textBox3.Text != "")
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_Hoz_MBP] SET kol_vo_r = (kol_vo_p - " + textBox3.Text + "), rashod = (cena * kol_vo_r) WHERE id = @id", con);//получаем значение из столбца cena в столбец rashod
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
                else if (textBox3.Text == "")
                {
                    label6.Visible = true;
                    label6.Text = "Установите кол-во";
                    MessageBox.Show("Установите кол-во!", "Внимание!");
                    textBox3.Select();//установить курсор
                }
                else if (dataGridView2.Rows.Count != 1)
                {
                    label6.Visible = true;
                    label6.Text = "Произведите поиск";
                    MessageBox.Show("Произведите поиск!", "Внимание!");
                }
                else if (dataGridView2.Rows.Count <= 0)
                {
                    label6.Text = "В базе не найдено!";
                    MessageBox.Show("В базе не найдено!", "Внимание!");
                }
            }
            disp_data();
            disp_data2();
            pereschet();
            ITOGI();//Итоги
            disp_data();
            disp_data2();
            textBox1.Text = "";//очистка текстовых полей
            textBox2.Text = "";//
            textBox3.Text = "";//
            textBox4.Text = "";//
            comboBox1.Text = "";
        }
        private void button6_Click(object sender, EventArgs e)//Удалить записи расхода
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
                {
                    //Действие №2 (удаление) основные средства
                    DateTime X = Convert.ToDateTime(dataGridView1.Rows[i].Cells[10].Value);//Переменная столбца даты
                    DateTime Y = DateTime.Today.AddMonths(-1);//Переменная текущей даты минус 1 месяц
                    int S = Convert.ToInt32(dataGridView1.Rows[i].Cells[8].Value);//Расход
                    if (X <= Y && S > 0 && (MessageBox.Show("Вы действительно хотите удалить записи расхода основных средств?", "Внимание! Рекомендую сохранить отчет", MessageBoxButtons.YesNo) == DialogResult.Yes))// Если дата меньше текущей на 1 месяц и Если расход больше 0
                    {
                        con.Open();//открыть соединение
                        SqlCommand cmd = new SqlCommand("DELETE FROM [Table_Hoz_Osnovnye] WHERE rashod = @rashod", con);//расход удаляется
                        cmd.Parameters.AddWithValue("@rashod", Convert.ToInt32(dataGridView1.Rows[i].Cells[8].Value));
                        cmd.ExecuteNonQuery();
                        con.Close();//закрыть соединение
                    }
                }
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
                {
                    //Действие №2 (удаление) МБП
                    DateTime W = Convert.ToDateTime(dataGridView2.Rows[i].Cells[12].Value);//Переменная столбца даты
                    DateTime Z = DateTime.Today.AddMonths(-1);//Переменная текущей даты минус 1 месяц
                    int S = Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value);//Расход
                    int K = Convert.ToInt32(dataGridView2.Rows[i].Cells[8].Value);//Кол-во расход
                    int L = Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value);//Кол-во
                    if (W <= Z && S > 0 && K == L && (MessageBox.Show("Вы действительно хотите удалить записи расхода МБП?", "Внимание! Рекомендую сохранить отчет", MessageBoxButtons.YesNo) == DialogResult.Yes))// Если дата меньше текущей на 1 месяц и Если расход больше 0 и кол-во расхода равно кол-во
                    {
                        con.Open();//открыть соединение
                        SqlCommand cmd = new SqlCommand("DELETE FROM [Table_2] WHERE rashod = @rashod", con);//расход удаляется
                        cmd.Parameters.AddWithValue("@rashod", Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value));
                        cmd.ExecuteNonQuery();
                        con.Close();//закрыть соединение
                    }
                }
            }
            disp_data();
            disp_data2();
            pereschet();
            ITOGI();//Итоги
            disp_data();
            disp_data2();
            textBox1.Text = "";//очистка текстовых полей
            textBox2.Text = "";//
            textBox3.Text = "";//
            textBox4.Text = "";//
            comboBox1.Text = "";
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)//Обработчик окраски строк в таблице
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                int S = Convert.ToInt32(dataGridView1.Rows[i].Cells[8].Value);//расход
                int C = Convert.ToInt32(dataGridView1.Rows[i].Cells[7].Value);//приход
                if (S > 0)// Eсли расход больше 0
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightSalmon;//окраска строк в красный цвет
                }
                else if (C > 0 & S <= 0)// Если приход больше 0 и расход равен 0
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;//окраска строк в зеленый цвет
                }
            }
        }
        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)//Обработчик окраски строк в таблице МБП
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
            {
                int S = Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value);//Расход
                int C = Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value);//Приход
                if (S > 0)// Eсли расход больше 0
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightSalmon;//окраска строк в красный цвет
                }
                else if (C > 0 & S <= 0)// Если приход больше 0 и расход равен 0
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;//окраска строк в зеленый цвет
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)//Выборка
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                DateTime date = new DateTime();
                date = dateTimePicker1.Value;
                DateTime date2 = dateTimePicker2.Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT * FROM [Table_Hoz_Osnovnye] WHERE (date BETWEEN @StartDate AND @EndDate)", con);//Выборка по датам
                cmd.Parameters.AddWithValue("StartDate", date);
                cmd.Parameters.AddWithValue("EndDate", date2);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView1.DataSource = dt;
                con.Close();//закрыть соединение
                pereschet();
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                DateTime date = new DateTime();
                date = dateTimePicker1.Value;
                DateTime date2 = dateTimePicker2.Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT * FROM [Table_Hoz_MBP] WHERE (date BETWEEN @StartDate AND @EndDate)", con);//Выборка по датам
                cmd.Parameters.AddWithValue("StartDate", date);
                cmd.Parameters.AddWithValue("EndDate", date2);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);
                dataGridView2.DataSource = dt;
                con.Close();//закрыть соединение
                pereschet();
            }
        }
        private void button4_Click(object sender, EventArgs e)//Отчеты
        {
            if (tabControl1.SelectedTab == tabPage1)//Если основные
            {
                pereschet();
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = "Отчет основных средств №.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Export_Data_To_Word(dataGridView1, sfd.FileName);
                }
            }
            else if (tabControl1.SelectedTab == tabPage2)//Если МБП
            {
                pereschet();
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = "Отчет МБП №.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Export_Data_To_Word2(dataGridView2, sfd.FileName);
                }
            }
        }
        public void Export_Data_To_Word(DataGridView dataGridView1, string filename)//Обработчик Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);

            rng.InsertBefore("Заголовок");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 9;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева

            if (dataGridView1.Rows.Count != 0)
            {
                int RowCount = dataGridView1.Rows.Count;
                int ColumnCount = dataGridView1.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = dataGridView1.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                //Добавление текста в документ
                string saldo_nachalo = Convert.ToString(textBox7.Text);//Сальдо начало
                string saldo_konec = Convert.ToString(textBox6.Text);//Сальдо конец
                oDoc.Content.SetRange(0, 0);
                oDoc.Content.Text = "Сальдо на начало:   " + saldo_nachalo + "           Сальдо на конец:   " + saldo_konec + Environment.NewLine +
                Environment.NewLine + "Выполнил__________________" + "              " + "Принял_____________________" + Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //заголовка стиль строки
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = dataGridView1.Columns[c].HeaderText;
                }
                //стиль таблицы               
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text = "Отчет №_";
                    headerRange.Font.Size = 12;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "ГП Служба специальной связи      " + Convert.ToString(Now.ToString("dd:MM:yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Export_Data_To_Word2(DataGridView dataGridView2, string filename)//Обработчик Word МБП
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);

            rng.InsertBefore("Заголовок");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 9;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева

            if (dataGridView2.Rows.Count != 0)
            {
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count;
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = dataGridView2.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                //Добавление текста в документ
                string saldo_nachalo = Convert.ToString(textBox8.Text);//Сальдо начало
                string saldo_konec = Convert.ToString(textBox9.Text);//Сальдо конец
                oDoc.Content.SetRange(0, 0);
                oDoc.Content.Text = "Сальдо на начало:   " + saldo_nachalo + "           Сальдо на конец:   " + saldo_konec + Environment.NewLine +
                Environment.NewLine + "Выполнил__________________" + "              " + "Принял_____________________" + Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //заголовка стиль строки
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;
                }
                //стиль таблицы               
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text = "Отчет №_";
                    headerRange.Font.Size = 12;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "ГП Служба специальной связи      " + Convert.ToString(Now.ToString("dd:MM:yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }


        //private void button9_Click(object sender, EventArgs e)//Печать МБП
        //{
        //    if (MessageBox.Show("Вы выполнили процедуры?", "Внимание! Перед тем как отправить на печать выполните процедуры", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //    {
        //        PrintDocument Document = new PrintDocument();
        //        Document.DefaultPageSettings.Landscape = true;//Альбомная ориентация
        //        Document.PrintPage += new PrintPageEventHandler(printDocument2_PrintPage);
        //        PrintPreviewDialog dlg = new PrintPreviewDialog();
        //        dlg.Document = Document;
        //        dlg.ShowDialog();
        //    }

        //}
        private void printDocument2_PrintPage(object sender, PrintPageEventArgs e)
        {
            Bitmap bmp2 = new Bitmap(dataGridView2.Size.Width + 10, dataGridView2.Size.Height + 10);
            dataGridView2.DrawToBitmap(bmp2, dataGridView2.Bounds);
            e.Graphics.DrawImage(bmp2, 0, 0);
        }
        //private void button4_Click(object sender, EventArgs e)//Печать основных средств
        //{
        //    if (MessageBox.Show("Вы выполнили процедуры?", "Внимание! Перед тем как отправить на печать выполните процедуры", MessageBoxButtons.YesNo) == DialogResult.Yes)
        //    {
        //        PrintDocument Document = new PrintDocument();
        //        Document.DefaultPageSettings.Landscape = true;//Альбомная ориентация
        //        Document.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);
        //        PrintPreviewDialog dlg = new PrintPreviewDialog();
        //        dlg.Document = Document;
        //        dlg.ShowDialog();
        //    }

        //}
        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)//Обработчик печати
        {
            Bitmap bmp = new Bitmap(dataGridView1.Size.Width + 10, dataGridView1.Size.Height + 10);
            dataGridView1.DrawToBitmap(bmp, dataGridView1.Bounds);
            e.Graphics.DrawImage(bmp, 0, 0);
        }
    }
}

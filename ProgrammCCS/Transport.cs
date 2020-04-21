using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excell = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Drawing.Printing;
using Word = Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using System.Deployment.Application;
using System.Reflection;
using System.Threading;
using System.Media;

namespace ProgramCCS
{
    public partial class Transport : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели
        public Transport()
        {  
            InitializeComponent();
            comboBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            comboBox2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            comboBox9.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox8.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox9.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
        }
        public void select_avto()//Вывод авто в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT avto FROM [avto]";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow row in dt.Rows)
            {
                comboBox1.Items.Add(row[0].ToString());
                comboBox4.Items.Add(row[0].ToString());
                comboBox5.Items.Add(row[0].ToString());
                comboBox6.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение
        }
        public void select_zapchast()//Вывод запчастей в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT zapchast FROM [zapchast]";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow row in dt.Rows)
            {
                comboBox2.Items.Add(row[0].ToString());
                comboBox3.Items.Add(row[0].ToString());
                comboBox7.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение
        }
        public void select_postavshik()//Вывод поставщиков в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT postavshik FROM [postavshik]";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow row in dt.Rows)
            {
                comboBox8.Items.Add(row[0].ToString());
                comboBox9.Items.Add(row[0].ToString());
                comboBox10.Items.Add(row[0].ToString());
                comboBox11.Items.Add(row[0].ToString());
                comboBox12.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение
        }
        public void disp_data()
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            //cmd.CommandText = "SELECT TOP 1000 * FROM [Table_1] ORDER BY data_zapisi DESC";//последние 1000 записей
            cmd.CommandText = "SELECT * FROM [Table_avto] WHERE (date_remont BETWEEN @StartDate AND @EndDate) ORDER BY date_remont DESC";
            cmd.Parameters.AddWithValue("@StartDate", DateTime.Today.AddMonths(-12));
            cmd.Parameters.AddWithValue("@EndDate", DateTime.Today);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
            signal();
            dataGridView1.Visible = true;
            dataGridView2.Visible = false;
        }
        public void signal()//Оповещания
        {
            for (int i = 0; i < dataGridView1.Rows.Count-1; i++)//Цикл
            {
                DateTime date_remont = Convert.ToDateTime(dataGridView1.Rows[i].Cells[5].Value);//Дата ремонта
                int service = Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);
                DateTime date = DateTime.Today.AddMonths(-service);
                if (date_remont <= date)
                {
                    SystemSounds.Beep.Play();
                    linkLabel1.Visible = true;
                    linkLabel1.Text = ("Внимание! Срок эксплуатации подходит к завершению, пробег пройден!");
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_avto] SET status = @status WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
                    cmd.Parameters.AddWithValue("@status", "пробег пройден");
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
                else if (date_remont.AddDays(-30) <= date)//Остался месяц
                {
                    SystemSounds.Beep.Play();
                    linkLabel2.Visible = true;
                    linkLabel2.Text = ("Внимание! Срок эксплуатации подходит к завершению, остался один месяц!");
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_avto] SET status = @status WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
                    cmd.Parameters.AddWithValue("@status", "внимание");
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
            }
            podschet();
        }
        public void podschet()//Произвести подсчет dataGridView1 и dataGridView5 и dataGridView2
        {
            if (dataGridView1.Visible == true)
            {
                //Сумма столбца стоимость
                double summa = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox4.Visible = true;
                textBox4.Text = summa.ToString() + " Сом";
                               
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1[i, j].Value != null)
                        {
                            textBox5.Text = Convert.ToString(dataGridView1.Rows.Count-1) + " Штук";// -1 это нижняя пустая строка
                            count++;
                            break;
                        }
                    }
                }
            }
            else if (dataGridView2.Visible == true)
            {
                //Сумма столбца стоимость
                double summa = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[6].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox4.Visible = true;
                textBox4.Text = summa.ToString() + " Сом";
                
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView2.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (dataGridView2[i, j].Value != null)
                        {
                            textBox5.Text = Convert.ToString(dataGridView2.Rows.Count-1) + " Штук";// -1 это нижняя пустая строка
                            count++;
                            break;
                        }
                    }
                }           
            }
        }
        private void linkLabel1_Click(object sender, EventArgs e)
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT avto,zapchast,service,probeg,date_remont,status,summ,postavshik,id FROM [Table_avto] WHERE status = @status ORDER BY date_remont DESC";
            cmd.Parameters.AddWithValue("@status", "пробег пройден");
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
            linkLabel1.Visible = false;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            podschet();
        }
        private void linkLabel2_Click(object sender, EventArgs e)
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT avto,zapchast,service,probeg,date_remont,status,summ,postavshik,id FROM [Table_avto] WHERE status = @status ORDER BY date_remont DESC";
            cmd.Parameters.AddWithValue("@status", "внимание");
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
            linkLabel2.Visible = false;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            podschet();
        }
        private void button1_Click(object sender, EventArgs e)//Добавить
        {
            if (textBox1.Text != "" & textBox8.Text != "" & textBox2.Text != "" & comboBox9.Text != "" & comboBox1.Text != "" & comboBox2.Text != "")
            {
                var summ = Convert.ToDouble(textBox2.Text.Replace(',', '.'));//запятую превратить в точку
                int service = Convert.ToInt32(textBox1.Text);
                service = int.Parse(textBox1.Text);
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_avto] (avto,zapchast,service,probeg,date_remont,status,summ,postavshik) VALUES (@avto,@zapchast,@service,@probeg,@date_remont,@status,@summ,@postavshik)", con);
                cmd.Parameters.AddWithValue("@avto", comboBox1.Text);
                cmd.Parameters.AddWithValue("@zapchast", comboBox2.Text);
                cmd.Parameters.AddWithValue("@service", Math.Round(service / 1.66666666));
                cmd.Parameters.AddWithValue("@probeg", textBox8.Text);
                cmd.Parameters.AddWithValue("@date_remont", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@status", "обслуживание не требуется");
                cmd.Parameters.AddWithValue("@summ", summ);
                cmd.Parameters.AddWithValue("@postavshik", comboBox9.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение 
                con.Open();//открыть соединение
                SqlCommand cmd1 = new SqlCommand("INSERT INTO [Table_avto_arhiv] (avto,zapchast,service,probeg,date_remont,status,summ,postavshik) VALUES (@avto,@zapchast,@service,@probeg,@date_remont,@status,@summ,@postavshik)", con);
                cmd1.Parameters.AddWithValue("@avto", comboBox1.Text);
                cmd1.Parameters.AddWithValue("@zapchast", comboBox2.Text);
                cmd1.Parameters.AddWithValue("@service", Math.Round(service / 1.66666666));
                cmd1.Parameters.AddWithValue("@probeg", textBox8.Text);
                cmd1.Parameters.AddWithValue("@date_remont", dateTimePicker1.Value);
                cmd1.Parameters.AddWithValue("@status", "обслуживание не требуется");
                cmd1.Parameters.AddWithValue("@summ", summ);
                cmd1.Parameters.AddWithValue("@postavshik", comboBox9.Text);
                cmd1.ExecuteNonQuery();
                con.Close();//закрыть соединение              
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //очистка текстовых полей
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            comboBox9.SelectedIndex = -1;
            comboBox10.SelectedIndex = -1;
            comboBox11.SelectedIndex = -1;
            textBox2.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            comboBox1.Select();//Установка курсора
            disp_data();
            podschet();
            disp_data();
        }
        private void button7_Click(object sender, EventArgs e)//Изменить
        {
            if (textBox7.Text == "Admin" & dataGridView2.Rows.Count == 2)
            {
                //double summ = double.Parse(textBox2.Text.Replace('.', ','));//ввод суммы
                //textBox2.Text = textBox2.Text.Replace(".", ",");
                var summ = Convert.ToDouble(textBox2.Text.Replace(',', '.'));//запятую превратить в точку
                int service = Convert.ToInt32(textBox1.Text);
                service = int.Parse(textBox1.Text);
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_avto] SET avto=@avto,zapchast=@zapchast,probeg=@probeg,date_remont=@date_remont,summ=@summ,postavshik=@postavshik WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.CurrentRow.Cells[8].Value));//первая строка в гриде
                if (comboBox1.Text != "") { cmd.Parameters.AddWithValue("@avto", comboBox1.Text); }
                else if (comboBox1.Text == "") { cmd.Parameters.AddWithValue("@avto", Convert.ToString(dataGridView2.CurrentRow.Cells[0].Value)); }
                if (comboBox2.Text != "") { cmd.Parameters.AddWithValue("@zapchast", comboBox2.Text); }
                else if (comboBox2.Text == "") { cmd.Parameters.AddWithValue("@zapchast", Convert.ToString(dataGridView2.CurrentRow.Cells[1].Value)); }
                if (textBox8.Text != "") { cmd.Parameters.AddWithValue("@probeg", textBox8.Text); }
                else if (textBox8.Text == "") { cmd.Parameters.AddWithValue("@probeg", Convert.ToString(dataGridView2.CurrentRow.Cells[3].Value)); }
                if (textBox2.Text != "") { cmd.Parameters.AddWithValue("@summ", summ); }
                else if (textBox2.Text == "") { cmd.Parameters.AddWithValue("@summ", Convert.ToString(dataGridView2.CurrentRow.Cells[6].Value)); }
                if (comboBox9.Text != "") { cmd.Parameters.AddWithValue("@postavshik", comboBox9.Text); }
                else if (comboBox9.Text == "") { cmd.Parameters.AddWithValue("@postavshik", Convert.ToString(dataGridView2.CurrentRow.Cells[7].Value)); }
                cmd.Parameters.AddWithValue("@date_remont", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение 
                con.Open();//открыть соединение
                SqlCommand cmd1 = new SqlCommand("UPDATE [Table_avto_arhiv] SET avto=@avto,zapchast=@zapchast,probeg=@probeg,date_remont=@date_remont,summ=@summ,postavshik=@postavshik WHERE id = @id", con);
                cmd1.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.CurrentRow.Cells[8].Value));//первая строка в гриде
                if (comboBox1.Text != "") { cmd1.Parameters.AddWithValue("@avto", comboBox1.Text); }
                else if (comboBox1.Text == "") { cmd1.Parameters.AddWithValue("@avto", Convert.ToString(dataGridView2.CurrentRow.Cells[0].Value)); }
                if (comboBox2.Text != "") { cmd1.Parameters.AddWithValue("@zapchast", comboBox2.Text); }
                else if (comboBox2.Text == "") { cmd1.Parameters.AddWithValue("@zapchast", Convert.ToString(dataGridView2.CurrentRow.Cells[1].Value)); }
                if (textBox8.Text != "") { cmd1.Parameters.AddWithValue("@probeg", textBox8.Text); }
                else if (textBox8.Text == "") { cmd1.Parameters.AddWithValue("@probeg", Convert.ToString(dataGridView2.CurrentRow.Cells[3].Value)); }
                if (textBox2.Text != "") { cmd1.Parameters.AddWithValue("@summ", summ); }
                else if (textBox2.Text == "") { cmd1.Parameters.AddWithValue("@summ", Convert.ToString(dataGridView2.CurrentRow.Cells[6].Value)); }
                if (comboBox9.Text != "") { cmd1.Parameters.AddWithValue("@postavshik", comboBox9.Text); }
                else if (comboBox9.Text == "") { cmd1.Parameters.AddWithValue("@postavshik", Convert.ToString(dataGridView2.CurrentRow.Cells[7].Value)); }
                cmd1.Parameters.AddWithValue("@date_remont", dateTimePicker1.Value);
                cmd1.ExecuteNonQuery();
                con.Close();//закрыть соединение 
                MessageBox.Show("Вы успешно изменили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                disp_data();
            }
            //очистка текстовых полей
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            comboBox9.SelectedIndex = -1;
            comboBox10.SelectedIndex = -1;
            comboBox11.SelectedIndex = -1;
            textBox2.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            comboBox1.Select();//Установка курсора
            disp_data();
            podschet();
            disp_data();
        }
        private void button6_Click(object sender, EventArgs e)//Удалить
        {
            if (textBox7.Text == "Admin" & dataGridView2.Rows.Count == 2)
            {
                if (MessageBox.Show("Вы хотите удалить эту запись?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("DELETE FROM [Table_avto] WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.CurrentRow.Cells[8].Value));//первая строка в гриде
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                    con.Open();//открыть соединение
                    SqlCommand cmd1 = new SqlCommand("DELETE FROM [Table_avto_arhiv] WHERE id = @id", con);
                    cmd1.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.CurrentRow.Cells[8].Value));//первая строка в гриде
                    cmd1.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                    MessageBox.Show("Запись успешно удалена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    disp_data();
                }
                else
                {
                    dataGridView1.Visible = true;
                    dataGridView2.Visible = false;
                    disp_data();
                    podschet();
                    disp_data();
                }
            }
        }
        private void button4_Click(object sender, EventArgs e)//Обновить
        {
            dataGridView1.Visible = true;
            dataGridView2.Visible = false;
            //очистка текстовых полей
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            comboBox9.SelectedIndex = -1;
            comboBox10.SelectedIndex = -1;
            comboBox11.SelectedIndex = -1;
            textBox2.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            disp_data();
            podschet();
            disp_data();
        }

        private void button2_Click(object sender, EventArgs e)//Найти
        {
            if (comboBox3.Text != "" || comboBox4.Text != "" || comboBox10.Text != "")
            {
                dataGridView1.Visible = false;
                dataGridView2.Visible = true;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT avto,zapchast,service,probeg,date_remont,status,summ,postavshik,id FROM [Table_avto]" +
                    " WHERE avto = @avto AND zapchast LIKE '%" + Convert.ToString(comboBox3.Text) + "%' AND postavshik LIKE '%" + Convert.ToString(comboBox10.Text) + "%' ORDER BY avto", con);
                cmd.Parameters.AddWithValue("@avto", comboBox4.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
            }
            else MessageBox.Show("Необходимо выбрать автомобиль или запчасть", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            podschet();
        }
        private void button3_Click(object sender, EventArgs e)//Произвести ремонт
        {
            if (textBox3.Text != "" & comboBox11.Text != "" & comboBox3.Text != "" & comboBox4.Text != "" & dataGridView2.Rows.Count == 2)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_avto] SET date_remont = @date_remont, status = @status, summ = @summ, postavshik = @postavshik WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[8].Value));//первая строка в гриде
                cmd.Parameters.AddWithValue("@date_remont", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("@status", "обслуживание не требуется");
                cmd.Parameters.AddWithValue("@summ", textBox3.Text);
                cmd.Parameters.AddWithValue("@postavshik", comboBox11.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                con.Open();//открыть соединение
                SqlCommand cmd1 = new SqlCommand("INSERT INTO [Table_avto_arhiv] (avto,zapchast,service,probeg,date_remont,status,summ,postavshik) VALUES (@avto,@zapchast,@service,@probeg,@date_remont,@status,@summ,@postavshik)", con);
                cmd1.Parameters.AddWithValue("@avto", Convert.ToString(dataGridView2.Rows[0].Cells[0].Value));
                cmd1.Parameters.AddWithValue("@zapchast", Convert.ToString(dataGridView2.Rows[0].Cells[1].Value));
                cmd1.Parameters.AddWithValue("@service", Convert.ToInt32(dataGridView2.Rows[0].Cells[2].Value));
                cmd1.Parameters.AddWithValue("@probeg", Convert.ToInt32(dataGridView2.Rows[0].Cells[3].Value));
                cmd1.Parameters.AddWithValue("@date_remont", dateTimePicker2.Value);
                cmd1.Parameters.AddWithValue("@status", "обслуживание не требуется");
                cmd1.Parameters.AddWithValue("@summ", textBox3.Text);
                cmd1.Parameters.AddWithValue("@postavshik", comboBox11.Text);
                cmd1.ExecuteNonQuery();
                con.Close();//закрыть соединение 
                MessageBox.Show("Ремонт успешно обновлен", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                comboBox4.SelectedIndex = -1;
                comboBox3.SelectedIndex = -1;
            }  
            else MessageBox.Show("Сначала произведите поиск", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            disp_data();
            podschet();
            disp_data();
        }
        private void Form3_avto_Load(object sender, EventArgs e)//Загрузка формы
        {
            dataGridView1.Visible = true;
            dataGridView2.Visible = false;          
            //-----------------Окраска Гридов-------------------//
            DataGridViewRow row1 = this.dataGridView1.RowTemplate;
            row1.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row1.Height = 5;
            row1.MinimumHeight = 17;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;//цвет заголовка
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            DataGridViewRow row2 = this.dataGridView2.RowTemplate;
            row2.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row2.Height = 5;
            row2.MinimumHeight = 17;
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSlateGray;//цвет заголовка
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            DataGridViewRow row3 = this.dataGridView3.RowTemplate;
            row3.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row3.Height = 5;
            row3.MinimumHeight = 17;
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            dataGridView3.EnableHeadersVisualStyles = false;
            dataGridView3.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSlateGray;//цвет заголовка
            dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            //----------------Окраска Гридов--------------------//           
            dateTimePicker1.Value = DateTime.Today.AddDays(0);
            dateTimePicker2.Value = DateTime.Today.AddDays(0);
                select_avto();//Вывод авто в Combobox
                select_zapchast();//Вывод запчастей в Combobox
                select_postavshik();//Вывод поставщиков в Combobox
            disp_data();
            podschet();
            disp_data();
        }
        private void Form3_avto_FormClosed(object sender, FormClosedEventArgs e)//Закрытие формы Выход
        {
            Application.Exit();
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)//окраска
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                DateTime date_remont = Convert.ToDateTime(dataGridView1.Rows[i].Cells[5].Value);//Дата ремонта
                int service = Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value);
                DateTime date = DateTime.Today.AddMonths(-service);
                if (date_remont <= date)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;//красный
                }
                else if (date_remont.AddDays(-30) <= date)//Остался месяц
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;//желтый                  
                }
            }
        }      
        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
            {
                //DateTime date_remont = Convert.ToDateTime(dataGridView2.Rows[i].Cells[4].Value);//Дата ремонта
                //int service = Convert.ToInt32(dataGridView2.Rows[i].Cells[2].Value);
                //DateTime date = DateTime.Today.AddMonths(service);
                //if (date_remont >= date)
                //{
                //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;//красный
                //}
                //else if (date_remont == date.AddDays(-30))//Остался месяц
                //{
                //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;//желтый                  
                //}
            }
        }

        private void button5_Click(object sender, EventArgs e)//отчет
        {
            button5.Enabled = false;
            button5.Text = "Ожидайте!";
            if (comboBox5.Text != "" | dateTimePicker3.Value <= DateTime.Today.AddDays(-1))
            {
                dataGridView1.Visible = false;
                dataGridView2.Visible = true;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT avto,zapchast,service,probeg,date_remont,status,summ,postavshik,id FROM [Table_avto_arhiv]" +
                    " WHERE avto LIKE '%" + Convert.ToString(comboBox5.Text) + "%' AND postavshik LIKE '%" + Convert.ToString(comboBox12.Text) + "%' AND date_remont BETWEEN @StartDate AND @EndDate ORDER BY avto", con);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker3.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker4.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение 
                podschet();

            //Выдача рееста в WORD
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Word Documents (*.docx)|*.docx";
            sfd.FileName = "Сводный Отчет УТБиМТО с " + Convert.ToString(dateTimePicker3.Value.ToString("dd.MM.yyyy")) + " по " + Convert.ToString(dateTimePicker4.Value.ToString("dd.MM.yyyy")) + ".docx";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
               Otchet_To_Word(dataGridView2, sfd.FileName);
            }
      }
            else MessageBox.Show("Выбирите диапазон", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            button5.Enabled = true;
            button5.Text = "Отчет";
        }
        public void Otchet_To_Word(DataGridView dataGridView2, string filename)//Метод экспорта в Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Отчет");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 12;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 8; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView2.Rows.Count != 0)
            {
                //удаление столбца
                //this.dataGridView1.Columns.RemoveAt(4);//дата записи

                string kol_vo = Convert.ToString(textBox5.Text);//кол-во
                string sum = Convert.ToString(textBox4.Text);//сумма
                DateTime month = Convert.ToDateTime(dataGridView2.Rows[0].Cells[4].Value);
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count - 1;// столбцы в гриде (-5 последних)             
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r + 1, c] = dataGridView2.Rows[r].Cells[c].Value;// +1 это первая строка в таблице над заголовком
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ

                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = "                                                                                   Кол-во: " + kol_vo + "     Сумма: " + sum +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + 
                Environment.NewLine + 
                Environment.NewLine + 
                Environment.NewLine +
                Environment.NewLine + "Начальник транспортного отдела " +
                Environment.NewLine +
                Environment.NewLine + "______________________________" +
                Environment.NewLine;

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
                //стиль строки заголовка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = "";
                    oDoc.Application.Selection.Tables[1].Cell(2, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;
                }
                //стиль таблицы   
                oDoc.Application.Selection.Tables[1].Columns[3].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Columns[3].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Columns[4].Delete();//Удалить столбец

                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 120;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 120;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 120;//ширина столбца
                //oDoc.Application.Selection.Tables[1].Columns[6].Width = 60;//ширина столбца
                //oDoc.Application.Selection.Tables[1].Columns[7].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -30;//Установка отступа слева               
                //oDoc.Application.Selection.Tables[1].Cell(1, 2).Range.Text = "текст в ячейке";
                //oDoc.Application.Selection.Tables[1].Cell(1, 2).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 4));//Объединение
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text =
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine + "Сводный отчет" +                    
                    Environment.NewLine + "по ремонту и обслуживанию автомобилей" +
                    Environment.NewLine + "ГП Службы специальной связи" +              
                    Environment.NewLine + "с " + Convert.ToString(dateTimePicker3.Value.ToString("dd.MM.yyyy")) + " по " + Convert.ToString(dateTimePicker4.Value.ToString("dd.MM.yyyy")) +
                    Environment.NewLine;

                    headerRange.Font.Size = 14;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "Служба специальной связи       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }

        private void button8_Click(object sender, EventArgs e)//Добавить авто,поставщик,запчасть
        {
            if (textBox6.Text == "Admin" & comboBox7.Text == "" & comboBox8.Text == "" & comboBox6.Text != "" & comboBox6.Text != Convert.ToString(dataGridView3.Rows[0].Cells[0].Value))
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [avto] (avto) VALUES (@avto)", con);
                cmd.Parameters.AddWithValue("@avto", comboBox6.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (textBox6.Text == "Admin" & comboBox7.Text != "" & comboBox8.Text == "" & comboBox6.Text == "" & comboBox7.Text != Convert.ToString(dataGridView3.Rows[0].Cells[0].Value))
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [zapchast] (zapchast) VALUES (@zapchast)", con);
                cmd.Parameters.AddWithValue("@zapchast", comboBox7.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (textBox6.Text == "Admin" & comboBox7.Text == "" & comboBox8.Text != "" & comboBox6.Text == "" & comboBox8.Text != Convert.ToString(dataGridView3.Rows[0].Cells[0].Value))
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [postavshik] (postavshik) VALUES (@postavshik)", con);
                cmd.Parameters.AddWithValue("@postavshik", comboBox8.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (comboBox6.Text == Convert.ToString(dataGridView3.Rows[0].Cells[0].Value) | dataGridView3.Rows.Count != 0)
            {
                MessageBox.Show("Данная запись уже существует в базе!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (comboBox7.Text == Convert.ToString(dataGridView3.Rows[0].Cells[0].Value) | dataGridView3.Rows.Count != 0)
            {
                MessageBox.Show("Данная запись уже существует в базе!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else if (comboBox8.Text == Convert.ToString(dataGridView3.Rows[0].Cells[0].Value) | dataGridView3.Rows.Count != 0)
            {
                MessageBox.Show("Данная запись уже существует в базе!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            else MessageBox.Show("Добавлять можно только по одному либо у Вас нет доступа к базе!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //очистка текстовых полей
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            comboBox7.SelectedIndex = -1;
            comboBox8.SelectedIndex = -1;
            comboBox9.SelectedIndex = -1;
            comboBox10.SelectedIndex = -1;
            comboBox11.SelectedIndex = -1;
        }

        private void comboBox6_TextChanged(object sender, EventArgs e)//
        {
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT avto FROM [avto]" +
                "WHERE avto LIKE '%" + Convert.ToString(comboBox6.Text) + "%'", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView3.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            dataGridView3.Columns[0].HeaderText = "Автомобиль";
            if (comboBox6.Text == "")//если поле очищено, отобразить базу
            {
                dt.Clear();//чистим DataTable, если он был не пуст
                foreach (DataRow row in dt.Rows)
                {
                    comboBox6.Items.Add(row[0].ToString());
                }
            }
        }
        private void comboBox7_TextChanged(object sender, EventArgs e)
        {
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT zapchast FROM [zapchast]" +
                "WHERE zapchast LIKE '%" + Convert.ToString(comboBox7.Text) + "%'", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView3.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            dataGridView3.Columns[0].HeaderText = "Запчасть";
            if (comboBox7.Text == "")//если поле очищено, отобразить базу
            {
                dt.Clear();//чистим DataTable, если он был не пуст
                foreach (DataRow row in dt.Rows)
                {
                    comboBox7.Items.Add(row[0].ToString());
                }
            }
        }
        private void comboBox8_TextChanged(object sender, EventArgs e)
        {
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT postavshik FROM [postavshik]" +
                "WHERE postavshik LIKE '%" + Convert.ToString(comboBox8.Text) + "%'", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView3.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            dataGridView3.Columns[0].HeaderText = "Поставщик";
            if (comboBox8.Text == "")//если поле очищено, отобразить базу
            {
                dt.Clear();//чистим DataTable, если он был не пуст
                foreach (DataRow row in dt.Rows)
                {
                    comboBox8.Items.Add(row[0].ToString());
                }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)//сумма выделенных строк и колл-во
        {
            //Колл-во
            textBox5.Text = Convert.ToString(dataGridView1.SelectedRows.Count) + " Штук";
            //Сумма столбца стоимость
            double summa = 0;
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                double incom;
                double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                summa += incom;
            }
            textBox4.Visible = true;
            textBox4.Text = summa.ToString() + " Сом" ; 
            
        }
        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            //Колл-во
            textBox5.Text = Convert.ToString(dataGridView2.SelectedRows.Count) + " Штук";
            //Сумма столбца стоимость
            double summa = 0;
            foreach (DataGridViewRow row in dataGridView2.SelectedRows)
            {
                double incom;
                double.TryParse((row.Cells[6].Value ?? "0").ToString().Replace(".", ","), out incom);
                summa += incom;
            }
            textBox4.Visible = true;
            textBox4.Text = summa.ToString() + " Сом";
        }
    }
}

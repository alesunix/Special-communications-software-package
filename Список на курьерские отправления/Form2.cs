using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Deployment.Application;
using System.Reflection;
using System.Data.SqlClient;

namespace Список_на_курьерские_отправления
{
    public partial class Form2 : Form
    {
        public string Data
        { get { return comboBoxF2.Text; } }
        public string tlc = "TLC-Express";
        public string osh = "Ошский филиал";
        public string dj = "Джалал-Абад филиал";
        public string avto = "Транспортный";
        public string sklad = "Склад";
        
        public Form2()
        {
            InitializeComponent();
            this.Text += "  Версия - " + CurrentVersion; //Добавляем в название программы, версию.
            textBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "Войти" с клавиатуры
            comboBoxF2.Text = tlc;//отображение TLC-Express 
        }     
        public string CurrentVersion//Версия программы
        {
            get
            {
                return ApplicationDeployment.IsNetworkDeployed
                      ? ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString()
                      : Assembly.GetExecutingAssembly().GetName().Version.ToString();
            }
        }
        private void button1_Click(object sender, EventArgs e)//Войти
        {
            Form1 form1 = new Form1();          
            form1.con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT * FROM [Table_Login] WHERE login = @login", form1.con);
            cmd.Parameters.AddWithValue("@login", comboBoxF2.Text);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            DGVF2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            form1.con.Close();//Закрываем соединение
            //Access._Access = DGVF2.Rows[0].Cells[3].Value.ToString();//Доступ
            //Access.Name = comboBoxF2.Text;//Login
            Person Access = new Person(comboBoxF2.Text, DGVF2.Rows[0].Cells[3].Value.ToString());// передать значения Конструктору в классе Person

            if (comboBoxF2.Text == tlc & textBox1.Text == "147258369" || comboBoxF2.Text == osh & textBox1.Text == "123789" || comboBoxF2.Text == dj & textBox1.Text == "123789" || comboBoxF2.Text == "root" & textBox1.Text == "root")
            {
                //Clipboard.SetText(comboBoxF2.Text);//Скопировать текст в буфер обмена
                //P.label1.Text = "Добро пожаловать! " + comboBoxF2.Text;
                form1.Show();
                this.Hide();
            }
            else if (comboBoxF2.Text == sklad & textBox1.Text == "159753")
            {
                //Clipboard.SetText(comboBoxF2.Text);//Скопировать текст в буфер обмена
                Sklad sklad = new Sklad();
                //P.label1.Text = "Добро пожаловать! " + comboBoxF2.Text;
                sklad.Show();
                this.Hide();
            }
            else if (comboBoxF2.Text == avto & textBox1.Text == "123654")
            {
                //Clipboard.SetText(comboBoxF2.Text);//Скопировать текст в буфер обмена
                Form3_avto avto = new Form3_avto();
                //P.label1.Text = "Добро пожаловать! " + comboBoxF2.Text;
                avto.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Неверный пароль", "Внимание!");
            }
        }
        void miganie(object sender, EventArgs e)//Метод мигания
        {
            label3.Visible = !label3.Visible;
        }
        public void Logins_select()//Вывод пользователей в Combobox
        {
            Form1 form1 = new Form1();
            form1.con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT login FROM [Table_Login] WHERE login NOT IN ('root') ORDER BY login", form1.con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            //DGVF1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            foreach (DataRow row in dt.Rows)
            {
                comboBoxF2.Items.Add(row[0].ToString());
            }
            form1.con.Close();//Закрываем соединение          
        }
        private void Form2_Load(object sender, EventArgs e)//Загрузка формы
        {
            textBox1.PasswordChar = '*';//Скрыть пароль
            DGVF2.Visible = false;
            Logins_select();
            label26.Text = "Версия - " + CurrentVersion;//Версия
            //Мигание кнопки
            Timer t = new Timer();
            t.Interval = 400;
            t.Tick += new EventHandler(miganie);
            t.Start();
            comboBoxF2.SelectedIndex = 0;
            textBox1.Select();//Установить курсор
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)// ссылка на страничку
        {
            System.Diagnostics.Process.Start("https://www.facebook.com/alesunix");
            linkLabel1.BackColor = Color.Transparent;
        }

        private void comboBoxF2_SelectedIndexChanged(object sender, EventArgs e)//Установить курсор после выбора
        {
            textBox1.Select();
        }

        private void Form2_FormClosed(object sender, FormClosedEventArgs e)//Закрытие формы
        {
            Application.Exit();
        }
    }
}

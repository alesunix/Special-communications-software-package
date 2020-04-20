using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Deployment.Application;
using System.Reflection;
using System.Data.SqlClient;

namespace ProgramCCS
{
    public partial class Login : Form
    {
        public SqlConnection con = new SqlConnection(@"Data Source=192.168.0.3;Initial Catalog=ccsbase;Persist Security Info=True;User ID=Lan;Password=Samsung0");

        public string tlc = "TLC-Express";
        public string osh = "Ошский филиал";
        public string dj = "Джалал-Абад филиал";
        public string avto = "Транспортный";
        public string sklad = "Склад";
        
        public Login()
        {
            InitializeComponent();
            //this.Text += "  Версия - " + CurrentVersion; //Добавляем в название программы, версию.
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
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT * FROM [Table_Login] WHERE login = @login", con);
            cmd.Parameters.AddWithValue("@login", comboBoxF2.Text);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            con.Close();//Закрываем соединение
            //Person.Access = dt.Rows[0][3].ToString();//Доступ
            //Person.Name = dt.Rows[0][1].ToString();//Login
            Person Access = new Person(comboBoxF2.Text, dt.Rows[0][3].ToString());// передать значения Конструктору в классе Person (имя и доступ)

            if (comboBoxF2.Text == tlc & textBox1.Text == dt.Rows[0][2].ToString() 
                || comboBoxF2.Text == osh & textBox1.Text == dt.Rows[0][2].ToString() 
                || comboBoxF2.Text == dj & textBox1.Text == dt.Rows[0][2].ToString() 
                || comboBoxF2.Text == "root" & textBox1.Text == dt.Rows[0][2].ToString())
            {
                //P.label1.Text = "Добро пожаловать! " + comboBoxF2.Text;
                TLC form1 = new TLC();
                form1.Show();
                this.Hide();
            }
            else if (comboBoxF2.Text == sklad & textBox1.Text == dt.Rows[0][2].ToString())
            {
                Inventory_control sklad = new Inventory_control();
                //P.label1.Text = "Добро пожаловать! " + comboBoxF2.Text;
                sklad.Show();
                this.Hide();
            }
            else if (comboBoxF2.Text == avto & textBox1.Text == dt.Rows[0][2].ToString())
            {
                Transport avto = new Transport();
                //P.label1.Text = "Добро пожаловать! " + comboBoxF2.Text;
                avto.Show();
                this.Hide();
            }
            else
            {
                MessageBox.Show("Неверный пароль", "Внимание!");
            }
        }
        
        public void Logins_select()//Вывод пользователей в Combobox
        {
            TLC form1 = new TLC();
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT login FROM [Table_Login] WHERE login NOT IN ('root') ORDER BY login", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow row in dt.Rows)
            {
                comboBoxF2.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение          
        }

        void Miganie(object sender, EventArgs e)//Метод мигания
        {
            label26.Visible = !label26.Visible;
        }
        private void Form2_Load(object sender, EventArgs e)//Загрузка формы
        {
            textBox1.PasswordChar = '*';//Скрыть пароль
            Logins_select();
            label26.Text = "Версия - " + CurrentVersion;//Версия
            //Мигание кнопки//
            Timer t = new Timer();
            t.Interval = 400;
            t.Tick += new EventHandler(Miganie);
            t.Start();
            comboBoxF2.SelectedIndex = 0;
            textBox1.Select();//Установить курсор
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)// ссылка на страничку
        {
            System.Diagnostics.Process.Start("https://alesunix.github.io/");
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

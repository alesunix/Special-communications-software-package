using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excell = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Deployment.Application;
using System.Reflection;
using System.Threading;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Data.Linq;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Diagnostics;
using System.Collections.Generic;

namespace ProgramCCS
{
    public partial class TLC : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели
        DataContext db = new DataContext(Connection.con);//Для работы LINQ to SQL
        
        private string fileName = string.Empty;

        Login formLogin = new Login();
        public object loker = new object();

        public TLC()
        {
            InitializeComponent();
            Text += "  Версия - " + CurrentVersion; //Добавляем в название программы, версию.
            //comboBox8.Text = Properties.Settings.Default.Prichina_vozvrat; // Загружаем ранее сохраненный текст
            //Properties.Settings.Default.Save();  // Сохраняем переменные.
            dataGridView2.KeyDown += (s, e) => { if (e.KeyCode == Keys.F5) Status_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dataGridView2.KeyDown += (s, e) => { if (e.KeyCode == Keys.F4) Edit_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dataGridView1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) Delete_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dataGridView2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) Delete_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dataGridView3.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) Delete_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dataGridView5.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) Delete_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
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
        void Miganie(object sender, EventArgs e)//Метод мигания
        {
            label_filial.Visible = !label_filial.Visible;
        }
        void TimeLabel(object sender, EventArgs e)//Часы
        {
            label10.Text = DateTime.Now.ToLongTimeString();
        }
        public void ProgressBar()//ПрогрессБар
        {
            lock (loker)
            {
                //---------------ПрогрессБар--------------------//
                progressBar1.Visible = true;
                progressBar1.Maximum = 101;
                progressBar1.Value = 0;
                Thread t = new Thread(new ThreadStart(delegate
                {
                    for (int i = 0; i < 101; i++)
                    {
                        Invoke(new ThreadStart(delegate
                        {
                            progressBar1.Value++;
                        }));
                    }
                }));
                t.Start();
            }

        }
        //---------------------------------------------------------------------//
        private void Form1_Load(object sender, EventArgs e)//Загрузка формы
        {
            label_filial.Text = Person.Name;
            //----------------------------------------//***
            if (Person.Access == "Low")
            {
                tabPage4.Enabled = false;
                button1.Enabled = false;
                toolStripButton3.Enabled = false;
                toolStripButton1.Enabled = false;
            }
            else if (Person.Access == "Medium")
            {
                tabPage4.Enabled = false;
                toolStripButton1.Enabled = false;
            }
            else if (Person.Access == "root")
            {
                tabPage4.Enabled = true;
                toolStripButton1.Enabled = true;
            }
            //Мигание кнопки и Обновление времени
            System.Windows.Forms.Timer timer = new System.Windows.Forms.Timer();
            timer.Interval = 777;
            timer.Tick += new EventHandler(Miganie);
            timer.Tick += new EventHandler(TimeLabel);
            timer.Start();
            //-----------------Окраска Гридов-------------------//
            DataGridViewRow row1 = this.dataGridView1.RowTemplate;
            row1.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row1.Height = 5;
            row1.MinimumHeight = 17;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView1.DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;//цвет заголовка
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            DataGridViewRow row2 = this.dataGridView2.RowTemplate;
            row2.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row2.Height = 5;
            row2.MinimumHeight = 17;
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView2.DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(83, 134, 166);//цвет заголовка
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            DataGridViewRow row5 = this.dataGridView5.RowTemplate;
            row5.DefaultCellStyle.BackColor = Color.LightSkyBlue;
            row5.Height = 5;
            row5.MinimumHeight = 17;
            dataGridView5.EnableHeadersVisualStyles = false;
            dataGridView5.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView5.DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView5.ColumnHeadersDefaultCellStyle.BackColor = Color.LightCoral;//цвет заголовка
            dataGridView5.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            DataGridViewRow row6 = this.dataGridView6.RowTemplate;
            row6.DefaultCellStyle.BackColor = Color.LightSkyBlue;
            row6.Height = 5;
            row6.MinimumHeight = 17;
            dataGridView6.EnableHeadersVisualStyles = false;
            dataGridView6.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView6.DefaultCellStyle.Font = new System.Drawing.Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView6.ColumnHeadersDefaultCellStyle.BackColor = Color.LightCoral;//цвет заголовка
            dataGridView6.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            //----------------Окраска Гридов--------------------//
            dataGridView2.Visible = true;
            dataGridView6.Visible = false;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;

            label26.Text = "Версия - " + CurrentVersion;

            dateTimePicker5.Value = DateTime.Today;
            //dataGridView2.Columns[7].DefaultCellStyle.Format = "dd.MM.yyyy";

            ////пример
            ////comboBox2.SelectedItem - это объект, программа не знает что это за объект, поэтому нам нужно
            ////явно указать что это объект класса ClassComboBox, а дальше работать с ним как с классом
            //comboBox2.SelectedIndex = 0;
            //string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
            //MessageBox.Show(comboitem.ToString());//проверка

            //Програмное добавление строк в ComboBox
            //Convert.ToString(comboBox3.Items.Add("ОсОО Тенгри"));
            //Convert.ToString(comboBox3.Items.Add("ИП 'JUMPER'"));
            //Convert.ToString(comboBox3.Items.Add("OcOO 'Кыргыз Сервис Логистик'"));
            //Convert.ToString(comboBox3.Items.Add("ОсОО 'Экспресс-Тайм'"));
            //Convert.ToString(comboBox3.Items.Add("ОсОО 'Альфа Вита'"));
            //Convert.ToString(comboBox3.Items.Add("ОсОО Мастер групп. Кейджизет"));
            //Convert.ToString(comboBox3.Items.Add("ОсОО Kyrgyz Express Post"));
            //Convert.ToString(comboBox3.Items.Add("Рокит профит"));
            //Convert.ToString(comboBox3.Items.Add("ИП Аргымбаев"));
            //Convert.ToString(comboBox3.Items.Add("ЧП Светофор"));
            //Convert.ToString(comboBox3.Items.Add("ИП Атантаева Н.Т."));
            //Convert.ToString(comboBox3.Items.Add("ОсОО Ангара курьер"));
            //Convert.ToString(comboBox3.Items.Add("Домашний магазин"));
            //Convert.ToString(comboBox3.Items.Add("TOO Sapar delivery"));
            //Convert.ToString(comboBox3.Items.Add("Физ. лицо"));

            //Справка
            label22.Text = "    В программе 10 видов выборки" +
            Environment.NewLine +
            Environment.NewLine + "1. Выборка на реестр-1 - Выбираем 'Статус + Обработка + Филиал + Клиент'." +
            Environment.NewLine + "2. Выборка на реестр-2 - Выбираем 'Статус + Обработка + Клиент'." +
            Environment.NewLine + "3. Выборка за период-1 (Дата Обработки) - Выбираем 'Статус + Период + Область + Клиент и поставте галочку Дата Обработки'." +
            Environment.NewLine + "4. Выборка за период-2 (Дата записи) - Выбираем 'Статус + Период + Область + Филиал + Клиент'." +
            Environment.NewLine + "5. Выборка за период-3 (Дата записи) - Выбираем 'Статус + Период + Область + Клиент'." +
            Environment.NewLine + "6. Выборка за период-4 (Дата Обработки) - Выбираем 'Статус + Период + Клиент'." +
            Environment.NewLine + "7. Выборка за период-5 (Дата записи) - Выбираем 'Статус + Период + Клиент'." +
            Environment.NewLine + "8. Выборка на накладную - Выбираем 'Дата + Область + Клиент'." +
            Environment.NewLine + "9. Выборка поиск Реестра - В поле № введите номер реестра + Клиент + Статус." +
            Environment.NewLine + "10. Выборка за период-5 (Дата Обработки) - Выбираем 'Период + Клиент и поставте галочку Дата Обработки'. (а так же можно выбрать область)" +
            Environment.NewLine + "11. Выборка за период-6 (Дата записи) - Выбираем 'Период + Клиент'. (а так же можно выбрать область)" +
            Environment.NewLine +
            Environment.NewLine + "Список принятых --- 1.Выбрать контрагент и рядом поставить №_номер 2.Выбрать контрагент(выдаст последний список) 3.Установить период и выбрать контрагент" +
            Environment.NewLine +
            Environment.NewLine +
            Environment.NewLine + "Каждый филиал видит только свои записи в базе!";

            //button2.Enabled = false;
            Disp_data();
            Podschet();//произвести подсчет по методу       
            Suffix_select();
            Partner_select();
            Logins_select();
            comboBox10.SelectedIndex = 0;
        }
        //----------------------------------------------------------------------//
        async Task DispdatabaseAsync()//Асинхронность (async, await) 
        {
            await Task.Run(() => Disp_data_all_base());
            Wanted_Pending_Replacement();
            MessageBox.Show("База данных отображена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void Wanted_Pending_Replacement()//Розыск, Ожидание, Замена (Группировка)
        {
            //Группировка Статусов
            var statusGroup = from table in db.GetTable<Table_1_incomplete>()
                              group table by table.Статус into g
                              select g.OrderByDescending(t => t.Статус).FirstOrDefault();
            dataGridView2.DataSource = statusGroup;
            db.Refresh(RefreshMode.OverwriteCurrentValues, statusGroup); //datacontext очистка 
            //Отображение найденных
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                if (Convert.ToString(dataGridView2.Rows[i].Cells[11].Value) == "Ожидание")
                {
                    linkLabel2.Visible = true;
                    linkLabel2.Text = ("В ожидании!");
                }
                if (Convert.ToString(dataGridView2.Rows[i].Cells[11].Value) == "Отправлено")
                {
                    linkLabel5.Visible = true;
                    linkLabel5.Text = ("Отправлено!");
                }
                if (Convert.ToString(dataGridView2.Rows[i].Cells[11].Value) == "Розыск")
                {
                    linkLabel3.Visible = true;
                    linkLabel3.Text = ("В розыске!");
                }
                if (Convert.ToString(dataGridView2.Rows[i].Cells[11].Value) == "Замена")
                {
                    linkLabel4.Visible = true;
                    linkLabel4.Text = ("На замену!");
                }
            }
        }
        private void LinkLabel2_Click(object sender, EventArgs e)//Отобразить список Ожидание!
        {
            dataGridView2.Visible = true;

            var command = from table in db.GetTable<Table_1_incomplete>()
                          where table.Статус == "Ожидание"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;

            linkLabel2.Visible = false;
            Podschet();
        }
        private void LinkLabel3_Click(object sender, EventArgs e)//Отобразить список Розыск!
        {
            dataGridView2.Visible = true;

            var command = from table in db.GetTable<Table_1_incomplete>()
                          where table.Статус == "Розыск"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;

            linkLabel3.Visible = false;
            Podschet();
        }
        private void LinkLabel4_Click(object sender, EventArgs e)//Отобразить список Замена!
        {
            dataGridView2.Visible = true;

            var command = from table in db.GetTable<Table_1_incomplete>()
                          where table.Статус == "Замена"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;

            linkLabel4.Visible = false;
            Podschet();
        }
        private void LinkLabel5_Click(object sender, EventArgs e)//Отобразить список Отправленно!
        {
            dataGridView2.Visible = true;

            var command = from table in db.GetTable<Table_1_incomplete>()
                          where table.Статус == "Отправлено"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;

            linkLabel5.Visible = false;
            Podschet();
        }
        public void Wait()//Отобразить список Ожидание! 
        {
            //Отобразить список Ожидание! 
            var command = from table in db.GetTable<Table_1_incomplete>()
                          where table.Статус == "Ожидание"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;
            Tarifs();//Т а р и ф ы
            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка command
        }

        public void UPOnSubmit()//Изминение строк в новом списке при загрузке через API
        {
            Table.Tarifs = new DataTable();//инициализируем DataTable
            //Отобразить список API и сделать Update 
            var command = from table in db.GetTable<Table_1_incomplete>()
                          where table.Статус == "API"
                          orderby table.Дата_записи descending
                          select table;
            foreach (Table_1_incomplete order in command)
            {
                order.Статус = "Ожидание";
                order.Обработка = "Не обработано";
                order.Филиал = Person.Name;
                order.Тарифы = Table.Tarifs.Rows[0][0].ToString();
            }
            db.SubmitChanges();
            dataGridView2.DataSource = command;
            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка command

            Wait();//Отобразить список Ожидание и Провести тарификацию!
            Podschet();//произвести подсчет
        }
        public void SelectData()//Группировка и Сортировка по дате записи (сначала новые)
        {
            UPOnSubmit();//Изминение строк в новом списке при загрузке через API
            Wanted_Pending_Replacement();//Розыск, Ожидание, Замена (Группировка)
            if (Person.Name == "root")
            {
                //Группировка по Филиалу (находим последнюю запись) сортируем по дате
                var maxDate = from table in db.GetTable<Table_1_incomplete>()
                              group table by table.Филиал into g
                              select g.OrderByDescending(t => t.Дата_записи).FirstOrDefault();
                dataGridView2.DataSource = maxDate;
                db.Refresh(RefreshMode.OverwriteCurrentValues, maxDate); //datacontext очистка 
                //последние записи по Дате
                var lastDays = from table in db.GetTable<Table_1_incomplete>()
                               where table.Дата_записи >= Convert.ToDateTime(dataGridView2.Rows[0].Cells[12].Value)
                               orderby table.Дата_записи descending
                               select table;
                dataGridView2.DataSource = lastDays;
                db.Refresh(RefreshMode.OverwriteCurrentValues, lastDays); //datacontext очистка 
                label1.Text = ("Отображены последние записи по всем филиалам");
            }
            else
            {
                var sevenDays = from table in db.GetTable<Table_1_incomplete>()
                                where table.Дата_записи >= DateTime.Now.AddDays(-7)
                                where table.Филиал == Person.Name
                                orderby table.Дата_записи descending
                                select table;
                dataGridView2.DataSource = sevenDays;
                db.Refresh(RefreshMode.OverwriteCurrentValues, sevenDays); //datacontext очистка 
                label1.Text = ("Отображена последняя неделя");
            }
            if (dataGridView2.Rows.Count == 0)
            {
                //Группировка по Филиалу (находим последнюю запись) сортируем по дате
                var maxDate = from table in db.GetTable<Table_1_incomplete>()
                              group table by table.Филиал into g
                              select g.OrderByDescending(t => t.Дата_записи).FirstOrDefault();
                dataGridView2.DataSource = maxDate;
                db.Refresh(RefreshMode.OverwriteCurrentValues, maxDate); //datacontext очистка 
                //последние записи по Дате
                var lastDays = from table in db.GetTable<Table_1_incomplete>()
                               where table.Дата_записи >= Convert.ToDateTime(dataGridView2.Rows[0].Cells[12].Value)
                               where table.Филиал == Person.Name
                               orderby table.Дата_записи descending
                               select table;
                dataGridView2.DataSource = lastDays;
                db.Refresh(RefreshMode.OverwriteCurrentValues, lastDays); //datacontext очистка command
                label1.Text = ("Отображены последние записи");               
            }
            dataGridView2.Columns[0].Visible = false;//Скрыть столбец ID
            dataGridView2.Columns[16].Visible = false;//Скрыть столбец Филиал
            dataGridView2.Columns[21].Visible = false;//Скрыть столбец Тарифы
        }
        public void Disp_data()//Отображает базу
        {
            button8.Text = "Ожидайте!";
            button8.Enabled = false;
            dataGridView2.Visible = true;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;

            SelectData(); //Группировка и Сортировка по дате записи (сначала новые) //Розыск, Ожидание, Замена (Группировка)            
            button8.Text = "Обновить";
            button8.Enabled = true;
            Podschet();
        }
        public void Disp_data_all_base()//Отображает всю базу и сортирует по дате записи
        {
            button9.Text = "Ожидайте!";
            button9.Enabled = false;
            dataGridView2.Visible = true;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;
            ProgressBar();

            if (Person.Name == "root")
            {
                var command = from table in db.GetTable<Table_1_incomplete>()
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            else
            {
                var command = from table in db.GetTable<Table_1_incomplete>()
                              where table.Филиал == Person.Name
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }

            label1.Text = ("База данных отображена");
            button9.Text = "Вся база";
            button9.Enabled = true;
            Podschet();//произвести подсчет по методу         
        }

        public void Podschet()//Произвести подсчет dataGridView1 и dataGridView5 и dataGridView2
        {
            if (dataGridView1.Visible == true)
            {
                //Сумма столбца стоимость
                double summa = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox5.Text = summa.ToString() + " Сом";
                Summ.Sum = summa.ToString() + " Сом";
                //Сумма столбца плата за услугу
                double summa_U = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_U += incom;
                }
                textBox15.Text = summa_U.ToString() + " Сом";
                Summ.SumService = summa_U.ToString() + " Сом";
                //Сумма столбца плата за возврат
                double summa_V = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[14].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_V += incom;
                }
                textBox21.Text = summa_V.ToString() + " Сом";
                Summ.SumReturn = summa_V.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1[i, j].Value != null)
                        {
                            textBox4.Text = Convert.ToString(dataGridView1.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            Summ.Quantity = Convert.ToString(dataGridView1.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            count++;
                            break;
                        }
                    }
                }
            }
            else if (dataGridView5.Visible == true)
            {
                //Сумма столбца стоимость
                double summa = 0;
                foreach (DataGridViewRow row in dataGridView5.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox5.Text = summa.ToString() + " Сом";
                Summ.Sum = summa.ToString() + " Сом";
                //Сумма столбца плата за услугу
                double summa_U = 0;
                foreach (DataGridViewRow row in dataGridView5.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[8].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_U += incom;
                }
                textBox15.Text = summa_U.ToString() + " Сом";
                Summ.SumService = summa_U.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView5.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView5.ColumnCount; i++)
                    {
                        if (dataGridView5[i, j].Value != null)
                        {
                            textBox4.Text = Convert.ToString(dataGridView5.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            Summ.Quantity = Convert.ToString(dataGridView5.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
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
                    double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox5.Text = summa.ToString() + " Сом";
                Summ.Sum = summa.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView2.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (dataGridView2[i, j].Value != null)
                        {
                            textBox4.Text = Convert.ToString(dataGridView2.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            Summ.Quantity = Convert.ToString(dataGridView2.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            count++;
                            break;
                        }
                    }
                }
            }
            else if (dataGridView6.Visible == true)//Админпанель
            {
                //Сумма столбца стоимость
                double summa = 0;
                foreach (DataGridViewRow row in dataGridView6.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox25.Text = summa.ToString() + " Сом";
                Summ.Sum = summa.ToString() + " Сом";
                //Сумма столбца плата за услугу
                double summa_U = 0;
                foreach (DataGridViewRow row in dataGridView6.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_U += incom;
                }
                textBox23.Text = summa_U.ToString() + " Сом";
                Summ.SumService = summa_U.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView6.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView6.ColumnCount; i++)
                    {
                        if (dataGridView6[i, j].Value != null)
                        {
                            textBox24.Text = Convert.ToString(dataGridView6.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            Summ.Quantity = Convert.ToString(dataGridView6.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            count++;
                            break;
                        }
                    }
                }
            }
        }
        public void Tarifs()//Т а р и ф ы
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)
            {
                int doplata = Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value);
                int stoimost = Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value);
                double tarif = Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value);
                string[] tarifs = { "Общий" };//Тарифы
                for (int y = 0; y < tarifs.Length; y++)
                {
                    if (Convert.ToString(dataGridView2.Rows[i].Cells[21].Value) == tarifs[y])//Т а р и ф для большинства организаций
                    {
                        double ob_cennost = (stoimost * 1.0 / 100);
                        double plata_za_nalog = (stoimost * 2.0 / 100);
                        con.Open();//открыть соединение
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu, ob_cennost = @ob_cennost, plata_za_nalog = @plata_za_nalog WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost + tarif + plata_za_nalog + doplata));
                        cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost));
                        cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                        cmd.ExecuteNonQuery();
                        con.Close();//закрыть соединение
                    }
                }
                if (Convert.ToString(dataGridView2.Rows[i].Cells[21].Value) == "по 1 проценту")//Т а р и ф для ИП 'JUMPER'
                {
                    double ob_cennost = (stoimost * 1.0 / 100);
                    double plata_za_nalog = (stoimost * 1.0 / 100);
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu, ob_cennost = @ob_cennost, plata_za_nalog = @plata_za_nalog WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost + tarif + plata_za_nalog + doplata));
                    cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost));
                    cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                    cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
                if (Convert.ToString(dataGridView2.Rows[i].Cells[21].Value) == "по 2 процента")//Т а р и ф для "ОсОО 'Экспресс-Тайм'"
                {
                    double ob_cennost = (stoimost * 2.0 / 100);
                    double plata_za_nalog = (stoimost * 2.0 / 100);
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu, ob_cennost = @ob_cennost, plata_za_nalog = @plata_za_nalog WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost + tarif + plata_za_nalog + doplata));
                    cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost));
                    cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                    cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение  
                }
                if (Convert.ToString(dataGridView2.Rows[i].Cells[21].Value) == "по 0 процентов")//Т а р и ф для "ОсОО 'Альфа Вита'"
                {
                    double ob_cennost = 0;
                    double plata_za_nalog = 0;
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu, ob_cennost = @ob_cennost, plata_za_nalog = @plata_za_nalog WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost + tarif + plata_za_nalog + doplata));
                    cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost));
                    cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                    cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
                if (Convert.ToString(dataGridView2.Rows[i].Cells[21].Value) == "1,5 процента")//Т а р и ф для ОсОО Kyrgyz Express Post
                {
                    double plata_za_nalog = (stoimost * 1.5 / 100);
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu, plata_za_nalog = @plata_za_nalog, ob_cennost = 0 WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(tarif + plata_za_nalog + doplata));
                    cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                    cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
            }

            for (int i = 0; i < dataGridView2.Rows.Count; i++)//Т а р и ф для Sapardelivery и ОсОО Тенгри
            {
                if (dataGridView2.Rows[i].Cells[21].Value.ToString() == "Сложный")
                {
                    int doplata = Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value);
                    int stoimost = Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value);
                    double tarif = Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value);
                    double ob_cennost_3000 = stoimost * 2.2 / 100;
                    double ob_cennost_6000 = stoimost * 1.7 / 100;
                    double ob_cennost_10000 = stoimost * 1.2 / 100;
                    double ob_cennost_20000 = stoimost * 0.7 / 100;
                    double ob_cennost_50000 = stoimost * 1.0 / 100;
                    double ob_cennost_50000i = stoimost * 0.4 / 100;
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu, ob_cennost = @ob_cennost, plata_za_nalog = @plata_za_nalog WHERE id = @id", con);
                    if (stoimost <= 1000)//Если стоимость До 1000 сом включительно
                    {
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", (tarif + 20 + 40 + doplata));
                        cmd.Parameters.AddWithValue("@ob_cennost", 20);
                        cmd.Parameters.AddWithValue("@plata_za_nalog", 40);
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    }
                    if (stoimost > 1000 && stoimost <= 3000)//Если стоимость От 1000 До 3000 сом включительно
                    {
                        double plata_za_nalog = stoimost * 4.4 / 100;
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost_3000 + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                        cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost_3000));
                        cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    }
                    if (stoimost > 3000 && stoimost <= 6000)//Если стоимость От 3000 До 6000 сом включительно
                    {
                        double plata_za_nalog = stoimost * 3.4 / 100;
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost_6000 + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                        cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost_6000));
                        cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    }
                    if (stoimost > 6000 && stoimost <= 10000)//Если стоимость От 6000 До 10000 сом включительно
                    {
                        double plata_za_nalog = stoimost * 2.4 / 100;
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost_10000 + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                        cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost_10000));
                        cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    }
                    if (stoimost > 10000 && stoimost <= 20000)//Если стоимость От 10000 До 20000 сом включительно
                    {
                        double plata_za_nalog = stoimost * 1.4 / 100;
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost_20000 + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                        cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost_20000));
                        cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    }
                    if (stoimost > 20000 && stoimost <= 50000)//Если стоимость От 20000 До 50000 сом включительно
                    {
                        double plata_za_nalog = stoimost * 1.0 / 100;
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost_50000 + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                        cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost_50000));
                        cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    }
                    if (stoimost > 50000)//Если стоимость свыше 50000
                    {
                        double plata_za_nalog = stoimost * 0.8 / 100;
                        cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost_50000i + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                        cmd.Parameters.AddWithValue("@ob_cennost", Math.Round(ob_cennost_50000i));
                        cmd.Parameters.AddWithValue("@plata_za_nalog", Math.Round(plata_za_nalog));
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                    }
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                }
            }
        }
        public void Tarif_Update()//Заново ищет №Заказа и делает пересчет (Для ускорения программы во время UPDATE)
        {
            int currRowIndex = dataGridView2.CurrentCell.RowIndex;//  Запоминаем строку, которую выбрал пользователь.
            if (dataGridView2.Rows.Count != 0)
            {
                int doplata = Convert.ToInt32(dataGridView2.Rows[0].Cells[7].Value);
                int tarif = Convert.ToInt32(dataGridView2.Rows[0].Cells[6].Value);
                double ob_cennost = Convert.ToInt32(dataGridView2.Rows[0].Cells[8].Value);
                double plata_za_nalog = Convert.ToInt32(dataGridView2.Rows[0].Cells[9].Value);
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                cmd.Parameters.AddWithValue("@id", dataGridView2.CurrentRow.Cells[0].Value);//выбранная строка в гриде
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
            }
            dataGridView2.CurrentCell = dataGridView2[0, currRowIndex];//  Выбираем нашу строку (именно выбираем, не выделяем).
        }

        public static DataTable ToDataTable<T>(IEnumerable<T> values)//Функция (передать результат LINQ таблице DataTable)
        {
            DataTable table = new DataTable();
            foreach (T value in values)
            {
                if (table.Columns.Count == 0)
                {
                    foreach (var p in value.GetType().GetProperties())
                    {
                        table.Columns.Add(p.Name);
                    }
                }

                DataRow dr = table.NewRow();
                foreach (var p in value.GetType().GetProperties())
                {
                    dr[p.Name] = p.GetValue(value, null) + "";

                }
                table.Rows.Add(dr);
            }
            return table;
        }
        public void Select_status_Nr()//(Для выдачи реестров)Выборка по статусу и сортировка по номеру реестра от больших значений к меньшим.
        {
            //con.Open();//Открываем соединение
            //SqlCommand cmd = new SqlCommand("SELECT MAX(id) AS ID, MAX(oblast) AS 'Область', MAX(punkt) AS 'Населенный пункт', MAX(familia) AS 'Ф.И.О'," +
            //    "MAX(summ) AS 'Стоимость',MAX(plata_za_uslugu) AS 'Услуга', MAX(tarif) AS 'Тариф', MAX(doplata) AS 'Доплата', MAX(ob_cennost) AS 'Обьяв.ценность', MAX(plata_za_nalog) AS 'Наложенный платеж'," +
            //    "MAX(N_zakaza) AS '№Заказа', MAX(status) AS 'Статус', MAX(data_zapisi) AS 'Дата записи', MAX(prichina) AS 'Причина', MAX(obrabotka) AS 'Обработка', MAX(data_obrabotki) AS 'Дата обработки'," +
            //    "MAX(filial) AS 'Филиал', MAX(client) AS 'Контрагент'," +
            //    "MAX(nomer_spiska) AS 'Список', MAX(nomer_nakladnoy) AS 'Накладная', MAX(nomer_reestra) AS 'Реестр', MAX(Ns) AS 'NS', MAX(Nn) AS 'NN', MAX(Nr) AS 'NR'" +
            //        "FROM [Table_1] WHERE status = @status GROUP BY Nr ORDER BY Nr DESC", con);
            //if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Выдано")
            //{
            //    cmd.Parameters.AddWithValue("@status", "Выдано");
            //}
            //else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Возврат")
            //{
            //    cmd.Parameters.AddWithValue("@status", "Возврат");
            //}
            //else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Розыск")
            //{
            //    cmd.Parameters.AddWithValue("@status", "Розыск");
            //}
            //else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Замена")
            //{
            //    cmd.Parameters.AddWithValue("@status", "Замена");
            //}
            //else MessageBox.Show("select_status", "Ошибка!");
            //cmd.ExecuteNonQuery();

            //DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            //SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            //dt.Clear();//чистим DataTable, если он был не пуст
            //da.Fill(dt);//заполняем данными созданный DataTable
            //dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //con.Close();//Закрываем соединение
            string status = "";
            if (Table.DtRegistry.Rows[0][5].ToString() == "Выдано")
            {
                status = "Выдано";
            }
            else if (Table.DtRegistry.Rows[0][5].ToString() == "Возврат")
            {
                status = "Возврат";
            }
            else if (Table.DtRegistry.Rows[0][5].ToString() == "Розыск")
            {
                status = "Розыск";
            }
            else if (Table.DtRegistry.Rows[0][5].ToString() == "Замена")
            {
                status = "Замена";
            }
            else MessageBox.Show("Select_status_Nr", "Ошибка!");

            var command = from table in db.GetTable<Table_1>()
                          where table.Статус == status
                          group table by table.Nr into g
                          select g.OrderByDescending(t => t.Nr).FirstOrDefault();
            //dataGridView2.DataSource = command;
            DataTable tableNr = ToDataTable(command); //передать результат LINQ таблице DataTable

            Number.Nr = Convert.ToInt32(tableNr.Rows[0][23].ToString()) + 1;
            //Number.Nr = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
            Number.Prefix_number = comboBox10.Text + Number.Nr;

            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка 
        }     
        public void Select_status_Nn()//(Для выдачи накладных)Выборка по статусу и сортировка по номеру накладеой от больших значений к меньшим.
        {
            //con.Open();//Открываем соединение
            //SqlCommand cmd = new SqlCommand("SELECT MAX(id) AS ID, MAX(oblast) AS 'Область', MAX(punkt) AS 'Населенный пункт', MAX(familia) AS 'Ф.И.О'," +
            //    "MAX(summ) AS 'Стоимость',MAX(plata_za_uslugu) AS 'Услуга', MAX(tarif) AS 'Тариф', MAX(doplata) AS 'Доплата', MAX(ob_cennost) AS 'Обьяв.ценность', MAX(plata_za_nalog) AS 'Наложеный платеж'," +
            //    "MAX(N_zakaza) AS '№Заказа', MAX(status) AS 'Статус', MAX(data_zapisi) AS 'Дата записи', MAX(prichina) AS 'Причина', MAX(obrabotka) AS 'Обработка', MAX(data_obrabotki) AS 'Дата обработки'," +
            //    "MAX(filial) AS 'Филиал', MAX(client) AS 'Контрагент'," +
            //    "MAX(nomer_spiska) AS 'Список', MAX(nomer_nakladnoy) AS 'Накладная', MAX(nomer_reestra) AS 'Реестр', MAX(Ns) AS 'NS', MAX(Nn) AS 'NN', MAX(Nr) AS 'NR'" +
            //        "FROM [Table_1] WHERE status = @status GROUP BY Nn ORDER BY Nn DESC", con);
            //if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Ожидание")//Для накладных
            //{
            //    cmd.Parameters.AddWithValue("@status", "Отправлено");//чтобы простовлять порядковый номер /не менять (все верно) могу забыть!
            //}
            //else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Отправлено")
            //{
            //    cmd.Parameters.AddWithValue("@status", "Отправлено");
            //}
            //else MessageBox.Show("select_status_nakladnoi", "Ошибка!");
            //cmd.ExecuteNonQuery();

            //DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            //SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            //dt.Clear();//чистим DataTable, если он был не пуст
            //da.Fill(dt);//заполняем данными созданный DataTable
            //dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //con.Close();//Закрываем соединение
            string status = "";
            if (Table.DtInvoice.Rows[0][5].ToString() == "Ожидание")
            {
                status = "Отправлено";//чтобы простовлять порядковый номер /не менять (все верно) могу забыть!
            }
            else if (Table.DtInvoice.Rows[0][5].ToString() == "Отправлено")
            {
                status = "Отправлено";
            }
            else MessageBox.Show("Select_status_Nn", "Ошибка!");

            var command = from table in db.GetTable<Table_1>()
                          where table.Статус == status
                          group table by table.Nn into g
                          select g.OrderByDescending(t => t.Nn).FirstOrDefault();
            //dataGridView2.DataSource = command;
            DataTable tableNn = ToDataTable(command); //передать результат LINQ таблице DataTable

            Number.Nn = Convert.ToInt32(tableNn.Rows[0][22].ToString()) + 1;
            //Number.Nn = Convert.ToInt32(dataGridView2.Rows[0].Cells[22].Value) + 1;
            Number.Prefix_number = comboBox10.Text + Number.Nn;

            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка 
        }
        public void Select_Ns()//(Для выдачи списка принятых)Выборка и сортировка по номеру от больших значений к меньшим.
        {
            //con.Open();//Открываем соединение
            //SqlCommand cmd = new SqlCommand("SELECT MAX(id) AS ID, MAX(oblast) AS 'Область', MAX(punkt) AS 'Населенный пункт', MAX(familia) AS 'Ф.И.О'," +
            //    "MAX(summ) AS 'Стоимость',MAX(plata_za_uslugu) AS 'Услуга', MAX(tarif) AS 'Тариф', MAX(doplata) AS 'Доплата', MAX(ob_cennost) AS 'Обьяв.ценность', MAX(plata_za_nalog) AS 'Наложеный платеж'," +
            //    "MAX(N_zakaza) AS '№Заказа', MAX(status) AS 'Статус', MAX(data_zapisi) AS 'Дата записи', MAX(prichina) AS 'Причина', MAX(obrabotka) AS 'Обработка', MAX(data_obrabotki) AS 'Дата обработки'," +
            //    "MAX(filial) AS 'Филиал', MAX(client) AS 'Контрагент'," +
            //    "MAX(nomer_spiska) AS 'Список', MAX(nomer_nakladnoy) AS 'Накладная', MAX(nomer_reestra) AS 'Реестр', MAX(Ns) AS 'NS', MAX(Nn) AS 'NN', MAX(Nr) AS 'NR'" +
            //        " FROM [Table_1] GROUP BY Ns ORDER BY Ns DESC", con);
            //cmd.ExecuteNonQuery();
            //DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            //SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            //dt.Clear();//чистим DataTable, если он был не пуст
            //da.Fill(dt);//заполняем данными созданный DataTable
            //dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //con.Close();//Закрываем соединение

            var command = from table in db.GetTable<Table_1>()
                          group table by table.Ns into g
                          select g.OrderByDescending(t => t.Ns).FirstOrDefault();
            //dataGridView2.DataSource = command;
            DataTable tableNs = ToDataTable(command); //передать результат LINQ таблице DataTable

            Number.Ns = Convert.ToInt32(tableNs.Rows[0][21].ToString()) + 1;
            //Number.Ns = Convert.ToInt32(dataGridView2.Rows[0].Cells[21].Value) + 1;
            Number.Prefix_number = comboBox10.Text + Number.Ns;

            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка 
        }
        public void Select_client()//Для сортировки принятых списков по клиенту
        {
            var command = from table in db.GetTable<Table_1>()
                          where table.Контрагент == Partner.Name
                          group table by table.Ns into g
                          select g.OrderByDescending(t => t.Ns).FirstOrDefault();
            //dataGridView2.DataSource = command;
            DataTable tableNs = ToDataTable(command); //передать результат LINQ таблице DataTable
            Number.Prefix_number = tableNs.Rows[0][18].ToString();

            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка 

            //con.Open();//Открываем соединение
            //SqlCommand cmd = new SqlCommand("SELECT MAX(id) AS ID, MAX(oblast) AS 'Область', MAX(punkt) AS 'Населенный пункт', MAX(familia) AS 'Ф.И.О'," +
            //    "MAX(summ) AS 'Стоимость',MAX(plata_za_uslugu) AS 'Услуга', MAX(tarif) AS 'Тариф', MAX(doplata) AS 'Доплата', MAX(ob_cennost) AS 'Обьяв.ценность', MAX(plata_za_nalog) AS 'Наложеный платеж'," +
            //    "MAX(N_zakaza) AS '№Заказа', MAX(status) AS 'Статус', MAX(data_zapisi) AS 'Дата записи', MAX(prichina) AS 'Причина', MAX(obrabotka) AS 'Обработка', MAX(data_obrabotki) AS 'Дата обработки'," +
            //    "MAX(filial) AS 'Филиал', MAX(client) AS 'Контрагент'," +
            //    "MAX(nomer_spiska) AS 'Список', MAX(nomer_nakladnoy) AS 'Накладная', MAX(nomer_reestra) AS 'Реестр', MAX(Ns) AS 'NS', MAX(Nn) AS 'NN', MAX(Nr) AS 'NR'" +
            //    " FROM [Table_1] WHERE client = @client GROUP BY Ns ORDER BY Ns DESC", con);
            //cmd.Parameters.AddWithValue("@client", comboBox5.Text);
            //cmd.ExecuteNonQuery();
            //DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            //SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            //dt.Clear();//чистим DataTable, если он был не пуст
            //da.Fill(dt);//заполняем данными созданный DataTable
            //dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //con.Close();//Закрываем соединение            
        }
        public void Suffix_select()//Вывод Суффикса в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT name FROM [Table_Suffix]", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow row in dt.Rows)
            {
                comboBox10.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение          
        }
        public void Partner_select()//Вывод Контрагентов в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT name FROM [Table_Partner] ORDER BY id", con);
            cmd.ExecuteNonQuery();
            Table.DtPartner = new DataTable();//инициализируем DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            Table.DtPartner.Clear();//чистим DataTable, если он был не пуст
            da.Fill(Table.DtPartner);//заполняем данными созданный DataTable
            foreach (DataRow column in Table.DtPartner.Rows)
            {
                comboBox6.Items.Add(column[0].ToString());
                comboBox3.Items.Add(column[0].ToString());
                comboBox5.Items.Add(column[0].ToString());
            }
            con.Close();//Закрываем соединение          
        }
        public void ComboBox5_TextChanged(object sender, EventArgs e)//поиск тарифа по контрагенту
        {
            Table.Tarifs = new DataTable();//инициализируем DataTable
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT tarif FROM [Table_Partner]" +
                "WHERE name = @name", con);
            cmd.Parameters.AddWithValue("@name", comboBox5.Text.ToString());
            cmd.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            Table.Tarifs.Clear();//чистим DataTable, если он был не пуст
            da.Fill(Table.Tarifs);//заполняем данными созданный DataTable
            con.Close();//закрыть соединение
            if (comboBox5.Text == "")//если поле очищено, отобразить базу
            {
                Table.Tarifs.Clear();//чистим DataTable, если он был не пуст
                foreach (DataRow column in Table.Tarifs.Rows)
                {
                    comboBox5.Items.Add(column[0].ToString());
                }
            }
        }
        //---------------------------------------------------------------------------------------------------------//
        private DataTableCollection tableCollection = null;
        private void OpenExcelFile(string path)//Считывание Excel таблицы в DataGridView
        {
            FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read);
            IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
            DataSet db = reader.AsDataSet(new ExcelDataSetConfiguration()
            {
                ConfigureDataTable = (x) => new ExcelDataTableConfiguration()
                {
                    UseHeaderRow = true
                }
            });
            tableCollection = db.Tables;

            comboBox13.Items.Clear();
            foreach (DataTable table in tableCollection)
            {
                comboBox13.Items.Add(table.TableName);
            }
            comboBox13.SelectedIndex = 0;
        }
        private void ComboBox13_SelectedIndexChanged(object sender, EventArgs e)
        {
            DataTable table = tableCollection[Convert.ToString(comboBox13.SelectedItem)];
            dataGridView3.DataSource = table;
        }
        private void Button1_Click(object sender, EventArgs e)//Открыть Excel и Загрузить в базу
        {
            if (dataGridView3.Rows.Count <= 0)//Если грид пустой
            {
                try
                {
                    dataGridView5.Visible = false;
                    dataGridView2.Visible = false;
                    dataGridView1.Visible = false;
                    dataGridView3.Visible = true;
                    OpenFileDialog ofd = new OpenFileDialog
                    {
                        Filter = "Excel|*.xlsx;*.xls"
                    };
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        fileName = ofd.FileName;
                        Text = fileName;
                        OpenExcelFile(fileName);
                        button1.Text = "Загрузить Excel";
                    }
                    else
                    {
                        throw new Exception("Файл не выбран!");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

                    comboBox13.SelectedIndex = -1;
                    comboBox13.Items.Clear();
                    button1.Text = "Открыть Excel";
                    dataGridView5.Visible = false;
                    dataGridView2.Visible = true;
                    dataGridView1.Visible = false;
                    dataGridView3.Visible = false;
                }
            }
            else if (dataGridView3.Rows.Count > 0)//Если грид не пустой
            {
                try
                {
                    if (comboBox5.Text != "")
                    {
                        if (MessageBox.Show("Вы уверенны что этот Реестр принадлежит " + comboBox5.Text, "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {
                            Partner.Name = comboBox5.Text;
                            ProgressBar();
                            button1.Text = "Ожидайте!";
                            button1.Enabled = false;
                            dataGridView2.Visible = true;
                            dataGridView1.Visible = false;
                            dataGridView5.Visible = false;
                            Select_Ns();//Выборка и сортировка по номеру от больших значений к меньшим.
                            con.Open();//открыть соединение
                            for (int i = 0; i < dataGridView3.Rows.Count; i++)
                            {
                                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_1] (oblast, filial, punkt, familia, summ, N_zakaza, data_zapisi, obrabotka, status, client," +
                                    " tarif, plata_za_uslugu, ob_cennost, plata_za_nalog, nomer_reestra, nomer_spiska, doplata, nomer_nakladnoy, Nr, Ns, Nn, tarifs) VALUES (@oblast, @filial, @punkt, @familia, @summ, @N_zakaza, @data_zapisi," +
                                    " @obrabotka, @status, @client, @tarif, @plata_za_uslugu, @ob_cennost, @plata_za_nalog, @nomer_reestra, @nomer_spiska, @doplata, @nomer_nakladnoy, @Nr, @Ns, @Nn, @tarifs)", con);

                                if (dataGridView3.Rows[i].Cells[1].Value != DBNull.Value)//если в EXCEL столбец область не пустой
                                { cmd.Parameters.AddWithValue("@oblast", Convert.ToString(dataGridView3.Rows[i].Cells[1].Value)); }
                                else if ((dataGridView3.Rows[i].Cells[1].Value) == DBNull.Value)//если в EXCEL столбец область пустой
                                { cmd.Parameters.AddWithValue("@oblast", "пусто"); MessageBox.Show("Область в одной из записей этого реестра не заполнена!", "Внимание! Заполните область и сформируйте новый список!"); }

                                cmd.Parameters.AddWithValue("@punkt", Convert.ToString(dataGridView3.Rows[i].Cells[2].Value));
                                cmd.Parameters.AddWithValue("@familia", Convert.ToString(dataGridView3.Rows[i].Cells[3].Value));

                                if (dataGridView3.Rows[i].Cells[4].Value != DBNull.Value)//если в EXCEL столбец стоимость не пустой
                                { cmd.Parameters.AddWithValue("@summ", Convert.ToString(dataGridView3.Rows[i].Cells[4].Value)); }
                                else if ((dataGridView3.Rows[i].Cells[4].Value) == DBNull.Value)//если в EXCEL столбец стоимость пустой
                                { cmd.Parameters.AddWithValue("@summ", 0); MessageBox.Show("Стоимость в одной из записей этого реестра не заполнена!", "Внимание! Заполните стоимость и сформируйте новый список!"); }
                                if ((dataGridView3.Rows[i].Cells[7].Value) != DBNull.Value)//если в EXCEL столбец тариф не пустой
                                { cmd.Parameters.AddWithValue("@tarif", Convert.ToString(dataGridView3.Rows[i].Cells[7].Value)); }
                                else if ((dataGridView3.Rows[i].Cells[7].Value) == DBNull.Value)//если в EXCEL столбец тариф пустой
                                { cmd.Parameters.AddWithValue("@tarif", 0); MessageBox.Show("Тариф в одной из записей этого реестра не заполнен!", "Внимание! Заполните тариф и сформируйте новый список!"); }
                                cmd.Parameters.AddWithValue("@plata_za_uslugu", 0);
                                cmd.Parameters.AddWithValue("@ob_cennost", 0);
                                cmd.Parameters.AddWithValue("@plata_za_nalog", 0);
                                if (dataGridView3.Rows[i].Cells[10].Value != DBNull.Value)//если в EXCEL столбец номер заказа не пустой
                                { cmd.Parameters.AddWithValue("@N_zakaza", Convert.ToString(dataGridView3.Rows[i].Cells[10].Value)); }
                                else if ((dataGridView3.Rows[i].Cells[10].Value) == DBNull.Value)//если в EXCEL столбец номер заказа пустой
                                { cmd.Parameters.AddWithValue("@N_zakaza", 0); MessageBox.Show("Номер заказа в одной из записей этого реестра не заполнен!", "Внимание! Заполните номер заказа и сформируйте новый список!"); }

                                cmd.Parameters.AddWithValue("@data_zapisi", DateTime.Today);
                                cmd.Parameters.AddWithValue("@obrabotka", "Не обработано");
                                cmd.Parameters.AddWithValue("@status", "Ожидание");
                                cmd.Parameters.AddWithValue("@client", Partner.Name);
                                cmd.Parameters.AddWithValue("@nomer_reestra", 0);
                                cmd.Parameters.AddWithValue("@nomer_spiska", Number.Prefix_number);
                                cmd.Parameters.AddWithValue("@nomer_nakladnoy", 0);
                                cmd.Parameters.AddWithValue("@Nr", 0);
                                cmd.Parameters.AddWithValue("@Ns", Number.Ns);
                                cmd.Parameters.AddWithValue("@Nn", 0);
                                cmd.Parameters.AddWithValue("@tarifs", Table.Tarifs.Rows[0][0].ToString());//tarif
                                cmd.Parameters.AddWithValue("@filial", Person.Name);

                                if ((dataGridView3.Rows[i].Cells[11].Value) != DBNull.Value)//если в EXCEL столбец доплата не пустой
                                { cmd.Parameters.AddWithValue("@doplata", Convert.ToString(dataGridView3.Rows[i].Cells[11].Value)); }
                                else if ((dataGridView3.Rows[i].Cells[11].Value) == DBNull.Value)//если в EXCEL столбец доплата пустой
                                { cmd.Parameters.AddWithValue("@doplata", 0); }
                                cmd.ExecuteNonQuery();
                            }
                            con.Close();//закрыть соединение
                            MessageBox.Show("Реестр успешно загружен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            label1.Text = "Реестр успешно загружен!";
                            dataGridView2.Visible = true;
                            dataGridView1.Visible = false;
                            dataGridView5.Visible = false;

                            //Отобразить список Ожидание! 
                            var command = from table in db.GetTable<Table_1_incomplete>()
                                          where table.Статус == "Ожидание"
                                          orderby table.Дата_записи descending
                                          select table;
                            dataGridView2.DataSource = command;
                            Tarifs();//Т а р и ф ы
                            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка command

                            Podschet();//произвести подсчет по методу
                        }
                        if (MessageBox.Show("Вы хотите получить список принятых?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
                        {
                            dataGridView5.Visible = true;
                            dataGridView1.Visible = false;
                            dataGridView2.Visible = false;
                            //-------------------------------------Выборка по последнему загруженному списку (по номеру)------------------------------------------------------------------//
                            Select_client();//Выборка и сортировка по номеру от больших значений к меньшим.
                            con.Open();//открыть соединение
                            SqlCommand cmd = new SqlCommand("SELECT oblast, punkt, familia, N_zakaza, data_zapisi, summ, tarif, doplata, plata_za_uslugu, ob_cennost, plata_za_nalog, id, nomer_spiska" +
                                " FROM [Table_1] WHERE (nomer_spiska = @nomer_spiska AND client = @client)", con);
                            cmd.Parameters.AddWithValue("nomer_spiska", Number.Prefix_number);
                            cmd.Parameters.AddWithValue("@client", Partner.Name);
                            cmd.ExecuteNonQuery();
                            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                            dt.Clear();//чистим DataTable, если он был не пуст
                            da.Fill(dt);//заполняем данными созданный DataTable
                            dataGridView5.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                            con.Close();//закрыть соединение
                                        //-------------------------------------Выборка по последнему загруженному списку (по номеру)--------------------------------------------------------------------//
                            Podschet();//произвести подсчет по методу 
                                       //Выдача в WORD
                            SaveFileDialog sfd = new SaveFileDialog();
                            sfd.Filter = "Word Documents (*.docx)|*.docx";
                            sfd.FileName = $"Список принятых № {Number.Prefix_number}.docx";
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                            }
                        }
                    }
                    else if (comboBox5.Text == "")
                    {
                        MessageBox.Show("Необходимо выбрать Контрагент", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                        MessageBox.Show("Откройте Excel!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception exp)
                {
                    MessageBox.Show("Ошибка! Excel файл содержит ошибку или неправельно сформирован! " + Environment.NewLine + exp, "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    con.Close();//закрыть соединение
                }
                dataGridView2.Visible = true;
                dataGridView1.Visible = false;
                dataGridView5.Visible = false;

                button1.Text = "Открыть Excel";
                button1.Enabled = true;
                //-------------Очистка грида---------------//
                int rowsCount = dataGridView3.Rows.Count;
                for (int i = 0; i < rowsCount; i++)
                {
                    dataGridView3.Rows.Remove(dataGridView3.Rows[0]);
                }
                //-------------Очистка грида---------------//
                //dataGridView3.Rows.Clear();
                //dataGridView3.Columns.Clear();
                comboBox13.SelectedIndex = -1;
                comboBox13.Items.Clear();
            }
        }
        //---------------------------------------------------------------------------------------------------------//
        private void toolStripButton3_Click(object sender, EventArgs e)//Ручной ввод
        {
            Form_manual_input FMI = new Form_manual_input();
            FMI.Show();
        }

        private void button11_Click(object sender, EventArgs e)//Пересчет
        {
            ProgressBar();
            button11.Text = "Ожидайте!";
            button11.Enabled = false;
            Tarifs();//Т а р и ф ы      
            Disp_data();
            button11.Text = "Пересчет";
            button11.Enabled = true;
            MessageBox.Show("Пересчет выполнен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void button28_Click(object sender, EventArgs e)//Поиск по ФИО
        {
            //con.Open();//открыть соединение
            //SqlCommand cmd = new SqlCommand("SELECT id AS ID, oblast AS 'Область', punkt AS 'Населенный пункт', familia AS 'Ф.И.О'," +
            //"summ AS 'Стоимость',plata_za_uslugu AS 'Услуга', tarif AS 'Тариф', doplata AS 'Доплата', ob_cennost AS 'Обьяв.ценность', plata_za_nalog AS 'Наложеный платеж'," +
            //    "N_zakaza AS '№Заказа', status AS 'Статус', data_zapisi AS 'Дата записи', prichina AS 'Причина', obrabotka AS 'Обработка', data_obrabotki AS 'Дата обработки'," +
            //    "filial AS 'Филиал', client AS 'Контрагент'," +
            //    "nomer_spiska AS 'Список', nomer_nakladnoy AS 'Накладная', nomer_reestra AS 'Реестр', tarifs AS 'Тарифы'" +
            //        "FROM [Table_1] WHERE familia LIKE N'%" + textBox3.Text.ToString() + "%'", con);
            ////cmd.Parameters.AddWithValue("@punkt", textBox2.Text);
            ////cmd.Parameters.AddWithValue("@familia", textBox2.Text);
            //cmd.ExecuteNonQuery();
            //DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            //SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            //dt.Clear();//чистим DataTable, если он был не пуст
            //da.Fill(dt);//заполняем данными созданный DataTable
            //dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //con.Close();//закрыть соединение

            if (textBox3.Text != "")
            {
                var command = from table in db.GetTable<Table_1_incomplete>()
                              where table.Ф_И_О.Contains(textBox3.Text.ToString())//Contains вместо LIKE
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;

                Podschet();//произвести подсчет по методу
                           //table1BindingSource.Filter = "[punkt] LIKE '%" + Convert.ToString(textBox2.Text) + "%' OR [familia] LIKE '%" + Convert.ToString(textBox2.Text) + "%'"; //Фильтр по гриду   
            }
            else MessageBox.Show("Введите ФИО в строке поиска!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }
        private void textBox3_TextChanged(object sender, EventArgs e)//Поиск по №Заказа
        {
            var command = from table in db.GetTable<Table_1_incomplete>()
                          where table.N_Заказа == textBox3.Text.ToString()
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;

            if (textBox3.Text == "")//если поле очищено, отобразить базу
            {
                Disp_data();
            }
            Podschet();//произвести подсчет по методу
        }
        private void button8_Click_1(object sender, EventArgs e)//Обновить (Калькуляция)
        {
            ProgressBar();
            Disp_data();
            Podschet();//произвести подсчет по методу 
        }
        private void button9_Click(object sender, EventArgs e)//Вся база данных
        {
            //DispdatabaseAsync();//Отображает всю базу асинхронно  
            Disp_data_all_base();
            MessageBox.Show("База данных отображена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void button3_Click(object sender, EventArgs e)//Возврат
        {
            ProgressBar();
            if (Person.Name == "root")
            {
                var command = from table in db.GetTable<Table_1_incomplete>()
                              where table.Статус == "Возврат"
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            else
            {
                var command = from table in db.GetTable<Table_1_incomplete>()
                              where table.Филиал == Person.Name & table.Статус == "Возврат"
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            Podschet();
        }
        private void button18_Click(object sender, EventArgs e)//Выдано
        {
            ProgressBar();
            if (Person.Name == "root")
            {
                var command = from table in db.GetTable<Table_1_incomplete>()
                              where table.Статус == "Выдано"
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            else
            {
                var command = from table in db.GetTable<Table_1_incomplete>()
                              where table.Филиал == Person.Name & table.Статус == "Выдано"
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            Podschet();
        }

        public void Print_Registy()//Печать Реестра
        {
            if (Table.DtRegistry != null)
                if (Table.DtRegistry == null)
                {
                    MessageBox.Show("Сделайте выборку, невозможно сгенерировать реестр!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            //Обработка и Выдача реестра
            if (Table.DtRegistry.Rows.Count > 0 && Table.DtRegistry.Rows[0][10].ToString() != "Обработано"
                    & Table.DtRegistry.Rows[0][5].ToString() != "Отправлено"
                    & Table.DtRegistry.Rows[0][5].ToString() != "Ожидание"
                    & Table.DtRegistry.Rows[0][5].ToString() != "Розыск"
                    & Table.DtRegistry.Rows[0][5].ToString() != "Замена")
            {
                Select_status_Nr();//Выборка по статусу и сортировка по номеру реестра от больших значений к меньшим.                                      
                if (MessageBox.Show("Вы хотите обработать эти записи?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    for (int i = 0; i < Table.DtRegistry.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET obrabotka = @obrabotka, data_obrabotki = @data_obrabotki, nomer_reestra = @nomer_reestra, Nr=@Nr WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@obrabotka", "Обработано");
                        cmd.Parameters.AddWithValue("@data_obrabotki", DateTime.Today.AddDays(0));
                        cmd.Parameters.AddWithValue("@id", Table.DtRegistry.Rows[i][11].ToString());
                        cmd.Parameters.AddWithValue("@nomer_reestra", Number.Prefix_number);
                        cmd.Parameters.AddWithValue("@Nr", Number.Nr);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение 
                    MessageBox.Show("Обработка выполнена / Присвоен № Реестра!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //------Ручная вставка номера реестра и обработки----------//
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
                    {
                        dataGridView1.Rows[i].Cells[12].Value = Number.Prefix_number;
                        dataGridView1.Rows[i].Cells[10].Value = "Обработано";
                    }
                    for (int i = 0; i < Table.DtRegistry.Rows.Count; i++)//Цикл
                    {
                        Table.DtRegistry.Rows[i][12] = Number.Prefix_number;
                        Table.DtRegistry.Rows[i][10] = "Обработано";
                    }
                    //------Ручная вставка номера реестра и обработки----------//
                }
                //Выдача рееста в WORD
                ExportReestr_ToPDF();
                string status = Convert.ToString(dataGridView1.Rows[0].Cells[5].Value);//Статус
                //string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Контрагент                
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (status != "Возврат" /*| kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри" & kontragent != "ИП 'JUMPER'"*/)
                    {
                        Export_Reestr_To_Word(dataGridView1, sfd.FileName);
                    }
                    else if (status == "Возврат" /*| kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри" & kontragent == "ИП 'JUMPER'"*/)
                    {
                        Export_Reestr_To_Word_vozvrat(dataGridView1, sfd.FileName);
                    }
                }
                //Выдача рееста в EXCEL
                if (status != "Возврат" /*| kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри" & kontragent != "ИП 'JUMPER'"*/)
                {
                    sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                    sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Reestr_To_Excel(dataGridView1, sfd.FileName);
                    }
                }
                else if (status == "Возврат" /*| kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри" & kontragent == "ИП 'JUMPER'"*/)
                {
                    sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                    sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Reestr_To_Excel_vozvrat(dataGridView1, sfd.FileName);
                    }
                }
            }
            else if (Table.DtRegistry.Rows.Count > 0 && Table.DtRegistry.Rows[0][10].ToString() == "Обработано")
            {
                if (MessageBox.Show("Вы хотите открыть этот Реестр?", "Внимание! Эти данные уже обработаны!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    ExportReestr_ToPDF();
                    //Выдача рееста в WORD
                    //string nomer = dataGridView1.Rows[0].Cells[12].Value.ToString();//№
                    string status = Convert.ToString(dataGridView1.Rows[0].Cells[5].Value);//Статус
                    //string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Контрагент
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        if (status != "Возврат" /*| kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри" & kontragent != "ИП 'JUMPER'"*/)
                        {
                            Export_Reestr_To_Word(dataGridView1, sfd.FileName);
                        }
                        else if (status == "Возврат" /*| kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри" & kontragent == "ИП 'JUMPER'"*/)
                        {
                            Export_Reestr_To_Word_vozvrat(dataGridView1, sfd.FileName);
                        }
                    }
                    //Выдача рееста в EXCEL
                    if (status != "Возврат" /*| kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри" & kontragent != "ИП 'JUMPER'"*/)
                    {
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Reestr_To_Excel(dataGridView1, sfd.FileName);
                        }
                    }
                    else if (status == "Возврат" /*| kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри" & kontragent == "ИП 'JUMPER'"*/)
                    {
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Reestr_To_Excel_vozvrat(dataGridView1, sfd.FileName);
                        }
                    }
                }
            }
            else if (Table.DtRegistry.Rows.Count > 0 && Table.DtRegistry.Rows[0][5].ToString() == "Розыск" | Table.DtRegistry.Rows[0][5].ToString() == "Замена")
            {
                if (MessageBox.Show("Вы хотите открыть этот Реестр?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    //Выдача рееста в WORD
                    ExportReestr_ToPDF();
                    //string nomer = dataGridView1.Rows[0].Cells[12].Value.ToString();//№
                    //string status = Convert.ToString(dataGridView1.Rows[0].Cells[5].Value);//Статус
                    //string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Контрагент
                    //SaveFileDialog sfd = new SaveFileDialog();
                    //sfd.Filter = "Word Documents (*.docx)|*.docx";
                    //sfd.FileName = $"Реестр № {nomer} на {status}.docx";
                    //if (sfd.ShowDialog() == DialogResult.OK)
                    //{
                    //    Export_Reestr_To_Word(dataGridView1, sfd.FileName);
                    //}
                }
            }
            else if (Table.DtRegistry.Rows.Count <= 0)
            {
                MessageBox.Show("Выборка не дала результатов, невозможно сгенерировать реестр!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Эти данные нельзя обработать", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            if (Table.DtRegistry != null)
                Table.DtRegistry.Clear();//чистим DataTable
        }
        public void Print_Invoice()//Печать Накладной и за период
        {
            if (Table.DtInvoice != null)
                if (Table.DtInvoice == null)
                {
                    MessageBox.Show("Сделайте выборку, невозможно сгенерировать файл!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            if (Table.DtInvoice.Rows.Count > 0 && Table.DtInvoice.Rows[0][5].ToString() == "Ожидание")
            {
                Select_status_Nn();//(Для выдачи накладных)Выборка по статусу и сортировка по номеру накладеой от больших значений к меньшим.               
                if (MessageBox.Show("Вы хотите получить 'Накладную'? Нажмите Нет если хотите получить 'Cписок за период'!", "Внимание! Статус изменится на 'Отправлено' и присвоется номер", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    for (int i = 0; i < Table.DtInvoice.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET nomer_nakladnoy = @nomer_nakladnoy, status = @status, Nn=@Nn, filial=@filial WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@id", Table.DtInvoice.Rows[i][11].ToString());
                        cmd.Parameters.AddWithValue("@status", "Отправлено");
                        cmd.Parameters.AddWithValue("@nomer_nakladnoy", Number.Prefix_number);
                        cmd.Parameters.AddWithValue("@Nn", Number.Nn);
                        cmd.Parameters.AddWithValue("@filial", Person.Name);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение

                    //Выдача накладной
                    ExportInvoice_ToPDF();
                    string region = Convert.ToString(dataGridView1.Rows[0].Cells[9].Value);//Область
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Накладная № {Number.Prefix_number} - {region}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Nakladnaya_To_Word(dataGridView1, sfd.FileName);
                    }
                }
                else//Список за период (Ожидание)
                {
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = "Список за период.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Spisok_To_Word(dataGridView1, sfd.FileName);
                    }
                }
            }
            else if (Table.DtInvoice.Rows.Count > 0 && Table.DtInvoice.Rows[0][5].ToString() == "Отправлено")
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = "Список за период.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Export_Spisok_To_Word(dataGridView1, sfd.FileName);
                }
            }
            else if (Table.DtInvoice.Rows.Count <= 0)
            {
                MessageBox.Show("Выборка не дала результатов, невозможно сгенерировать накладную!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Эти данные нельзя обработать", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }

            if (Table.DtInvoice != null)
                Table.DtInvoice.Clear();//чистим DataTable
        }

        private void dataGridView2_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)//Окраска статусов dataGridView2
        {
            for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
            {
                string status = Convert.ToString(dataGridView2.Rows[i].Cells[11].Value);//статус
                string obrabotka = Convert.ToString(dataGridView2.Rows[i].Cells[14].Value);//Обработка
                string V = Convert.ToString("Выдано");
                string L = Convert.ToString("Возврат");
                string K = Convert.ToString("Ожидание");
                string W = Convert.ToString("Отправлено");

                if (status == V & obrabotka == "Обработано")
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;//Выдано и Обработано                
                }
                else if (status == L & obrabotka == "Обработано")
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightSalmon;//Возврат и Обработано                  
                }
                else if (status == V)
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightSeaGreen;//Выдано
                }
                else if (status == L)
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Salmon;//Возврат                 
                }
                else if (status == W)
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;//Отправлено                
                }
            }
        }
        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)//Окраска статусов dataGridView1
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                string S = Convert.ToString(dataGridView1.Rows[i].Cells[5].Value);//статус
                string V = Convert.ToString("Выдано");
                string L = Convert.ToString("Возврат");
                string K = Convert.ToString("Ожидание");
                string W = Convert.ToString("Отправлено");
                if (S == V)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;//Выдано
                }
                else if (S == L)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightSalmon;//Возврат                  
                }
                else if (S == W)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;//Отправлено                
                }
            }
        }

        public void Export_Reestr_To_Word(DataGridView dataGridView1, string filename)//Метод экспорта в Word Реестр
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Реестр");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 9;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(9);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 8; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView1.Rows.Count != 0)
            {
                //удалить столбцы
                this.dataGridView1.Columns.RemoveAt(4);//дата записи
                int RowCount = dataGridView1.Rows.Count;
                int ColumnCount = dataGridView1.Columns.Count - 0;// столбцы в гриде (-3 последних)id и обработанные и клиент не нужны
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
                string status = Convert.ToString(dataGridView1.Rows[0].Cells[4].Value);//Статус
                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = $"Итого:    {Summ.Quantity}                    {Summ.Sum}" +
                //Environment.NewLine + " Сумма за услугу " + Summ.SumService +
                Environment.NewLine +
                Environment.NewLine + "Проверила____________________________" + Environment.NewLine;

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
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру                
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -35;//Установка отступа слева              
                if (Convert.ToString(dataGridView1.Rows[0].Cells[4].Value) == "Выдано")
                {
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();//Удалить столбец причина на Выданных реестрах
                    oDoc.Application.Selection.Tables[1].Rows.LeftIndent = 20;//Установка отступа слева
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[6].Delete();
                }
                if (Convert.ToString(dataGridView1.Rows[0].Cells[4].Value) == "Возврат")
                {
                    oDoc.Application.Selection.Tables[1].Columns[6].Width = 100;
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                    oDoc.Application.Selection.Tables[1].Columns[7].Delete();
                }
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 150;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 80;
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 65;
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 60;
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 55;
                //текст заголовка
                //int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
                //string prefix_number = comboBox10.Text + number;
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
                    {
                        headerRange.Text = $"Реестр №  {Number.Prefix_number}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {Partner.Name}" +
                        Environment.NewLine;
                    }
                    else if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) == "Обработано")
                    {
                        string Reestr = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        headerRange.Text = $"Реестр №  {Reestr}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {Partner.Name}" +
                        Environment.NewLine;
                    }
                    headerRange.Font.Size = 12;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "TLC-Express       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Export_Reestr_To_Word_vozvrat(DataGridView dataGridView1, string filename)//Метод экспорта в Word Реестра на возврат для TOO Sapar delivery
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Реестр");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 9;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(9);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 8; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView1.Rows.Count != 0)
            {
                //удаление столбца
                this.dataGridView1.Columns.RemoveAt(4);//дата записи
                string status = Convert.ToString(dataGridView1.Rows[0].Cells[4].Value);//Статус
                //string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[7].Value);//Контрагент
                int RowCount = dataGridView1.Rows.Count;
                int ColumnCount = dataGridView1.Columns.Count - 0;// столбцы в гриде (-3 последних)id и обработанные и клиент не нужны              
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

                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = $"Итого:    {Summ.Quantity}                    {Summ.Sum}" +
                Environment.NewLine + $"Сумма за возврат   {Summ.SumReturn}" +
                Environment.NewLine +
                Environment.NewLine + "Проверила____________________________" + Environment.NewLine;

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
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -35;//Установка отступа слева  

                oDoc.Application.Selection.Tables[1].Columns[7].Delete();//Удалить столбец плата за услугу
                oDoc.Application.Selection.Tables[1].Columns[7].Delete();//Удалить столбец контрагент
                oDoc.Application.Selection.Tables[1].Columns[7].Delete();//Удалить столбец область
                oDoc.Application.Selection.Tables[1].Columns[7].Delete();//Удалить столбец обработка               
                oDoc.Application.Selection.Tables[1].Columns[7].Delete();//Удалить столбец id
                oDoc.Application.Selection.Tables[1].Columns[8].Delete();//Удалить столбец за наложеный

                oDoc.Application.Selection.Tables[1].Columns[1].Width = 120;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 65;
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 60;
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 50;
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 70;
                oDoc.Application.Selection.Tables[1].Columns[6].Width = 70;
                oDoc.Application.Selection.Tables[1].Columns[7].Width = 40;
                oDoc.Application.Selection.Tables[1].Columns[8].Width = 40;
                //текст заголовка
                //int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
                //string prefix_number = comboBox10.Text + number;
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
                    {
                        headerRange.Text = $"Реестр №  {Number.Prefix_number}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {Partner.Name}" +
                        Environment.NewLine;
                    }
                    else if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) == "Обработано")
                    {
                        string Reestr = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        headerRange.Text = $"Реестр №  {Reestr}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {Partner.Name}" +
                        Environment.NewLine;
                    }
                    headerRange.Font.Size = 12;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "TLC-Express       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Export_Reestr_To_Excel(DataGridView dataGridView1, string filename)//Метод экспорта в Excel Реестр
        {
            //int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
            //string prefix_number = comboBox10.Text + number;
            Excell._Application app = new Excell.Application();// Создание Excel Application
            Excell._Workbook workbook = app.Workbooks.Add(Type.Missing);// создание новой книги внутри приложения Excel
            Excell._Worksheet worksheet = null;// создание нового листа Excel в книге
            app.Visible = true;// увидеть Excel лист за программу
                               //worksheet = workbook.Sheets["Книга1"];// Получить ссылку на первом листе. По умолчанию его имя Лист1.
            worksheet = workbook.ActiveSheet;// Сохранить свою ссылку на таблицу

            // Хранение часть заголовка в Excel
            for (int i = 1; i < dataGridView1.Columns.Count - 5; i++)// заголовки в гриде (-5 последних) не нужны
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            // Хранение каждой строки и столбца значение для Excel лист (получает данные из DataGridView и заполняет клетки.)
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count - 6; j++)// столбцы в гриде (-6 последних) не нужны
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            //Добавление Суммы и Кол-во
            DateTime Now = DateTime.Today;
            worksheet.Cells[2, "I"] = "Сумма " + Summ.Sum;
            worksheet.Cells[3, "I"] = "Кол-во " + Summ.Quantity;
            if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
            {
                string Reestr = Number.Prefix_number;
                worksheet.Cells[4, "I"] = "Реестр№ " + Reestr + " от " + Now;
                worksheet.Name = "Реестр№ " + Reestr + " TLC -Express";// Изменение названия активного листа
            }
            else if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) == "Обработано")
            {
                string Reestr = dataGridView1.Rows[0].Cells[11].Value.ToString();
                worksheet.Cells[4, "I"] = "Реестр№ " + Reestr + " от " + Now;
                worksheet.Name = "Реестр№ " + Reestr + " TLC -Express";// Изменение названия активного листа
            }
            worksheet.Columns.AutoFit();//Автоматическая ширина колонок
            worksheet.Rows[1].Font.Bold = true; //Жирный шрифт
                                                //----------------------------------------------//
                                                // Сохранить приложение
            workbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excell.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //app.Quit();// Выход из приложения           
            Marshal.ReleaseComObject(app);// Уничтожение объекта Excel.          
            GC.GetTotalMemory(true);// Вызываем сборщик мусора для немедленной очистки памяти
            label1.Text = "Реестр сохранен";
        }
        public void Export_Reestr_To_Excel_vozvrat(DataGridView dataGridView1, string filename)//Метод экспорта в Excel Реестр на возврат для TOO Sapar delivery
        {
            //int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
            //string prefix_number = comboBox10.Text + number;
            Excell._Application app = new Excell.Application();// Создание Excel Application
            Excell._Workbook workbook = app.Workbooks.Add(Type.Missing);// создание новой книги внутри приложения Excel
            Excell._Worksheet worksheet = null;// создание нового листа Excel в книге
            app.Visible = true;// увидеть Excel лист за программу
                               //worksheet = workbook.Sheets["Книга1"];// Получить ссылку на первом листе. По умолчанию его имя Лист1.
            worksheet = workbook.ActiveSheet;// Сохранить свою ссылку на таблицу

            // Хранение часть заголовка в Excel
            for (int i = 1; i < dataGridView1.Columns.Count - 6; i++)// заголовки в гриде (-5 последних) не нужны
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            // Хранение каждой строки и столбца значение для Excel лист (получает данные из DataGridView и заполняет клетки.)
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count - 7; j++)// столбцы в гриде (-6 последних) не нужны
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }
            //Добавление Суммы и Кол-во
            DateTime Now = DateTime.Today;
            worksheet.Cells[2, "I"] = "Сумма " + Summ.Sum;
            worksheet.Cells[3, "I"] = "Кол-во " + Summ.Quantity;
            if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
            {
                string Reestr = Number.Prefix_number;
                worksheet.Cells[4, "I"] = "Реестр№ " + Reestr + " от " + Now;
                worksheet.Name = "Реестр№ " + Reestr + " TLC -Express";// Изменение названия активного листа
            }
            else if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) == "Обработано")
            {
                string Reestr = dataGridView1.Rows[0].Cells[11].Value.ToString();
                worksheet.Cells[4, "I"] = "Реестр№ " + Reestr + " от " + Now;
                worksheet.Name = "Реестр№ " + Reestr + " TLC -Express";// Изменение названия активного листа
            }
            worksheet.Columns.AutoFit();//Автоматическая ширина колонок
            worksheet.Rows[1].Font.Bold = true; //Жирный шрифт
                                                //----------------------------------------------//
                                                // Сохранить приложение
            workbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excell.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //app.Quit();// Выход из приложения           
            Marshal.ReleaseComObject(app);// Уничтожение объекта Excel.          
            GC.GetTotalMemory(true);// Вызываем сборщик мусора для немедленной очистки памяти
            label1.Text = "Реестр сохранен";
        }
        public void Export_Nakladnaya_To_Word(DataGridView dataGridView1, string filename)//Метод экспорта в Word Накладная
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);

            rng.InsertBefore("Накладная");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 9;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(5);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 8; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView1.Rows.Count != 0)
            {
                //удалить столбцы
                this.dataGridView1.Columns.RemoveAt(4);//дата записи
                int RowCount = dataGridView1.Rows.Count;
                int ColumnCount = dataGridView1.Columns.Count - 9;// столбцы в гриде (-7 последних)
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

                string region = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Область
                //string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[7].Value);//Контрагент
                oDoc.Content.SetRange(0, 0);
                oDoc.Content.Text = $"                             Итого:    {Summ.Quantity}                    {Summ.Sum}" +
                Environment.NewLine +
                Environment.NewLine + $"Принял__________________              Сдал_____________________" + Environment.NewLine;

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
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -35;//Установка отступа слева
                //текст заголовка
                //int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[22].Value) + 1;
                //string prefix_number = comboBox10.Text + number;
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text = Partner.Name + Environment.NewLine + "Накладная № " + Number.Prefix_number + " от " + Convert.ToString(Now.ToString("dd.MM.yyyy")) + " куда " + region +
                    Environment.NewLine;
                    headerRange.Font.Size = 16;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "TLC-Express       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Export_Spisok_To_Word(DataGridView dataGridView1, string filename)//Метод экспорта в Word Список за период
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);

            rng.InsertBefore("Список за период");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 9;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(5);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 8; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView1.Rows.Count != 0)
            {
                int RowCount = dataGridView1.Rows.Count;
                int ColumnCount = dataGridView1.Columns.Count - 7;// столбцы в гриде (-7 последних)
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
                //string client = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Клиент
                //DateTime DatePriem = Convert.ToDateTime(dataGridView2.Rows[0].Cells[8].Value);
                oDoc.Content.SetRange(0, 0);
                oDoc.Content.Text = $"                             Итого:     {Summ.Quantity}                    {Summ.Sum}" +
                Environment.NewLine +
                Environment.NewLine + $"Принял__________________              Сдал_____________________" + Environment.NewLine;

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
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру              
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 85;
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 60;
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 55;
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 45;
                oDoc.Application.Selection.Tables[1].Columns[6].Width = 50;
                oDoc.Application.Selection.Tables[1].Columns[7].Width = 70;
                oDoc.Application.Selection.Tables[1].Columns[8].Width = 40;
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -35;//Установка отступа слева
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул

                    headerRange.Text = "Список за период " +
                    Environment.NewLine + $"c  {Convert.ToString(Dates.StartDate.ToString("dd.MM.yyyy "))}  по  {Convert.ToString(Dates.EndDate.ToString(" dd.MM.yyyy"))}" +
                    Environment.NewLine +
                    Environment.NewLine + $"Отправитель  {Partner.Name}" +
                    Environment.NewLine;

                    headerRange.Font.Size = 16;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "TLC-Express       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Export_Spisok_Prinyatyh_To_Word(DataGridView dataGridView5, string filename)//Метод экспорта в Word СПИСКА ПРИНЯТЫХ
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);

            rng.InsertBefore("Список принятых");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 9;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(9);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 8; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView5.Rows.Count != 0)
            {
                //удалить столбцы
                this.dataGridView5.Columns.RemoveAt(4);//дата приема
                int RowCount = dataGridView5.Rows.Count;
                int ColumnCount = dataGridView5.Columns.Count - 2;// столбцы в гриде (-2 последних)
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r, c] = dataGridView5.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ
                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = $"Итого: {Summ.Quantity}" +
                Environment.NewLine + $"Сумма объявленной ценности  {Summ.Sum}" +
                Environment.NewLine + $"Сумма за услугу  {Summ.SumService}" + Environment.NewLine +
                Environment.NewLine + "Исполнитель____________________________" + Environment.NewLine +
                Environment.NewLine + "Отправитель____________________________" + Environment.NewLine;

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
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = dataGridView5.Columns[c].HeaderText;
                }
                //стиль таблицы          
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру     
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 60;//Ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 60;
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 90;
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 60;
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 50;
                oDoc.Application.Selection.Tables[1].Columns[6].Width = 40;
                oDoc.Application.Selection.Tables[1].Columns[7].Width = 30;
                oDoc.Application.Selection.Tables[1].Columns[8].Width = 40;
                oDoc.Application.Selection.Tables[1].Columns[9].Width = 40;
                oDoc.Application.Selection.Tables[1].Columns[10].Width = 40;
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -35;//Установка отступа слева
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    DateTime DatePriem = Convert.ToDateTime(dataGridView2.Rows[0].Cells[12].Value);
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    //int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[21].Value) + 1;
                    //string prefix_number = comboBox10.Text + number;
                    if (dataGridView5.Rows[0].Cells[11].Value.ToString() == "0")
                    {
                        headerRange.Text = $"СПИСОК № {Number.Prefix_number}" +
                        Environment.NewLine + $"от {Convert.ToString(DatePriem.ToString("dd.MM.yyyy"))} принятых в ТЛЦ ГП 'Спецсвязь' " +
                        Environment.NewLine +
                        Environment.NewLine + $"Отправитель {Partner.Name}" +
                        Environment.NewLine;
                    }
                    else if (dataGridView5.Rows[0].Cells[11].Value.ToString() != "0")
                    {
                        string nomer = dataGridView5.Rows[0].Cells[11].Value.ToString();//№
                        headerRange.Text = $"СПИСОК № {nomer}" +
                        Environment.NewLine + $"от {Convert.ToString(DatePriem.ToString("dd.MM.yyyy"))} принятых в ТЛЦ ГП 'Спецсвязь' " +
                        Environment.NewLine +
                        Environment.NewLine + $"Отправитель {Partner.Name}" +
                        Environment.NewLine;
                    }
                    headerRange.Font.Size = 12;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    footerRange.Text = "TLC-Express       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Export_Spisok_Prinyatyh_To_Excel(DataGridView dataGridView5, string filename)//Метод экспорта в Excel СПИСКА ПРИНЯТЫХ
        {
            Excell._Application app = new Excell.Application();// Создание Excel Application
            Excell._Workbook workbook = app.Workbooks.Add(Type.Missing);// создание новой книги внутри приложения Excel
            Excell._Worksheet worksheet = null;// создание нового листа Excel в книге
            app.Visible = true;// увидеть Excel лист за программу
            worksheet = workbook.ActiveSheet;// Сохранить свою ссылку на таблицу
            int x = 0;
            string client = Convert.ToString(dataGridView2.Rows[x].Cells[15].Value);//Клиент
            worksheet.Name = "Список принятых ";// Изменение названия активного листа
                                                // Хранение часть заголовка в Excel
            for (int i = 1; i < dataGridView5.Columns.Count - 3; i++)// заголовки в гриде (-2 последних) не нужны
            {
                worksheet.Cells[1, i] = dataGridView5.Columns[i - 1].HeaderText;
            }
            // Хранение каждой строки и столбца значение для Excel лист (получает данные из DataGridView и заполняет клетки.)
            for (int i = 0; i < dataGridView5.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView5.Columns.Count - 4; j++)// столбцы в гриде (-3 последних) не нужны
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView5.Rows[i].Cells[j].Value.ToString();
                }
            }
            //Добавление Суммы и Кол-во
            DateTime Now = DateTime.Today;
            worksheet.Cells[2, "k"] = "Сумма " + Summ.Sum;
            worksheet.Cells[3, "k"] = "Кол-во " + Summ.Quantity;
            worksheet.Columns.AutoFit();//Автоматическая ширина колонок
            worksheet.Rows[1].Font.Bold = true; //Жирный шрифт
                                                //----------------------------------------------//
                                                // Сохранить приложение
            workbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excell.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //app.Quit();// Выход из приложения           
            Marshal.ReleaseComObject(app);// Уничтожение объекта Excel.          
            GC.GetTotalMemory(true);// Вызываем сборщик мусора для немедленной очистки памяти
        }

        private void ExportReestr_ToPDF()//Метод экспорта Реестра в PDF 
        {
            string status = Table.DtRegistry.Rows[0][5].ToString();//Статус
            //string kontragent = Table.DtRegistry.Rows[0][8].ToString();//Контрагент
            DateTime Now = DateTime.Now;
            //int number = Number.Nr;
            //string prefix_number = comboBox10.Text + number;
            string processing = Table.DtRegistry.Rows[0][10].ToString();
            string Reestr = "";
            string Heading = "";
            //Определение шрифта необходимо для сохранения кириллического текста
            //Иначе мы не увидим кириллический текст
            //Если мы работаем только с англоязычными текстами, то шрифт можно не указывать
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\Arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 9, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.BOLD);

            PdfPTable table = null;
            //Обход по всем таблицам датасета
            for (int i = 0; i < Table.DtRegistry.Rows.Count; i++)
            {
                //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                table = new PdfPTable(Table.DtRegistry.Columns.Count);
                table.DefaultCell.Padding = 1;
                table.WidthPercentage = 100;
                if (status == "Выдано")
                {
                    float[] widths = new float[] { 150f, 80f, 65f, 60f, 0f, 60f, 0f, 0f, 0f, 0f, 0f, 0f, 0f, 0f, 0f };
                    table.SetWidths(widths);
                }
                else if (status == "Возврат")
                {
                    float[] widths = new float[] { 150f, 80f, 65f, 60f, 0f, 60f, 60f, 0f, 0f, 0f, 0f, 0f, 60f, 0f, 60f };
                    table.SetWidths(widths);
                }
                table.HorizontalAlignment = Element.ALIGN_LEFT;
                table.DefaultCell.BorderWidth = 1;

                //Добавим в таблицу общий заголовок                
                if (processing != "Обработано")
                {
                    Reestr = Number.Prefix_number;                   
                }
                else if (processing == "Обработано")
                {
                    Reestr = Table.DtRegistry.Rows[0][12].ToString();
                }
                    Heading = $"Реестр №  {Reestr}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. " +
                    Environment.NewLine + $"отправлений с наложенным платежом от {Partner.Name}" +
                    Environment.NewLine + Environment.NewLine;

                PdfPCell cell = new PdfPCell(new Phrase(Heading, fontBold));
                cell.Colspan = Table.DtRegistry.Columns.Count;
                cell.HorizontalAlignment = 1;
                //Убираем границу первой ячейки, чтобы была как заголовок
                cell.Border = 0;
                table.AddCell(cell);
                //Сначала добавляем заголовки таблицы
                for (int k = 0; k < Table.DtRegistry.Columns.Count; k++)
                {
                    cell = new PdfPCell(new Phrase(Table.DtRegistry.Columns[k].ColumnName, font));
                    //Фоновый цвет (необязательно, просто сделаем по красивее)
                    cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    //cell.Border = 0;
                    table.AddCell(cell);
                }
                //Добавляем все остальные ячейки
                for (int x = 0; x < Table.DtRegistry.Rows.Count; x++)
                {
                    for (int j = 0; j < Table.DtRegistry.Columns.Count; j++)
                    {
                        table.AddCell(new Phrase(Table.DtRegistry.Rows[x][j].ToString(), font));
                    }
                }
            }
            //----------------------------------------------------------------------------------------------------------//
            //Exporting to PDF
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Pdf File |*.pdf";
            sfd.FileName = $"Реестр № {Reestr} на {status}.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create))
                {
                    Document Doc = new Document(PageSize.A4, 40f, 40f, 20f, 20f);//Для А4
                                                                                 //Document Doc = new Document(new iTextSharp.text.Rectangle(Width, Height), 0, 0, 0, 0);
                                                                                 //Document Doc = new Document(new iTextSharp.text.Rectangle(120, 1000), 0f, 0f, 0f, 0f);//для чека
                    PdfWriter.GetInstance(Doc, stream);
                    Doc.Open();
                    DateTime date = DateTime.Now;

                    Doc.Add(table);
                    Doc.Add(new Paragraph(Environment.NewLine));
                    if (status == "Выдано")
                    {
                        Doc.Add(new Paragraph($"Итого:    {Summ.Quantity}                    {Summ.Sum}" +
                //Environment.NewLine + " Сумма за услугу " + Summ.SumService +                       
                Environment.NewLine +
                Environment.NewLine + "Проверил(а)____________________________" + Environment.NewLine, font));
                    }
                    if (status == "Возврат")
                    {
                        Doc.Add(new Paragraph($"Итого:    {Summ.Quantity}                    {Summ.Sum}" +
                Environment.NewLine + $"Сумма за возврат   {Summ.SumReturn}" +
                Environment.NewLine +
                Environment.NewLine + "Проверил(а)____________________________" + Environment.NewLine, font));

                    }
                    Doc.Close();
                    stream.Close();
                }
            }

            // Печать на устройство, установленное используемым по умолчанию
            Process printJob = new Process();
            printJob.StartInfo.FileName = sfd.FileName;//Открыть документ
            printJob.StartInfo.UseShellExecute = true;
            //printJob.StartInfo.Verb = "print";
            printJob.Start();

            //printJob.WaitForInputIdle();
            //printJob.Kill();
        }
        private void ExportInvoice_ToPDF()//Метод экспорта Накладной в PDF 
        {
            string status = Table.DtInvoice.Rows[0][5].ToString();//Статус
            //string kontragent = Table.DtInvoice.Rows[0][8].ToString();//Контрагент
            string region = Table.DtInvoice.Rows[0][9].ToString();//Область
            DateTime Now = DateTime.Now;
            //int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
            //string prefix_number = comboBox10.Text + number;
            string Heading = "";
            //Определение шрифта необходимо для сохранения кириллического текста
            //Иначе мы не увидим кириллический текст
            //Если мы работаем только с англоязычными текстами, то шрифт можно не указывать
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\Arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 9, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.BOLD);

            PdfPTable table = null;
            //Обход по всем таблицам датасета
            for (int i = 0; i < Table.DtInvoice.Rows.Count; i++)
            {
                //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                table = new PdfPTable(Table.DtInvoice.Columns.Count);
                table.DefaultCell.Padding = 1;
                table.WidthPercentage = 100;
                float[] widths = new float[] { 150f, 80f, 65f, 60f, 0f, 60f, 0f, 0f, 0f, 0f, 0f, 0f, 0f, 0f, 0f };
                table.SetWidths(widths);
                table.HorizontalAlignment = Element.ALIGN_LEFT;
                table.DefaultCell.BorderWidth = 1;
                //Добавим в таблицу общий заголовок                
                Heading = Partner.Name + Environment.NewLine + "Накладная № " + Number.Prefix_number + " от " + Convert.ToString(Now.ToString("dd.MM.yyyy")) + " куда " + region +
                Environment.NewLine;

                PdfPCell cell = new PdfPCell(new Phrase(Heading, fontBold));
                cell.Colspan = Table.DtInvoice.Columns.Count;
                cell.HorizontalAlignment = 1;
                //Убираем границу первой ячейки, чтобы была как заголовок
                cell.Border = 0;
                table.AddCell(cell);
                //Сначала добавляем заголовки таблицы
                for (int k = 0; k < Table.DtInvoice.Columns.Count; k++)
                {
                    cell = new PdfPCell(new Phrase(Table.DtInvoice.Columns[k].ColumnName, font));
                    //Фоновый цвет (необязательно, просто сделаем по красивее)
                    cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                    //cell.Border = 0;
                    table.AddCell(cell);
                }
                //Добавляем все остальные ячейки
                for (int x = 0; x < Table.DtInvoice.Rows.Count; x++)
                {
                    for (int j = 0; j < Table.DtInvoice.Columns.Count; j++)
                    {
                        table.AddCell(new Phrase(Table.DtInvoice.Rows[x][j].ToString(), font));
                    }
                }
            }
            //----------------------------------------------------------------------------------------------------------//
            //Exporting to PDF
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Pdf File |*.pdf";
            sfd.FileName = $"Накладная № {Number.Prefix_number} - {region}.pdf";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                using (FileStream stream = new FileStream(sfd.FileName, FileMode.Create))
                {
                    Document Doc = new Document(PageSize.A4, 40f, 40f, 20f, 20f);//Для А4
                                                                                 //Document Doc = new Document(new iTextSharp.text.Rectangle(Width, Height), 0, 0, 0, 0);
                                                                                 //Document Doc = new Document(new iTextSharp.text.Rectangle(120, 1000), 0f, 0f, 0f, 0f);//для чека
                    PdfWriter.GetInstance(Doc, stream);
                    Doc.Open();
                    DateTime date = DateTime.Now;

                    Doc.Add(table);
                    Doc.Add(new Paragraph(Environment.NewLine));
                    Doc.Add(new Paragraph($" Итого:    {Summ.Quantity}                    {Summ.Sum}" +
                    Environment.NewLine +
                    Environment.NewLine + $"Принял__________________              Сдал_____________________" + Environment.NewLine));

                    Doc.Close();
                    stream.Close();
                }
            }

            // Печать на устройство, установленное используемым по умолчанию
            Process printJob = new Process();
            printJob.StartInfo.FileName = sfd.FileName;//Открыть документ
            printJob.StartInfo.UseShellExecute = true;
            //printJob.StartInfo.Verb = "print";
            printJob.Start();

            //printJob.WaitForInputIdle();
            //printJob.Kill();
        }

        private void button17_Click(object sender, EventArgs e)//Выход
        {
            if (MessageBox.Show("Вы действительно хотите выйти?!", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                formLogin.Show();
                this.Hide();
            }
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)//Закрытие формы Выход
        {
            Application.Exit();
        }
        private void linkLabel1_LinkClicked_1(object sender, LinkLabelLinkClickedEventArgs e)// ссылка на страничку
        {
            System.Diagnostics.Process.Start("https://alesunix.github.io/");
            linkLabel1.BackColor = Color.Transparent;
        }

        //---------------------Админпанель------------------------//
        public void Logins_select()//Вывод пользователей в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT login FROM [Table_Login] WHERE login NOT IN ('root') ORDER BY login";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow row in dt.Rows)
            {
                comboBox11.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение
        }
        public static void EnableTab(TabPage page, bool enable)//Класс TabPage скрывает свойство Enabled.(Открываем это свойство!)
        {
            foreach (Control ctl in page.Controls) ctl.Enabled = enable;
        }
        private void button15_Click(object sender, EventArgs e)//Выборка в Админпанели
        {
            //10. Выборка за период (Дата записи) - 'Период + Клиент'.
            if (comboBox3.Text != "" & textBox6.Text == "" & comboBox9.Text == "")
            {
                dataGridView6.Visible = true;
                button15.Enabled = true;
                DateTime date = new DateTime();
                date = dateTimePicker4.Value;
                DateTime date2 = dateTimePicker5.Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia, punkt, N_zakaza, summ, data_zapisi, status, prichina, plata_za_uslugu, client, oblast, obrabotka, id," +
                    "nomer_reestra, plata_za_nalog, plata_za_vozvrat = (plata_za_uslugu - plata_za_nalog)" +
                    "FROM [Table_1] WHERE (data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                cmd.Parameters.AddWithValue("StartDate", date);
                cmd.Parameters.AddWithValue("EndDate", date2);
                cmd.Parameters.AddWithValue("@client", comboBox3.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView6.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
            }
            else if (textBox6.Text != "" & comboBox3.Text != "" & comboBox9.Text == "")
            {
                dataGridView6.Visible = true;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia, punkt, N_zakaza, summ, data_zapisi, status, prichina, plata_za_uslugu, client, oblast, obrabotka, id," +
                    "nomer_spiska, plata_za_nalog, plata_za_vozvrat = (plata_za_uslugu - plata_za_nalog)" +
                    "FROM [Table_1] WHERE nomer_spiska = @nomer_spiska AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@nomer_spiska", textBox6.Text);
                cmd.Parameters.AddWithValue("@client", comboBox3.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView6.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
            }
            else if (textBox6.Text != "" & comboBox3.Text != "" & comboBox9.Text != "")
            {
                dataGridView6.Visible = true;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia, punkt, N_zakaza, summ, data_zapisi, status, prichina, plata_za_uslugu, client, oblast, obrabotka, id," +
                    "nomer_reestra, plata_za_nalog, plata_za_vozvrat = (plata_za_uslugu - plata_za_nalog)" +
                    "FROM [Table_1] WHERE nomer_reestra = @nomer_reestra AND client = @client AND status = @status ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@nomer_reestra", textBox6.Text);
                cmd.Parameters.AddWithValue("@client", comboBox3.Text);
                cmd.Parameters.AddWithValue("@status", comboBox9.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView6.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
            }
            else
            {
                MessageBox.Show("Выберите контрагента!");
            }
            Podschet();//произвести подсчет из метода
            textBox6.Text = "";
            comboBox9.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
        }
        private void button19_Click(object sender, EventArgs e)//Выгрузка в Админпанели
        {
            //Список за период АРХИВ
            if (dataGridView6.Rows.Count > 0)
            {
                button19.Enabled = false;
                button19.Text = "Ожидайте идет выгрузка!";
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = "Список за период АРХИВ" + ".docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Export_Spisok_To_Word(dataGridView6, sfd.FileName);
                }
            }
            else if (dataGridView6.Rows.Count <= 0)
            {
                MessageBox.Show("Выборка не дала результатов!", "Внимание!");
            }
            button19.Text = "Выгрузить";
            button19.Enabled = true;
        }
        private void button20_Click(object sender, EventArgs e)//Удалить из базы в Админпанели
        {
            if (dataGridView6.Rows.Count > 0)
            {
                if (MessageBox.Show("Вы действительно хотите удалить эти записи?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    for (int i = 0; i < dataGridView6.Rows.Count; i++)
                    {
                        SqlCommand cmd = new SqlCommand("DELETE FROM [Table_1] WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView6.Rows[i].Cells[11].Value));
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение 
                    MessageBox.Show("Записи успешно удалены!", "Внимание!");
                }
            }
            else if (dataGridView6.Rows.Count <= 0)
            {
                MessageBox.Show("Выборка не дала результатов!", "Внимание!");
            }


        }
        private void button21_Click(object sender, EventArgs e)//Изменить номер Реестра
        {
            if (dataGridView6.Rows.Count > 0)
            {
                if (textBox6.Text != "")
                {
                    con.Open();//открыть соединение
                    for (int i = 0; i < dataGridView6.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET nomer_reestra = @nomer_reestra WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView6.Rows[i].Cells[11].Value));
                        cmd.Parameters.AddWithValue("@nomer_reestra", textBox6.Text);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение 
                    MessageBox.Show("Реестр изменен!", "Внимание!");
                }
                else MessageBox.Show("Введите номер реестра на который хотите изменить!", "Внимание!");
            }
            else MessageBox.Show("Произведите выборку!", "Внимание!");
        }
        private void button25_Click(object sender, EventArgs e)//Изменить суффикс
        {
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("UPDATE [Table_Suffix] SET name = @name WHERE id = @id", con);
            cmd.Parameters.AddWithValue("@id", 1);
            cmd.Parameters.AddWithValue("@name", comboBox10.Text);
            cmd.ExecuteNonQuery();
            con.Close();//закрыть соединение 
        }
        private void button24_Click(object sender, EventArgs e)//test суффикс
        {
            string MyString = "a10";
            char[] suffix = { 'a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't', 'u', 'v', 'w', 'x', 'y', 'z', '!', ' ' };
            string NewString = MyString.TrimStart(suffix);
            int number = Convert.ToInt32(NewString) + 1;
            string reestr = comboBox10.Text + number;
            button24.Text = reestr.ToString();
        }
        private void button23_Click(object sender, EventArgs e)//Добавить столбец в таблицу базы
        {
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("ALTER TABLE [Table_1] ADD number INT NULL", con);
            cmd.ExecuteNonQuery();
            con.Close();//закрыть соединение
            MessageBox.Show("Столбец добавлен в таблицу!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void button16_Click(object sender, EventArgs e)//Начать нумерацию заново
        {
            Disp_data_all_base();
            con.Open();//открыть соединение
            for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
            {
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET Nr = @Nr, Ns = @Ns, Nn = @Nn WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[i].Cells[0].Value));
                cmd.Parameters.AddWithValue("@Nr", 0);
                cmd.Parameters.AddWithValue("@Ns", 0);
                cmd.Parameters.AddWithValue("@Nn", 0);
                cmd.ExecuteNonQuery();
            }
            con.Close();//закрыть соединение 
            MessageBox.Show("0 присвоен!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void button13_Click(object sender, EventArgs e)//Добавить контрагента и тариф
        {
            if (comboBox6.Text != "" & comboBox7.Text != "")
            {
                con.Open();
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_Partner] (name, tarif) VALUES (@name, @tarif)", con);
                cmd.Parameters.AddWithValue("@name", comboBox6.Text);
                cmd.Parameters.AddWithValue("@tarif", comboBox7.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Контрагент добавлен!", "Внимание!");
            }
            else
            {
                MessageBox.Show("Контрагент не добавлен!", "Внимание!");
            }

        }
        private void button27_Click(object sender, EventArgs e)//Добавить юзера
        {
            if (comboBox11.Text != "" & textBox7.Text != "" & comboBox12.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_Login] (login, pass, access) VALUES (@login, @pass, @access)", con);
                cmd.Parameters.AddWithValue("@login", comboBox11.Text);
                cmd.Parameters.AddWithValue("@pass", textBox7.Text);
                cmd.Parameters.AddWithValue("@access", comboBox12.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Вы успешно добавили нового юзера!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)//График
        {
            Graph graph = new Graph();
            graph.Show();
        }


        private void накладнаяToolStripMenuItem_Click(object sender, EventArgs e)//Накладная
        {
            Invoice Invoice = new Invoice();
            Invoice.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            Invoice.Show();
        }
        private void реестрToolStripMenuItem_Click(object sender, EventArgs e)//Реестр
        {
            Registry Registry = new Registry(this.dataGridView1, this.dataGridView2);// передаем ссылку на грид в форму Registry
            Registry.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            Registry.Show();
        }
        private void периодToolStripMenuItem_Click(object sender, EventArgs e)//Период
        {
            Period Period = new Period(this.dataGridView1, this.dataGridView2);// передаем ссылку на грид в форму Period
            Period.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            Period.Show();
        }
        private void toolStripButton4_Click(object sender, EventArgs e)//Поиск реестра
        {
            Search_registry Search_registry = new Search_registry(this.dataGridView1, this.dataGridView2);// передаем ссылку на грид в форму Search_registry
            Search_registry.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            Search_registry.Show();
        }
        private void списокПринятыхToolStripMenuItem_Click(object sender, EventArgs e)//Список принятых
        {
            List_of_accepted List_of_accepted = new List_of_accepted();
            List_of_accepted.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            List_of_accepted.Show();
        }
        private void toolStripButton1_Click(object sender, EventArgs e)//Поиск
        {
            Search Search = new Search(this.dataGridView1, this.dataGridView2);// передаем ссылку на грид в форму Editor
            Search.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            Search.Show();
        }
        private void dataGridView2_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)//Контекстное меню в гриде по нажатию мыши
        {
            ContextMenuStrip contextMenuStrip = new ContextMenuStrip();
            // Игнорировать нажатие на заголовок столбца или строки
            if (e.RowIndex != -1 && e.ColumnIndex != -1)
            {
                if (e.Button == MouseButtons.Right)
                {
                    DataGridViewCell clickedCell = (sender as DataGridView).Rows[e.RowIndex].Cells[e.ColumnIndex];

                    // Здесь вы можете делать с ячейкой все, что хотите
                    this.dataGridView2.CurrentCell = clickedCell;  // Выберите, например, ячейку, по которой щелкнули

                    // Получить положение мыши относительно сетки транспортных средств
                    var relativeMousePosition = dataGridView2.PointToClient(Cursor.Position);

                    // Показать контекстное меню
                    contextMenuStrip.Items.Add("Статус - F5").Click += new EventHandler(Status_Click);
                    contextMenuStrip.Items.Add("Редактирование - F4").Click += new EventHandler(Edit_Click);
                    contextMenuStrip.Items.Add("Удалить строку - Delete").Click += new EventHandler(Delete_Click);
                    contextMenuStrip.Show(dataGridView2, relativeMousePosition);
                }
            }
        }

        private void Edit_Click(object sender, EventArgs e)//Редактирование записей
        {
            Editor Editor = new Editor(this.dataGridView1, this.dataGridView2);// передаем ссылку на грид в форму Editor
            Editor.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            Editor.Show();
        }
        private void Status_Click(object sender, EventArgs e)//Статус
        {
            Status_Change Status_Change = new Status_Change(this.dataGridView2);// передаем ссылку на грид в форму Status_Change
            Status_Change.Owner = this;//Передаём ссылку на первую форму через свойство Owner //Вызов метода формы из другой формы
            Status_Change.Show();
        }
        private void Delete_Click(object sender, EventArgs e)//удаление строк из dataGridView1 и dataGridView3 и dataGridView2
        {
            // удаляем выделенные строки из dataGridView1
            foreach (DataGridViewRow row in dataGridView1.SelectedRows)
            {
                dataGridView1.Rows.Remove(row);
                label1.Text = ("Удалена строка");
                Podschet();//произвести подсчет по методу       
            }
            // удаляем выделенные строки из dataGridView2
            foreach (DataGridViewRow row in dataGridView2.SelectedRows)
            {
                dataGridView2.Rows.Remove(row);
                label1.Text = ("Удалена строка");
                Podschet();//произвести подсчет по методу       
            }
            // удаляем выделенные строки из dataGridView3
            foreach (DataGridViewRow row in dataGridView3.SelectedRows)
            {
                dataGridView3.Rows.Remove(row);
                label1.Text = ("Удалена строка");
                Podschet();//произвести подсчет по методу       
            }
            // удаляем выделенные строки из dataGridView5
            foreach (DataGridViewRow row in dataGridView5.SelectedRows)
            {
                dataGridView5.Rows.Remove(row);
                label1.Text = ("Удалена строка");
                Podschet();//произвести подсчет по методу       
            }
        }

        
    }
}




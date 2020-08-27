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
using System.Linq;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using ExcelDataReader;
using System.Data.Linq;

namespace ProgramCCS
{
    public partial class TLC : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели
        DataContext db = new DataContext(Connection.con);//Для работы LINQ to SQL
        //public SqlConnection con = new SqlConnection(@"Data Source=192.168.0.3;Initial Catalog=ccsbase;Persist Security Info=True;User ID=Lan;Password=Samsung0");
        MySqlConnection mycon = new MySqlConnection("SERVER= хостинг_сервер;" + "DATABASE= имя_базы;" + "UID= логин;" + "PASSWORD=пароль;" + "connection timeout = 180");

        public DataTable dtTarif = new DataTable();//создаем экземпляр класса DataTable
        private string fileName = string.Empty;
        
        

        int[] massiv1 = { 723504, 724508, 720114, 725000, 721100, 723500, 720306, 723500, 723100 };
        int[] massiv2 = { 724002, 723509, 725000, 722200, 723330, 723307, 723500, 723503, 723507, 721100 };
        int[] massiv3 = { 721901, 723504, 720016, 720300, 722300, 721200, 720800, 720300, 724321, 723510, 724604, 723510, 723503, 722200, 724104, 723800, 724913, 720601, 723509, 720803 };
        int[] massiv4 = { 724002, 723509, 725000, 722200, 723330, 723307, 723500, 723503, 723507, 721100 };
        int[] massiv5 = { 724002, 723509, 725000, 722200, 723330, 723307, 723500, 723503, 723507, 721100 };
        int[] massiv6 = { /*кызылкия*/720300, 720302, 720303, 720304, 720305, };
        int[] massiv7 = { 724002, 723509, 725000, 722200, 723330, 723307, 723500, 723503, 723507, 721100 };
        string[] spisok = { "Кок - Ой", "Байзак", "Жумгал", "Кызарт", "Куйругук", "Багышан", "Туголь - Сай", "Баш - Куванды", "Баетово", "Терек", "Бай - Гончок", "Улут", "Ак - Тал", "ТоголокМолдо", "Куртка", "Ак-Жар", "Ага Каинды", "Баш Каинды", "Бирлик", "Ак Муз", "Калинин",
            "Кара-Суу","Байгазак", "Орто -Саз", "Чет-Нура", "Орто-Нура", "Ийрисуу", "Алыш", "Доболуу", "Каинды", "Мин-Булак", "Ички-Башы", "Оттук", "Жергетал", "Эмгекчи", "Жан-Булак", "Достук", "Куланак", "Учкун", "Пионер", "Арсы", "Кара - Тоо", "Ак - Жар",
            "Кум - Добо", "Семиз - Бель", "Ормон Хан", "Кара - Суу", "Исакеев", "Кок - Жар", "Чекилдек", "Туз", "Чолпон", "Мантыш", "Дон - Алыш", "Советский", "Бешик-Жон", "Кочкор-Ата", "Момбеково", "Кыпчак-Талас", "Кок-Таш", "Кызыл-Туу", "Аксы", "Шамалды-Сай", "Кош-Тобо",
            "Шарк", "Нариман", "Кашгар-Кыштак", "Отуз-Адыр", "Жаны-Арык", "Куршаб", "Шералы", "Ильичевка", "Мырза-Аки", "Жекерчи", "Ылай-Талаа", "Кара-Кочкор", "1-Май", "Тоготой", "Жар-Кыштак", "Кыр-Кол", "КурманжанДатка", "Жылуу-Суу", "Россия", "Бель-Орук", "Кок-Жар", "Кенеш",
            "Уч-Коргон", "Марказ", "Кок-Талаа", "Халмион", "Торговый", "Бурганды", "Уч-Коргон", "Марказ", "Кок-Талаа", "Халмион", "Торговый", "Бурганды", "Орозбеково", "Кызыл-Булак", "Караван", "Чон-Гара", "Жаны-Жер", "Ак-Таш", "Ак-Татыр", "Бужум", "Кызыл-Жол", "Чек",
            "Ак-Терек", "Коргон","Тогуз-Булак", "Кара-Булак", "Чимген", "Ново-Павловка", "Военно-Антоновка", "Гавриловка", "Романовка", "Шопоков", "Александровка", "Садовое", "Петровка", "Полтавка", "Ново-Николаевка", "Петропавловка", "Калининское", "Алексеевка", "Вознесеновка",
            "Лебединовка", "Ново-Покровка", "Киршелк", "Люксембург", "Дмитриевка", "Буденовка", "Кенеш", "Красная Речка", "Ивановка", "Кенбулун", "Гидростроитель", "Арал", "Искра", "Чемкургон", "Бообек", "Жаналыш", "Акбекет", "Каскелен",
            "Сары-Ой", "Кара-Ой", "Чолпон-Ата", "Бакту-Долонотуу", "Бозтери", "Кен-Арал", "Озгорут", "Ак-Добо", "Кызыл-Сай", "Мин-Булак", "Боо-Терек", "Бакыян", "Тамчы-Булак", "Бейшеке", "Кичи-Кировка", "Кировка, Жийде", "Пушкин", "Кок-Токой", "Жон-Арык", "Кок-Ой" };

        Login formLogin = new Login();
        public object loker = new object();
      
        public TLC()
        {
            InitializeComponent();
            Text += "  Версия - " + CurrentVersion; //Добавляем в название программы, версию.
            //comboBox8.Text = Properties.Settings.Default.Prichina_vozvrat; // Загружаем ранее сохраненный текст
            //Properties.Settings.Default.Save();  // Сохраняем переменные.
            comboBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button10_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            comboBox2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button10_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            comboBox8.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button10_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox8.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button10_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox16.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button10_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox18.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button10_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox19.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button10_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button28_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dataGridView2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) button4_Click(new object(), new EventArgs()); };//Нажатие кнопки "Удалить строку" с клавиатуры
            dataGridView1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) button4_Click(new object(), new EventArgs()); };//Нажатие кнопки "Удалить строку" с клавиатуры
            dataGridView5.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) button4_Click(new object(), new EventArgs()); };//Нажатие кнопки "Удалить строку" с клавиатуры
            dataGridView3.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) button4_Click(new object(), new EventArgs()); };//Нажатие кнопки "Удалить строку" с клавиатуры
        }
        private void Button22_Click(object sender, EventArgs e)//Запись и чтение из файла
        {
            //Properties.Settings.Default.Prichina_vozvrat = comboBox8.Text; // Записываем содержимое comboBox8 в Prichina_vozvrat
            //Properties.Settings.Default.Save(); // Сохраняем переменные.
            //MessageBox.Show("Текст сохранен", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            //Запись в файл потоком
            using (var sw = new StreamWriter("Prichina_vozvrat.txt", true, Encoding.UTF8))
            {
                sw.WriteLine(comboBox8.Text);
                MessageBox.Show("Причина сохранена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            comboBox8.Items.Clear();//очистим перед чтением чтобы текст не дублировался
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
            //очищаем редактируемое поле
            comboBox8.Text = string.Empty;
        }
        private void button29_Click(object sender, EventArgs e)//Удалить строку, чтение из файла
        {
            string[] lines = File.ReadAllLines("Prichina_vozvrat.txt");
            using (var sw = new StreamWriter("Prichina_vozvrat.txt"))
            {
                foreach (var line in lines.Where(x => x != comboBox8.Text))
                    sw.WriteLine(line);
                MessageBox.Show("Причина Удалена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

            comboBox8.Items.Clear();//очистим перед чтением чтобы текст не дублировался
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
            //очищаем редактируемое поле
            comboBox8.Text = string.Empty;
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
  
            comboBox4.Text = Person.Name;
            label_filial.Text = Person.Name;
            //----------------------------------------//***
            if (Person.Access == "Low")
            {
                tabPage4.Enabled = false;
                button1.Enabled = false;
                button2.Enabled = false;
                button14.Enabled = false;
                button6.Enabled = false;
                toolStripButton3.Enabled = false;
            }
            else if (Person.Access == "Medium")
            {
                tabPage4.Enabled = false;
                button14.Enabled = false;
            }
            else if (Person.Access == "root")
            {
                tabPage4.Enabled = true;
                button14.Enabled = true;
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
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView1.DefaultCellStyle.Font = new Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;//цвет заголовка
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            DataGridViewRow row2 = this.dataGridView2.RowTemplate;
            row2.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row2.Height = 5;
            row2.MinimumHeight = 17;
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView2.DefaultCellStyle.Font = new Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSlateGray;//цвет заголовка
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            DataGridViewRow row5 = this.dataGridView5.RowTemplate;
            row5.DefaultCellStyle.BackColor = Color.LightSkyBlue;
            row5.Height = 5;
            row5.MinimumHeight = 17;
            dataGridView5.EnableHeadersVisualStyles = false;
            dataGridView5.ColumnHeadersDefaultCellStyle.Font = new Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView5.DefaultCellStyle.Font = new Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView5.ColumnHeadersDefaultCellStyle.BackColor = Color.LightCoral;//цвет заголовка
            dataGridView5.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            DataGridViewRow row6 = this.dataGridView6.RowTemplate;
            row6.DefaultCellStyle.BackColor = Color.LightSkyBlue;
            row6.Height = 5;
            row6.MinimumHeight = 17;
            dataGridView6.EnableHeadersVisualStyles = false;
            dataGridView6.ColumnHeadersDefaultCellStyle.Font = new Font("Verdana", 10F, FontStyle.Bold, GraphicsUnit.Pixel);//Шрифт заголовка
            dataGridView6.DefaultCellStyle.Font = new Font("Tahoma", 11F, GraphicsUnit.Pixel);//Шрифт строк
            dataGridView6.ColumnHeadersDefaultCellStyle.BackColor = Color.LightCoral;//цвет заголовка
            dataGridView6.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            //----------------Окраска Гридов--------------------//
            dataGridView2.Visible = true;
            dataGridView6.Visible = false;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;
            button12.Enabled = false;
            button14.Enabled = false;
            
            label26.Text = "Версия - " + CurrentVersion;

            toolTip1.SetToolTip(checkBox1, "Установите галочку если хотите сделать Выборку по дате обработки");
            toolTip1.SetToolTip(comboBox2, "Область");
            toolTip1.SetToolTip(comboBox5, "Контрагент");
            toolTip1.SetToolTip(comboBox4, "Филиал");
            toolTip1.SetToolTip(comboBox1, "Статус");
            toolTip1.SetToolTip(comboBox8, "Укажите причину возврата");
            toolTip1.SetToolTip(textBox8, "Присвойте тариф");
            toolTip1.SetToolTip(textBox3, "Поиск по №Заказа");
            toolTip1.SetToolTip(textBox14, "(Введите номер реестра) или (номер списка принятых)");
            toolTip1.SetToolTip(button10, "Присвоить или изменить (Статус, Область, Стоимость, Тариф или Доплата)");
            toolTip1.SetToolTip(button4, "Удалить строку");
            toolTip1.SetToolTip(button6, "Удалить запись из базы");

            dateTimePicker1.Value = DateTime.Today.AddDays(0);
            dateTimePicker2.Value = DateTime.Today.AddDays(0);
            dateTimePicker5.Value = DateTime.Today;
            //dataGridView2.Columns[7].DefaultCellStyle.Format = "dd.MM.yyyy";

            // инициализация         
            comboBox2.Items.Add(new ClassComboBoxOblast("Чу", "Чуйская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ош", "Ошская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Та", "Таласская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Жал", "Джалал - Абадская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Батк", "Баткенская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("Ис", "Иссык - Кульская область"));
            comboBox2.Items.Add(new ClassComboBoxOblast("На", "Нарынская область"));
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
            Environment.NewLine + "10. Выборка за период-5 (Дата Обработки) - Выбираем 'Период + Клиент и поставте галочку Дата Обработки'." +
            Environment.NewLine + "11. Выборка за период-6 (Дата записи) - Выбираем 'Период + Клиент'." +
            Environment.NewLine +
            Environment.NewLine + "Список принятых --- 1.Выбрать контрагент и рядом поставить №_номер 2.Выбрать контрагент(выдаст последний список) 3.Установить период и выбрать контрагент" +
            Environment.NewLine +
            Environment.NewLine +
            Environment.NewLine + "Каждый филиал видит только свои записи в базе!";

            //button2.Enabled = false;
            Disp_data();
            Podschet();//произвести подсчет по методу       
            comboBox4.SelectedIndex = -1;
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
            var statusGroup = from table in db.GetTable<Table_1>()
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
            button14.Enabled = false;
        }
        private void LinkLabel2_Click(object sender, EventArgs e)//Отобразить список Ожидание!
        {
            dataGridView2.Visible = true;

            var command = from table in db.GetTable<Table_1>()
                          where table.Статус == "Ожидание"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;

            linkLabel2.Visible = false;
            Podschet();
            button14.Enabled = true;
        }
        private void LinkLabel3_Click(object sender, EventArgs e)//Отобразить список Розыск!
        {
            dataGridView2.Visible = true;

            var command = from table in db.GetTable<Table_1>()
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

            var command = from table in db.GetTable<Table_1>()
                          where table.Статус == "Замена"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;

            linkLabel4.Visible = false;
            Podschet();
        }

        public void Wait()
        {
            //Отобразить список Ожидание! 
            var command = from table in db.GetTable<Table_1>()
                          where table.Статус == "Ожидание"
                          orderby table.Дата_записи descending
                          select table;
            dataGridView2.DataSource = command;
            Tarifs();//Т а р и ф ы
            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка command
        }

        public void SelectData()//Группировка и Сортировка по дате записи (сначала новые)
        {
            Wanted_Pending_Replacement();//Розыск, Ожидание, Замена (Группировка)
            if (Person.Name == "root")
            {
                //Группировка по Филиалу (находим последнюю запись) сортируем по дате
                var maxDate = from table in db.GetTable<Table_1>()
                              group table by table.Филиал into g
                              select g.OrderByDescending(t => t.Дата_записи).FirstOrDefault();
                dataGridView2.DataSource = maxDate;
                db.Refresh(RefreshMode.OverwriteCurrentValues, maxDate); //datacontext очистка 
                //последние записи по Дате
                var lastDays = from table in db.GetTable<Table_1>()
                              where table.Дата_записи >= Convert.ToDateTime(dataGridView2.Rows[0].Cells[12].Value)
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = lastDays;
                db.Refresh(RefreshMode.OverwriteCurrentValues, lastDays); //datacontext очистка 
                label1.Text = ("Отображены последние записи по всем филиалам");
            }
            else
            {
                var sevenDays = from table in db.GetTable<Table_1>()
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
                var maxDate = from table in db.GetTable<Table_1>()
                              group table by table.Филиал into g
                              select g.OrderByDescending(t => t.Дата_записи).FirstOrDefault();
                dataGridView2.DataSource = maxDate;
                db.Refresh(RefreshMode.OverwriteCurrentValues, maxDate); //datacontext очистка 
                //последние записи по Дате
                var lastDays = from table in db.GetTable<Table_1>()
                               where table.Дата_записи >= DateTime.Now.AddMonths(-8)
                               where table.Филиал == Person.Name
                               orderby table.Дата_записи descending
                               select table;
                dataGridView2.DataSource = lastDays;
                db.Refresh(RefreshMode.OverwriteCurrentValues, lastDays); //datacontext очистка command
                label1.Text = ("Отображены последние записи");
            }
        }
        public void Disp_data()//Отображает базу
        {
            button8.Text = "Ожидайте!";
            button8.Enabled = false;
            button2.Enabled = true;
            dataGridView2.Visible = true;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;

            SelectData(); //Группировка и Сортировка по дате записи (сначала новые) //Розыск, Ожидание, Замена (Группировка)            
            button12.Enabled = false;
            button8.Text = "Обновить";
            button8.Enabled = true;
            comboBox4.SelectedIndex = -1;

            //--------------------------Погода и курс валют-------------------//
            pictureBox4.ImageLocation = "http://www.informer.kg/cur/pngs/informer11.png";
            pictureBox4.Width = 120;
            pictureBox4.Height = 160;
            //--------------------------Погода и курс валют-------------------//       
        }
        public void Disp_data_all_base()//Отображает всю базу и сортирует по дате записи
        {
            button9.Text = "Ожидайте!";
            button9.Enabled = false;
            button2.Enabled = true;
            dataGridView2.Visible = true;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;
            ProgressBar();

            if (Person.Name == "root")
            {
                var command = from table in db.GetTable<Table_1>()
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            else
            {
                var command = from table in db.GetTable<Table_1>()
                              where table.Филиал == Person.Name
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }

            label1.Text = ("База данных отображена");
            button12.Enabled = false;
            button9.Text = "Вся база";
            button9.Enabled = true;
            comboBox4.SelectedIndex = -1;
            Podschet();//произвести подсчет по методу         
        }
        public void Mydisp_data()
        {
            mycon.Open();
            MySqlCommand mycmd = mycon.CreateCommand();
            mycmd.CommandType = CommandType.Text;
            mycmd.CommandText = "SELECT * FROM [Table_1]";
            mycmd.ExecuteNonQuery();

            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter();//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение

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
                textBox5.Visible = true;
                textBox5.Text = summa.ToString() + " Сом";
                //Сумма столбца плата за услугу
                double summa_U = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_U += incom;
                }
                textBox15.Visible = true;
                textBox15.Text = summa_U.ToString() + " Сом";
                //Сумма столбца плата за возврат
                double summa_V = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[14].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_V += incom;
                }
                textBox21.Visible = true;
                textBox21.Text = summa_V.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView1.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView1.ColumnCount; i++)
                    {
                        if (dataGridView1[i, j].Value != null)
                        {
                            textBox4.Text = Convert.ToString(dataGridView1.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
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
                textBox5.Visible = true;
                textBox5.Text = summa.ToString() + " Сом";
                //Сумма столбца плата за услугу
                double summa_U = 0;
                foreach (DataGridViewRow row in dataGridView5.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[8].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_U += incom;
                }
                textBox15.Visible = true;
                textBox15.Text = summa_U.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView5.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView5.ColumnCount; i++)
                    {
                        if (dataGridView5[i, j].Value != null)
                        {
                            textBox4.Text = Convert.ToString(dataGridView5.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
                            count++;
                            break;
                        }
                    }
                }
            }
            else if (dataGridView2.Visible == true)
            {
                //Сумма столбца
                double summa = 0;
                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox5.Visible = true;
                textBox5.Text = summa.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView2.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (dataGridView2[i, j].Value != null)
                        {
                            textBox4.Text = Convert.ToString(dataGridView2.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
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
                //Сумма столбца плата за услугу
                double summa_U = 0;
                foreach (DataGridViewRow row in dataGridView6.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa_U += incom;
                }
                textBox23.Text = summa_U.ToString() + " Сом";
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView6.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView6.ColumnCount; i++)
                    {
                        if (dataGridView6[i, j].Value != null)
                        {
                            textBox24.Text = Convert.ToString(dataGridView6.Rows.Count/*-1*/) + " Штук";// -1 это нижняя пустая строка
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
                    if (Convert.ToString(dataGridView2.Rows[i].Cells[24].Value) == tarifs[y])//Т а р и ф для большинства организаций
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
                if (Convert.ToString(dataGridView2.Rows[i].Cells[24].Value) == "по 1 проценту")//Т а р и ф для ИП 'JUMPER'
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
                if (Convert.ToString(dataGridView2.Rows[i].Cells[24].Value) == "по 2 процента")//Т а р и ф для "ОсОО 'Экспресс-Тайм'"
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
                if (Convert.ToString(dataGridView2.Rows[i].Cells[24].Value) == "по 0 процентов")//Т а р и ф для "ОсОО 'Альфа Вита'"
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
                if (Convert.ToString(dataGridView2.Rows[i].Cells[24].Value) == "1,5 процента")//Т а р и ф для ОсОО Kyrgyz Express Post
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
                if (dataGridView2.Rows[i].Cells[24].Value.ToString() == "Сложный")
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
            if (dataGridView2.Rows.Count != 0 & dataGridView2.Rows.Count == 1)
            {
                var command = from table in db.GetTable<Table_1>()
                              where table.N_Заказа == Convert.ToString(textBox3.Text)
                              select table;
                dataGridView2.DataSource = command;
                db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка 

                int doplata = Convert.ToInt32(dataGridView2.Rows[0].Cells[7].Value);
                int tarif = Convert.ToInt32(dataGridView2.Rows[0].Cells[6].Value);
                double ob_cennost = Convert.ToInt32(dataGridView2.Rows[0].Cells[8].Value);
                double plata_za_nalog = Convert.ToInt32(dataGridView2.Rows[0].Cells[9].Value);
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET plata_za_uslugu = @plata_za_uslugu WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@plata_za_uslugu", Math.Round(ob_cennost + tarif + plata_za_nalog + doplata));//Math.Round округляет до целого
                cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[0].Cells[0].Value);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
            }

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
            if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Выдано")
            {
                status = "Выдано";
            }
            else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Возврат")
            {
                status = "Возврат";
            }
            else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Розыск")
            {
                status = "Розыск";
            }
            else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Замена")
            {
                status = "Замена";
            }
            else MessageBox.Show("Select_status_Nr", "Ошибка!");

            var command = from table in db.GetTable<Table_1>()
                          where table.Статус == status
                          group table by table.Nr into g
                          select g.OrderByDescending(t => t.Nr).FirstOrDefault();
            dataGridView2.DataSource = command;

            Number.Nr = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
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
            if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Ожидание")
            {
                status = "Отправлено";//чтобы простовлять порядковый номер /не менять (все верно) могу забыть!
            }
            else if (Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Отправлено")
            {
                status = "Отправлено";
            }
            else MessageBox.Show("Select_status_Nn", "Ошибка!");

            var command = from table in db.GetTable<Table_1>()
                          where table.Статус == status
                          group table by table.Nn into g
                          select g.OrderByDescending(t => t.Nn).FirstOrDefault();
            dataGridView2.DataSource = command;

            Number.Nn = Convert.ToInt32(dataGridView2.Rows[0].Cells[22].Value) + 1;
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
            dataGridView2.DataSource = command;

            Number.Ns = Convert.ToInt32(dataGridView2.Rows[0].Cells[21].Value) + 1;
            Number.Prefix_number = comboBox10.Text + Number.Ns;

            db.Refresh(RefreshMode.OverwriteCurrentValues, command); //datacontext очистка 
        }
        public void Select_client()//Для сортировки принятых списков по клиенту
        {
            var command = from table in db.GetTable<Table_1>()
                          where table.Контрагент == comboBox5.Text
                          group table by table.Ns into g
                          select g.OrderByDescending(t => t.Ns).FirstOrDefault();
            dataGridView2.DataSource = command;

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
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow column in dt.Rows)
            {
                comboBox6.Items.Add(column[0].ToString());
                comboBox3.Items.Add(column[0].ToString());
                comboBox5.Items.Add(column[0].ToString());
            }
            con.Close();//Закрываем соединение          
        }
        public void ComboBox5_TextChanged(object sender, EventArgs e)//поиск тарифа по контрагенту
        {
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT tarif FROM [Table_Partner]" +
                "WHERE name = @name", con);
            cmd.Parameters.AddWithValue("@name", comboBox5.Text.ToString());
            cmd.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dtTarif.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dtTarif);//заполняем данными созданный DataTable
            con.Close();//закрыть соединение
            if (comboBox5.Text == "")//если поле очищено, отобразить базу
            {
                dtTarif.Clear();//чистим DataTable, если он был не пуст
                foreach (DataRow column in dtTarif.Rows)
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
                                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                                cmd.Parameters.AddWithValue("@nomer_reestra", 0);
                                cmd.Parameters.AddWithValue("@nomer_spiska", Number.Prefix_number);
                                cmd.Parameters.AddWithValue("@nomer_nakladnoy", 0);
                                cmd.Parameters.AddWithValue("@Nr", 0);
                                cmd.Parameters.AddWithValue("@Ns", Number.Ns);
                                cmd.Parameters.AddWithValue("@Nn", 0);
                                cmd.Parameters.AddWithValue("@tarifs", dtTarif.Rows[0][0].ToString());//tarif
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
                            var command = from table in db.GetTable<Table_1>()
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
                            cmd.Parameters.AddWithValue("nomer_spiska", dataGridView2.Rows[0].Cells[18].Value.ToString());
                            cmd.Parameters.AddWithValue("@client", comboBox5.Text);
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
                            button2.Text = "Ожидайте!";
                            SaveFileDialog sfd = new SaveFileDialog();
                            sfd.Filter = "Word Documents (*.docx)|*.docx";
                            sfd.FileName = $"Список принятых № {dataGridView2.Rows[0].Cells[18].Value.ToString()}.docx";
                            if (sfd.ShowDialog() == DialogResult.OK)
                            {
                                Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                            }
                            button2.Text = "Список принятых";
                        }
                    }
                    else if (comboBox5.Text == "")
                    {
                        MessageBox.Show("Необходимо выбрать клиента", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        private void button7_Click(object sender, EventArgs e)//Выборка
        {
            //1.Выборка на реестр-1 - 'Статус + Обработка + Филиал + Клиент'.
            if (checkBox1.Checked == false & checkBox2.Checked == false && comboBox1.Text != "" & comboBox4.Text != "" & comboBox2.Text == "" & textBox14.Text == "")
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
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
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
                label1.Text = "1.Выборка на реестр-1 - 'Статус + Обработка + Филиал + Клиент'.";
            }
            //2.Выборка на реестр-2 - 'Статус + Обработка + Клиент'.
            else if (checkBox1.Checked == false & checkBox2.Checked == false && comboBox1.Text != "" & comboBox5.Text != "" & comboBox2.Text == "" & comboBox4.Text == "" & textBox14.Text == "")
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
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
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
                label1.Text = "2.Выборка на реестр-2 - 'Статус + Обработка + Клиент'.";
            }
            //3. Выборка за период-1 (Дата обработки) - 'Статус + Период + Область + Клиент'.
            else if (comboBox1.Text != "" & comboBox2.Text != "" & comboBox5.Text != "" & comboBox4.Text == "" & checkBox1.Checked & textBox14.Text == "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_obrabotki BETWEEN @StartDate AND @EndDate AND oblast LIKE '%" + comboitem.ToString() + "%' AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение   
                label1.Text = "3. Выборка за период-1 (Дата обработки) - 'Статус + Период + Область + Клиент'.";
            }
            //4. Выборка за период-2 (Дата записи) - 'Статус + Период + Область + Филиал + Клиент'.
            else if (comboBox1.Text != "" & comboBox2.Text != "" & comboBox4.Text != "" & textBox14.Text == "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_zapisi BETWEEN @StartDate AND @EndDate AND oblast LIKE '%" + comboitem.ToString() + "%'" +
                    " AND filial = @filial AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@filial", comboBox4.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение   
                label1.Text = "4. Выборка за период-2 (Дата записи) - 'Статус + Период + Область + Филиал + Клиент'.";
            }
            //5. Выборка за период-3 (Дата записи) - 'Статус + Период + Область + Клиент +- Пункт'.
            else if (comboBox1.Text != "" & comboBox2.Text != "" & comboBox5.Text != "" & comboBox4.Text == "" & textBox14.Text == "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_zapisi BETWEEN @StartDate AND @EndDate AND oblast LIKE '%" + comboitem.ToString() + "%'" +
                    " AND client = @client AND punkt LIKE '%" + Convert.ToString(textBox18.Text) + "%' ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение      
                label1.Text = "5. Выборка за период-3 (Дата записи) - 'Статус + Период + Область + Клиент +- Пункт'.";
            }
            //6. Выборка за период-4 (Дата Обработки) - 'Статус + Период + Клиент'.
            else if (checkBox1.Checked && comboBox1.Text != "" & comboBox2.Text == "" & comboBox5.Text != "" & textBox14.Text == "")
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_obrabotki BETWEEN @StartDate AND @EndDate AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение     
                label1.Text = "6. Выборка за период-4 (Дата Обработки) - 'Статус + Период + Клиент'.";
            }
            //7. Выборка за период-5 (Дата записи) - 'Статус + Период + Клиент'.
            else if (checkBox2.Checked && comboBox1.Text != "" & comboBox5.Text != "" & comboBox2.Text == "" & textBox14.Text == "")
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE status = @status AND data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение     
                label1.Text = "7. Выборка за период-5 (Дата Записи) - 'Статус + Период + Клиент'.";
            }
            //8. Выборка на накладную - 'Дата + Область + Клиент +- Пункт'.
            else if (checkBox1.Checked == false & checkBox2.Checked == false && comboBox1.Text == "" & comboBox2.Text != "" & textBox14.Text == "")
            {
                string comboitem = ((ClassComboBoxOblast)comboBox2.SelectedItem).Value;
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE oblast LIKE '%" + comboitem.ToString() + "%' AND data_zapisi = @data_zapisi AND client = @client" +
                    " AND punkt LIKE '%" + Convert.ToString(textBox18.Text) + "%' ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@data_zapisi", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
                label1.Text = "8. Выборка на накладную - 'Дата + Область + Клиент +- Пункт'.";
            }
            //9. Выборка поиск Реестра
            else if (checkBox1.Checked == false & checkBox2.Checked == false && textBox14.Text != "" & comboBox5.Text != "" & comboBox1.Text != "" & comboBox2.Text == "" & comboBox4.Text == "")
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE nomer_reestra = @nomer_reestra AND client = @client AND status = @status ORDER BY N_zakaza", con);
                cmd.Parameters.AddWithValue("@nomer_reestra", textBox14.Text);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
                label1.Text = "9. Выборка поиск Реестра";
            }
            //10. Выборка за период-5 (Дата Обработки) - 'Период + Клиент'.
            else if (comboBox5.Text != "" & checkBox1.Checked)
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                button2.Enabled = false;
                DateTime date = new DateTime();
                date = dateTimePicker1.Value;
                DateTime date2 = dateTimePicker2.Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE (data_obrabotki BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                cmd.Parameters.AddWithValue("StartDate", date2);
                cmd.Parameters.AddWithValue("EndDate", date);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
                label1.Text = "10. Выборка за период-5 (Дата Обработки) - 'Период + Клиент'.";
            }
            //11. Выборка за период-6 (Дата записи) - 'Период + Клиент'.
            else if (comboBox5.Text != "")
            {
                dataGridView1.Visible = true;
                dataGridView2.Visible = false;
                button12.Enabled = true;
                DateTime date = new DateTime();
                date = dateTimePicker1.Value;
                DateTime date2 = dateTimePicker2.Value;
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT familia AS 'Ф.И.О', punkt AS 'Населенный пункт', N_zakaza AS '№Заказа', summ AS 'Стоимость', data_zapisi AS 'Дата записи', status AS 'Статус'," +
                    " prichina AS 'Причина', plata_za_uslugu AS 'Плата за услугу', client AS 'Контрагент', oblast AS 'Область', obrabotka AS 'Обработка', id AS ID, nomer_reestra AS 'Реестр'," +
                    " plata_za_nalog AS 'Наложеный платеж', (plata_za_uslugu - plata_za_nalog) AS 'Плата за возврат' FROM [Table_1]" +
                    " WHERE (data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                cmd.Parameters.AddWithValue("StartDate", date2);
                cmd.Parameters.AddWithValue("EndDate", date);
                cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение    
                label1.Text = "11. Выборка за период-6 (Дата записи) - 'Период + Клиент'.";
            }
            else
            {
                label1.Text = "Выборка!";
            }
            Podschet();//произвести подсчет из метода
            textBox2.Text = "";//очистка текстовых полей 
            //textBox1.Text = "";
            textBox3.Text = "";
            textBox14.Text = "";
            textBox18.Text = "";
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            if (dataGridView1.Rows.Count <= 0)
            {
                button12.Enabled = false;
            }
        }
        private void TextBox3_TextChanged(object sender, EventArgs e)//Поиск по №Заказа
        {
            dataGridView2.Visible = true;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;

            var command = from table in db.GetTable<Table_1>()
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
        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            if (textBox2.Text == "")//если поле очищено, отобразить базу
            {
                Disp_data();
            }
        }
        private void button28_Click(object sender, EventArgs e)//Поиск по Ф.И.О
        {
            if (textBox2.Text != "")
            {
                dataGridView2.Visible = true;
                dataGridView1.Visible = false;
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
                dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
                Podschet();//произвести подсчет по методу
                //table1BindingSource.Filter = "[punkt] LIKE '%" + Convert.ToString(textBox2.Text) + "%' OR [familia] LIKE '%" + Convert.ToString(textBox2.Text) + "%'"; //Фильтр по гриду
            }
        }
        private void button10_Click(object sender, EventArgs e)//Статус, Область, Тариф или Доплату (изминение)
        {
            if (comboBox1.Text != "" & textBox8.Text == "" & dataGridView2.Rows.Count == 1)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET status = @status, prichina = @prichina, filial = @filial WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                if (Convert.ToString(dataGridView2.Rows[0].Cells[11].Value) == "Ожидание" ||
                    Convert.ToString(dataGridView2.Rows[0].Cells[11].Value) == "Отправлено" ||
                    Convert.ToString(dataGridView2.Rows[0].Cells[11].Value) == "Розыск" ||
                    Convert.ToString(dataGridView2.Rows[0].Cells[11].Value) == "Замена" ||
                    comboBox1.Text == "Розыск" || comboBox1.Text == "Замена")
                {
                    cmd.Parameters.AddWithValue("@status", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@prichina", comboBox8.Text);
                    cmd.Parameters.AddWithValue("@filial", Person.Name);
                    MessageBox.Show("Статус успешно обновлен", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    label1.Text = "Статус успешно обновлен";
                    cmd.ExecuteNonQuery();
                }
                else MessageBox.Show("Изминение статуса невозможно, так как статус уже присвоен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                con.Close();//закрыть соединение             
                textBox3.Select();//Установка курсора
            }
            else if (textBox8.Text != "" & comboBox1.Text == "" & dataGridView2.Rows.Count == 1)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET tarif = @tarif WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                cmd.Parameters.AddWithValue("@tarif", textBox8.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Тариф успешно обновлен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label1.Text = "Тариф успешно обновлен!";
                textBox3.Select();//Установка курсора
            }
            else if (textBox16.Text != "" & textBox8.Text == "" & comboBox1.Text == "" & dataGridView2.Rows.Count == 1)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET doplata = @doplata WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                cmd.Parameters.AddWithValue("@doplata", textBox16.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Доплата успешно обновлена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label1.Text = "Доплата успешно обновлена!";
                textBox3.Select();//Установка курсора
            }
            else if (comboBox2.Text != "" & dataGridView2.Rows.Count == 1)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET oblast = @oblast WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                cmd.Parameters.AddWithValue("@oblast", comboBox2.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Область успешно изменена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label1.Text = "Область успешно изменена!";
                textBox3.Select();//Установка курсора
            }
            else if (textBox18.Text != "" & dataGridView2.Rows.Count == 1)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET punkt = @punkt WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                cmd.Parameters.AddWithValue("@punkt", textBox18.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Населенный пункт успешно изменен!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label1.Text = "Населенный пункт успешно изменен!";
                textBox3.Select();//Установка курсора
            }
            else if (textBox19.Text != "" & dataGridView2.Rows.Count == 1)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET summ = @summ WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                cmd.Parameters.AddWithValue("@summ", textBox19.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Стоимость успешно изменена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                label1.Text = "Стоимость успешно изменена!";
                textBox3.Select();//Установка курсора
            }
            else if (textBox8.Text == "" & textBox16.Text == "" & comboBox1.Text == "" & comboBox2.Text == "" & dataGridView2.Rows.Count == 1)
            {
                label1.Text = "Введите (сумму тарифа или доплаты или стоимость) - или - Выбирите статус - или - Выбирите область";
                MessageBox.Show("Введите (сумму тарифа или доплаты или стоимость) - или - Выбирите статус - или - Выбирите область", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (dataGridView2.Rows.Count != 1)
            {
                label1.Text = "Чтобы изменить (Статус, Область, Стоимость, Тариф или Доплату), Произведите поиск по №Заказа или по Фамилии";
                MessageBox.Show("Произведите поиск по №Заказа или по Фамилии", "Внимание! Чтобы изменить (Статус, Область, Стоимость, Тариф или Доплату)", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (dataGridView2.Rows.Count <= 0)
            {
                label1.Text = "В базе не найдено отправление";
                MessageBox.Show("В базе не найдено отправление", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.None);
            }

            Tarif_Update();//Заново ищет №Заказа и делает пересчет
            textBox3.Text = "";//очистка текстовых полей
            textBox8.Text = "";
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox16.Text = "";
            textBox19.Text = "";
            Disp_data();
        }
        private void button6_Click(object sender, EventArgs e)//Удалить из базы данных
        {
            if (textBox3.Text != "" & dataGridView2.Rows.Count == 1)
            {
                if (MessageBox.Show("Вы хотите удалить эту запись?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("DELETE FROM [Table_1] WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[0].Cells[0].Value));//первая строка в гриде
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                    MessageBox.Show("Запись успешно удалена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    label1.Text = "Запись успешно удалена!";
                    Disp_data();
                    textBox3.Select();//Установка курсора
                }
                else
                {
                    Disp_data();
                    textBox3.Select();//Установка курсора
                }
            }
            else if (dataGridView2.Rows.Count != 1)
            {
                label1.Text = "Чтобы удалить запись из базы данных, Произведите поиск по №Заказа или по Фамилии";
                MessageBox.Show("Произведите поиск по №Заказа или по Фамилии", "Внимание! Чтобы удалить запись из базы данных", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (dataGridView2.Rows.Count <= 0)
            {
                label1.Text = "В базе не найдено отправление";
                MessageBox.Show("В базе не найдено отправление", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            textBox3.Text = "";//очистка текстовых полей
            textBox8.Text = "";
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox16.Text = "";
            textBox19.Text = "";
            Disp_data();
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
        private void button14_Click(object sender, EventArgs e)//Объединить
        {
            if (Person.Name == "root" & textBox14.Text != "")
            {
                if (MessageBox.Show("Вы хотите объединить эти записи?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        con.Open();
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET nomer_spiska = @nomer_spiska WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@nomer_spiska", textBox14.Text);
                        cmd.Parameters.AddWithValue("@id", dataGridView2.Rows[i].Cells[0].Value);
                        cmd.ExecuteNonQuery();
                        con.Close();//закрыть соединение
                    }
                    MessageBox.Show("Объединение выполнено!", "Внимание!");
                }
                Disp_data();
            }
            else if (textBox14.Text == "")
            {
                MessageBox.Show("Введите номер списка в который хотите объединить!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else MessageBox.Show("Вы не администратор!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
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
                var command = from table in db.GetTable<Table_1>()
                              where table.Статус == "Возврат"
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            else
            {
                var command = from table in db.GetTable<Table_1>()
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
                var command = from table in db.GetTable<Table_1>()
                              where table.Статус == "Выдано"
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            else
            {
                var command = from table in db.GetTable<Table_1>()
                              where table.Филиал == Person.Name & table.Статус == "Выдано"
                              orderby table.Дата_записи descending
                              select table;
                dataGridView2.DataSource = command;
            }
            Podschet();
        }

        private void button4_Click(object sender, EventArgs e)// кнопка удаления строк из dataGridView1 и dataGridView3 и dataGridView2
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
        //private void dataGridView2_SelectionChanged(object sender, EventArgs e)//сумма выделенных строк и колл-во
        //{
        //    //Колл-во
        //    textBox4.Text = Convert.ToString(dataGridView2.SelectedRows.Count) + " Штук";
        //    //Сумма столбца стоимость
        //    double summa = 0;
        //    foreach (DataGridViewRow row in dataGridView2.SelectedRows)
        //    {
        //        double incom;
        //        double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom);
        //        summa += incom;
        //    }
        //    textBox5.Visible = true;
        //    textBox5.Text = summa.ToString() + " Сом";

        //}

        private void button12_Click(object sender, EventArgs e)//Реестр - Накладная и Обработка
        {
            //Обработка и Выдача реестра
            if (dataGridView1.Rows.Count > 0 & Convert.ToString(dataGridView1.Rows[0].Cells[10].Value) != "Обработано"
                & Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) != "Отправлено"
                & Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) != "Ожидание"
                & Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) != "Розыск"
                & Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) != "Замена")
            {
                Select_status_Nr();//Выборка по статусу и сортировка по номеру реестра от больших значений к меньшим.                                      
                if (MessageBox.Show("Вы хотите обработать эти записи?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    button12.Enabled = false;
                    button12.Text = "Ожидайте идет обработка!";
                    con.Open();//открыть соединение
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET obrabotka = @obrabotka, data_obrabotki = @data_obrabotki, nomer_reestra = @nomer_reestra, Nr=@Nr WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@obrabotka", "Обработано");
                        cmd.Parameters.AddWithValue("@data_obrabotki", DateTime.Today.AddDays(0));
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value));
                        cmd.Parameters.AddWithValue("@nomer_reestra", Number.Prefix_number);
                        cmd.Parameters.AddWithValue("@Nr", Number.Nr);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение 
                    label1.Text = ("Обработка выполнена - Присвоен № Реестра!");
                    MessageBox.Show("Обработка выполнена / Присвоен № Реестра!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //------Ручная вставка номера реестра и обработки----------//
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
                    {
                        dataGridView1.Rows[i].Cells[12].Value = Number.Prefix_number;
                        dataGridView1.Rows[i].Cells[10].Value = "Обработано";
                    }
                    //------Ручная вставка номера реестра и обработки----------//
                }
                //Выдача рееста в WORD
                string status = Convert.ToString(dataGridView1.Rows[0].Cells[5].Value);//Статус
                string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Контрагент                
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                    {
                        Export_Reestr_To_Word(dataGridView1, sfd.FileName);
                    }
                    else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                    {
                        Export_Reestr_To_Word_vozvrat(dataGridView1, sfd.FileName);
                    }
                }
                //Выдача рееста в EXCEL
                if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                {
                    sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                    sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Reestr_To_Excel(dataGridView1, sfd.FileName);
                    }
                }
                else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                {
                    sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                    sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Reestr_To_Excel_vozvrat(dataGridView1, sfd.FileName);
                    }
                }
            }
            else if (dataGridView1.Rows.Count > 0 & Convert.ToString(dataGridView1.Rows[0].Cells[10].Value) == "Обработано")
            {
                if (MessageBox.Show("Вы хотите открыть этот Реестр?", "Внимание! Эти данные уже обработаны!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    button12.Enabled = false;
                    button12.Text = "Ожидайте идет выгрузка!";
                    //Выдача рееста в WORD
                    string nomer = dataGridView1.Rows[0].Cells[12].Value.ToString();//№
                    string status = Convert.ToString(dataGridView1.Rows[0].Cells[5].Value);//Статус
                    string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Контрагент
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Реестр № {nomer} на {status}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                        {
                            Export_Reestr_To_Word(dataGridView1, sfd.FileName);
                        }
                        else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                        {
                            Export_Reestr_To_Word_vozvrat(dataGridView1, sfd.FileName);
                        }
                    }
                    //Выдача рееста в EXCEL
                    if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                    {
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Реестр № {nomer} на {status}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Reestr_To_Excel(dataGridView1, sfd.FileName);
                        }
                    }
                    else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                    {
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Реестр № {nomer} на {status}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Reestr_To_Excel_vozvrat(dataGridView1, sfd.FileName);
                        }
                    }
                }
            }
            else if (dataGridView1.Rows.Count > 0 & Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Розыск" || Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Замена")
            {
                if (MessageBox.Show("Вы хотите открыть этот Реестр?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    button12.Enabled = false;
                    button12.Text = "Ожидайте идет выгрузка!";
                    //Выдача рееста в WORD
                    string nomer = dataGridView1.Rows[0].Cells[12].Value.ToString();//№
                    string status = Convert.ToString(dataGridView1.Rows[0].Cells[5].Value);//Статус
                    string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Контрагент
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Реестр № {nomer} на {status}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Reestr_To_Word(dataGridView1, sfd.FileName);
                    }
                }
            }
            //Выдача накладной
            else if (dataGridView1.Rows.Count > 0 & Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Ожидание")
            {
                Select_status_Nn();//(Для выдачи накладных)Выборка по статусу и сортировка по номеру накладеой от больших значений к меньшим.               
                if (MessageBox.Show("Вы хотите получить 'Накладную'? Нажмите Нет если хотите получить 'Cписок за период'!", "Внимание! Статус изменится на 'Отправлено' и присвоется номер", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    button12.Enabled = false;
                    button12.Text = "Ожидайте идет выгрузка!";
                    con.Open();//открыть соединение
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET nomer_nakladnoy = @nomer_nakladnoy, status = @status, Nn=@Nn, filial=@filial WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[11].Value));
                        cmd.Parameters.AddWithValue("@status", "Отправлено");
                        cmd.Parameters.AddWithValue("@nomer_nakladnoy", Number.Prefix_number);
                        cmd.Parameters.AddWithValue("@Nn", Number.Nn);
                        cmd.Parameters.AddWithValue("@filial", Person.Name);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение 
                    label1.Text = ("Присвоен № Накладной!");

                    string oblast = Convert.ToString(dataGridView1.Rows[0].Cells[9].Value);//Область
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Накладная № {Number.Prefix_number} - {oblast}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Nakladnaya_To_Word(dataGridView1, sfd.FileName);
                    }
                }
                else//Список за период (Ожидание)
                {
                    button12.Enabled = false;
                    button12.Text = "Ожидайте идет выгрузка!";
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = "Список за период.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        Export_Spisok_To_Word(dataGridView1, sfd.FileName);
                    }
                }
            }//-------------------------------------------------------------------------------------------------------------//
             //Список за период (Отправлено)
            else if (dataGridView1.Rows.Count > 0 & Convert.ToString(dataGridView1.Rows[0].Cells[5].Value) == "Отправлено")
            {
                button12.Enabled = false;
                button12.Text = "Ожидайте идет выгрузка!";
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = "Список за период.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    Export_Spisok_To_Word(dataGridView1, sfd.FileName);
                }
            }
            else if (dataGridView1.Rows.Count <= 0)
            {
                MessageBox.Show("Выборка не дала результатов, невозможно сгенерировать реестр!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("Эти данные нельзя обработать", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            textBox14.Text = "";//Очистка поля №_
            textBox15.Text = "";
            textBox21.Text = "";
            button12.Text = "Реестр Накладная";
            Disp_data();
            dataGridView2.Visible = true;
            dataGridView1.Visible = false;
            dataGridView5.Visible = false;
        }
        private void button2_Click(object sender, EventArgs e)//Список принятых
        {
            if (comboBox5.Text != "")//только при выбранном контрагенте
            {
                dataGridView5.Visible = true;
                dataGridView1.Visible = false;
                dataGridView2.Visible = false;
                if (dateTimePicker2.Value <= DateTime.Today.AddDays(-1))//За диапазон
                {
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("SELECT oblast, punkt, familia, N_zakaza, data_zapisi, summ, tarif, doplata, plata_za_uslugu, ob_cennost, plata_za_nalog, id, nomer_spiska " +
                        "FROM [Table_1] WHERE (data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                    cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                    cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                    cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                    SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                    dt.Clear();//чистим DataTable, если он был не пуст
                    da.Fill(dt);//заполняем данными созданный DataTable
                    dataGridView5.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                    con.Close();//закрыть соединение    
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() != "0")
                    {
                        Podschet();//произвести подсчет по методу 
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        sfd.FileName = $"Список принятых с  {Convert.ToString(dateTimePicker2.Value.ToString("dd.MM.yyyy"))}  по  {Convert.ToString(dateTimePicker1.Value.ToString("dd.MM.yyyy"))}.docx";
                        button2.Text = "Ожидайте!";
                        //Выдача в WORD 
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                        }
                        //Выдача рееста в EXCEL
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Список принятых с  {Convert.ToString(dateTimePicker2.Value.ToString("dd.MM.yyyy"))}  по  {Convert.ToString(dateTimePicker1.Value.ToString("dd.MM.yyyy"))}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Excel(dataGridView5, sfd.FileName);
                        }
                        button2.Text = "Список принятых";
                    }
                    else if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() == "0")
                    {
                        MessageBox.Show("Вы не обработали один из списков!", "Внимание! Дальнейшие действия запрещены!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        MessageBox.Show("За этот период не найдено списков!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else if (dateTimePicker2.Value == DateTime.Today.AddDays(0))//За текущий день
                {
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    Select_client();//Для сортировки принятых списков по клиенту
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("SELECT oblast, punkt, familia, N_zakaza, data_zapisi, summ, tarif, doplata, plata_za_uslugu, ob_cennost, plata_za_nalog, id, nomer_spiska" +
                        " FROM [Table_1] WHERE (nomer_spiska = @nomer_spiska AND client = @client)", con);
                    if (textBox14.Text != "") { cmd.Parameters.AddWithValue("nomer_spiska", textBox14.Text); }//Ввести номер списка
                    else if (textBox14.Text == "") { cmd.Parameters.AddWithValue("nomer_spiska", dataGridView2.Rows[0].Cells[18].Value.ToString()); }//Если не ввести номер то выдаст последний
                    cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                    SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                    dt.Clear();//чистим DataTable, если он был не пуст
                    da.Fill(dt);//заполняем данными созданный DataTable
                    dataGridView5.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                    con.Close();//закрыть соединение
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    Podschet();//произвести подсчет по методу 
                    if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() == "0")
                    {
                        Select_Ns();//Выборка и сортировка по номеру от больших значений к меньшим.
                        con.Open();//открыть соединение
                        for (int i = 0; i < dataGridView5.Rows.Count; i++)//Цикл
                        {
                            SqlCommand cmd1 = new SqlCommand("UPDATE [Table_1] SET nomer_spiska = @nomer_spiska, Ns=@Ns WHERE id = @id", con);
                            cmd1.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView5.Rows[i].Cells[11].Value));
                            cmd1.Parameters.AddWithValue("@nomer_spiska", Number.Prefix_number);
                            cmd1.Parameters.AddWithValue("@Ns", Number.Ns);
                            cmd1.ExecuteNonQuery();
                        }
                        con.Close();//закрыть соединение 
                        label1.Text = ("Присвоен № Списка!");
                        //----------------------------------------//
                        button2.Text = "Ожидайте!";
                        //Выдача в WORD
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        sfd.FileName = $"Список принятых № {Number.Prefix_number}.docx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                        }
                        button2.Text = "Список принятых";
                    }
                    else if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() != "0")
                    {
                        //Выдача в WORD
                        button2.Text = "Ожидайте!";
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        if (textBox14.Text == "") { string number = dataGridView5.Rows[0].Cells[12].Value.ToString(); sfd.FileName = $"Список принятых № {number}.docx"; }//№}
                        else if (textBox14.Text != "") { string nomer = textBox14.Text; sfd.FileName = $"Список принятых № {nomer}.docx"; }
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                        }
                        button2.Text = "Список принятых";
                    }
                }
                else
                {
                    MessageBox.Show($"Список по контрагенту {comboBox5.Text} не найден", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                dataGridView2.Visible = true;
                dataGridView1.Visible = false;
                dataGridView5.Visible = false;
                Disp_data();
                textBox14.Text = "";//Очистка поля
                comboBox5.SelectedIndex = -1;
            }
            else
            {
                MessageBox.Show("Необходимо выбрать контрагента", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                string kol_vo = Convert.ToString(textBox4.Text);//кол-во
                string sum = Convert.ToString(textBox5.Text);//сумма
                string plata_za_usluguSumm = Convert.ToString(textBox15.Text);//Плата за услугу Сумма
                string status = Convert.ToString(dataGridView1.Rows[0].Cells[4].Value);//Статус
                string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[7].Value);//Контрагент
                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = $"Итого:    {kol_vo}                    {sum}" +
                //Environment.NewLine + " Сумма за услугу " + plata_za_usluguSumm +
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
                int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
                string prefix_number = comboBox10.Text + number;
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
                    {
                        string Reestr = prefix_number;
                        headerRange.Text = $"Реестр №  {Reestr}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {kontragent}" +
                        Environment.NewLine;
                    }
                    else if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) == "Обработано")
                    {
                        string Reestr = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        headerRange.Text = $"Реестр №  {Reestr}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {kontragent}" +
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

                string kol_vo = Convert.ToString(textBox4.Text);//кол-во
                string sum = Convert.ToString(textBox5.Text);//сумма
                string plata_za_usluguSumm = Convert.ToString(textBox15.Text);//Плата за услугу Сумма
                string status = Convert.ToString(dataGridView1.Rows[0].Cells[4].Value);//Статус
                string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[7].Value);//Контрагент
                string plata_za_vozvrat = Convert.ToString(textBox21.Text);//Плата за возврат сумма

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
                oDoc.Content.Text = $"Итого:    {kol_vo}                    {sum}" +
                Environment.NewLine + $"Сумма за возврат   {plata_za_vozvrat}" +
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
                int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
                string prefix_number = comboBox10.Text + number;
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
                    {
                        string Reestr = prefix_number;
                        headerRange.Text = $"Реестр №  {Reestr}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {kontragent}" +
                        Environment.NewLine;
                    }
                    else if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) == "Обработано")
                    {
                        string Reestr = dataGridView1.Rows[0].Cells[11].Value.ToString();
                        headerRange.Text = $"Реестр №  {Reestr}  на  {status}  от  {Convert.ToString(Now.ToString("dd.MM.yyyy"))} г. отправлений с наложенным платежом" +
                        Environment.NewLine + $"от  {kontragent}" +
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
            int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
            string prefix_number = comboBox10.Text + number;
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
            string kol_vo = Convert.ToString(textBox4.Text);//кол-во
            string sum = Convert.ToString(textBox5.Text);//сумма
            DateTime Now = DateTime.Today;
            worksheet.Cells[2, "I"] = "Сумма " + sum;
            worksheet.Cells[3, "I"] = "Кол-во " + kol_vo;
            if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
            {
                string Reestr = prefix_number;
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
            int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[23].Value) + 1;
            string prefix_number = comboBox10.Text + number;
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
            string kol_vo = Convert.ToString(textBox4.Text);//кол-во
            string sum = Convert.ToString(textBox5.Text);//сумма
            DateTime Now = DateTime.Today;
            worksheet.Cells[2, "I"] = "Сумма " + sum;
            worksheet.Cells[3, "I"] = "Кол-во " + kol_vo;
            if (Convert.ToString(dataGridView1.Rows[0].Cells[9].Value) != "Обработано")
            {
                string Reestr = prefix_number;
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

                string kol_vo = Convert.ToString(textBox4.Text);//кол-во
                string sum = Convert.ToString(textBox5.Text);//сумма               
                string oblast = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Область
                string kontragent = Convert.ToString(dataGridView1.Rows[0].Cells[7].Value);//Контрагент
                oDoc.Content.SetRange(0, 0);
                oDoc.Content.Text = $"                             Итого:    {kol_vo}                    {sum}" +
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
                int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[22].Value) + 1;
                string prefix_number = comboBox10.Text + number;
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text = kontragent + Environment.NewLine + "Накладная № " + prefix_number + " от " + Convert.ToString(Now.ToString("dd.MM.yyyy")) + " куда " + oblast +
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
                string kol_vo = Convert.ToString(textBox4.Text);//кол-во
                string sum = Convert.ToString(textBox5.Text);//сумма
                textBox23.Text = Convert.ToString(textBox15.Text);//передача услуга в Админпанель
                textBox24.Text = Convert.ToString(textBox4.Text);//передача кол-во в Админпанель
                textBox25.Text = Convert.ToString(textBox5.Text);//передача суммы в Админпанель
                string oblast = Convert.ToString(dataGridView1.Rows[0].Cells[9].Value);//Область
                string client = Convert.ToString(dataGridView1.Rows[0].Cells[8].Value);//Клиент
                //DateTime DatePriem = Convert.ToDateTime(dataGridView2.Rows[0].Cells[8].Value);
                oDoc.Content.SetRange(0, 0);
                oDoc.Content.Text = $"                             Итого:     {kol_vo}                    {sum}" +
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
                    Environment.NewLine + $"c  {Convert.ToString(dateTimePicker2.Value.ToString("dd.MM.yyyy "))}  по  {Convert.ToString(dateTimePicker1.Value.ToString(" dd.MM.yyyy"))}" +
                    Environment.NewLine +
                    Environment.NewLine + $"Отправитель  {client}" +
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
                string client = comboBox5.Text;//Клиент
                string kol_vo = Convert.ToString(textBox4.Text);//кол-во
                string sum = Convert.ToString(textBox5.Text);//сумма
                string plata_za_usluguSumm = Convert.ToString(textBox15.Text);//Плата за услугу Сумма
                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = $"Итого: {kol_vo}" +
                Environment.NewLine + $"Сумма объявленной ценности  {sum}" +
                Environment.NewLine + $"Сумма за услугу  {plata_za_usluguSumm}" + Environment.NewLine +
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
                    int number = Convert.ToInt32(dataGridView2.Rows[0].Cells[21].Value) + 1;
                    string prefix_number = comboBox10.Text + number;
                    if (dataGridView5.Rows[0].Cells[11].Value.ToString() == "0")
                    {
                        headerRange.Text = $"СПИСОК № {prefix_number}" +
                        Environment.NewLine + $"от {Convert.ToString(DatePriem.ToString("dd.MM.yyyy"))} принятых в ТЛЦ ГП 'Спецсвязь' " +
                        Environment.NewLine +
                        Environment.NewLine + $"Отправитель {client}" +
                        Environment.NewLine;
                    }
                    else if (dataGridView5.Rows[0].Cells[11].Value.ToString() != "0")
                    {
                        string nomer = dataGridView5.Rows[0].Cells[11].Value.ToString();//№
                        headerRange.Text = $"СПИСОК № {nomer}" +
                        Environment.NewLine + $"от {Convert.ToString(DatePriem.ToString("dd.MM.yyyy"))} принятых в ТЛЦ ГП 'Спецсвязь' " +
                        Environment.NewLine +
                        Environment.NewLine + $"Отправитель {client}" +
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
            string kol_vo = Convert.ToString(textBox4.Text);//кол-во
            string sum = Convert.ToString(textBox5.Text);//сумма
            DateTime Now = DateTime.Today;
            worksheet.Cells[2, "k"] = "Сумма " + sum;
            worksheet.Cells[3, "k"] = "Кол-во " + kol_vo;
            worksheet.Columns.AutoFit();//Автоматическая ширина колонок
            worksheet.Rows[1].Font.Bold = true; //Жирный шрифт
                                                //----------------------------------------------//
                                                // Сохранить приложение
            workbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excell.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //app.Quit();// Выход из приложения           
            Marshal.ReleaseComObject(app);// Уничтожение объекта Excel.          
            GC.GetTotalMemory(true);// Вызываем сборщик мусора для немедленной очистки памяти
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
            Invoice Invoice = new Invoice(this.dataGridView1, this.dataGridView2);// передаем ссылку на грид в форму Invoice
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
            if (comboBox5.Text != "")//только при выбранном контрагенте
            {
                dataGridView5.Visible = true;
                dataGridView1.Visible = false;
                dataGridView2.Visible = false;
                if (dateTimePicker2.Value <= DateTime.Today.AddDays(-1))//За диапазон
                {
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("SELECT oblast, punkt, familia, N_zakaza, data_zapisi, summ, tarif, doplata, plata_za_uslugu, ob_cennost, plata_za_nalog, id, nomer_spiska " +
                        "FROM [Table_1] WHERE (data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                    cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                    cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                    cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                    SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                    dt.Clear();//чистим DataTable, если он был не пуст
                    da.Fill(dt);//заполняем данными созданный DataTable
                    dataGridView5.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                    con.Close();//закрыть соединение    
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() != "0")
                    {
                        Podschet();//произвести подсчет по методу 
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        sfd.FileName = $"Список принятых с  {Convert.ToString(dateTimePicker2.Value.ToString("dd.MM.yyyy"))}  по  {Convert.ToString(dateTimePicker1.Value.ToString("dd.MM.yyyy"))}.docx";
                        button2.Text = "Ожидайте!";
                        //Выдача в WORD 
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                        }
                        //Выдача рееста в EXCEL
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Список принятых с  {Convert.ToString(dateTimePicker2.Value.ToString("dd.MM.yyyy"))}  по  {Convert.ToString(dateTimePicker1.Value.ToString("dd.MM.yyyy"))}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Excel(dataGridView5, sfd.FileName);
                        }
                        button2.Text = "Список принятых";
                    }
                    else if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() == "0")
                    {
                        MessageBox.Show("Вы не обработали один из списков!", "Внимание! Дальнейшие действия запрещены!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    }
                    else
                    {
                        MessageBox.Show("За этот период не найдено списков!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else if (dateTimePicker2.Value == DateTime.Today.AddDays(0))//За текущий день
                {
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    Select_client();//Для сортировки принятых списков по клиенту
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("SELECT oblast, punkt, familia, N_zakaza, data_zapisi, summ, tarif, doplata, plata_za_uslugu, ob_cennost, plata_za_nalog, id, nomer_spiska" +
                        " FROM [Table_1] WHERE (nomer_spiska = @nomer_spiska AND client = @client)", con);
                    if (textBox14.Text != "") { cmd.Parameters.AddWithValue("nomer_spiska", textBox14.Text); }//Ввести номер списка
                    else if (textBox14.Text == "") { cmd.Parameters.AddWithValue("nomer_spiska", dataGridView2.Rows[0].Cells[18].Value.ToString()); }//Если не ввести номер то выдаст последний
                    cmd.Parameters.AddWithValue("@client", comboBox5.Text);
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                    SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                    dt.Clear();//чистим DataTable, если он был не пуст
                    da.Fill(dt);//заполняем данными созданный DataTable
                    dataGridView5.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                    con.Close();//закрыть соединение
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    Podschet();//произвести подсчет по методу 
                    if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() == "0")
                    {
                        Select_Ns();//Выборка и сортировка по номеру от больших значений к меньшим.
                        con.Open();//открыть соединение
                        for (int i = 0; i < dataGridView5.Rows.Count; i++)//Цикл
                        {
                            SqlCommand cmd1 = new SqlCommand("UPDATE [Table_1] SET nomer_spiska = @nomer_spiska, Ns=@Ns WHERE id = @id", con);
                            cmd1.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView5.Rows[i].Cells[11].Value));
                            cmd1.Parameters.AddWithValue("@nomer_spiska", Number.Prefix_number);
                            cmd1.Parameters.AddWithValue("@Ns", Number.Ns);
                            cmd1.ExecuteNonQuery();
                        }
                        con.Close();//закрыть соединение 
                        label1.Text = ("Присвоен № Списка!");
                        //----------------------------------------//
                        button2.Text = "Ожидайте!";
                        //Выдача в WORD
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        sfd.FileName = $"Список принятых № {Number.Prefix_number}.docx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                        }
                        button2.Text = "Список принятых";
                    }
                    else if (dataGridView5.Rows.Count != 0 && dataGridView5.Rows[0].Cells[12].Value.ToString() != "0")
                    {
                        //Выдача в WORD
                        button2.Text = "Ожидайте!";
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        if (textBox14.Text == "") { string number = dataGridView5.Rows[0].Cells[12].Value.ToString(); sfd.FileName = $"Список принятых № {number}.docx"; }//№}
                        else if (textBox14.Text != "") { string nomer = textBox14.Text; sfd.FileName = $"Список принятых № {nomer}.docx"; }
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            Export_Spisok_Prinyatyh_To_Word(dataGridView5, sfd.FileName);
                        }
                        button2.Text = "Список принятых";
                    }
                }
                else
                {
                    MessageBox.Show($"Список по контрагенту {comboBox5.Text} не найден", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                dataGridView2.Visible = true;
                dataGridView1.Visible = false;
                dataGridView5.Visible = false;
                Disp_data();
                textBox14.Text = "";//Очистка поля
                comboBox5.SelectedIndex = -1;
            }
            else
            {
                MessageBox.Show("Необходимо выбрать контрагента", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

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
                    contextMenuStrip.Items.Add("Редактирование").Click += new EventHandler(Edit_Click);
                    contextMenuStrip.Items.Add("Статус").Click += new EventHandler(Status_Click);
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
        private void Status_Click(object sender, EventArgs e)
        {

        }
    }
}




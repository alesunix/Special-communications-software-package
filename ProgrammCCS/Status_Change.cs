using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProgramCCS
{
    public partial class Status_Change : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1
        public Status_Change(DataGridView dgv2)
        {
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
            InitializeComponent();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int currRowIndex = dgv2_TLC.CurrentCell.RowIndex;//  Запоминаем строку, которую выбрал пользователь.
            if (comboBox1.Text != "")
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
            }
            else if (comboBox1.Text == "")
            {
                MessageBox.Show("Выбирите статус", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            comboBox1.SelectedIndex = -1;
            F1.Disp_data();
            dgv2_TLC.CurrentCell = dgv2_TLC[0, currRowIndex];//  Выбираем нашу строку (именно выбираем, не выделяем).
        }

        private void Status_Change_Load(object sender, EventArgs e)
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
        }

        private void Status_Change_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }
    }
}

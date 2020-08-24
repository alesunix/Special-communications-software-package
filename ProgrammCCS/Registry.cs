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
    public partial class Registry : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели

        private DataGridView dgv1_TLC; // эта переменная будет содержать ссылку на грид dataGridView1 из формы Form1
        private DataGridView dgv2_TLC; // эта переменная будет содержать ссылку на грид dataGridView2 из формы Form1
        public Registry(DataGridView dgv1, DataGridView dgv2)
        {
            dgv1_TLC = dgv1;// теперь dgv1_TLC будет ссылкой на грид dataGridView1
            dgv2_TLC = dgv2;// теперь dgv1_TLC2 будет ссылкой на грид dataGridView2
            InitializeComponent();
        }
        TLC F1 = new TLC();

        private void button2_Click(object sender, EventArgs e)//Выборка
        {
            //1.Выборка на реестр-1 - 'Статус + Обработка + Филиал + Клиент'.
            if (comboBox1.Text != "" & comboBox5.Text != "" & comboBox4.Text != "")
            {
                dgv1_TLC.Visible = true;
                dgv2_TLC.Visible = false;
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
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
            }
            //2.Выборка на реестр-2 - 'Статус + Обработка + Клиент'.
            else if (comboBox1.Text != "" & comboBox5.Text != "" & comboBox4.Text == "")
            {
                dgv1_TLC.Visible = true;
                dgv1_TLC.Visible = false;
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
                dgv1_TLC.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение
            }
            F1.Podschet();//произвести подсчет из метода
        }
        private void button1_Click(object sender, EventArgs e)//Печать
        {
            //Обработка и Выдача реестра
            if (dgv1_TLC.Rows.Count > 0 & Convert.ToString(dgv1_TLC.Rows[0].Cells[10].Value) != "Обработано"
                & Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value) != "Отправлено"
                & Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value) != "Ожидание"
                & Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value) != "Розыск"
                & Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value) != "Замена")
            {
                F1.Select_status_Nr();//Выборка по статусу и сортировка по номеру реестра от больших значений к меньшим.                                      
                if (MessageBox.Show("Вы хотите обработать эти записи?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    for (int i = 0; i < dgv1_TLC.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_1] SET obrabotka = @obrabotka, data_obrabotki = @data_obrabotki, nomer_reestra = @nomer_reestra, Nr=@Nr WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@obrabotka", "Обработано");
                        cmd.Parameters.AddWithValue("@data_obrabotki", DateTime.Today.AddDays(0));
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dgv1_TLC.Rows[i].Cells[11].Value));
                        cmd.Parameters.AddWithValue("@nomer_reestra", Number.Prefix_number);
                        cmd.Parameters.AddWithValue("@Nr", Number.Nr);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение 
                    MessageBox.Show("Обработка выполнена / Присвоен № Реестра!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //------Ручная вставка номера реестра и обработки----------//
                    for (int i = 0; i < dgv1_TLC.Rows.Count; i++)//Цикл
                    {
                        dgv1_TLC.Rows[i].Cells[12].Value = Number.Prefix_number;
                        dgv1_TLC.Rows[i].Cells[10].Value = "Обработано";
                    }
                    //------Ручная вставка номера реестра и обработки----------//
                }
                //Выдача рееста в WORD
                string status = Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value);//Статус
                string kontragent = Convert.ToString(dgv1_TLC.Rows[0].Cells[8].Value);//Контрагент                
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                    {
                        F1.Export_Reestr_To_Word(dgv1_TLC, sfd.FileName);
                    }
                    else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                    {
                        F1.Export_Reestr_To_Word_vozvrat(dgv1_TLC, sfd.FileName);
                    }
                }
                //Выдача рееста в EXCEL
                if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                {
                    sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                    sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        F1.Export_Reestr_To_Excel(dgv1_TLC, sfd.FileName);
                    }
                }
                else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                {
                    sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                    sfd.FileName = $"Реестр № {Number.Prefix_number} на {status}.xlsx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        F1.Export_Reestr_To_Excel_vozvrat(dgv1_TLC, sfd.FileName);
                    }
                }
            }
            else if (dgv1_TLC.Rows.Count > 0 & Convert.ToString(dgv1_TLC.Rows[0].Cells[10].Value) == "Обработано")
            {
                if (MessageBox.Show("Вы хотите открыть этот Реестр?", "Внимание! Эти данные уже обработаны!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    //Выдача рееста в WORD
                    string nomer = dgv1_TLC.Rows[0].Cells[12].Value.ToString();//№
                    string status = Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value);//Статус
                    string kontragent = Convert.ToString(dgv1_TLC.Rows[0].Cells[8].Value);//Контрагент
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Реестр № {nomer} на {status}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                        {
                            F1.Export_Reestr_To_Word(dgv1_TLC, sfd.FileName);
                        }
                        else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                        {
                            F1.Export_Reestr_To_Word_vozvrat(dgv1_TLC, sfd.FileName);
                        }
                    }
                    //Выдача рееста в EXCEL
                    if (status != "Возврат" | kontragent != "TOO Sapar delivery" & kontragent != "ОсОО Тенгри")
                    {
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Реестр № {nomer} на {status}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            F1.Export_Reestr_To_Excel(dgv1_TLC, sfd.FileName);
                        }
                    }
                    else if (status == "Возврат" | kontragent == "TOO Sapar delivery" & kontragent == "ОсОО Тенгри")
                    {
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Реестр № {nomer} на {status}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            F1.Export_Reestr_To_Excel_vozvrat(dgv1_TLC, sfd.FileName);
                        }
                    }
                }
            }
            else if (dgv1_TLC.Rows.Count > 0 & Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value) == "Розыск" || Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value) == "Замена")
            {
                if (MessageBox.Show("Вы хотите открыть этот Реестр?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    //Выдача рееста в WORD
                    string nomer = dgv1_TLC.Rows[0].Cells[12].Value.ToString();//№
                    string status = Convert.ToString(dgv1_TLC.Rows[0].Cells[5].Value);//Статус
                    string kontragent = Convert.ToString(dgv1_TLC.Rows[0].Cells[8].Value);//Контрагент
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "Word Documents (*.docx)|*.docx";
                    sfd.FileName = $"Реестр № {nomer} на {status}.docx";
                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        F1.Export_Reestr_To_Word(dgv1_TLC, sfd.FileName);
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

        private void Registry_Load(object sender, EventArgs e)//Загрузка формы
        {
            Partner_select();
        }

        private void Registry_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }

        
    }
}

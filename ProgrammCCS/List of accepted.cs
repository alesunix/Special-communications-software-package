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
    public partial class List_of_accepted : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели

        public List_of_accepted()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)//Список принятых
        {
            //Передадим данные в конструктор класса Dates
            Dates.StartDate = dateTimePicker2.Value;
            Dates.EndDate = dateTimePicker1.Value;
            Partner.Name = comboBox5.Text;

            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            if (comboBox5.Text != "")//только при выбранном контрагенте
            {
                F1.dataGridView5.Visible = true;
                F1.dataGridView1.Visible = false;
                F1.dataGridView2.Visible = false;
                if (dateTimePicker2.Value <= DateTime.Today.AddDays(-1))//За диапазон
                {
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("SELECT oblast, punkt, familia, N_zakaza, data_zapisi, summ, tarif, doplata, plata_za_uslugu, ob_cennost, plata_za_nalog, id, nomer_spiska " +
                        "FROM [Table_1] WHERE (data_zapisi BETWEEN @StartDate AND @EndDate AND client = @client)", con);
                    cmd.Parameters.AddWithValue("StartDate", dateTimePicker2.Value);
                    cmd.Parameters.AddWithValue("EndDate", dateTimePicker1.Value);
                    cmd.Parameters.AddWithValue("@client", Partner.Name);
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                    SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                    dt.Clear();//чистим DataTable, если он был не пуст
                    da.Fill(dt);//заполняем данными созданный DataTable
                    F1.dataGridView5.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                    con.Close();//закрыть соединение    
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    if (F1.dataGridView5.Rows.Count != 0 && F1.dataGridView5.Rows[0].Cells[12].Value.ToString() != "0")
                    {
                        F1.Podschet();//произвести подсчет по методу 
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        sfd.FileName = $"Список принятых с  {Convert.ToString(dateTimePicker2.Value.ToString("dd.MM.yyyy"))}  по  {Convert.ToString(dateTimePicker1.Value.ToString("dd.MM.yyyy"))}.docx";
                        //Выдача в WORD 
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            F1.Export_Spisok_Prinyatyh_To_Word(F1.dataGridView5, sfd.FileName);
                        }
                        //Выдача рееста в EXCEL
                        sfd.Filter = "Книга Execl (*.xlsx)|*.xlsx";
                        sfd.FileName = $"Список принятых с  {Convert.ToString(dateTimePicker2.Value.ToString("dd.MM.yyyy"))}  по  {Convert.ToString(dateTimePicker1.Value.ToString("dd.MM.yyyy"))}.xlsx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            F1.Export_Spisok_Prinyatyh_To_Excel(F1.dataGridView5, sfd.FileName);
                        }
                    }
                    else if (F1.dataGridView5.Rows.Count != 0 && F1.dataGridView5.Rows[0].Cells[12].Value.ToString() == "0")
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
                    F1.Select_client();//Для сортировки принятых списков по клиенту
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("SELECT oblast, punkt, familia, N_zakaza, data_zapisi, summ, tarif, doplata, plata_za_uslugu, ob_cennost, plata_za_nalog, id, nomer_spiska" +
                        " FROM [Table_1] WHERE (nomer_spiska = @nomer_spiska AND client = @client)", con);
                    if (textBox14.Text != "") { cmd.Parameters.AddWithValue("nomer_spiska", textBox14.Text); }//Ввести номер списка
                    else if (textBox14.Text == "") { cmd.Parameters.AddWithValue("nomer_spiska", F1.dataGridView2.Rows[0].Cells[18].Value.ToString()); }//Если не ввести номер то выдаст последний
                    cmd.Parameters.AddWithValue("@client", Partner.Name);
                    cmd.ExecuteNonQuery();
                    DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                    SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                    dt.Clear();//чистим DataTable, если он был не пуст
                    da.Fill(dt);//заполняем данными созданный DataTable
                    F1.dataGridView5.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                    con.Close();//закрыть соединение
                    //-------------------------------------Выборка--------------------------------------------------------------------------------//
                    F1.Podschet();//произвести подсчет по методу 
                    if (F1.dataGridView5.Rows.Count != 0 && F1.dataGridView5.Rows[0].Cells[12].Value.ToString() == "0")
                    {
                        F1.Select_Ns();//Выборка и сортировка по номеру от больших значений к меньшим.
                        con.Open();//открыть соединение
                        for (int i = 0; i < F1.dataGridView5.Rows.Count; i++)//Цикл
                        {
                            SqlCommand cmd1 = new SqlCommand("UPDATE [Table_1] SET nomer_spiska = @nomer_spiska, Ns=@Ns WHERE id = @id", con);
                            cmd1.Parameters.AddWithValue("@id", Convert.ToInt32(F1.dataGridView5.Rows[i].Cells[11].Value));
                            cmd1.Parameters.AddWithValue("@nomer_spiska", Number.Prefix_number);
                            cmd1.Parameters.AddWithValue("@Ns", Number.Ns);
                            cmd1.ExecuteNonQuery();
                        }
                        con.Close();//закрыть соединение 
                        //----------------------------------------//
                        //Выдача в WORD
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        sfd.FileName = $"Список принятых № {Number.Prefix_number}.docx";
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            F1.Export_Spisok_Prinyatyh_To_Word(F1.dataGridView5, sfd.FileName);
                        }
                    }
                    else if (F1.dataGridView5.Rows.Count != 0 && F1.dataGridView5.Rows[0].Cells[12].Value.ToString() != "0")
                    {
                        //Выдача в WORD
                        SaveFileDialog sfd = new SaveFileDialog();
                        sfd.Filter = "Word Documents (*.docx)|*.docx";
                        if (textBox14.Text == "") { string number = F1.dataGridView5.Rows[0].Cells[12].Value.ToString(); sfd.FileName = $"Список принятых № {number}.docx"; }//№}
                        else if (textBox14.Text != "") { string nomer = textBox14.Text; sfd.FileName = $"Список принятых № {nomer}.docx"; }
                        if (sfd.ShowDialog() == DialogResult.OK)
                        {
                            F1.Export_Spisok_Prinyatyh_To_Word(F1.dataGridView5, sfd.FileName);
                        }
                    }
                }
                else
                {
                    MessageBox.Show($"Список по контрагенту {comboBox5.Text} не найден", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                F1.dataGridView2.Visible = true;
                F1.dataGridView1.Visible = false;
                F1.dataGridView5.Visible = false;
                F1.Disp_data();
                textBox14.Text = "";//Очистка поля
                comboBox5.SelectedIndex = -1;
            }
            else
            {
                MessageBox.Show("Необходимо выбрать контрагента", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
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
        private void List_of_accepted_Load(object sender, EventArgs e)
        {
            Partner_select();

            dateTimePicker2.Value = DateTime.Today.AddDays(0);
            dateTimePicker1.Value = DateTime.Today.AddMonths(-1);
        }

        private void List_of_accepted_FormClosed(object sender, FormClosedEventArgs e)
        {
            Hide();
        }
    }
}

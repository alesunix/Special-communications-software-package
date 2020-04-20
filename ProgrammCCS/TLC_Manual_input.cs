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
    public partial class Form_manual_input : Form
    {
        public Form_manual_input()
        {
            InitializeComponent();
        }
        public Form_manual_input(TLC f1)
        {
            InitializeComponent();
        }
        TLC F1 = new TLC();
        Login formLogin = new Login();

        private void button13_Click(object sender, EventArgs e)
        {
            if (textBox7.Text != "" & textBox9.Text != "" & textBox10.Text != "" & textBox11.Text != "" & textBox12.Text != "" & comboBox6.Text != "" & comboBox7.Text != "")
            {
                var summ = Convert.ToDouble(textBox10.Text.Replace(',', '.'));//запятую превратить в точку
                formLogin.con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_1] (oblast, punkt, familia, summ, N_zakaza, data_zapisi, obrabotka, status, client, tarif, nomer_reestra," +
                    " doplata, nomer_spiska, nomer_nakladnoy, Nr,Nn,Ns,tarifs) VALUES (@oblast, @punkt, @familia, @summ, @N_zakaza, @data_zapisi, @obrabotka, @status, @client," +
                    " @tarif, @nomer_reestra, @doplata, @nomer_spiska, @nomer_nakladnoy,@Nr,@Nn,@Ns,@tarifs)", formLogin.con);
                cmd.Parameters.AddWithValue("@oblast", comboBox6.Text);
                cmd.Parameters.AddWithValue("@punkt", textBox7.Text);
                cmd.Parameters.AddWithValue("@familia", textBox9.Text);
                cmd.Parameters.AddWithValue("@summ", summ);
                cmd.Parameters.AddWithValue("@N_zakaza", textBox11.Text);
                cmd.Parameters.AddWithValue("@data_zapisi", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@obrabotka", "Не обработано");
                cmd.Parameters.AddWithValue("@status", "Ожидание");
                cmd.Parameters.AddWithValue("@client", comboBox7.Text);
                cmd.Parameters.AddWithValue("@tarif", textBox12.Text);
                cmd.Parameters.AddWithValue("@nomer_reestra", 0);
                cmd.Parameters.AddWithValue("@nomer_nakladnoy", 0);
                cmd.Parameters.AddWithValue("@Nr", 0);
                cmd.Parameters.AddWithValue("@Nn", 0);
                cmd.Parameters.AddWithValue("@tarifs", F1.dataTable.Rows[0][0].ToString());//tarif
                if (textBox17.Text != "") { cmd.Parameters.AddWithValue("@doplata", textBox17.Text); }
                else if (textBox17.Text == "") { cmd.Parameters.AddWithValue("@doplata", 0); }
                if (textBox20.Text != "") { cmd.Parameters.AddWithValue("@nomer_spiska", textBox20.Text); cmd.Parameters.AddWithValue("@Ns", textBox20.Text); }
                else if (textBox20.Text == "") { cmd.Parameters.AddWithValue("@nomer_spiska", 0); cmd.Parameters.AddWithValue("@Ns", 0); }
                cmd.ExecuteNonQuery();
                formLogin.con.Close();//закрыть соединение
                textBox9.Text = "";//очистка текстовых полей
                textBox7.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                textBox12.Text = "";
                textBox17.Text = "";
                textBox20.Text = "";
                comboBox6.Select();//Установка курсора
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                F1.Podschet();//произвести подсчет по методу       
                F1.Disp_data();
            }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            F1.Disp_data();
            F1.Tarifs();//Т а р и ф ы   
            F1.Disp_data();
        }
        public void Partner_select()//Вывод Контрагентов в Combobox
        {
            formLogin.con.Open();//Открываем соединение
            SqlCommand cmd = formLogin.con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT name FROM [Table_Partner] ORDER BY id";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            //DGVF1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            foreach (DataRow column in dt.Rows)
            {
                comboBox7.Items.Add(column[0].ToString());
            }
            formLogin.con.Close();//Закрываем соединение          
        }
        private void Form_manual_input_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Hide();
        }

        private void Form_manual_input_Load(object sender, EventArgs e)
        {
            Partner_select();
        }

        private void comboBox7_TextChanged(object sender, EventArgs e)
        {
            formLogin.con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT tarif FROM [Table_Partner]" +
                "WHERE name = @name", formLogin.con);
            cmd.Parameters.AddWithValue("@name", comboBox7.Text.ToString());
            cmd.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            F1.dataTable.Clear();//чистим DataTable, если он был не пуст
            da.Fill(F1.dataTable);//заполняем данными созданный DataTable
            formLogin.con.Close();//закрыть соединение
            if (comboBox7.Text == "")//если поле очищено, отобразить базу
            {
                F1.dataTable.Clear();//чистим DataTable, если он был не пуст
                foreach (DataRow column in F1.dataTable.Rows)
                {
                    comboBox7.Items.Add(column[0].ToString());
                }
            }
        }
    }
}

using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace ProgramCCS
{
    public partial class Form_manual_input : Form
    {
        public SqlConnection con = Connection.con;//Получить строку соединения из класса модели

        public Form_manual_input()
        {
            InitializeComponent();
        }
        
        private void Button13_Click(object sender, EventArgs e)
        {
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            if (textBox7.Text != "" & textBox9.Text != "" & textBox10.Text != "" & textBox11.Text != "" & textBox12.Text != "" & comboBox6.Text != "" & comboBox7.Text != "")
            {
                var summ = Convert.ToDouble(textBox10.Text.Replace(',', '.'));//запятую превратить в точку
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_1] (oblast, punkt, familia, summ, N_zakaza, data_zapisi, obrabotka, status, client, tarif, plata_za_uslugu, ob_cennost, plata_za_nalog, nomer_reestra," +
                    " doplata, nomer_spiska, nomer_nakladnoy, Nr,Nn,Ns,tarifs) VALUES (@oblast, @punkt, @familia, @summ, @N_zakaza, @data_zapisi, @obrabotka, @status, @client," +
                    " @tarif, @plata_za_uslugu, @ob_cennost, @plata_za_nalog, @nomer_reestra, @doplata, @nomer_spiska, @nomer_nakladnoy,@Nr,@Nn,@Ns,@tarifs)", con);
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
                cmd.Parameters.AddWithValue("@plata_za_uslugu", 0);
                cmd.Parameters.AddWithValue("@ob_cennost", 0);
                cmd.Parameters.AddWithValue("@plata_za_nalog", 0);
                cmd.Parameters.AddWithValue("@nomer_reestra", 0);
                cmd.Parameters.AddWithValue("@nomer_nakladnoy", 0);
                cmd.Parameters.AddWithValue("@Nr", 0);
                cmd.Parameters.AddWithValue("@Nn", 0);
                cmd.Parameters.AddWithValue("@tarifs", F1.dtTarif.Rows[0][0].ToString());//tarif
                if (textBox17.Text != "") { cmd.Parameters.AddWithValue("@doplata", textBox17.Text); }
                else if (textBox17.Text == "") { cmd.Parameters.AddWithValue("@doplata", 0); }
                if (textBox20.Text != "") { cmd.Parameters.AddWithValue("@nomer_spiska", textBox20.Text); cmd.Parameters.AddWithValue("@Ns", textBox20.Text); }
                else if (textBox20.Text == "") { cmd.Parameters.AddWithValue("@nomer_spiska", 0); cmd.Parameters.AddWithValue("@Ns", 0); }
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                textBox9.Text = "";//очистка текстовых полей
                textBox7.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                textBox12.Text = "";
                textBox17.Text = "";
                textBox20.Text = "";
                comboBox6.Select();//Установка курсора
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);               
            }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);               
                F1.Wait();
            }
            //Отобразить список Ожидание! 
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
            //DGVF1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            foreach (DataRow column in dt.Rows)
            {
                comboBox7.Items.Add(column[0].ToString());
            }
            con.Close();//Закрываем соединение          
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
            TLC F1 = this.Owner as TLC;//Получаем ссылку на первую форму //Вызов метода формы из другой формы
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT tarif FROM [Table_Partner]" +
                "WHERE name = @name", con);
            cmd.Parameters.AddWithValue("@name", comboBox7.Text.ToString());
            cmd.ExecuteNonQuery();
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            F1.dtTarif.Clear();//чистим DataTable, если он был не пуст
            da.Fill(F1.dtTarif);//заполняем данными созданный DataTable
            con.Close();//закрыть соединение
            if (comboBox7.Text == "")//если поле очищено, отобразить базу
            {
                F1.dtTarif.Clear();//чистим DataTable, если он был не пуст
                foreach (DataRow column in F1.dtTarif.Rows)
                {
                    comboBox7.Items.Add(column[0].ToString());
                }
            }
        }
    }
}

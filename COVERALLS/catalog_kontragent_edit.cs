using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Media;
using System.Text.RegularExpressions;

namespace COVERALLS
{
    public partial class catalog_kontragent_edit : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        public catalog_kontragent_edit()
        {
            InitializeComponent();
        }

        public static bool isValid(string email)
        {
            string pattern = "[.\\-_a-z0-9]+@([a-z0-9][\\-a-z0-9]+\\.)+[a-z]{2,6}";
            Match isMatch = Regex.Match(email, pattern, RegexOptions.IgnoreCase);
            return isMatch.Success;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введен ИНН/КПП", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (textBox2.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введено Наименование КА", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (textBox3.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введен Юр. адрес", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (textBox4.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введен Почтовый адрес", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (textBox5.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены Контакты", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (textBox6.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введен E-mail", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (!isValid(textBox6.Text))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Введен не корректный E-mail", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            catalog_kontragent catalog_kontragent = (catalog_kontragent)this.Owner;

            /////Вставка данных в таблицу
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "update KONTRAGENT set INN_KPP=@inn, NAIMENOVAN=@naim, UR_ADDRESS=@ur_ad, POST_ADDRESS=@pst_ad, CONTACTS=@cont, EMAIL=@em where ID=@id";
            scmd4.Parameters.AddWithValue("inn",textBox1.Text);
            scmd4.Parameters.AddWithValue("naim", textBox2.Text);
            scmd4.Parameters.AddWithValue("ur_ad", textBox3.Text);
            scmd4.Parameters.AddWithValue("pst_ad", textBox4.Text);
            scmd4.Parameters.AddWithValue("cont", textBox5.Text);
            scmd4.Parameters.AddWithValue("em", textBox6.Text);
            scmd4.Parameters.AddWithValue("id", catalog_kontragent.idd);
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////

            catalog_kontragent.refill();

            this.Close();
        }

        private void catalog_kontragent_edit_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from KONTRAGENT where ID=@id", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("id",catalog_kontragent.idd);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "KONTRAGENT");

            textBox1.Text = ds.Tables[0].Rows[0][1].ToString();
            textBox2.Text = ds.Tables[0].Rows[0][2].ToString();
            textBox3.Text = ds.Tables[0].Rows[0][3].ToString();
            textBox4.Text = ds.Tables[0].Rows[0][4].ToString();
            textBox5.Text = ds.Tables[0].Rows[0][5].ToString();
            textBox6.Text = ds.Tables[0].Rows[0][6].ToString();

        }
    }
}

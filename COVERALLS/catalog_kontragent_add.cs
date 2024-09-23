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
    public partial class catalog_kontragent_add : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        public catalog_kontragent_add()
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
            scmd4.CommandText = "INSERT into KONTRAGENT (INN_KPP, NAIMENOVAN, UR_ADDRESS, POST_ADDRESS, CONTACTS, EMAIL) values (@inn, @naim, @u_ad, @p_ad, @cont, @em)";
            scmd4.Parameters.AddWithValue("inn", textBox1.Text);
            scmd4.Parameters.AddWithValue("naim", textBox2.Text);
            scmd4.Parameters.AddWithValue("u_ad", textBox3.Text);
            scmd4.Parameters.AddWithValue("p_ad", textBox4.Text);
            scmd4.Parameters.AddWithValue("cont", textBox5.Text);
            scmd4.Parameters.AddWithValue("em", textBox6.Text);            
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
    }
}

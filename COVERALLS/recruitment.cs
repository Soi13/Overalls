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

namespace COVERALLS
{
    public partial class recruitment : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
     
        public recruitment()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if ((textBox1.Text.Length == 0) || (comboBox1.Text.Length == 0) || (comboBox2.Text.Length == 0) || (comboBox3.Text.Length == 0) || (textBox2.Text.Length == 0) || (textBox3.Text.Length == 0) || (textBox4.Text.Length == 0) || (comboBox4.Text.Length == 0))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Заполнены не все поля!", "Внимание", MessageBoxButtons.OK);
                return;
            }

            //Проверка на уже существующего сотрудника в БД при приеме
            SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where FIO='"+textBox1.Text+"'", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "SOTRUDNIKI");

            if (ds.Tables[0].Rows.Count>0)
            {
                MessageBox.Show("Сотрудник "+textBox1.Text+" уже существует в базе! Повторный прием невозможен!!!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                SystemSounds.Beep.Play();
                return;
            }

            int siz = 0;
            if (checkBox6.Checked == true)
            {
                siz = 1;
            }
            else
            {
                siz = 0;
            }

            int br = 0;
            if (comboBox2.Text == "Красноярский филиал ОАО \"СибЭР\"") { br = 4; }
            if (comboBox2.Text == "Абаканский филиал ОАО \"СибЭР\"") { br = 1; }
            if (comboBox2.Text == "Кемеровский филиал ОАО \"СибЭР\"") { br = 3; }
            if (comboBox2.Text == "Барнаульский филиал ОАО \"СибЭР\"") { br = 2; }
            if (comboBox2.Text == "Исполнительный аппарат") { br = 5; }

            /////////Вставка данных
            SqlCommand scmd = conn.CreateCommand();
            scmd.CommandText = "INSERT INTO SOTRUDNIKI (FIO,PROFESSION,BRANCH,DEPARTMENT,USER_ID,USERS_BRANCH_ID,NEED_SIZ,SIZE_SHOES,SIZE_CLOTHES,HEIGHT,SEX) VALUES (" + "'" + textBox1.Text + "', '" + comboBox1.Text + "', '" + comboBox2.Text + "', '" + comboBox3.Text + "', '" + Form2.val + "', '" + br + "', '" + siz + "', '" + textBox2.Text + "', '"+textBox3.Text+"', '"+textBox4.Text+ "', '"+comboBox4.Text+"')";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader;
            reader = scmd.ExecuteReader();
            conn.Close();
            //////////////////
                        
            try
            {
                conn.Open();
            }
            catch (Exception)
            {
                MessageBox.Show("Не удалось установить соединение с БД.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                SystemSounds.Beep.Play();
                return;
            }

          
            MessageBox.Show("Сотрудник принят удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            SystemSounds.Beep.Play();
        }

        private void recruitment_Load(object sender, EventArgs e)
        {
            //Заполнение поля профессия, данными из БД
            SqlCommand command2 = conn.CreateCommand();
            command2.CommandText = "select distinct PROFESSION from SOTRUDNIKI order by PROFESSION";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader1;
            reader1 = command2.ExecuteReader();
            while (reader1.Read())
            {
                try
                {
                    string result1 = reader1.GetString(0);
                    comboBox1.Items.Add(result1);                    
                }
                catch { }

            }
            conn.Close();
            ////////////////////////

            //Заполнение поля Отдел, данными из БД
            SqlCommand command3 = conn.CreateCommand();
            command3.CommandText = "select distinct DEPARTMENT from SOTRUDNIKI order by DEPARTMENT";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader2;
            reader2 = command3.ExecuteReader();
            while (reader2.Read())
            {
                try
                {
                    string result2 = reader2.GetString(0);
                    comboBox3.Items.Add(result2);                    
                }
                catch { }

            }
            conn.Close();
            ////////////////////////
        }

        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void textBox3_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void comboBox4_TextChanged(object sender, EventArgs e)
        {
            if ((comboBox4.Text != "М") && (comboBox4.Text != "Ж"))
            {
                MessageBox.Show("Произвольный ввод информации в данное поле запрещен! Возможен только выбор из списка.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                comboBox4.Text = "";
                SystemSounds.Beep.Play();
                return;
            }

        }
    }
}

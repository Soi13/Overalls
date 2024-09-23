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
    public partial class useradd : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        public useradd()
        {
            InitializeComponent();
        }

        private void useradd_Load(object sender, EventArgs e)
        {
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
                    comboBox2.Items.Add(result2);                    
                }
                catch { }

            }
            conn.Close();
            ////////////////////////




        }

        private void button1_Click(object sender, EventArgs e)
        {
            ////////////////////////////////
            if (textBox1.Text.Length == 0) 
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не заполнено поле \"Логин\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////////////////////////

            ////////////////////////////////
            if (textBox2.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не заполнено поле \"Полное имя\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ///////////////////////////////
            
            ////////////////////////////////
            if (comboBox2.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран Отдел!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ///////////////////////////////

            ////////////////////////////////
            if (comboBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран Филиал!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ///////////////////////////////

        
            //Проверка существует ли уже вводимы пользователь в БД
            SqlCommand command = new SqlCommand("select * from USERS where USER_NAME='" + textBox1.Text+"'", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "USERS");
            
            if (ds.Tables[0].Rows.Count != 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Cотрудник c логином \""+ textBox1.Text+ "\"" + "уже существует в базе!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            /////////////

            int status = 0;
            if (checkBox1.Checked == true) 
            {
                status = 1;
            } 
            else
            {
                status = 0;
            }

            int br = 0;
            if (comboBox1.Text == "Красноярский филиал ОАО \"СибЭР\"") { br = 4; }
            if (comboBox1.Text == "Абаканский филиал ОАО \"СибЭР\"") { br = 1; }
            if (comboBox1.Text == "Кемеровский филиал ОАО \"СибЭР\"") { br = 3; }
            if (comboBox1.Text == "Барнаульский филиал ОАО \"СибЭР\"") { br = 2; }
            if (comboBox1.Text == "Исполнительный аппарат") { br = 5; }

            /////Вставка пользовтеля в БД
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "insert into USERS (USER_NAME,FULL_NAME,PASSW,DEPARTMENT,BRANCH,EMAIL,ADMINISTRATION,BRANCH_ID) VALUES (" + "'" + textBox1.Text + "', '" + textBox2.Text + "', '6ece4fd51bc113942692637d9d4b860e', '" + comboBox2.Text + "', '" + comboBox1.Text + "', '" + textBox3.Text + "', '" + status + "', '" + br + "')";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////

            SystemSounds.Beep.Play();
            MessageBox.Show("Пользователь добавлен удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
           
        }
    }
}

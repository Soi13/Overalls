using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;


namespace COVERALLS
{
    public partial class user_change_data : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        public static int id;

        public user_change_data()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["USER_NAME"].HeaderText = "Логин";
            dataGridView1.Columns["USER_NAME"].Width = 100;
            dataGridView1.Columns["FULL_NAME"].HeaderText = "Имя пользователя";
            dataGridView1.Columns["FULL_NAME"].Width = 200;
            dataGridView1.Columns["PASSW"].HeaderText = "PASSW";
            dataGridView1.Columns["PASSW"].Width = 40;
            dataGridView1.Columns["PASSW"].Visible = false;            
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Отдел";
            dataGridView1.Columns["DEPARTMENT"].Width = 200;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 200;
            dataGridView1.Columns["EMAIL"].HeaderText = "EMAIL";
            dataGridView1.Columns["EMAIL"].Width = 200;
            dataGridView1.Columns["ADMINISTRATION"].HeaderText = "ADMINISTRATION";
            dataGridView1.Columns["ADMINISTRATION"].Width = 40;
            dataGridView1.Columns["ADMINISTRATION"].Visible = false;
            dataGridView1.Columns["BRANCH_ID"].HeaderText = "BRANCH_ID";
            dataGridView1.Columns["BRANCH_ID"].Width = 40;
            dataGridView1.Columns["BRANCH_ID"].Visible = false;            
            

        }
        //////////////////////////////////////////////////////

        private void user_change_data_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from USERS order by USER_NAME", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "USERS");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();

            //Заполнение поля отдел/участок, данными из БД
            SqlCommand command2 = conn.CreateCommand();
            command2.CommandText = "select distinct DEPARTMENT from USERS order by DEPARTMENT";
            try
            {
                conn.Open();
            }
            catch {}
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
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button2.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
            label5.Visible = true;
            textBox1.Visible = true;
            textBox2.Visible = true;
            textBox3.Visible = true;
            comboBox1.Visible = true;
            comboBox2.Visible = true;
                     
            id = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            textBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            textBox2.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[2].Value);
            comboBox1.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[4].Value);
            comboBox2.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[5].Value);
            textBox3.Text = Convert.ToString(dataGridView1.CurrentRow.Cells[6].Value);

            this.Height = 761;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            int k = 0;
            if (comboBox2.Text == "Красноярский филиал ОАО \"СибЭР\"") { k = 4; }
            if (comboBox2.Text == "Кемеровский филиал ОАО \"СибЭР\"") { k = 3; }
            if (comboBox2.Text == "Барнаульский филиал ОАО \"СибЭР\"") { k = 2; }
            if (comboBox2.Text == "Абаканский филиал ОАО \"СибЭР\"") { k = 1; }
            if (comboBox2.Text == "Исполнительный аппрарат ОАО \"СибЭР\"") { k = 5; }
            
            /////Вставка данных в таблицу журнала вход/выход
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "update USERS set USER_NAME='"+textBox1.Text+"', FULL_NAME='"+textBox2.Text+"', DEPARTMENT='"+comboBox1.Text+"', BRANCH='"+comboBox2.Text+"', EMAIL='"+textBox3.Text+"', BRANCH_ID="+ k +" where ID="+id;
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////

            //Обновление данных после UPDATE
            SqlCommand command = new SqlCommand("select * from USERS order by USER_NAME", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "USERS");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();

            //Заполнение поля отдел/участок, данными из БД
            SqlCommand command2 = conn.CreateCommand();
            command2.CommandText = "select distinct DEPARTMENT from USERS order by DEPARTMENT";
            try
            {
                conn.Open();
            }
            catch { }
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

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            
            label1.Visible = false;
            label2.Visible = false;
            label3.Visible = false;
            label4.Visible = false;
            label5.Visible = false;
            textBox1.Visible = false;
            textBox2.Visible = false;
            textBox3.Visible = false;
            comboBox1.Visible = false;
            comboBox2.Visible = false;
            button2.Visible = false;
            this.Height = 679;
            
        }
    }
}

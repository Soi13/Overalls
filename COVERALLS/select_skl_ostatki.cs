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
    public partial class select_skl_ostatki : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");


        public select_skl_ostatki()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["CODE_TMC"].HeaderText = "Код ТМЦ";
            dataGridView1.Columns["CODE_TMC"].Width = 100;
            dataGridView1.Columns["CODE_R3"].HeaderText = "Код R3";
            dataGridView1.Columns["CODE_R3"].Width = 100;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 200;
            dataGridView1.Columns["NAIMENOVAN_TMC"].HeaderText = "Наименование ТМЦ";
            dataGridView1.Columns["NAIMENOVAN_TMC"].Width = 350;
            dataGridView1.Columns["ZAYAVITEL"].HeaderText = "Заявитель";
            dataGridView1.Columns["ZAYAVITEL"].Width = 100;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед.изм.";
            dataGridView1.Columns["ED_IZM"].Width = 70;
            dataGridView1.Columns["KONECH_OSTATOK"].HeaderText = "Конечн. остаток";
            dataGridView1.Columns["KONECH_OSTATOK"].Width = 100;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 20;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 20;
            dataGridView1.Columns["USER_ID"].Visible = false;
            dataGridView1.Columns["BRANCH_ID"].HeaderText = "BRANCH_ID";
            dataGridView1.Columns["BRANCH_ID"].Width = 20;
            dataGridView1.Columns["BRANCH_ID"].Visible = false;
        }
        //////////////////////////////////////////////////////

        private void select_skl_ostatki_Load(object sender, EventArgs e)
        {
            if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ или Шепилева
            {
                SqlCommand command = new SqlCommand("select * from SKL_OSTATKI", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SKL_OSTATKI");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }
            else
            {
                SqlCommand command = new SqlCommand("select * from SKL_OSTATKI where branch_id=@br", conn);
                command.Parameters.AddWithValue("br", Form2.branch);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SKL_OSTATKI");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }

            //Заполнение поля Филиал для импорта
            SqlCommand command2 = conn.CreateCommand();
            command2.CommandText = "select distinct BRANCH from BRANCHES order by BRANCH";
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
                    comboBox3.Items.Add(result1);
                }
                catch { }

            }
            conn.Close();
            ////////////////////////
        }

        private void comboBox3_TextChanged(object sender, EventArgs e)
        {
            if (comboBox3.Text.Length != 0)
            {
                //Проверка на изменение поля Филиал для импорта остатков
                SqlCommand command = new SqlCommand("select BRANCH from BRANCHES where BRANCH=@br", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                command.Parameters.AddWithValue("br", comboBox3.Text);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "BRANCHES");

                checkBox3.Checked = false;

                if (ds.Tables[0].Rows.Count == 0)
                {
                    comboBox3.Text = "";
                    checkBox3.Enabled = false;
                    button2.Enabled = false;
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Ручной ввод в поле запрещен!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    checkBox3.Enabled = true;
                    button2.Enabled = true;
                }
            }
        }

        private void checkBox3_Click(object sender, EventArgs e)
        {
            if (checkBox3.Checked == true)
            {
                comboBox3.Text = "";
                button2.Enabled = false;

                SqlCommand command2 = new SqlCommand("select * from SKL_OSTATKI", conn);
                SqlDataAdapter da2 = new SqlDataAdapter(command2);//Переменная объявлена как глобальная
                SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                DataSet ds2 = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da2.Fill(ds2, "SKL_OSTATKI");
                dataGridView1.DataSource = ds2.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds2.Tables[0].Rows.Count);

                fill_gridview();
            }
            if ((checkBox3.Checked == false) && (comboBox3.Text.Length == 0))
            {
                checkBox3.Enabled = false;
                button2.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SqlCommand command1 = new SqlCommand("select * from SKL_OSTATKI where BRANCH=@b", conn);
            command1.Parameters.AddWithValue("b", comboBox3.Text);
            SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
            DataSet ds1 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds1, "SKL_OSTATKI");
            dataGridView1.DataSource = ds1.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds1.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void checkBox1_Click(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;
            }
        }

        private void checkBox2_Click(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                checkBox2.Checked = false;

                SqlCommand command2 = new SqlCommand("select * from SKL_OSTATKI where NAIMENOVAN_TMC like '%" + textBox1.Text + "%' order by NAIMENOVAN_TMC", conn);
                SqlDataAdapter da2 = new SqlDataAdapter(command2);//Переменная объявлена как глобальная
                SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                DataSet ds2 = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da2.Fill(ds2, "SKL_OSTATKI");
                dataGridView1.DataSource = ds2.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds2.Tables[0].Rows.Count);

                fill_gridview();
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                checkBox1.Checked = false;

                SqlCommand command3 = new SqlCommand("select * from SKL_OSTATKI where CODE_R3 like '" + textBox2.Text + "%'  order by NAIMENOVAN_TMC", conn);
                SqlDataAdapter da3 = new SqlDataAdapter(command3);//Переменная объявлена как глобальная
                SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                DataSet ds3 = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da3.Fill(ds3, "SKL_OSTATKI");
                dataGridView1.DataSource = ds3.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds3.Tables[0].Rows.Count);

                fill_gridview();
            }
        }
    }
}

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
    public partial class spr_iais : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
       

        public spr_iais()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["CODE_IAIS"].HeaderText = "Код ИАИС";
            dataGridView1.Columns["CODE_IAIS"].Width = 70;
            dataGridView1.Columns["NAIMEN_IAIS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_IAIS"].Width = 255;
            dataGridView1.Columns["ED_IZM_IAIS"].HeaderText = "Ед.изм.";
            dataGridView1.Columns["ED_IZM_IAIS"].Width = 50;
            dataGridView1.Columns["CODE_GROUP_IAIS"].HeaderText = "Код группы";
            dataGridView1.Columns["CODE_GROUP_IAIS"].Width = 50;
            dataGridView1.Columns["GROUP_IAIS"].HeaderText = "Наименование группы";
            dataGridView1.Columns["GROUP_IAIS"].Width = 200;
            dataGridView1.Columns["CODE_SAP"].HeaderText = "Код SAP";
            dataGridView1.Columns["CODE_SAP"].Width = 70;
            dataGridView1.Columns["NAIMEN_SAP"].HeaderText = "Наименование SAP";
            dataGridView1.Columns["NAIMEN_SAP"].Width = 200;
            dataGridView1.Columns["ED_IZM_SAP"].HeaderText = "Ед.изм SAP";
            dataGridView1.Columns["ED_IZM_SAP"].Width = 50;
            dataGridView1.Columns["CODE_GROUP_SAP"].HeaderText = "Код группы SAP";
            dataGridView1.Columns["CODE_GROUP_SAP"].Width = 70;
            dataGridView1.Columns["GROUP_SAP"].HeaderText = "Наименование группы SAP";
            dataGridView1.Columns["GROUP_SAP"].Width = 200;
            dataGridView1.Columns["PRICE"].HeaderText = "Цена";
            dataGridView1.Columns["PRICE"].Width = 100;          

        }
        //////////////////////////////////////////////////////

        private void spr_iais_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from IAIS order by NAIMEN_IAIS", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "IAIS");
            dataGridView1.DataSource = ds.Tables[0];
            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true) { checkBox2.Checked = false; }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true) { checkBox1.Checked = false; }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {

                if (textBox1.Text == "")
                {
                    SqlCommand command1 = new SqlCommand("select * from IAIS order by NAIMEN_IAIS", conn);
                    SqlDataAdapter da1 = new SqlDataAdapter(command1);
                    SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                    DataSet ds1 = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da1.Fill(ds1, "IAIS");
                    dataGridView1.DataSource = ds1.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds1.Tables[0].Rows.Count);

                    fill_gridview();
                    return;
                }

                SqlCommand command = new SqlCommand("select * from IAIS where CODE_IAIS='" + textBox1.Text + "'", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "IAIS");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
                
            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Поиск по кодам ИАИС не включен! Включите.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox1.Clear();
                return;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {

                if (textBox2.Text == "")
                {
                    SqlCommand command1 = new SqlCommand("select * from IAIS order by NAIMEN_IAIS", conn);
                    SqlDataAdapter da1 = new SqlDataAdapter(command1);
                    SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                    DataSet ds1 = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da1.Fill(ds1, "IAIS");
                    dataGridView1.DataSource = ds1.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds1.Tables[0].Rows.Count);

                    fill_gridview();
                }

                SqlCommand command = new SqlCommand("select * from IAIS where NAIMEN_IAIS like '%" + textBox2.Text + "%'", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "IAIS");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();

            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Поиск по наименованию ИАИС не включен! Включите.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox2.Clear();
                return;
            }
        }
    }
}

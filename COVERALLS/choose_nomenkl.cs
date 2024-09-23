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
    public partial class choose_nomenkl : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        public choose_nomenkl()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview1()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["NORMS_ID"].HeaderText = "NORMS_ID";
            dataGridView1.Columns["NORMS_ID"].Width = 40;
            dataGridView1.Columns["NORMS_ID"].Visible = false;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView1.Columns["CHARACTERISTIC_OVERALLS"].Width = 200;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView1.Columns["ED_IZM"].Width = 50;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 50;
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView1.Columns["PRICE_EDENIC_PLAN"].HeaderText = "Стоим. за ед. план.";
            dataGridView1.Columns["PRICE_EDENIC_PLAN"].Width = 80;
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоим. за комплект план.";
            dataGridView1.Columns["PRICE_KOMPLEKT_PLAN"].Width = 80;
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "ИАИС Наименование";
            dataGridView1.Columns["IAIS_NAIMEN_OVERALLS"].Width = 230;
            dataGridView1.Columns["CODE_GROUP_IAIS"].HeaderText = "CODE_GROUP_IAIS";
            dataGridView1.Columns["CODE_GROUP_IAIS"].Width = 40;
            dataGridView1.Columns["CODE_GROUP_IAIS"].Visible = false;
            dataGridView1.Columns["KONTRAGENT_ID"].HeaderText = "KONTRAGENT_ID";
            dataGridView1.Columns["KONTRAGENT_ID"].Width = 10;
            dataGridView1.Columns["KONTRAGENT_ID"].Visible = false;
        }
        //////////////////////////////////////////////////////

        private void choose_nomenkl_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from NORMS_LIST where KONTRAGENT_ID is null order by NAIMEN_OVERALLS", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da1);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds, "NORMS_LIST");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview1();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {
                if (textBox2.Text.Length > 0)
                {
                    SqlCommand command = new SqlCommand("select * from NORMS_LIST where where KONTRAGENT_ID is null and NAIMEN_OVERALLS like '%" + textBox2.Text + "%'  order by NAIMEN_OVERALLS", conn);
                    SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "NORMS_LIST");
                    dataGridView1.DataSource = ds.Tables[0];

                    fill_gridview1();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from NORMS_LIST where KONTRAGENT_ID is null order by NAIMEN_OVERALLS", conn);
                    SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "NORMS_LIST");
                    dataGridView1.DataSource = ds.Tables[0];

                    fill_gridview1();
                }

            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Поиск не включен! Включите.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox2.Clear();
                return;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
            }
            catch
            {
            }
            for (int u=0; u<=dataGridView1.Rows.Count-1; u++)
            {                
                if (dataGridView1.Rows[u].Selected)
                {
                    SqlCommand scmd = conn.CreateCommand();
                    scmd.CommandText = "update NORMS_LIST set KONTRAGENT_ID=@id where ID=@id_str";
                    scmd.Parameters.AddWithValue("id", catalog_kontragent.idd);
                    scmd.Parameters.AddWithValue("id_str", dataGridView1.Rows[u].Cells[0].Value);
                    SqlDataReader reader;
                    reader = scmd.ExecuteReader();
                    reader.Dispose();
                }
            }
            conn.Close();

            SystemSounds.Beep.Play();
            MessageBox.Show("Номенклатура привязана удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
    }
}

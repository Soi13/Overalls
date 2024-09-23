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
    public partial class show_tie_nomenkl : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
       
        public show_tie_nomenkl()
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

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void show_tie_nomenkl_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from NORMS_LIST where KONTRAGENT_ID=@id order by NAIMEN_OVERALLS", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da1);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("id",catalog_kontragent.idd);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds, "NORMS_LIST");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview1();
        }
    }
}

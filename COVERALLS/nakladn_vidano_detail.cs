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
    public partial class nakladn_vidano_detail : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        public nakladn_vidano_detail()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["GUID_NAKLADN_VIDANO"].HeaderText = "GUID_NAKLADN_VIDANO";
            dataGridView1.Columns["GUID_NAKLADN_VIDANO"].Width = 40;
            dataGridView1.Columns["GUID_NAKLADN_VIDANO"].Visible = false;
            dataGridView1.Columns["NORMS_LIST_ID"].HeaderText = "NORMS_LIST_ID";
            dataGridView1.Columns["NORMS_LIST_ID"].Width = 20;
            dataGridView1.Columns["NORMS_LIST_ID"].Visible = false;
            dataGridView1.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование СИЗ";
            dataGridView1.Columns["NAIMEN_OVERALLS"].Width = 300;
            dataGridView1.Columns["ED_IZM"].HeaderText = "Ед.изм";
            dataGridView1.Columns["ED_IZM"].Width = 60;
            dataGridView1.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView1.Columns["KOLVO"].Width = 80;
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView1.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView1.Columns["DATE_VIDACHI"].HeaderText = "Дата выдачи";
            dataGridView1.Columns["DATE_VIDACHI"].Width = 100;
            dataGridView1.Columns["DATE_SLED_VIDACHI"].HeaderText = "Дата след. выдачи";
            dataGridView1.Columns["DATE_SLED_VIDACHI"].Width = 120;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 20;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;

        }

        private void nakladn_vidano_detail_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from NAKLADN_VIDANO_DETAIL where GUID_NAKLADN_VIDANO=@GUID", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("GUID", nakladn_vidano.g);
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NAKLADN_VIDANO_DETAIL");
            dataGridView1.DataSource = ds.Tables[0];
                        
            fill_gridview();
        }
    }
}

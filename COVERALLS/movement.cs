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
    public partial class movement : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
       
        public movement()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 20;
            dataGridView1.Columns["USER_ID"].Visible = false;
            dataGridView1.Columns["SOTRUDNIKI_ID"].HeaderText = "SOTRUDNIKI_ID";
            dataGridView1.Columns["SOTRUDNIKI_ID"].Width = 20;
            dataGridView1.Columns["SOTRUDNIKI_ID"].Visible = false;
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 150;            
            dataGridView1.Columns["OLD_PROFESSION"].HeaderText = "Старая профессия/должность";
            dataGridView1.Columns["OLD_PROFESSION"].Width = 200;
            dataGridView1.Columns["OLD_DEPARTMENT"].HeaderText = "Старый Отдел/участок/цех";
            dataGridView1.Columns["OLD_DEPARTMENT"].Width = 200;
            dataGridView1.Columns["NEW_PROFESSION"].HeaderText = "Новая профессия/должность";
            dataGridView1.Columns["NEW_PROFESSION"].Width = 200;
            dataGridView1.Columns["NEW_DEPARTMENT"].HeaderText = "Новый Отдел/участок/цех";
            dataGridView1.Columns["NEW_DEPARTMENT"].Width = 200;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 20;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;
            dataGridView1.Columns["DATE_MOVEMENT"].HeaderText = "Дата перемещения";
            dataGridView1.Columns["DATE_MOVEMENT"].Width = 100;
            dataGridView1.Columns["NOTES"].HeaderText = "NOTES";
            dataGridView1.Columns["NOTES"].Width = 20;
            dataGridView1.Columns["NOTES"].Visible = false;

        }
        //////////////////////////////////////////////////////

        private void button1_Click(object sender, EventArgs e)
        {
            movement_param movement_param = new movement_param();
            movement_param.Owner = this;
            movement_param.ShowDialog();
        }

        private void movement_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select MOVEMENT.ID, MOVEMENT.USER_ID, MOVEMENT.SOTRUDNIKI_ID, SOTRUDNIKI.FIO, MOVEMENT.OLD_PROFESSION, MOVEMENT.OLD_DEPARTMENT, MOVEMENT.NEW_PROFESSION, MOVEMENT.NEW_DEPARTMENT, MOVEMENT.DATETIME_CREATE,MOVEMENT.DATE_MOVEMENT, MOVEMENT.NOTES from MOVEMENT, SOTRUDNIKI where MOVEMENT.SOTRUDNIKI_ID=SOTRUDNIKI.ID", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "MOVEMENT");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            richTextBox1.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();
        }
    }
}

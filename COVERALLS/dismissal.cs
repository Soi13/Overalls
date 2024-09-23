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
    public partial class dismissal : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
       
        public dismissal()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 200;
            dataGridView1.Columns["PROFESSION"].HeaderText = "Профессия/должность";
            dataGridView1.Columns["PROFESSION"].Width = 200;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 200;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Отдел/участок/цех";
            dataGridView1.Columns["DEPARTMENT"].Width = 200;
            dataGridView1.Columns["SPISOK_ODEJD"].HeaderText = "Норматив";
            dataGridView1.Columns["SPISOK_ODEJD"].Width = 200;
            dataGridView1.Columns["SPISOK_ODEJD"].Visible = false;         
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 40;
            dataGridView1.Columns["USER_ID"].Visible = false;
            dataGridView1.Columns["NORMS_KEY_N"].HeaderText = "NORMS_KEY_N";
            dataGridView1.Columns["NORMS_KEY_N"].Width = 20;
            dataGridView1.Columns["NORMS_KEY_N"].Visible = false;
            dataGridView1.Columns["USERS_BRANCH_ID"].HeaderText = "USERS_BRANCH_ID";
            dataGridView1.Columns["USERS_BRANCH_ID"].Width = 20;
            dataGridView1.Columns["USERS_BRANCH_ID"].Visible = false;
            dataGridView1.Columns["DATE_CREATE"].HeaderText = "DATE_CREATE";
            dataGridView1.Columns["DATE_CREATE"].Width = 20;
            dataGridView1.Columns["DATE_CREATE"].Visible = false;
            dataGridView1.Columns["DATE_DISMISSAL"].HeaderText = "Дата увольнения";
            dataGridView1.Columns["DATE_DISMISSAL"].Width = 100;
            dataGridView1.Columns["NOTES"].HeaderText = "Примечания";
            dataGridView1.Columns["NOTES"].Width = 200;
            dataGridView1.Columns["SIZE_SHOES"].HeaderText = "Размер обуви";
            dataGridView1.Columns["SIZE_SHOES"].Width = 70;
            dataGridView1.Columns["SIZE_SHOES"].Visible = false;
            dataGridView1.Columns["SIZE_CLOTHES"].HeaderText = "Размер одежды";
            dataGridView1.Columns["SIZE_CLOTHES"].Width = 70;
            dataGridView1.Columns["SIZE_CLOTHES"].Visible = false;
            dataGridView1.Columns["HEIGHT"].HeaderText = "Рост";
            dataGridView1.Columns["HEIGHT"].Width = 40;
            dataGridView1.Columns["HEIGHT"].Visible = false;
            dataGridView1.Columns["SEX"].HeaderText = "Пол";
            dataGridView1.Columns["SEX"].Width = 40;
            dataGridView1.Columns["SEX"].Visible = false;
            
        }
        //////////////////////////////////////////////////////

        private void dismissal_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;

            SqlCommand command = new SqlCommand("select * from DISMISSAL", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "DISMISSAL");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dismissal_select dismissal_select = new dismissal_select();
            dismissal_select.ShowDialog();
        }
    }
}

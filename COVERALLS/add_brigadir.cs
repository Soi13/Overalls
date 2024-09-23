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
    public partial class add_brigadir : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;

        public add_brigadir()
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
            
        }
        //////////////////////////////////////////////////////

        private void add_brigadir_Load(object sender, EventArgs e)
        {
            if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ или Шепилева
            {             
                SqlCommand command = new SqlCommand("select ID, FIO, PROFESSION, BRANCH, DEPARTMENT from SOTRUDNIKI", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];
                             
                fill_gridview();
            }
            else
            {
                SqlCommand command = new SqlCommand("select ID, FIO, PROFESSION, BRANCH, DEPARTMENT from SOTRUDNIKI where USERS_BRANCH_ID=" + Form2.branch, conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];
                             
                fill_gridview();
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            brigades_add brigades_add = (brigades_add)this.Owner;

            brigades_add.dataGridView1.Rows.Add(dataGridView1.CurrentRow.Cells[0].Value, dataGridView1.CurrentRow.Cells[1].Value, dataGridView1.CurrentRow.Cells[2].Value, dataGridView1.CurrentRow.Cells[3].Value, dataGridView1.CurrentRow.Cells[4].Value);
            this.Close();
        }
    }
}

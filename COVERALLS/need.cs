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
    public partial class need : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        public static string gd;

        public need()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 20;
            dataGridView1.Columns["USER_ID"].Visible = false;
            dataGridView1.Columns["KIND_ZAYAV"].HeaderText = "Тип заявки";
            dataGridView1.Columns["KIND_ZAYAV"].Width = 100;
            dataGridView1.Columns["PODR"].HeaderText = "Отдел/участок/цех";
            dataGridView1.Columns["PODR"].Width = 200;
            dataGridView1.Columns["PERIOD"].HeaderText = "Период заявки";
            dataGridView1.Columns["PERIOD"].Width = 150;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 20;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;
        }
        //////////////////////////////////////////////////////

        private void need_Load(object sender, EventArgs e)
        {
            if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ или Шепилева
            {
                SqlCommand command = new SqlCommand("select * from NEED_MAIN", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_MAIN");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }
            else
            {
                SqlCommand command = new SqlCommand("select * from NEED_MAIN where USER_ID=" + Form2.val, conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NEED_MAIN");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            gd = dataGridView1.CurrentRow.Cells[1].Value.ToString();

            need_details need_details = new need_details();
            need_details.ShowDialog();
        }
    }
}

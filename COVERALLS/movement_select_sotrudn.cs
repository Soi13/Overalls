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
    public partial class movement_select_sotrudn : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
      
        public movement_select_sotrudn()
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
            dataGridView1.Columns["FIO"].Width = 150;
            dataGridView1.Columns["PROFESSION"].HeaderText = "Профессия/должность";
            dataGridView1.Columns["PROFESSION"].Width = 200;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Отдел/участок/цех";
            dataGridView1.Columns["DEPARTMENT"].Width = 200;           

        }
        //////////////////////////////////////////////////////

        private void movement_select_sotrudn_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select id, fio,PROFESSION,DEPARTMENT from SOTRUDNIKI order by FIO", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "SOTRUDNIKI");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            movement_param movement_param = (movement_param)this.Owner;

            if (movement_param.dataGridView1.Rows.Count > 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Сотрудник уже выбран!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string id = dataGridView1.CurrentRow.Cells[0].Value.ToString();
            string fio = dataGridView1.CurrentRow.Cells[1].Value.ToString();
            string profession = dataGridView1.CurrentRow.Cells[2].Value.ToString();
            string department = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                        
            movement_param.dataGridView1.Rows.Add(id, fio, profession, department);
            this.Close();
        }
    }
}

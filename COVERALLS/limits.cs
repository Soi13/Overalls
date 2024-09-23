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
    public partial class limits : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        SqlDataAdapter da;

        public limits()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 10;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 250;
            dataGridView1.Columns["LIMIT"].HeaderText = "Лимит";
            dataGridView1.Columns["LIMIT"].Width = 100;
        }
        //////////////////////////////////////////////////////

        private void limits_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select ID, BRANCH, LIMIT from LIMITS order by BRANCH", conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "LIMITS");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            label2.Visible = true;
            button2.Visible = true;
            button3.Visible = true;
            dataGridView1.ReadOnly = false;
            dataGridView1.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
            button1.Enabled = false;

        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
            dataGridView1.ReadOnly = true;
            button1.Enabled = true;
            label2.Visible = false;
            button2.Visible = false;
            button3.Visible = false; 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                da.Update((System.Data.DataTable)dataGridView1.DataSource);
                dataGridView1.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
                button1.Enabled = true;
                label2.Visible = false;
                button2.Visible = false;                

                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в лимитах выполнены.", "Уведомление о результатах", MessageBoxButtons.OK);

            }
            catch (Exception)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в базе данных выполнить не удалось!", "Уведомление о результатах", MessageBoxButtons.OK);
            }
        }
    }
}

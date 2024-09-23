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
    public partial class user_remove : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        public user_remove()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["USER_NAME"].HeaderText = "Логин";
            dataGridView1.Columns["USER_NAME"].Width = 100;
            dataGridView1.Columns["FULL_NAME"].HeaderText = "Имя пользователя";
            dataGridView1.Columns["FULL_NAME"].Width = 200;
            dataGridView1.Columns["PASSW"].HeaderText = "PASSW";
            dataGridView1.Columns["PASSW"].Width = 40;
            dataGridView1.Columns["PASSW"].Visible = false;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Отдел";
            dataGridView1.Columns["DEPARTMENT"].Width = 200;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 200;
            dataGridView1.Columns["EMAIL"].HeaderText = "EMAIL";
            dataGridView1.Columns["EMAIL"].Width = 200;
            dataGridView1.Columns["ADMINISTRATION"].HeaderText = "ADMINISTRATION";
            dataGridView1.Columns["ADMINISTRATION"].Width = 40;
            dataGridView1.Columns["ADMINISTRATION"].Visible = false;
            dataGridView1.Columns["BRANCH_ID"].HeaderText = "BRANCH_ID";
            dataGridView1.Columns["BRANCH_ID"].Width = 40;
            dataGridView1.Columns["BRANCH_ID"].Visible = false;


        }
        //////////////////////////////////////////////////////

        private void user_remove_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from USERS order by USER_NAME", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "USERS");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что необходимо удалить пользователя " + dataGridView1.CurrentRow.Cells[2].Value + "?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                /////////Удаление пользователя (фактически перенос человека в архивную таблицу USERS_ARCHIVE)
                SqlCommand cm = conn.CreateCommand();
                cm.CommandText = "BEGIN TRANSACTION " +
                                 "insert into USERS_ARCHIVE (USER_NAME,FULL_NAME,PASSW,DEPARTMENT,BRANCH,EMAIL,ADMINISTRATION,BRANCH_ID) values ('" + dataGridView1.CurrentRow.Cells[1].Value + "',  '" + dataGridView1.CurrentRow.Cells[2].Value + "', '" + dataGridView1.CurrentRow.Cells[3].Value + "', '" + dataGridView1.CurrentRow.Cells[4].Value + "', '" + dataGridView1.CurrentRow.Cells[5].Value + "', '" + dataGridView1.CurrentRow.Cells[6].Value + "', '" + dataGridView1.CurrentRow.Cells[7].Value + "', '" + dataGridView1.CurrentRow.Cells[8].Value +"') " +
                                 "delete from USERS where ID=" + dataGridView1.CurrentRow.Cells[0].Value +
                                 " COMMIT TRANSACTION";
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader1;
                reader1 = cm.ExecuteReader();
                conn.Close();
                ///////////////////////////////

                SystemSounds.Beep.Play();
                MessageBox.Show("Сотрудник " + dataGridView1.CurrentRow.Cells[1].Value + " удален успешно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();
            }
        }
    }
}

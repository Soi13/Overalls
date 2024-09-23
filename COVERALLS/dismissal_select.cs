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
    public partial class dismissal_select : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
      
        public dismissal_select()
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
            dataGridView1.Columns["NEED_SIZ"].HeaderText = "NEED_SIZ";
            dataGridView1.Columns["NEED_SIZ"].Width = 20;
            dataGridView1.Columns["NEED_SIZ"].Visible = false;
            dataGridView1.Columns["SIZE_SHOES"].HeaderText = "Размер обуви";
            dataGridView1.Columns["SIZE_SHOES"].Width = 70;
            dataGridView1.Columns["SIZE_SHOES"].Visible = false;
            dataGridView1.Columns["SIZE_CLOTHES"].HeaderText = "Размер одежды";
            dataGridView1.Columns["SIZE_CLOTHES"].Width = 70;
            dataGridView1.Columns["SIZE_CLOTHES"].Visible = false;
            dataGridView1.Columns["HEIGHT"].HeaderText = "Рост";
            dataGridView1.Columns["HEIGHT"].Width = 40;
            dataGridView1.Columns["HEIGHT"].Visible = false;

        }
        //////////////////////////////////////////////////////

        private void dismissal_select_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;

            SqlCommand command = new SqlCommand("select * from SOTRUDNIKI", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "SOTRUDNIKI");
            dataGridView1.DataSource = ds.Tables[0];

            fill_gridview();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where FIO like '%" + textBox1.Text + "%'", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "SOTRUDNIKI");
            dataGridView1.DataSource = ds.Tables[0];
                        
            fill_gridview();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что необходимо уволить сотрудника " + dataGridView1.CurrentRow.Cells[1].Value+"?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                //Проверка есть ли у человека уже выданная спецодежда для ее отображения
                SqlCommand command = new SqlCommand("select * from vidano where SOTRUDNIK_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
                SqlDataAdapter da = new SqlDataAdapter(command);
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                da.Fill(ds, "VIDANO");
                if (ds.Tables[0].Rows.Count > 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Данному сотруднику произведена выдача спецодежды, для его увольнения необходимо произвести списание выданных СИЗ.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                ///////////////////

                ///Проверка даты увольнения. Если она больше текущей, то выводится ошибка.
                if (DateTime.Now.Date < dateTimePicker1.Value.Date)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Увольнение будущим периодом невозможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                /////////Увольнение (перенос человека в таблицу Dismissal и удаление из таблицы Sotrudniki)
                /////после удаления работника из таблицы Sotrudniki,триггером "dismissal_trigger" удаляется привязка нормы к ээтому сотруднику. Триггер лежит в базе и срабатывает на событие УДАЛЕНИЕ по таблице SOTRUDNIKI
                SqlCommand cm = conn.CreateCommand();
                cm.CommandText = "BEGIN TRANSACTION " +
                                 "insert into DISMISSAL (FIO,PROFESSION,BRANCH,DEPARTMENT,SPISOK_ODEJD,USER_ID,NORMS_KEY_N,USERS_BRANCH_ID,DATE_CREATE,DATE_DISMISSAL,NOTES,SIZE_SHOES,SIZE_CLOTHES,HEIGHT) values ('" + dataGridView1.CurrentRow.Cells[1].Value + "',  '" + dataGridView1.CurrentRow.Cells[2].Value + "', '" + dataGridView1.CurrentRow.Cells[3].Value + "', '" + dataGridView1.CurrentRow.Cells[4].Value + "', '" + dataGridView1.CurrentRow.Cells[5].Value + "', '" + Form2.val + "', '" + dataGridView1.CurrentRow.Cells[7].Value + "', '" + dataGridView1.CurrentRow.Cells[8].Value + "', convert(datetime,'" + DateTime.Now.ToString() + "', 103), convert(datetime,'" + dateTimePicker1.Value.Date + "', 103), '" + richTextBox1.Text + "', '"+dataGridView1.CurrentRow.Cells[10].Value+"', '"+dataGridView1.CurrentRow.Cells[11].Value+"', '"+dataGridView1.CurrentRow.Cells[12].Value+"') " +
                                 "delete from SOTRUDNIKI where ID=" + dataGridView1.CurrentRow.Cells[0].Value +
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
                MessageBox.Show("Сотрудник " + dataGridView1.CurrentRow.Cells[1].Value+" уволен успешно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close();

            }
        }
    }
}

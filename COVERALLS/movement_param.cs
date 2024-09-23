using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;
using System.Data.SqlClient;

namespace COVERALLS
{
    public partial class movement_param : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        public movement_param()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            movement_select_sotrudn movement_select_sotrudn = new movement_select_sotrudn();
            movement_select_sotrudn.Owner = this;
            movement_select_sotrudn.ShowDialog();
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView1.Rows.Remove(dataGridView1.CurrentRow);
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count > 0)
            {
                button2.Enabled = true;
                button3.Enabled = true;
                button4.Enabled = true;                
                button6.Enabled = true;
                //button7.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
                button3.Enabled = false;
                button4.Enabled = false;                
                button6.Enabled = false;
                //button7.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            movement_select_profession movement_select_profession = new movement_select_profession();
            movement_select_profession.Owner = this;
            movement_select_profession.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбран сотрудник! Перемещение не возможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (dataGridView2.Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбраны профессия/должность и подразделение/участок! Перемещение не возможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;            
            }

            if (dataGridView2.Rows[0].Cells[0].Value.ToString().Length==0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбрана профессия/должность! Перемещение не возможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (dataGridView2.Rows[0].Cells[1].Value.ToString().Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбрано подразделение/участок! Перемещение не возможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (label4.Text == "label4")
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не выбрана норма! Перемещение не возможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            //Проверка положена ли СИЗ перемещаемому человеку. Если нет, то запрет на перемещение.
            SqlCommand command = new SqlCommand("select NEED_SIZ from SOTRUDNIKI where ID=" + dataGridView1.Rows[0].Cells[0].Value.ToString(), conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "SOTRUDNIKI");

            if (ds.Tables[0].Rows.Count==0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Данному сотруднику СИЗ не поожены! Перемещение не возможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            movement movement = (movement)this.Owner;

            /////////Обновление данных в БД
            SqlCommand cm = conn.CreateCommand();
            cm.CommandText = "BEGIN TRANSACTION " +
                             " insert into MOVEMENT (USER_ID, SOTRUDNIKI_ID, OLD_PROFESSION,OLD_DEPARTMENT, NEW_PROFESSION, NEW_DEPARTMENT, DATETIME_CREATE, DATE_MOVEMENT,NOTES) " +  
                             "VALUES (@us_id, @sotr_id, @old_prof, @old_dep, @new_prof, @new_dep, GETDATE(), @date_movement, @notes) " +
                             "update SOTRUDNIKI SET PROFESSION=@new_prof, DEPARTMENT=@new_dep where ID=@sotr_id " +
                             "update NORMS_PERSONAL SET NORMS_ID=(select ID from NORMS where profession=@norma) where SOTRUDNIK_ID=@sotr_id " +
                             " COMMIT TRANSACTION";
            //Параметры
            cm.Parameters.AddWithValue("us_id",Form2.val);
            cm.Parameters.AddWithValue("sotr_id", dataGridView1.Rows[0].Cells[0].Value.ToString());
            cm.Parameters.AddWithValue("old_prof", dataGridView1.Rows[0].Cells[2].Value.ToString());
            cm.Parameters.AddWithValue("old_dep", dataGridView1.Rows[0].Cells[3].Value.ToString());
            cm.Parameters.AddWithValue("new_prof", dataGridView2.Rows[0].Cells[0].Value.ToString());
            cm.Parameters.AddWithValue("new_dep", dataGridView2.Rows[0].Cells[1].Value.ToString());
            cm.Parameters.AddWithValue("date_movement", dateTimePicker1.Value);
            cm.Parameters.AddWithValue("notes", richTextBox1.Text);
            cm.Parameters.AddWithValue("norma", label4.Text);
            
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader1;
            reader1 = cm.ExecuteReader();
            conn.Close();
            ///////////////////////////////

            
            //Обновление данных на форме Перемещение сотрудников
            SqlCommand command1 = new SqlCommand("select MOVEMENT.ID, MOVEMENT.USER_ID, MOVEMENT.SOTRUDNIKI_ID, SOTRUDNIKI.FIO, MOVEMENT.OLD_PROFESSION, MOVEMENT.OLD_DEPARTMENT, MOVEMENT.NEW_PROFESSION, MOVEMENT.NEW_DEPARTMENT, MOVEMENT.DATETIME_CREATE,MOVEMENT.DATE_MOVEMENT, MOVEMENT.NOTES from MOVEMENT, SOTRUDNIKI where MOVEMENT.SOTRUDNIKI_ID=SOTRUDNIKI.ID", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
            DataSet ds1 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds1, "MOVEMENT");
            movement.dataGridView1.DataSource = ds1.Tables[0];

            movement.statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds1.Tables[0].Rows.Count);

            movement.fill_gridview();

            SystemSounds.Beep.Play();
            MessageBox.Show("Перемещение проведено удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);

            this.Close();

        }

        private void button3_Click(object sender, EventArgs e)
        {
            movement_select_department movement_select_department = new movement_select_department();
            movement_select_department.Owner = this;
            movement_select_department.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView2.Rows.Remove(dataGridView2.CurrentRow);
            }
        }

        private void удалитьToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView2.Rows.Remove(dataGridView2.CurrentRow);
            }
        }

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count > 0)
            {
                button7.Enabled = true;
            }
            else
            {
                button7.Enabled = false;
            }                    
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            movement_select_norms movement_select_norms = new movement_select_norms();
            movement_select_norms.Owner = this;
            movement_select_norms.ShowDialog();
        }
    }
}

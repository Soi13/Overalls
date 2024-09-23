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
    public partial class cancel_vidano : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;
        DataSet ds;

        public cancel_vidano()
        {
            InitializeComponent();
        }

        private void cancel_vidano_Load(object sender, EventArgs e)
        {
            dataGridView2.AllowUserToAddRows = false;
            label2.Text = "Сотрудник: " + Form1.sotrudn_fio;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                if (dataGridView2.Rows.Count == 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Список пуст. Удалять нечего!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                    dataGridView2.Rows.Remove(dataGridView2.CurrentRow);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите отказаться от операции?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView2.Rows.Clear();
                this.Close();
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //Проверка есть ли у человека уже выданная спецодежда для ее отображения
            SqlCommand command = new SqlCommand("select * from vidano where SOTRUDNIK_ID=" + Form1.sotrudn_id, conn);
            da = new SqlDataAdapter(command);
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            da.Fill(ds, "VIDANO");
            if (ds.Tables[0].Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Данному сотруднику не производилась выдача спецодежды! Отображать нечего.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ///////////////////

            Form8 form8 = new Form8();
            form8.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            cancel_vidano_add_positions cancel_vidano_add_positions = new cancel_vidano_add_positions(this);
            cancel_vidano_add_positions.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Произвести отмену выдачи невозможно, т.к. не выбраны позиции к выдаче.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("Вы уверены, что необходимо произвести отмену выдачи спецодежды?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                try
                {
                    conn.Open();
                }
                catch
                {
                }

                for (int u = 0; u <= dataGridView2.Rows.Count - 1; u++)
                {
                    string id = Convert.ToString(dataGridView2.Rows[u].Cells[0].Value);
                    string sotrudnik_id = Convert.ToString(dataGridView2.Rows[u].Cells[1].Value);
                    string naim = Convert.ToString(dataGridView2.Rows[u].Cells[2].Value);
                    string kolvo = Convert.ToString(dataGridView2.Rows[u].Cells[4].Value);

                    //Выборка позиции отменяемой спецодежды
                    SqlCommand command = new SqlCommand("select * from vidano where id='" + id +"' and sotrudnik_id='" + sotrudnik_id+ "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "VIDANO");

                    //Проверка на превышение кол-ва к отмене с отпущенным кол-вом
                    if (Convert.ToInt16(ds.Tables[0].Rows[0][5]) < Convert.ToInt16(kolvo))
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Превышено кол-во для отмены по " + naim + ". Отпущенное кол-во " + Convert.ToString(ds.Tables[0].Rows[0][5]) + ", а для отмены вы указали " + kolvo + ".", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    ///////////////////////////////////////////////////////////

                    //Проверка на указание нулевого кол-ва или меньше 0 к отмене
                    if (Convert.ToInt16(kolvo)<=0)
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Заданное кол-во для отмены по " + naim + " меньше или равно 0. Отмену произвести невозможно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    ///////////////////////////////////////////////////////////

                    //Проверка на равенсто отменяемого кол-ва и отпущенного кол-ва. Если они равны, то позиция просто удаляется.
                    if (Convert.ToInt16(ds.Tables[0].Rows[0][5]) == Convert.ToInt16(kolvo))
                    {
                        SqlCommand cmd = conn.CreateCommand();
                        cmd.CommandText = "delete from vidano where id='"+id+"' and sotrudnik_id='"+sotrudnik_id+"'";
                        try
                        {
                            conn.Open();
                        }
                        catch {}
                        SqlDataReader reader;
                        reader = cmd.ExecuteReader();                        
                    }
                    ///////////////////////////////////////////////////////////

                    //Проверка на превышение отменяемого кол-ва и отпущенного кол-ва. Если отменяемое кол-во меньше отпущенного, то позиция происходит вычитание и обновление кол-ва в данной позиции.
                    if (Convert.ToInt16(ds.Tables[0].Rows[0][5]) > Convert.ToInt16(kolvo))
                    {
                        int res=(Convert.ToInt16(ds.Tables[0].Rows[0][5])) - (Convert.ToInt16(kolvo));
                        SqlCommand cmd = conn.CreateCommand();
                        cmd.CommandText = "update vidano set kolvo='"+ res + "' where id='" + id + "' and sotrudnik_id='" + sotrudnik_id + "'";
                        try
                        {
                            conn.Open();
                        }
                        catch { }
                        SqlDataReader reader;
                        reader = cmd.ExecuteReader();                        
                    }
                    ///////////////////////////////////////////////////////////
                    conn.Close();                                           
                }

                conn.Close();
                SystemSounds.Beep.Play();
                MessageBox.Show("Отмена произведена успешно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dataGridView2.Rows.Clear();
            }
        }
    }
}

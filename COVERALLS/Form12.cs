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
    public partial class Form12 : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;
        DataSet ds;

        public Form12()
        {
            InitializeComponent();
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

        private void button1_Click(object sender, EventArgs e)
        {            
            Form13 form13 = new Form13(this);
            form13.ShowDialog();
        }

        private void Form12_Load(object sender, EventArgs e)
        {
            dataGridView2.AllowUserToAddRows = false;
            label2.Text = "Сотрудник: " + Form1.sotrudn_fio;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Сформировать акт на списание невозможно, т.к. не выбраны позиции к списанию.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (MessageBox.Show("Вы уверены, что необходимо произвести списание спецодежды?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                              
                //Проверка позиций на превышение списываемого кол-ва перед вставкой самого акта на списание (заголовка), а также проверка дат при списании
                for (int p = 0; p <= dataGridView2.Rows.Count - 1; p++)
                {
                    string id_main = Convert.ToString(dataGridView2.Rows[p].Cells[0].Value);
                    string sotrudnik_id_main = Convert.ToString(dataGridView2.Rows[p].Cells[1].Value);
                    string naim_main = Convert.ToString(dataGridView2.Rows[p].Cells[3].Value);
                    string kolvo_main = Convert.ToString(dataGridView2.Rows[p].Cells[5].Value);
                    string data_vidachi = Convert.ToString(dataGridView2.Rows[p].Cells[7].Value);

                    //Выборка позиции отменяемой спецодежды
                    SqlCommand command1 = new SqlCommand("select * from vidano where id='" + id_main + "' and sotrudnik_id='" + sotrudnik_id_main + "'", conn);
                    da = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb1 = new SqlCommandBuilder(da);
                    DataSet ds1 = new DataSet();
                    da.Fill(ds1, "VIDANO");
                    if (Convert.ToInt16(ds1.Tables[0].Rows[0][5]) < Convert.ToInt16(kolvo_main))
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Превышено кол-во для списания по " + naim_main + ". Кол-во в наличие " + Convert.ToString(ds1.Tables[0].Rows[0][5]) + ", а для списания вы указали " + kolvo_main + ".", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }

                    //Проверка дат. Если дата акта меньше, чем дата выдачи списываемых СИЗ, то списание не возможно, т.к. логически это СИЗ еще не выдано.
                    if (dateTimePicker1.Value.Date < Convert.ToDateTime(data_vidachi))
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Дата выдачи СИЗ по \"" + naim_main + "\" меньше даты данного акта. Невозможно произвести списание СИЗ датой ниже, чем ее выдали!!!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    ///////////////////////////////

                    //Проверка кол-ва на меньше 0.
                    if (Convert.ToInt32(kolvo_main)<=0)
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Количество к списанию по позиции \"" + naim_main + "\" меньше или равно 0!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        return;
                    }
                    ///////////////////////////////
                }         
                
                /////////Вставка самого акта (заголовка)
                SqlCommand cm = conn.CreateCommand();
                cm.CommandText = "BEGIN TRANSACTION " +
                                  "DECLARE @m as int " +
                                  "SET @m=(select max(NUMBER) from AKT_SPISAN) " +
                                  "if @m is null set @m=1 else set @m=@m+1 " +
                                  "insert into AKT_SPISAN (SOTRUDNIK_ID, NUMBER, DATE_SPISAN, NOTES, DATE_CREATE, USER_ID) values ('" + Form1.sotrudn_id + "', @m, convert(datetime,'" + dateTimePicker1.Value.Date + "', 103), '" + richTextBox1.Text + "', convert(datetime,'" + DateTime.Now.ToString() + "', 103), '" + Form2.val + "') " +
                                  "COMMIT TRANSACTION";
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader1;
                reader1 = cm.ExecuteReader();
                conn.Close();
                ///////////////////////////////

                //////////Вычисление максимального значения ID в актах на списание для присваивания этого значения позициям акта
                try
                {
                    conn.Open();
                }
                catch { }
                SqlCommand command2 = new SqlCommand("select max(ID) from AKT_SPISAN", conn);
                da = new SqlDataAdapter(command2);//Переменная объявлена как глобальная
                SqlCommandBuilder cb2 = new SqlCommandBuilder(da);
                DataSet ds2 = new DataSet();
                da.Fill(ds2, "AKT_SPISAN");
                conn.Close();

                string ID_value_akt_spis=ds2.Tables[0].Rows[0][0].ToString();
                /////////

                //Вставка наполнения акта (списанные позиции)
                for (int u = 0; u <= dataGridView2.Rows.Count - 1; u++)
                {
                    string id = Convert.ToString(dataGridView2.Rows[u].Cells[0].Value);
                    string sotrudnik_id = Convert.ToString(dataGridView2.Rows[u].Cells[1].Value);
                    string norms_list_id = Convert.ToString(dataGridView2.Rows[u].Cells[2].Value);
                    string naim = Convert.ToString(dataGridView2.Rows[u].Cells[3].Value);
                    string edizm = Convert.ToString(dataGridView2.Rows[u].Cells[4].Value);
                    string kolvo = Convert.ToString(dataGridView2.Rows[u].Cells[5].Value);
                    string period_isp = Convert.ToString(dataGridView2.Rows[u].Cells[6].Value);
                    string data_vidachi = Convert.ToString(dataGridView2.Rows[u].Cells[7].Value);
                    string data_sled_vidachi = Convert.ToString(dataGridView2.Rows[u].Cells[8].Value);
                    string size = Convert.ToString(dataGridView2.Rows[u].Cells[9].Value);
                    string user_id = Convert.ToString(dataGridView2.Rows[u].Cells[10].Value);
                    string naim_iais = Convert.ToString(dataGridView2.Rows[u].Cells[11].Value);
                    string id_iais = Convert.ToString(dataGridView2.Rows[u].Cells[12].Value);
                    
                    //Выборка позиции отменяемой спецодежды
                    SqlCommand command = new SqlCommand("select * from vidano where id='" + id + "' and sotrudnik_id='" + sotrudnik_id + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    da.Fill(ds, "VIDANO");

                    //Проверка на равенсто отменяемого кол-ва и отпущенного кол-ва. Если они равны, то позиция просто удаляется.
                    if (Convert.ToInt16(ds.Tables[0].Rows[0][5]) == Convert.ToInt16(kolvo))
                    {                      
                        SqlCommand cmd = conn.CreateCommand();
                        cmd.CommandText = "BEGIN TRANSACTION " +
                                          "insert into AKT_SPISAN_NOMENKL (AKT_SPISAN_ID, SOTRUDNIK_ID, NORMS_LIST_ID, NAIMEN_OVERALLS, ED_IZM, KOLVO, PERIOD_ISPOLZOVAN, DATE_VIDACHI, DATE_SLED_VIDACHI, SIZE, USER_ID, DATETIME_CREATE, IAIS_NAIMEN_OVERALLS, ID_IAIS) values ('"+ ID_value_akt_spis+"', '" + sotrudnik_id + "', '" + norms_list_id + "', '" + naim + "', '" + edizm + "', '" + kolvo + "', '" + period_isp + "', '" +data_vidachi + "', '" +data_sled_vidachi+ "', '" + size + "', '" + user_id + "', convert(datetime,'" + DateTime.Now.ToString() + "', 103), '" + naim_iais + "', '" + id_iais + "') " +
                                          "delete from vidano where id='" + id + "' and sotrudnik_id='" + sotrudnik_id + "' "+
                                          "COMMIT TRANSACTION";                                           
                        try
                        {
                            conn.Open();
                        }
                        catch { }
                        SqlDataReader reader;
                        reader = cmd.ExecuteReader();
                        conn.Close();
                    }
                    ///////////////////////////////////////////////////////////


                    //Проверка на равенсто отменяемого кол-ва и отпущенного кол-ва. Если кол-во к отмене меньше отпущенного, то происходит UPDATE.
                    if (Convert.ToInt32(ds.Tables[0].Rows[0][5]) > Convert.ToInt32(kolvo))
                    {

                        int ress = Convert.ToInt32(ds.Tables[0].Rows[0][5]) - Convert.ToInt32(kolvo);//Подсчет разницы
                        SqlCommand cmd = conn.CreateCommand();
                        cmd.CommandText = "BEGIN TRANSACTION " +
                                          "insert into AKT_SPISAN_NOMENKL (AKT_SPISAN_ID, SOTRUDNIK_ID, NORMS_LIST_ID, NAIMEN_OVERALLS, ED_IZM, KOLVO, PERIOD_ISPOLZOVAN, DATE_VIDACHI, DATE_SLED_VIDACHI, SIZE, USER_ID, DATETIME_CREATE, IAIS_NAIMEN_OVERALLS, ID_IAIS) values ('" + ID_value_akt_spis + "', '" + sotrudnik_id + "', '" + norms_list_id + "', '" + naim + "', '" + edizm + "', '" + kolvo + "', '" + period_isp + "', '" + data_vidachi + "', '" + data_sled_vidachi + "', '" + size + "', '" + user_id + "', convert(datetime,'" + DateTime.Now.ToString() + "', 103), '" + naim_iais + "', '" + id_iais + "') " +
                                          "update vidano set kolvo='"+ ress + "' where id='" + id + "' and sotrudnik_id='" + sotrudnik_id + "' " +
                                          "COMMIT TRANSACTION";
                        try
                        {
                            conn.Open();
                        }
                        catch { }
                        SqlDataReader reader;
                        reader = cmd.ExecuteReader();
                        conn.Close();
                    }
                    ///////////////////////////////////////////////////////////
                                        
                }

                SystemSounds.Beep.Play();
                MessageBox.Show("Акт сформирован удачно!", "Уведомление", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dataGridView2.Rows.Clear();
            }
        }

        private void dataGridView2_KeyPress(object sender, KeyPressEventArgs e)
        {
            //Ввод только цифр
            if (!Char.IsDigit(e.KeyChar))
            {
                e.Handled = true;
            }
        }
    }
}

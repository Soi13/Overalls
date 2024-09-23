using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Media;

namespace COVERALLS
{
    public partial class Form7 : Form
    {
        
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        SqlDataAdapter da;
        //DataSet ds;
        string date_sled_vidachi;

                  
        public Form7()
        {
            InitializeComponent();
        }

        private void Form7_Load(object sender, EventArgs e)
        {
            dataGridView2.AllowUserToAddRows = false;
            label2.Text = "Сотрудник: " + Form1.sotrudn_fio;
                        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView2.Rows.Count == 0)
            {
                MessageBox.Show("Произвести выдачу невозможно, т.к. не выбраны позиции к выдаче.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;                
            }
            
            if (MessageBox.Show("Вы уверены, что необходимо произвести выдачу спецодежды?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {

                if (textBox1.Text.Length == 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Выдача не возможна, т.к. она не подтверждена личной цифровой подписью.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    textBox1.Focus();                
                    return;
                }

                //Проверка на существование считываемой карты в системе
                SqlCommand command3 = new SqlCommand("select CARD_UIN from CARDS where CARD_UIN=@uin", conn);
                SqlDataAdapter da3 = new SqlDataAdapter(command3);//Переменная объявлена как глобальная
                SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                DataSet ds3 = new DataSet();
                command3.Parameters.AddWithValue("uin", textBox1.Text);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da3.Fill(ds3, "CARDS");

                if (ds3.Tables[0].Rows.Count == 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Выдача не возможна! Данная карта не зарегистрирована в системе! Сначала присвойте данную карту сотруднику, а затем производите выдачу", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    textBox1.Focus();
                    return;
                }
                ////////////////////////////////////////////////////////


                //////////////Проверка на подпись своей картой, а не чужой, т.е. чтобы выдать СИЗ было невозможно если подписываешь другой картой, не принадлежащей данному сторуднику
                
                //Определение UIN карты сотруднику кому выдаем СИЗ
                SqlCommand command4 = new SqlCommand("select CARD_UIN from CARDS where SOTRUDNIK_ID=@id", conn);
                SqlDataAdapter da4 = new SqlDataAdapter(command4);//Переменная объявлена как глобальная
                SqlCommandBuilder cb4 = new SqlCommandBuilder(da4);
                DataSet ds4 = new DataSet();
                command4.Parameters.AddWithValue("id", Form1.sotrudn_id);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da4.Fill(ds4, "CARDS");

                if (ds4.Tables[0].Rows.Count != 0)
                {
                    if (ds4.Tables[0].Rows[0][0].ToString() != textBox1.Text)
                    {
                        SystemSounds.Beep.Play();
                        MessageBox.Show("Выдача не возможна! Попытка подписать выдачу картой принадлещей другому сотруднику!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        textBox1.Focus();
                        return;
                    }


                    try
                    {
                        conn.Open();
                    }
                    catch
                    {
                    }

                    for (int u = 0; u <= dataGridView2.Rows.Count - 1; u++)
                    {
                        string naim = Convert.ToString(dataGridView2.Rows[u].Cells[1].Value);
                        string kolvo = Convert.ToString(dataGridView2.Rows[u].Cells[4].Value);
                        //string size = Convert.ToString(dataGridView2.Rows[u].Cells[8].Value);
                        //string iais = Convert.ToString(dataGridView2.Rows[u].Cells[9].Value);
                        if ((kolvo.Length == 0) || (kolvo == "0"))
                        {
                            SystemSounds.Beep.Play();
                            MessageBox.Show("Не введено кол-во для отпуска в позиции \"" + naim + "\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }

                        /* if ((size.Length == 0) || (kolvo == "0"))
                         {
                             SystemSounds.Beep.Play();
                             MessageBox.Show("Не введен Размер выдаваемой позиции \"" + naim + "\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                             return;
                         }*/

                        if (Convert.ToInt32(kolvo) <= 0)
                        {
                            SystemSounds.Beep.Play();
                            MessageBox.Show("Кол-во к выдаче по позиции \"" + naim + "\" меньше или равно 0. Отпуск невозможен!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }

                        /* if (iais.Length == 0)
                         {
                             SystemSounds.Beep.Play();
                             MessageBox.Show("Не выбрано наименование из справочника ИАИС!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                             return;
                         }*/
                    }

                    /////////Выдача спецодежды (ввод данных в таблицу VIDANO)
                    for (int j = 0; j <= dataGridView2.Rows.Count - 1; j++)//Цикл только по выделенным строкам в Datagrid, для того чтобы выдачу можно было производить выборочно
                    {
                        string id_position_coverall = Convert.ToString(dataGridView2.Rows[j].Cells[0].Value);
                        string naimen = Convert.ToString(dataGridView2.Rows[j].Cells[1].Value);
                        string charact = Convert.ToString(dataGridView2.Rows[j].Cells[2].Value);
                        string edizm = Convert.ToString(dataGridView2.Rows[j].Cells[3].Value);
                        string kol_vo = Convert.ToString(dataGridView2.Rows[j].Cells[4].Value);
                        string per_isp = Convert.ToString(dataGridView2.Rows[j].Cells[5].Value);
                        string date_posled_vidachi = Convert.ToString(dataGridView2.Rows[j].Cells[6].Value);
                        string date_sled_vidachi = Convert.ToString(dataGridView2.Rows[j].Cells[7].Value);

                        //Проверка кол-ва выдаваемой спецодежды на превышение нормы
                        SqlCommand command = new SqlCommand("select norms_list.* from norms_personal,norms_list where norms_personal.norms_id=norms_list.norms_id and norms_personal.sotrudnik_id=" + Form1.sotrudn_id + " and norms_list.id=" + id_position_coverall, conn);
                        da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb = new SqlCommandBuilder(da);
                        DataSet ds = new DataSet();
                        //Заполнение DataGridView наименованиями полей 
                        da.Fill(ds, "NORMS_LIST");
                        if (Convert.ToInt16(ds.Tables[0].Rows[0][5]) < Convert.ToInt16(kol_vo))
                        {
                            SystemSounds.Beep.Play();
                            MessageBox.Show("Превышена норма по " + naimen + ". Кол-во по норме " + Convert.ToString(ds.Tables[0].Rows[0][5]) + ", а вы указали " + kol_vo + ".", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            return;
                        }
                        ///////////////////////////////////////////////////////////


                        //Проверка кол-ва на уже выданную спецодежду
                        SqlCommand command1 = new SqlCommand("select SUM(kolvo) from vidano where norms_list_id=" + id_position_coverall + " and SOTRUDNIK_ID=" + Form1.sotrudn_id, conn);
                        SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                        SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                        DataSet ds1 = new DataSet();
                        //conn.Close();
                        //Заполнение DataGridView наименованиями полей 
                        da1.Fill(ds1, "VIDANO");
                        if ((ds1.Tables[0].Rows.Count != 0) && (ds1.Tables[0].Rows[0][0] != DBNull.Value))
                        {
                            if ((Convert.ToInt32(ds1.Tables[0].Rows[0][0]) + Convert.ToInt32(kol_vo)) > Convert.ToInt32(ds.Tables[0].Rows[0][5]))
                            {
                                SystemSounds.Beep.Play();
                                MessageBox.Show("Сумма кол-ва уже выданной номенклатуры \"" + naimen + "\" и указанное кол-во к выдаче превышает норму. Кол-во по норме=" + Convert.ToString(ds.Tables[0].Rows[0][5]) + ". Уже выдано " + Convert.ToInt32(ds1.Tables[0].Rows[0][0]) + ", а вы указываете " + kol_vo + ". В сумме это " + (Convert.ToInt32(ds1.Tables[0].Rows[0][0]) + Convert.ToInt32(kol_vo)), "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                return;
                            }
                        }
                        ///////////////////////////////////////////////////////////

                    }

                    ////Непосредственно цикл по вставке данных в бд (непосредственно выдача)
                    for (int k = 0; k <= dataGridView2.Rows.Count - 1; k++)
                    {
                        string id_position_coverall1 = Convert.ToString(dataGridView2.Rows[k].Cells[0].Value);
                        string naimen1 = Convert.ToString(dataGridView2.Rows[k].Cells[1].Value);
                        string charact1 = Convert.ToString(dataGridView2.Rows[k].Cells[2].Value);
                        string edizm1 = Convert.ToString(dataGridView2.Rows[k].Cells[3].Value);
                        string kol_vo1 = Convert.ToString(dataGridView2.Rows[k].Cells[4].Value);
                        string per_isp1 = Convert.ToString(dataGridView2.Rows[k].Cells[5].Value);
                        string date_posled_vidachi1 = Convert.ToString(dataGridView2.Rows[k].Cells[6].Value);
                        string date_sled_vidachi1 = Convert.ToString(dataGridView2.Rows[k].Cells[7].Value);
                        //string size1 = Convert.ToString(dataGridView2.Rows[k].Cells[8].Value);
                        //string iais_naimen = Convert.ToString(dataGridView2.Rows[k].Cells[9].Value);
                        //string iais_id = Convert.ToString(dataGridView2.Rows[k].Cells[10].Value);

                        //Непосредственно вставка данных в таблицу VIDANO, т.е. выдача спецодежды
                        SqlCommand scmd = conn.CreateCommand();
                        scmd.CommandText = "insert into VIDANO (SOTRUDNIK_ID,NORMS_LIST_ID,NAIMEN_OVERALLS,ED_IZM,KOLVO,PERIOD_ISPOLZOVAN,DATE_VIDACHI,DATE_SLED_VIDACHI,USER_ID,DATETIME_CREATE,CARD_UIN) VALUES ('" + Form1.sotrudn_id + "', '" + id_position_coverall1 + "', '" + naimen1 + "', '" + edizm1 + "', '" + kol_vo1 + "', '" + per_isp1 + "', convert(datetime,'" + date_posled_vidachi1 + "', 103), convert(datetime,'" + date_sled_vidachi1 + "',103), '" + Form2.val + "', convert(datetime,'" + DateTime.Now + "',103), '" + textBox1.Text + "')";
                        SqlDataReader reader;
                        reader = scmd.ExecuteReader();
                        reader.Dispose();
                    }
                    //////////////////////////////////////////////


                    //Получение GUID для 
                    SqlCommand command2 = new SqlCommand("select NEWID()", conn);
                    SqlDataAdapter da2 = new SqlDataAdapter(command2);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                    DataSet ds2 = new DataSet();
                    //Заполнение DataGridView наименованиями полей 
                    da2.Fill(ds2, "NEWID");
                    string GUID = ds2.Tables[0].Rows[0][0].ToString();

                    //Создание накладной на выдачу в БД                              
                    SqlCommand scmd1 = conn.CreateCommand();
                    scmd1.CommandText = "insert into NAKLADN_VIDANO (GUID, SOTRUDNIK_ID, USER_ID, DATETIME_CREATE) VALUES (@G, @SOTR_ID, @US_ID, GETDATE())";
                    scmd1.Parameters.AddWithValue("G", GUID);
                    scmd1.Parameters.AddWithValue("SOTR_ID", Form1.sotrudn_id);
                    scmd1.Parameters.AddWithValue("US_ID", Form2.val);
                    SqlDataReader reader1;
                    reader1 = scmd1.ExecuteReader();
                    reader1.Dispose();

                    ////Цикл по вставке позиций по накладной в таблицу
                    for (int k1 = 0; k1 <= dataGridView2.Rows.Count - 1; k1++)
                    {
                        string id_position_coverall1 = Convert.ToString(dataGridView2.Rows[k1].Cells[0].Value);
                        string naimen1 = Convert.ToString(dataGridView2.Rows[k1].Cells[1].Value);
                        string charact1 = Convert.ToString(dataGridView2.Rows[k1].Cells[2].Value);
                        string edizm1 = Convert.ToString(dataGridView2.Rows[k1].Cells[3].Value);
                        string kol_vo1 = Convert.ToString(dataGridView2.Rows[k1].Cells[4].Value);
                        string per_isp1 = Convert.ToString(dataGridView2.Rows[k1].Cells[5].Value);
                        string date_posled_vidachi1 = Convert.ToString(dataGridView2.Rows[k1].Cells[6].Value);
                        string date_sled_vidachi1 = Convert.ToString(dataGridView2.Rows[k1].Cells[7].Value);

                        //Непосредственно вставка данных в таблицу VIDANO, т.е. выдача спецодежды
                        SqlCommand scmd3 = conn.CreateCommand();
                        scmd3.CommandText = "insert into NAKLADN_VIDANO_DETAIL (GUID_NAKLADN_VIDANO,NORMS_LIST_ID,NAIMEN_OVERALLS,ED_IZM,KOLVO,PERIOD_ISPOLZOVAN,DATE_VIDACHI,DATE_SLED_VIDACHI,DATETIME_CREATE,CARD_UIN) VALUES (@G,@NR, @NAIM, @ED, @KOL, @PER_ISP, @DT_VIDACH, @DT_SLED, GETDATE(),@UIN)";
                        scmd3.Parameters.AddWithValue("G", GUID);
                        scmd3.Parameters.AddWithValue("NR", id_position_coverall1);
                        scmd3.Parameters.AddWithValue("NAIM", naimen1);
                        scmd3.Parameters.AddWithValue("ED", edizm1);
                        scmd3.Parameters.AddWithValue("KOL", kol_vo1);
                        scmd3.Parameters.AddWithValue("PER_ISP", per_isp1);
                        scmd3.Parameters.AddWithValue("DT_VIDACH", Convert.ToDateTime(date_posled_vidachi1));
                        scmd3.Parameters.AddWithValue("DT_SLED", Convert.ToDateTime(date_sled_vidachi1));
                        scmd3.Parameters.AddWithValue("UIN", textBox1.Text);
                        SqlDataReader reader3;
                        reader3 = scmd3.ExecuteReader();
                        reader3.Dispose();
                    }
                    //////////////////////////////////////////////


                    conn.Close();

                    SystemSounds.Beep.Play();
                    MessageBox.Show("Выдача спецодежды для " + Form1.sotrudn_fio + " произведена!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dataGridView2.Rows.Clear();
                }
                else
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Данному сотруднику карта не присвоена! Присвойте карту и повторите выдачу.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;                    
                }

            }
          
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите отменить выдачу?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView2.Rows.Clear();
                this.Close();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Form9 form9 = new Form9(this);
            form9.ShowDialog();
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                dataGridView2.Rows.Remove(dataGridView2.CurrentRow);
            }
        }

        private void button4_Click(object sender, EventArgs e)
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

        private void Form7_FormClosing(object sender, FormClosingEventArgs e)
        {
               dataGridView2.Rows.Clear();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Form10 form10 = new Form10();
            form10.ShowDialog();   
        }

       

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //Определение срока использования носки, если он меньше месяца, например 0.7, то используется функция AddDays, если более месяца (целое число), то AddMonth
            string srok = dataGridView2.CurrentRow.Cells[5].Value.ToString();
            string decimal_sep = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator;
            srok = srok.Replace(".", decimal_sep);
            srok = srok.Replace(",", decimal_sep);
            double srok_p = Convert.ToDouble(srok);

            if (srok_p < 1)
            {
                double rs = srok_p * Convert.ToDouble(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                date_sled_vidachi = Convert.ToString(DateTime.Now.Date.AddDays(rs));//Вычисление даты следующей выдачи спецодежы            
            }
            else
            {
                date_sled_vidachi = Convert.ToString(DateTime.Now.Date.AddMonths(Convert.ToInt16(dataGridView2.CurrentRow.Cells[5].Value)));//Вычисление даты следующей выдачи спецодежы
            }
            ////////////////////////////////////////

            dataGridView2.CurrentRow.Cells[7].Value = Convert.ToDateTime(date_sled_vidachi);  //Вычисление даты следующей выдачи спецодежы            
            textBox1.Focus();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //Form11 form11 = new Form11();
            //form11.ShowDialog(); 
        }

        private void dataGridView2_CellMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (dataGridView2.CurrentRow.Cells[6].Selected)
            {
                Form11 form11 = new Form11(this);
                form11.ShowDialog(); 
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            textBox1.ReadOnly = true;
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 5)
            {                
                timer1.Enabled = true;
            }
        }

        private void dataGridView2_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            textBox1.Enabled = true;
        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            textBox1.ReadOnly = false;
            textBox1.Clear();
            textBox1.Focus();
        }

       

        

      
    }
}

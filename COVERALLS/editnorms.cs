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
    public partial class editnorms : Form
    {

        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        SqlDataAdapter da;
        SqlDataAdapter da1;
        
        public editnorms()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["PROFESSION"].HeaderText = "Профессия/должность";
            dataGridView1.Columns["PROFESSION"].Width = 500;
            dataGridView1.Columns["KEY_N"].HeaderText = "KEY_N";
            dataGridView1.Columns["KEY_N"].Width = 20;
            dataGridView1.Columns["KEY_N"].Visible = false;

        }
        //////////////////////////////////////////////////////

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview1()
        {
            dataGridView2.Columns["ID"].HeaderText = "ID";
            dataGridView2.Columns["ID"].Width = 40;
            dataGridView2.Columns["ID"].Visible = false;
            dataGridView2.Columns["NORMS_ID"].HeaderText = "NORMS_ID";
            dataGridView2.Columns["NORMS_ID"].Width = 40;
            dataGridView2.Columns["NORMS_ID"].Visible = false;
            dataGridView2.Columns["NAIMEN_OVERALLS"].HeaderText = "Наименование";
            dataGridView2.Columns["NAIMEN_OVERALLS"].Width = 200;
            dataGridView2.Columns["CHARACTERISTIC_OVERALLS"].HeaderText = "Характеристики";
            dataGridView2.Columns["CHARACTERISTIC_OVERALLS"].Width = 200;
            dataGridView2.Columns["ED_IZM"].HeaderText = "Ед. изм.";
            dataGridView2.Columns["ED_IZM"].Width = 50;
            dataGridView2.Columns["KOLVO"].HeaderText = "Кол-во";
            dataGridView2.Columns["KOLVO"].Width = 50;
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].HeaderText = "Период испол-я";
            dataGridView2.Columns["PERIOD_ISPOLZOVAN"].Width = 100;
            dataGridView2.Columns["PRICE_EDENIC_PLAN"].HeaderText = "Стоим. за ед. план.";
            dataGridView2.Columns["PRICE_EDENIC_PLAN"].Width = 80;
            dataGridView2.Columns["PRICE_KOMPLEKT_PLAN"].HeaderText = "Стоим. за комплект план.";
            dataGridView2.Columns["PRICE_KOMPLEKT_PLAN"].Width = 80;
            dataGridView2.Columns["IAIS_NAIMEN_OVERALLS"].HeaderText = "ИАИС Наименование";
            dataGridView2.Columns["IAIS_NAIMEN_OVERALLS"].Width = 230;
            dataGridView2.Columns["CODE_GROUP_IAIS"].HeaderText = "CODE_GROUP_IAIS";
            dataGridView2.Columns["CODE_GROUP_IAIS"].Width = 40;
            dataGridView2.Columns["CODE_GROUP_IAIS"].Visible = false;
            dataGridView2.Columns["KONTRAGENT_ID"].HeaderText = "KONTRAGENT_ID";
            dataGridView2.Columns["KONTRAGENT_ID"].Width = 40;
            dataGridView2.Columns["KONTRAGENT_ID"].Visible = false;
            dataGridView2.Columns["CHANGED"].HeaderText = "CHANGED";
            dataGridView2.Columns["CHANGED"].Width = 40;
            dataGridView2.Columns["CHANGED"].Visible = false;
        }
        //////////////////////////////////////////////////////

        private void editnorms_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую
            dataGridView2.AllowUserToAddRows = false;//Запретить пользователю добавлять строки из грида напрямую

            SqlCommand command = new SqlCommand("select * from NORMS order by PROFESSION", conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS");
            dataGridView1.DataSource = ds.Tables[0];
                        
            fill_gridview();
        }

              

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow.Cells[0].Value != DBNull.Value)
            {
                SqlCommand command = new SqlCommand("select * from NORMS_LIST where NORMS_ID=" + dataGridView1.CurrentRow.Cells[0].Value + " order by NAIMEN_OVERALLS", conn);
                da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da1);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da1.Fill(ds, "NORMS_LIST");
                dataGridView2.DataSource = ds.Tables[0];

                fill_gridview1();
            }
            else
            {
                dataGridView2.DataSource = null;
                dataGridView2.Rows.Clear();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            button1.Text = "Включен режим редактирования";
            button2.Visible = true;
            button3.Visible = true;
            dataGridView2.ReadOnly = false;
            dataGridView2.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
            button1.Enabled = false;

        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                da1.Update((System.Data.DataTable)dataGridView2.DataSource);
                dataGridView2.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
                button1.Enabled = true;
                button1.Text = "Редактирование";
                button2.Visible = false;

                //Установка статуса Изменено в 1
                SqlCommand scmd4 = conn.CreateCommand();
                scmd4.CommandText = "update NORMS_LIST set CHANGED=1 where ID=@idd";
                scmd4.Parameters.AddWithValue("idd", dataGridView2.CurrentRow.Cells[0].Value);
                try
                {
                    conn.Open();
                }
                catch
                {
                    MessageBox.Show("Ошибка соединения с базой данных");
                }
                SqlDataReader reader4;
                reader4 = scmd4.ExecuteReader();
                conn.Close();
                ////////////////////////

                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в нормах выполнены.", "Уведомление о результатах", MessageBoxButtons.OK);

            }
            catch (Exception)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в базе данных выполнить не удалось!", "Уведомление о результатах", MessageBoxButtons.OK);
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView2.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
            dataGridView2.ReadOnly = true;
            button1.Enabled = true;
            button1.Text = "Редактирование";
            button2.Visible = false;
            button3.Visible = false;        
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Height = 794;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Наименование спецодежды, спецобуви и других СИЗ\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////
            if (textBox2.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Наименование спецодежды, спецобуви и других СИЗ из ИАИС\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////
            if (textBox3.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Характеристика спецодежды, спецобуви и других СИЗ\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////
            if (comboBox1.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Ед. изм.\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////
            if (textBox5.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Кол-во\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////
            if (textBox6.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Срок испол-я\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////
            if (textBox7.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Цена плановая за единицу\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////
            if (textBox8.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Не введены данные в поле \"Цена плановая за нормокомплект\"", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            ////////////

            textBox7.Text.Replace(",", ".");
            textBox8.Text.Replace(",", ".");
                        
            /////Добавление данных в таблицу NORMS_LIST
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "INSERT into NORMS_LIST (NORMS_ID,NAIMEN_OVERALLS,CHARACTERISTIC_OVERALLS,ED_IZM,KOLVO,PERIOD_ISPOLZOVAN,PRICE_EDENIC_PLAN,PRICE_KOMPLEKT_PLAN,IAIS_NAIMEN_OVERALLS) VALUES (" + "'" + dataGridView1.CurrentRow.Cells[0].Value + "', '" + textBox1.Text + "', '" + textBox3.Text + "', '"+comboBox1.Text+"', '"+textBox5.Text+"', '"+textBox6.Text+"', '"+textBox7.Text+"', '"+textBox8.Text+"', '"+textBox2.Text+"')";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////

            //Обновление данных в гриде после добавления записи
            SqlCommand command = new SqlCommand("select * from NORMS_LIST where NORMS_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
            da1 = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da1);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds, "NORMS_LIST");
            dataGridView2.DataSource = ds.Tables[0];
            fill_gridview1();
            /////////////////////////

            this.Height = 694;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            comboBox1.Text = "";
            this.Height = 694;        
        }

        private void удалитьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить текущую позицию?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                conn.Open();
                SqlCommand mycommand = new SqlCommand("select * from VIDANO where NORMS_LIST_ID='" + dataGridView2.CurrentRow.Cells[0].Value +"'", conn);
                SqlDataAdapter da = new SqlDataAdapter(mycommand);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                da.Fill(ds, "VIDANO");

                if (ds.Tables[0].Rows.Count > 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Невозможно удалить данную позицию/СИЗ, т.к. она участвовала к выдаче.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                else
                {
                    SystemSounds.Beep.Play();
                    dataGridView2.Rows.Remove(dataGridView2.CurrentRow);
                }
            }
        }

        private void textBox6_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8))
            {
                e.Handled = true;
            }
        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!Char.IsDigit(e.KeyChar) && e.KeyChar != Convert.ToChar(8) && e.KeyChar != '.')
            {
                e.Handled = true;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (textBox4.Text.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Вы не ввели наименование нормы!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }


            //Проверка существует ли уже вводимая норма в БД
            SqlCommand command = new SqlCommand("select * from NORMS where PROFESSION='" + textBox4.Text+"'", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS");
            if (ds.Tables[0].Rows.Count != 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Данное наименование нормы уже присутствует в базе!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }
            /////////////

            /////Добавление данных в таблицу NORMS
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "declare @b as int " +
                                "SET @b=(select max(ID) from NORMS) " +
                                "Set @b=@b+1 " +
                                "INSERT into NORMS (PROFESSION, KEY_N) VALUES (" + "'" + textBox4.Text + "', @b)";
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////
        }

        private void button8_Click(object sender, EventArgs e)
        {
            button8.Text = "Включен режим редактирования";
            button9.Visible = true;
            button10.Visible = true;
            dataGridView1.ReadOnly = false;
            dataGridView1.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
            button8.Enabled = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
            dataGridView1.ReadOnly = true;
            button8.Enabled = true;
            button8.Text = "Редактирование";
            button9.Visible = false;
            button10.Visible = false; 
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                da.Update((System.Data.DataTable)dataGridView1.DataSource);
                dataGridView1.AllowUserToAddRows = false;//Запретить/разрешить пользователю добавлять строки из грида напрямую
                button8.Enabled = true;
                button8.Text = "Редактирование";
                button9.Visible = false;
                button10.Visible = false;                
                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в нормах выполнены.", "Уведомление о результатах", MessageBoxButtons.OK);

            }
            catch (Exception)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Изменения в базе данных выполнить не удалось!", "Уведомление о результатах", MessageBoxButtons.OK);
            }
        }

        private void скопироватьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите создать копию данной нормы?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                //Сохранение в ds Наименования нормы
                SqlCommand command = new SqlCommand("select * from NORMS where ID='" + dataGridView1.CurrentRow.Cells[0].Value+ "'", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NORMS");

                //Сохранение в ds1 содержания нормы
                SqlCommand command1 = new SqlCommand("select * from NORMS_LIST where NORMS_ID='" + dataGridView1.CurrentRow.Cells[0].Value + "'", conn);
                SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                DataSet ds1 = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da1.Fill(ds1, "NORMS_LIST");

                /////Добавление данных в таблицу NORMS
                SqlCommand scmd4 = conn.CreateCommand();
                scmd4.CommandText = "INSERT into NORMS (PROFESSION) VALUES (" + "'" + ds.Tables[0].Rows[0][1] + "')";
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader4;
                reader4 = scmd4.ExecuteReader();
                conn.Close();
                //////////////////

                ///Добавление наполнения в введенной норме
                for (int k = 0; k <= ds1.Tables[0].Rows.Count - 1; k++)
                {
                    string naim = ds1.Tables[0].Rows[k][2].ToString();
                    string charact = ds1.Tables[0].Rows[k][3].ToString();
                    string edizm = ds1.Tables[0].Rows[k][4].ToString();
                    string kolvo = ds1.Tables[0].Rows[k][5].ToString();
                    string per_isp = ds1.Tables[0].Rows[k][6].ToString();
                    string price_plan = ds1.Tables[0].Rows[k][7].ToString().Replace(",", ".");
                    string price_fact = ds1.Tables[0].Rows[k][8].ToString().Replace(",", ".");
                    string iais = ds1.Tables[0].Rows[k][9].ToString();
                    string iais_code = ds1.Tables[0].Rows[k][10].ToString();
                    
                    

                    /////Добавление данных в таблицу NORMS_LIST
                    SqlCommand scmd = conn.CreateCommand();
                    scmd.CommandText = "INSERT into NORMS_LIST (NORMS_ID,NAIMEN_OVERALLS,CHARACTERISTIC_OVERALLS,ED_IZM,KOLVO,PERIOD_ISPOLZOVAN,PRICE_EDENIC_PLAN,PRICE_KOMPLEKT_PLAN,IAIS_NAIMEN_OVERALLS,CODE_GROUP_IAIS) VALUES ((select max(ID) from NORMS), '" + naim + "', '" + charact + "', '" + edizm + "', '" + kolvo + "', '" + per_isp + "', '" + price_plan + "', '" + price_fact + "', '" + iais + "', '" + iais_code + "')";
                    try
                    {
                        conn.Open();
                    }
                    catch {}
                    SqlDataReader reader;
                    reader = scmd.ExecuteReader();
                    conn.Close();
                    //////////////////
                                        
                }

                //Обновление данных в гриде после добавления записи
                SqlCommand command2 = new SqlCommand("select * from NORMS order by PROFESSION", conn);
                SqlDataAdapter da2 = new SqlDataAdapter(command2);//Переменная объявлена как глобальная
                SqlCommandBuilder cb2 = new SqlCommandBuilder(da2);
                DataSet ds2 = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da2.Fill(ds2, "NORMS");
                dataGridView1.DataSource = ds2.Tables[0];
                fill_gridview();
                /////////////////////////


            }
        }

        private void удалитьНормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите удалить данную норму?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                /////Удаление заголовка нормы, само же удаление наполнения нормы выполнятеся триггером
                SqlCommand scmd = conn.CreateCommand();
                scmd.CommandText = "delete from NORMS where ID="+dataGridView1.CurrentRow.Cells[0].Value;
                try
                {
                    conn.Open();
                }
                catch { }
                SqlDataReader reader;
                reader = scmd.ExecuteReader();
                conn.Close();
                //////////////////

                //Обновление данных в гриде после добавления записи
                SqlCommand command3 = new SqlCommand("select * from NORMS order by PROFESSION", conn);
                SqlDataAdapter da3 = new SqlDataAdapter(command3);//Переменная объявлена как глобальная
                SqlCommandBuilder cb3 = new SqlCommandBuilder(da3);
                DataSet ds3 = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da3.Fill(ds3, "NORMS");
                dataGridView1.DataSource = ds3.Tables[0];
                fill_gridview();
                /////////////////////////
            }
        }

        
    }
}

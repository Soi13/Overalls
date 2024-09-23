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
using System.Reflection;


namespace COVERALLS
{
    public partial class Form1 : Form
    {
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        SqlDataAdapter da;
        DataSet ds;
        public string NORMS_LIST_ID;

        Form2 form2=new Form2();
        Form3 form3 = new Form3();
        Form4 form4 = new Form4();
        Form5 form5 = new Form5();
        Form6 form6 = new Form6();
        Form7 form7 = new Form7();
        Form8 form8 = new Form8();
        public static int sotrudn_id;
        public static int sotrudn_id_4_need_siz;
        public static string sotrudn_fio_4_need_siz;
        public static string code_need_siz;
        public static string sotrudn_fio;
        public static int i;

        public Form1()
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
            dataGridView1.Columns["SPISOK_ODEJD"].Width = 150;
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
            dataGridView1.Columns["SIZE_CLOTHES"].HeaderText = "Размер одежды";
            dataGridView1.Columns["SIZE_CLOTHES"].Width = 70;
            dataGridView1.Columns["HEIGHT"].HeaderText = "Рост";
            dataGridView1.Columns["HEIGHT"].Width = 40;
            dataGridView1.Columns["SEX"].HeaderText = "Пол";
            dataGridView1.Columns["SEX"].Width = 40;
            dataGridView1.Columns["SEX"].Visible = false;        
        }
        //////////////////////////////////////////////////////

          
        private void Form1_Load(object sender, EventArgs e)
        {
            //Отображение раздела "Администрирование" взависимости от прав пользователя.
            if (Form2.administration == "0") 
               {
                //menuStrip1.Items[5].Visible = false;
                администрированиеToolStripMenuItem.Visible = false;
                привязатьНормуToolStripMenuItem.Enabled = false;
                удалитьПриявязаннуюНормуToolStripMenuItem.Enabled = false;
                необходимостьВСИЗToolStripMenuItem.Enabled = false;
                лимитыToolStripMenuItem.Enabled = false;
                каталогКАToolStripMenuItem.Enabled = false;
                бригадыToolStripMenuItem.Enabled = false;
                складскиеОстаткиToolStripMenuItem.Enabled = false;
                сформироватьПотребностьToolStripMenuItem.Enabled = false;
                загрузкаИтоговПослеТорговToolStripMenuItem.Enabled = false;
                }

            if (Form2.administration == "1")
               {
                //menuStrip1.Items[5].Visible = true;
                администрированиеToolStripMenuItem.Visible = true;
                привязатьНормуToolStripMenuItem.Enabled=true;
                удалитьПриявязаннуюНормуToolStripMenuItem.Enabled = true;
                необходимостьВСИЗToolStripMenuItem.Enabled = true;
                лимитыToolStripMenuItem.Enabled = true;
                каталогКАToolStripMenuItem.Enabled = true;
                бригадыToolStripMenuItem.Enabled = true;
                складскиеОстаткиToolStripMenuItem.Enabled = true;
                сформироватьПотребностьToolStripMenuItem.Enabled = true;
                загрузкаИтоговПослеТорговToolStripMenuItem.Enabled = true;
               }
            /////

            dataGridView1.AllowUserToAddRows = true;
            label12.Text = Form2.name_user;
            this.Text = "СИЗ - " + Assembly.GetExecutingAssembly().GetName().Version;

            if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ или Шепилева
            {
                каталогКАToolStripMenuItem.Visible = true;

                SqlCommand command = new SqlCommand("select * from SOTRUDNIKI", conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();

                //Проверяем были ли изменения в нормах, если были, то выводим список где изменения произошли
                SqlCommand command3_ = new SqlCommand("select changed from norms_list where changed=1", conn);
                SqlDataAdapter da3_ = new SqlDataAdapter(command3_);//Переменная объявлена как глобальная
                SqlCommandBuilder cb3_ = new SqlCommandBuilder(da3_);
                DataSet ds3_ = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da3_.Fill(ds3_, "NORMS_LIST");

                if (ds3_.Tables[0].Rows.Count > 0)
                {
                    alert_norms_changed alert_norms_changed = new alert_norms_changed();
                    alert_norms_changed.ShowDialog();
                }

            }
            else
            {
                каталогКАToolStripMenuItem.Visible = false;
                
                SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where USERS_BRANCH_ID=" + Form2.branch, conn);
                da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "SOTRUDNIKI");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }

            //Заполнение поля профессия, данными из БД
            SqlCommand command2 = conn.CreateCommand();
            command2.CommandText = "select distinct PROFESSION from SOTRUDNIKI order by PROFESSION";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader1;
            reader1 = command2.ExecuteReader();
            while (reader1.Read())
            {
                try
                {
                    string result1 = reader1.GetString(0);
                    comboBox4.Items.Add(result1);
                }
                catch { }

            }
            conn.Close();
            ////////////////////////

            //Заполнение поля Отдел, данными из БД
            SqlCommand command3 = conn.CreateCommand();
            command3.CommandText = "select distinct DEPARTMENT from SOTRUDNIKI order by DEPARTMENT";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader2;
            reader2 = command3.ExecuteReader();
            while (reader2.Read())
            {
                try
                {
                    string result2 = reader2.GetString(0);
                    comboBox6.Items.Add(result2);
                }
                catch { }

            }
            conn.Close();
            ////////////////////////

            /////Выделение строк цветом в DataGridview если у сотрудника покаким-то отпущеным позициям спецодежды подходит дата следующей выдачи
            try
            {
                conn.Open();
            }
            catch { }

            //Проверка дат у выданных позиций с просроченным сроком выдачи и пометка их красным цветом
            SqlCommand command4 = new SqlCommand("exec alarm_elapsed_SIZ_all_sotrudniki", conn);
            da = new SqlDataAdapter(command4);//Переменная объявлена как глобальная
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da);
            DataSet ds3 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds3, "VIDANO");
            for (int a = 0; a <= ds3.Tables[0].Rows.Count - 1; a++)
            {
                //string idd = Convert.ToString(ds3.Tables[0].Rows[a][0]);
                
                    for (int s = 0; s <= dataGridView1.Rows.Count - 1; s++)
                    {
                        if (Convert.ToInt16(dataGridView1.Rows[s].Cells[0].Value) == Convert.ToInt16(ds3.Tables[0].Rows[a][0]))
                        {
                            dataGridView1[1, s].Style.BackColor = Color.Red;
                            dataGridView1[2, s].Style.BackColor = Color.Red;
                            dataGridView1[3, s].Style.BackColor = Color.Red;
                            dataGridView1[4, s].Style.BackColor = Color.Red;
                            dataGridView1[5, s].Style.BackColor = Color.Red;

                        }
                    }
                
            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            //Проверка дат у выданных позиций и подстветка желтым за 5 дней до окончания использования
            SqlCommand command5 = new SqlCommand("exec alarm_before_5days_all_sotrudniki", conn);
            SqlDataAdapter da5 = new SqlDataAdapter(command5);//Переменная объявлена как глобальная
            SqlCommandBuilder cb5 = new SqlCommandBuilder(da5);
            DataSet ds5 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da5.Fill(ds5, "VIDANO");
            for (int a1 = 0; a1 <= ds5.Tables[0].Rows.Count - 1; a1++)
            {
                //string idd1 = Convert.ToString(ds5.Tables[0].Rows[a1][0]);

                for (int s1 = 0; s1 <= dataGridView1.Rows.Count - 1; s1++)
                {
                    if (Convert.ToInt16(dataGridView1.Rows[s1].Cells[0].Value) == Convert.ToInt16(ds5.Tables[0].Rows[a1][0]))
                    {
                        dataGridView1[1, s1].Style.BackColor = Color.Yellow;
                        dataGridView1[2, s1].Style.BackColor = Color.Yellow;
                        dataGridView1[3, s1].Style.BackColor = Color.Yellow;
                        dataGridView1[4, s1].Style.BackColor = Color.Yellow;
                        dataGridView1[5, s1].Style.BackColor = Color.Yellow;

                    }
                }

            }
            ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            
            /////Вставка данных в таблицу журнала вход/выход
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "INSERT into JOURNAL (USER_ID,USER_FULL_NAME,EVENT_DATETIME,EVENT_STATUS,MACHINE_NAME,SYSTEM_NAME) VALUES (" + "'" + Form2.val + "', (select FULL_NAME from USERS where ID=" + Form2.val + "), convert(datetime,'" + DateTime.Now.ToString() + "', 103),'Вход','"+Environment.MachineName+"','"+Environment.UserName+"')";
            try
            {
                conn.Open();
            }
            catch {}
            SqlDataReader reader4;
            reader4 = scmd4.ExecuteReader();
            conn.Close();
            //////////////////

            /*
            /////Вставка данных в таблицу ACTIVE_USERS для определения того, что он залогинился в систему
            SqlCommand scmd5 = conn.CreateCommand();
            scmd5.CommandText = "INSERT into ACTIVE_USERS (ID_USER, DATETIME_LOGIN) VALUES (" + "'" + Form2.val + "', '" + DateTime.Now.ToString() + "')";
            try
            {
                conn.Open();
            }
            catch {}
            SqlDataReader reader5;
            reader5 = scmd5.ExecuteReader();
            conn.Close();
            //////////////////   */

            if (Form2.psw == "6ece4fd51bc113942692637d9d4b860e")
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Вам рекомендуется сменить пароль. Это можно сделать в меню \"Инструменты\"", "Внимание", MessageBoxButtons.OK);
            }
        }

      
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        
        private void показатьПривязанныйНормативToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sotrudn_id = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
            sotrudn_fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);

            //Проверка существует ли уже привязанная норма для показа
            SqlCommand command = new SqlCommand("select * from NORMS_PERSONAL where SOTRUDNIK_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS_PERSONAL");
            if (ds.Tables[0].Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("К данному сотруднику не привязано ни одной нормы! Отображать нечего.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            /////////////

            form4.ShowDialog(); 
        }

        private void удалитьПривязаннуюНормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что хотите выйти из программы?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private void сменитьПарольToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form5.ShowDialog();
        }

        private void письмоРазработчикуПОToolStripMenuItem_Click(object sender, EventArgs e)
        {
            form6.ShowDialog();
        }

      

        private void просмотрВыданнойСпецодеждыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sotrudn_id = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
            sotrudn_fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);

            //Проверка есть ли у человека уже выданная спецодежда для ее отображения
            SqlCommand command = new SqlCommand("select * from vidano where SOTRUDNIK_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
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

            form8.ShowDialog();
        }

        private void привязатьНормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sotrudn_id = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
            sotrudn_fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);

            //Проверка существует ли уже привязанная норма, перед привязкой
            SqlCommand command = new SqlCommand("select * from NORMS_PERSONAL where SOTRUDNIK_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS_PERSONAL");
            if (ds.Tables[0].Rows.Count > 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("К данному сотруднику уже привязана норма! Если желаете привязать другую, то сначала удалите существующую!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            /////////////

            form3.ShowDialog();
        }

        private void удалитьПриявязаннуюНормуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            if (MessageBox.Show("Вы уверены, что хотите удалить привязанный норматив?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
            {

                /////////Удаление нормы у сотрудника
                SqlCommand scmd = conn.CreateCommand();
                scmd.CommandText = "delete from NORMS_PERSONAL where SOTRUDNIK_ID=" + dataGridView1.CurrentRow.Cells[0].Value;
                try
                {
                    conn.Open();
                }
                catch
                {
                    MessageBox.Show("Ошибка соединения с базой данных");
                }
                SqlDataReader reader;
                reader = scmd.ExecuteReader();
                conn.Close();
                //////////////////

                SystemSounds.Beep.Play();
                MessageBox.Show("У сотрудника " + dataGridView1.CurrentRow.Cells[1].Value + " норма удалена успешно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void выдачаСпецодеждыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sotrudn_id = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
            sotrudn_fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);

            //Проверка существует ли уже привязанная норма для выдачи спецодежды
            SqlCommand command = new SqlCommand("select * from NORMS_PERSONAL where SOTRUDNIK_ID=" + dataGridView1.CurrentRow.Cells[0].Value, conn);
            da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NORMS_PERSONAL");
            if (ds.Tables[0].Rows.Count == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Невозможно произвести выдачу, т.к. к данному сотруднику не привязана норма. Привяжите норму.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            /////////////
                     
            form7.ShowDialog();
        }

       
        private void просмотрОстатковПоНормеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sotrudn_id = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
            sotrudn_fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);

            Form10 form10 = new Form10();
            form10.ShowDialog();  
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //Фильтр по Профессии
            if (checkBox1.Checked == true)
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "' and USERS_BRANCH_ID="+Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////

            //Фильтр по Филиалу
            if (checkBox2.Checked == true)
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where BRANCH='" + comboBox5.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where BRANCH='" + comboBox5.Text + "' and USERS_BRANCH_ID="+Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////

            //Фильтр по Отделу
            if (checkBox3.Checked == true)
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where DEPARTMENT='" + comboBox6.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where DEPARTMENT='" + comboBox6.Text + "' and USERS_BRANCH_ID="+Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////

            //Фильтр по Профессии и филиалу
            if ((checkBox1.Checked == true) && (checkBox2.Checked == true))
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "' and BRANCH='" + comboBox5.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "' and BRANCH='"+ comboBox5.Text +"' and USERS_BRANCH_ID=" + Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////

            //Фильтр по Филиалу и отделу
            if ((checkBox2.Checked == true) && (checkBox3.Checked == true))
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where BRANCH='" + comboBox5.Text + "' and DEPARTMENT='" + comboBox6.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where BRANCH='" + comboBox5.Text + "' and DEPARTMENT='" + comboBox6.Text + "' and USERS_BRANCH_ID=" + Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////

            //Фильтр по Профессии и отделу
            if ((checkBox1.Checked == true) && (checkBox3.Checked == true))
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "' and DEPARTMENT='" + comboBox6.Text + "'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "' and DEPARTMENT='" + comboBox6.Text + "' and USERS_BRANCH_ID=" + Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////

            //Фильтр по Профессии, филиалу и отделу
            if ((checkBox1.Checked == true) && (checkBox2.Checked == true)  && (checkBox3.Checked == true))
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "' and DEPARTMENT='" + comboBox6.Text + "' and BRANCH='"+ comboBox5.Text +"'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where PROFESSION='" + comboBox4.Text + "' and DEPARTMENT='" + comboBox6.Text +"' and BRANCH='"+ comboBox5.Text + "' and USERS_BRANCH_ID=" + Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////
            
            //Снять все фильтры
            if (checkBox4.Checked == true)
            {
                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where USERS_BRANCH_ID="+Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
            }
            /////////////////////////////////////////////
        }

        
        private void checkBox4_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
        }

        private void dataGridView1_RowStateChanged(object sender, DataGridViewRowStateChangedEventArgs e)
        {
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            /////Вставка данных в таблицу журнала вход/выход
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "INSERT into JOURNAL (USER_ID,USER_FULL_NAME,EVENT_DATETIME,EVENT_STATUS,MACHINE_NAME,SYSTEM_NAME) VALUES (" + "'" + Form2.val + "', (select FULL_NAME from USERS where ID=" + Form2.val + "), convert(datetime,'" + DateTime.Now.ToString() + "', 103),'Выход','"+Environment.MachineName+"','"+Environment.UserName+"')";
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

            /*
            /////Удаление данных из таблицы ACTIVE_USERS для определения того, что пользователь вышел из системы
            SqlCommand scmd5 = conn.CreateCommand();
            scmd5.CommandText = "delete from ACTIVE_USERS where id_user=" + "'" + Form2.val + "'";
            try
            {
                conn.Open();
            }
            catch { }
            SqlDataReader reader5;
            reader5 = scmd5.ExecuteReader();
            conn.Close();
            //////////////////     */

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            if (checkBox5.Checked == true)
            {

                if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where FIO like '%" + textBox2.Text + "%'", conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }
                else  //Если не Петушков и Скворцов, то показываем исходя из того из каког филиала пользователь производит поиск
                {
                    SqlCommand command = new SqlCommand("select * from SOTRUDNIKI where FIO like '%" + textBox2.Text + "%' and USERS_BRANCH_ID="+Form2.branch, conn);
                    da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                    SqlCommandBuilder cb = new SqlCommandBuilder(da);
                    DataSet ds = new DataSet();
                    conn.Close();
                    //Заполнение DataGridView наименованиями полей 
                    da.Fill(ds, "SOTRUDNIKI");
                    dataGridView1.DataSource = ds.Tables[0];

                    statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                    fill_gridview();
                }

            }
            else
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Поиск не включен! Включите.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                textBox2.Clear();
                return;
            }
        }

        private void добавлениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            useradd useradd = new useradd();
            useradd.ShowDialog();
        }

        private void работаСНормамиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            editnorms editnorms = new editnorms();
            editnorms.ShowDialog();
        }

        private void отменаВыдачиСпецодеждыToolStripMenuItem_Click(object sender, EventArgs e)
        {            
                sotrudn_id = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
                sotrudn_fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);

                //Проверка возможности отмены выдачи спецодежды. Если отменить пытается пользовтль, которые не отпускал ее, то программа не дает это сделать    
                SqlCommand cmdd = new SqlCommand("select * from vidano where SOTRUDNIK_ID=" + sotrudn_id + " and USER_ID="+Form2.val, conn);
                da = new SqlDataAdapter(cmdd);
                SqlCommandBuilder cb1 = new SqlCommandBuilder(da);
                DataSet ds1 = new DataSet();
                conn.Close();
                da.Fill(ds1, "VIDANO");
                if (ds1.Tables[0].Rows.Count == 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Вам запрещено производить отмету выдачи спецодежды по данному сотруднику, т.к. выдача производилась другими пользователем!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }

                cancel_vidano cancel_vidano = new cancel_vidano();
                cancel_vidano.ShowDialog();            
        }

        

        private void полныйОВыданныхСИЗToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2) || (Form2.val == 1)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ
            {
                conn.Open();
                SqlCommand mycommand = new SqlCommand("select * from SOTRUDNIKI order by FIO", conn);
                da = new SqlDataAdapter(mycommand);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                conn.Close();
                da.Fill(ds, "SOTRUDNIKI");
            }
            else
            {
                conn.Open();
                SqlCommand mycommand = new SqlCommand("select * from SOTRUDNIKI where USERS_BRANCH_ID=" + Form2.branch + " order by FIO", conn);
                da = new SqlDataAdapter(mycommand);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                ds = new DataSet();
                conn.Close();
                da.Fill(ds, "SOTRUDNIKI");
            }
           
            int cnt = 6;
            int cnt1 = 5;
            int cnt_=0;

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Application.Workbooks.Add(Type.Missing);

            //делаем временно неактивным документ
            ExcelApp.Interactive = false;
            ExcelApp.EnableEvents = false;

            ExcelApp.Columns[1].ColumnWidth = 50;
            ExcelApp.Columns[2].ColumnWidth = 8.45;
            ExcelApp.Columns[3].ColumnWidth = 6.71;
            ExcelApp.Columns[4].ColumnWidth = 15;
            ExcelApp.Columns[5].ColumnWidth = 18;
            ExcelApp.Columns[6].ColumnWidth = 18;
            ExcelApp.Columns[7].ColumnWidth = 7.86;

            ExcelApp.Rows[1].Font.Bold = true; //Установка жирного шрифта на первой строчке
            ExcelApp.Rows[1].Font.Size = 20;
            ExcelApp.Rows[3].Font.Bold = true;


            ExcelApp.Cells[1, 1] = "Полный отчет о выданных СИЗ на " + DateTime.Now;
            ExcelApp.Cells[1, 1].Font.Color = Color.Blue;
            
            ExcelApp.Cells[3, 1] = "Наименование";
            ExcelApp.Cells[3, 2] = "Ед.изм.";
            ExcelApp.Cells[3, 3] = "Кол-во";
            ExcelApp.Cells[3, 4] = "Период испол-я";
            ExcelApp.Cells[3, 5] = "Дата выдачи";
            ExcelApp.Cells[3, 6] = "Дата след. выдачи";
            ExcelApp.Cells[3, 7] = "Размер";

            progressBar1.Value = 0;
            progressBar1.Maximum = ds.Tables[0].Rows.Count;
            progressBar1.Visible = true;
            label6.Visible = true;

            for (int h = 0; h <= ds.Tables[0].Rows.Count-1; h++)
            {
                int id = Convert.ToInt16(ds.Tables[0].Rows[h][0]);
                string fio = Convert.ToString(ds.Tables[0].Rows[h][1]);
                string profession = Convert.ToString(ds.Tables[0].Rows[h][2]);
                string branch = Convert.ToString(ds.Tables[0].Rows[h][3]);
                string department = Convert.ToString(ds.Tables[0].Rows[h][4]);

                //Выбор выданных СИЗ по конкретному работнику и вывод его в Excel
                conn.Open();
                SqlCommand mycommand1 = new SqlCommand("select * from VIDANO where SOTRUDNIK_ID="+id, conn);
                SqlDataAdapter da1 = new SqlDataAdapter(mycommand1);//Переменная объявлена как глобальная
                SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                DataSet ds1 = new DataSet();
                conn.Close();
                da1.Fill(ds1, "VIDANO");


                ExcelApp.Cells[cnt1, 1].Font.Bold = true;
                ExcelApp.Cells[cnt1, 1] = "Сотрудник: "+fio;

                if (ds1.Tables[0].Rows.Count != 0)
                {
                    cnt_ = 0;
                    for (i = 0; i <= ds1.Tables[0].Rows.Count - 1; i++)
                    {
                        ExcelApp.Cells[i + cnt, 1] = Convert.ToString(ds1.Tables[0].Rows[i][3]);
                        ExcelApp.Cells[i + cnt, 2] = Convert.ToString(ds1.Tables[0].Rows[i][4]);
                        ExcelApp.Cells[i + cnt, 3] = Convert.ToString(ds1.Tables[0].Rows[i][5]);
                        ExcelApp.Cells[i + cnt, 4] = Convert.ToString(ds1.Tables[0].Rows[i][6]);
                        ExcelApp.Cells[i + cnt, 5] = Convert.ToString(ds1.Tables[0].Rows[i][7]);
                        ExcelApp.Cells[i + cnt, 6] = Convert.ToString(ds1.Tables[0].Rows[i][8]);
                        ExcelApp.Cells[i + cnt, 7] = Convert.ToString(ds1.Tables[0].Rows[i][9]);
                        cnt_ = cnt_ + 1;

                    }

                    cnt1 = cnt + cnt_ + 3;
                    cnt = cnt1 + 1;
                    label6.Text = "Обработано сотрудников " + Convert.ToString(h);
                    progressBar1.Value = h;
                }
            }

            //Показываем ексель
            ExcelApp.Visible = true;
            ExcelApp.Interactive = true;
            ExcelApp.ScreenUpdating = true;
            ExcelApp.UserControl = true;
            
            progressBar1.Visible = false;
            label6.Visible = false;
        }

        private void актНаСписаниеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sotrudn_id = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
            sotrudn_fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);

            Form12 form12 = new Form12();
            form12.ShowDialog();
        }

        private void актыСверкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            view_akt_spis view_akt_spis = new view_akt_spis();
            view_akt_spis.ShowDialog();
        }

        private void движениеВыдачаСИЗToolStripMenuItem_Click(object sender, EventArgs e)
        {
            report_vidano report_vidano = new report_vidano();
            report_vidano.ShowDialog();
        }

        private void иАИСToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spr_iais spr_iais = new spr_iais();
            spr_iais.ShowDialog();
        }

        private void нормыСИЗToolStripMenuItem_Click(object sender, EventArgs e)
        {
            spr_norns spr_norns = new spr_norns();
            spr_norns.ShowDialog();
        }

        private void кодыСоответствияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            add_codes add_codes = new add_codes();
            add_codes.ShowDialog();
        }

        private void изменениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            user_change_data user_change_data = new user_change_data();
            user_change_data.ShowDialog();
        }

        private void увольнениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dismissal dismissal = new dismissal();
            dismissal.ShowDialog();
        }

        private void удалениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            user_remove user_remove = new user_remove();
            user_remove.ShowDialog();
        }

        private void сИЗСИстекшимСрокомНоскиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int color =dataGridView1.CurrentCell.Style.BackColor.ToArgb(); //определение цвета
            
            /////////
            if (color == -65536) //это красный цвет
            {
                elapsed_period_siz elapsed_period_siz = new elapsed_period_siz();

                //Просроченные к выдаче по людям
                SqlCommand command1 = new SqlCommand("alarm_elapsed_SIZ_detail  " + dataGridView1.CurrentRow.Cells[0].Value, conn);
                SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                DataSet ds1 = new DataSet();
                conn.Close();
                da1.Fill(ds1, "SOTRUDNIKI");
                elapsed_period_siz.dataGridView1.DataSource = ds1.Tables[0];
                
                elapsed_period_siz.ShowDialog();
            }

            ////////
            if (color == -256) //это желтый цвет
            {
                elapsed_period_siz elapsed_period_siz = new elapsed_period_siz();

                //Просроченные к выдаче по людям
                SqlCommand command1 = new SqlCommand("alarm_before_5days_detail  " + dataGridView1.CurrentRow.Cells[0].Value, conn);
                SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
                SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
                DataSet ds1 = new DataSet();
                conn.Close();
                da1.Fill(ds1, "SOTRUDNIKI");
                elapsed_period_siz.dataGridView1.DataSource = ds1.Tables[0];

                elapsed_period_siz.ShowDialog();
            }
        }

        private void приемToolStripMenuItem_Click(object sender, EventArgs e)
        {
            recruitment recruitment = new recruitment();
            recruitment.ShowDialog();
        }

        private void формированиеЗаявкиНаОбеспечениеСИЗToolStripMenuItem1_Click(object sender, EventArgs e)
        {
           
        }

        private void переводToolStripMenuItem_Click(object sender, EventArgs e)
        {
            movement movement = new movement();
            movement.ShowDialog();
        }

        private void импортПринятыхСотрудниковИзExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            export_sotrudniki export_sotrudniki = new export_sotrudniki();
            export_sotrudniki.ShowDialog();
        }

        private void необходимостьВСИЗToolStripMenuItem_Click(object sender, EventArgs e)
        {
            sotrudn_fio_4_need_siz = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            sotrudn_id_4_need_siz = Convert.ToInt16(dataGridView1.CurrentRow.Cells[0].Value);
            code_need_siz = Convert.ToString(dataGridView1.CurrentRow.Cells[9].Value);

            need_siz need_siz = new need_siz();
            need_siz.Owner = this;
            need_siz.ShowDialog();
        }

        private void сформироватьЗаявкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            kind_zayavka kind_zayavka = new kind_zayavka();
            kind_zayavka.ShowDialog();
        }

        private void накладнаяНаВыдачуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            nakladn_vidano nakladn_vidano = new nakladn_vidano();
            nakladn_vidano.ShowDialog();
        }

        private void каталогКАToolStripMenuItem_Click(object sender, EventArgs e)
        {
            catalog_kontragent catalog_kontragent = new catalog_kontragent();
            catalog_kontragent.ShowDialog();
        }

        private void привязкаКартПользователямToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cards_add cards_add = new cards_add();
            cards_add.ShowDialog();
        }

        private void заявкиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            view_zayavki view_zayavki = new view_zayavki();
            view_zayavki.ShowDialog();
        }

        private void бригадыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            brigades brigades = new brigades();
            brigades.ShowDialog();
        }

        private void складскиеОстаткиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            skl_ostatki skl_ostatki = new skl_ostatki();
            skl_ostatki.ShowDialog();
        }

        private void установкаДефлятораToolStripMenuItem_Click(object sender, EventArgs e)
        {
            set_deflator set_deflator = new set_deflator();
            set_deflator.ShowDialog();
        }

        private void лимитыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            limits limits = new limits();
            limits.ShowDialog();
        }

        private void потребностьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            need need = new need();
            need.ShowDialog();
        }

        private void сформироватьПотребностьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            kind_need kind_need = new kind_need();
            kind_need.ShowDialog();
        }

        private void загрузкаИтоговПослеТорговToolStripMenuItem_Click(object sender, EventArgs e)
        {
            load_itogi load_itogi = new load_itogi();
            load_itogi.ShowDialog();
        }

        

       
    }
}

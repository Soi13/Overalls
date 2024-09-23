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
    public partial class cards_add : Form
    {
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        bool myText = false;
        public static int idd_sotr;
        public static string fio;
        public static string uin;
        
        public cards_add()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 20;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 200;
            dataGridView1.Columns["PROFESSION"].HeaderText = "Должность/Профессия";
            dataGridView1.Columns["PROFESSION"].Width = 200;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Отдел/Участок/Цех";
            dataGridView1.Columns["DEPARTMENT"].Width = 200;
            dataGridView1.Columns["BRANCH"].HeaderText = "Филиал";
            dataGridView1.Columns["BRANCH"].Width = 200;
            dataGridView1.Columns["CARD_UIN"].HeaderText = "Номер карты";
            dataGridView1.Columns["CARD_UIN"].Width = 200;           

        }
        //////////////////////////////////////////////////////

        private void cards_add_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select SOTRUDNIKI.ID, SOTRUDNIKI.FIO, SOTRUDNIKI.PROFESSION,SOTRUDNIKI.DEPARTMENT,SOTRUDNIKI.BRANCH, CARDS.CARD_UIN from SOTRUDNIKI left join CARDS on SOTRUDNIKI.ID=CARDS.SOTRUDNIK_ID order by FIO", conn);
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

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (myText == false)
            {                
                this.label2.ForeColor = Color.WhiteSmoke;
                myText = true;
            }
            else
            {
                this.label2.ForeColor = Color.Red;
                myText = false;
            }
        }

        private void присвоитьНомерКартыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Проверка на уже присвоенный номера карты
            SqlCommand command = new SqlCommand("select SOTRUDNIK_ID, CARD_UIN from CARDS where SOTRUDNIK_ID=@sotr_id", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            command.Parameters.AddWithValue("sotr_id", dataGridView1.CurrentRow.Cells[0].Value);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "CARDS");
            if ((ds.Tables[0].Rows.Count>0) && (ds.Tables[0].Rows[0][1] != null))
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Данному сотруднику уже присвоена карта. Если необходимо заменить карту, воспользуйтесь специальным пунктом меню.", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            idd_sotr = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            button1.Visible = true;
            button2.Visible = true;
            button3.Visible = true;
            timer1.Enabled = true;
            label2.Visible = true;
            textBox1.Visible = true;
            textBox1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //Проверка на дублировние номеров карт, т.е. присвоение уже существующей карты
            SqlCommand command = new SqlCommand("select CARD_UIN from CARDS where CARD_UIN=@cd_uin", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            command.Parameters.AddWithValue("cd_uin", textBox1.Text);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "CARDS");
            if (ds.Tables[0].Rows.Count >0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Данная карта уже привязана в системе! Дублирование карт невозможно!!!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int ind = dataGridView1.CurrentRow.Index;  //Запоминание позиции строки в DataGrid

            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "INSERT into CARDS (SOTRUDNIK_ID, CARD_UIN, USER_ID, DATETIME_CREATE) VALUES (@sotr_id, @card_uin, @us_id, GETDATE())";
            scmd4.Parameters.AddWithValue("sotr_id", idd_sotr);
            scmd4.Parameters.AddWithValue("card_uin", textBox1.Text);
            scmd4.Parameters.AddWithValue("us_id", Form2.val);
            
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

            timer1.Enabled = false;
            timer2.Enabled = false;
            label2.Visible = false;
            textBox1.ReadOnly = false;
            textBox1.Clear();
            textBox1.Visible = false;
            button2.Visible = false;
            button1.Visible = false;
            button3.Visible = false;

            //Обновление инфы в DataGrid
            SqlCommand command1 = new SqlCommand("select SOTRUDNIKI.ID, SOTRUDNIKI.FIO, SOTRUDNIKI.PROFESSION,SOTRUDNIKI.DEPARTMENT,SOTRUDNIKI.BRANCH, CARDS.CARD_UIN from SOTRUDNIKI left join CARDS on SOTRUDNIKI.ID=CARDS.SOTRUDNIK_ID order by FIO", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
            DataSet ds1 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds1, "SOTRUDNIKI");
            dataGridView1.DataSource = ds1.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds1.Tables[0].Rows.Count);

            fill_gridview();

            dataGridView1.CurrentCell = dataGridView1[1, ind]; //Перемещение к той записи, к которой привязывали карту
            
            SystemSounds.Beep.Play();
            MessageBox.Show("Номер карты присвоен удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer2.Enabled = false;
            textBox1.ReadOnly = false;            
            textBox1.Clear();
            textBox1.Focus();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 5)
            {
                timer2.Enabled = true;
            }
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            textBox1.ReadOnly = true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            timer2.Enabled = false;
            label2.Visible = false;
            textBox1.ReadOnly = false;
            textBox1.Clear();
            textBox1.Visible = false;
            button2.Visible = false;
            button1.Visible = false;
            button3.Visible = false;
        }

        private void отвязатьКартуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int ind = dataGridView1.CurrentRow.Index;  //Запоминание позиции строки в DataGrid

            //Проверка производилась ли выдача с данной подписью
            SqlCommand command = new SqlCommand("select SOTRUDNIK_ID, CARD_UIN from VIDANO where SOTRUDNIK_ID=@sotr_id and CARD_UIN=@cd_uin", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            command.Parameters.AddWithValue("sotr_id", dataGridView1.CurrentRow.Cells[0].Value);
            command.Parameters.AddWithValue("cd_uin", dataGridView1.CurrentRow.Cells[5].Value);            
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "VIDANO");

            if (ds.Tables[0].Rows.Count > 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Отвязать карту невозможно, т.к. по ней производилась выдача СИЗ! Возможно только заменить карту на другую.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            
            
            //Отвызывание карты
            SqlCommand scmd4 = conn.CreateCommand();
            scmd4.CommandText = "delete from CARDS where sotrudnik_id=@id";
            scmd4.Parameters.AddWithValue("id", dataGridView1.CurrentRow.Cells[0].Value);            
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

            //Обновление инфы в DataGrid
            SqlCommand command1 = new SqlCommand("select SOTRUDNIKI.ID, SOTRUDNIKI.FIO, SOTRUDNIKI.PROFESSION,SOTRUDNIKI.DEPARTMENT,SOTRUDNIKI.BRANCH, CARDS.CARD_UIN from SOTRUDNIKI left join CARDS on SOTRUDNIKI.ID=CARDS.SOTRUDNIK_ID order by FIO", conn);
            SqlDataAdapter da1 = new SqlDataAdapter(command1);//Переменная объявлена как глобальная
            SqlCommandBuilder cb1 = new SqlCommandBuilder(da1);
            DataSet ds1 = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da1.Fill(ds1, "SOTRUDNIKI");
            dataGridView1.DataSource = ds1.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds1.Tables[0].Rows.Count);

            fill_gridview();

            dataGridView1.CurrentCell = dataGridView1[1, ind]; //Перемещение к той записи, к которой привязывали карту
            
            SystemSounds.Beep.Play();
            MessageBox.Show("Карта отвязана удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void заменитьКартуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //Проверка есть ли присвоенный номер карты
            SqlCommand command = new SqlCommand("select SOTRUDNIK_ID, CARD_UIN from CARDS where SOTRUDNIK_ID=@sotr_id", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            command.Parameters.AddWithValue("sotr_id", dataGridView1.CurrentRow.Cells[0].Value);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "CARDS");
            if (ds.Tables[0].Rows.Count == 0) 
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("Данному сотруднику карта не присвоена! Изменять нечего!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            idd_sotr = Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value);
            fio = Convert.ToString(dataGridView1.CurrentRow.Cells[1].Value);
            uin = Convert.ToString(dataGridView1.CurrentRow.Cells[5].Value);

            change_card change_card = new change_card();
            change_card.ShowDialog();
        }
                
    }
}


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
    public partial class change_card : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
     
        public change_card()
        {
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            textBox1.ReadOnly = false;
            textBox1.Clear();
            button2.Visible = false;
            button1.Visible = false;
            this.Close();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            if (textBox1.Text.Length > 5)
            {
                timer1.Enabled = true;
                button1.Visible = true;
                button2.Visible = true;
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            textBox1.ReadOnly = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            textBox1.ReadOnly = false;
            textBox1.Clear();
            textBox1.Focus();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что необходимо заменить карту сотруднику \""+cards_add.fio+"\"?", "Предупреждение", MessageBoxButtons.OKCancel, MessageBoxIcon.Stop) == DialogResult.OK)
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
                if (ds.Tables[0].Rows.Count > 0)
                {
                    SystemSounds.Beep.Play();
                    MessageBox.Show("Данная карта уже привязана в системе! Дублирование карт невозможно!!!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                
                //Замена карты, изменение CARD_UIN в таблицах VIDANO,  
                SqlCommand scmd4 = conn.CreateCommand();
                scmd4.CommandText = "update CARDS set CARD_UIN=@new_uin where SOTRUDNIK_ID=@sotr_id and CARD_UIN=@uin " +
                                    "update VIDANO set CARD_UIN=@new_uin where SOTRUDNIK_ID=@sotr_id and CARD_UIN=@uin " +
                                    "update NAKLADN_VIDANO_DETAIL set CARD_UIN=@new_uin where CARD_UIN=@uin";
                scmd4.Parameters.AddWithValue("new_uin", textBox1.Text);
                scmd4.Parameters.AddWithValue("sotr_id", cards_add.idd_sotr);
                scmd4.Parameters.AddWithValue("uin", cards_add.uin);
                
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

                SystemSounds.Beep.Play();
                MessageBox.Show("Сотруднику "+cards_add.fio+" карта изменена.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}

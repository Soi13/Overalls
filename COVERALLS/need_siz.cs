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
    public partial class need_siz : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
      
        public need_siz()
        {
            InitializeComponent();
        }

        private void need_siz_Load(object sender, EventArgs e)
        {
            groupBox1.Text = Form1.sotrudn_fio_4_need_siz;
            if (Form1.code_need_siz == "0") { radioButton2.Checked = true; radioButton1.Checked = false; }
            if (Form1.code_need_siz == "1") { radioButton1.Checked = true; radioButton2.Checked = false; }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int code=0;
            if (radioButton1.Checked==true) {code=1;}
            if (radioButton2.Checked==true) {code=0;}
            /////////Корректировка необходимости выдачи СИЗ
            SqlCommand scmd = conn.CreateCommand();
            scmd.CommandText = "update SOTRUDNIKI set NEED_SIZ=@code where ID=@id";
            scmd.Parameters.AddWithValue("code",code);
            scmd.Parameters.AddWithValue("id", Form1.sotrudn_id_4_need_siz);            
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
            MessageBox.Show("Изменения внесены удачно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            this.Close();
                      
        }
    }
}

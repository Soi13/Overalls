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
    public partial class Form3 : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");
        
        SqlDataAdapter da;
        Form2 form2 = new Form2();

        public Form3()
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
            dataGridView1.Columns["PROFESSION"].Width = 250;
            
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            dataGridView1.AllowUserToAddRows = false;
            
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

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 form1 = new Form1();

            /////////Вставка данных
            SqlCommand scmd = conn.CreateCommand();
            scmd.CommandText = "INSERT INTO NORMS_PERSONAL (NORMS_ID,SOTRUDNIK_ID,USER_ID) VALUES ((select id from norms where profession=" + "'" + dataGridView1.CurrentRow.Cells[1].Value + "'),'" + Form1.sotrudn_id + "','"+Form2.val+ "')";
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

            /////////UPDATE записи в таблице SOTRUDNIKI в поле SPISOK_ODEJD - вставка наименования профессии норматива
            SqlCommand scmd1 = conn.CreateCommand();
            scmd1.CommandText = "update SOTRUDNIKI set SPISOK_ODEJD='"+dataGridView1.CurrentRow.Cells[1].Value+"' where ID=" + "'" + Form1.sotrudn_id+"'";
            try
            {
                conn.Open();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с базой данных");
            }
            SqlDataReader reader1;
            reader1 = scmd1.ExecuteReader();
            conn.Close();
            //////////////////

            SystemSounds.Beep.Play();
            MessageBox.Show("Для сотрудника "+Form1.sotrudn_fio+ " нормы привязаны успешно!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.Close();
        }
        
    }
}

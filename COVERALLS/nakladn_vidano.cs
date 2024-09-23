using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace COVERALLS
{
    public partial class nakladn_vidano : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");
        //SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=OLEG_PC\MSSQLSERVER2012");

        public static string g;
        Label T;

        public nakladn_vidano()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "№п/п";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 40;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["FIO"].HeaderText = "ФИО";
            dataGridView1.Columns["FIO"].Width = 200;
            dataGridView1.Columns["DEPARTMENT"].HeaderText = "Подразделение";
            dataGridView1.Columns["DEPARTMENT"].Width = 300;
            dataGridView1.Columns["SOTRUDNIK_ID"].HeaderText = "SOTRUDNIK_ID";
            dataGridView1.Columns["SOTRUDNIK_ID"].Width = 20;
            dataGridView1.Columns["SOTRUDNIK_ID"].Visible = false;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 20;
            dataGridView1.Columns["USER_ID"].Visible = false;            
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "Дата создания";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 100;            
        }

        private void nakladn_vidano_Load(object sender, EventArgs e)
        {
            SqlCommand command = new SqlCommand("select NAKLADN_VIDANO.ID, NAKLADN_VIDANO.GUID, SOTRUDNIKI.FIO, SOTRUDNIKI.DEPARTMENT, NAKLADN_VIDANO.SOTRUDNIK_ID, NAKLADN_VIDANO.USER_ID, NAKLADN_VIDANO.DATETIME_CREATE from NAKLADN_VIDANO, SOTRUDNIKI where NAKLADN_VIDANO.SOTRUDNIK_ID=SOTRUDNIKI.ID", conn);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "NAKLADN_VIDANO");
            dataGridView1.DataSource = ds.Tables[0];

            statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

            fill_gridview();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            g = dataGridView1.CurrentRow.Cells[1].Value.ToString();

            nakladn_vidano_detail nakladn_vidano_detail = new nakladn_vidano_detail();
            nakladn_vidano_detail.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Сформировать накладную?", "Вопрос", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                SqlCommand command = new SqlCommand("select * from NAKLADN_VIDANO_DETAIL where GUID_NAKLADN_VIDANO=@g", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("g", dataGridView1.CurrentRow.Cells[1].Value);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "NAKLADN_VIDANO_DETAIL");
                
                /////Создание объекта задание на платеж
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(Environment.CurrentDirectory + @"\template\MB-7.xlsm", 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                /////////

                ObjWorkSheet.Cells[6, 1] = "ВЕДОМОСТЬ N"+ dataGridView1.CurrentRow.Cells[0].Value.ToString();                
                ObjWorkSheet.Cells[13, 3] = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                ObjWorkSheet.Cells[16, 10] = dataGridView1.CurrentRow.Cells[3].Value.ToString();

                /*
                 Заполянем форму. Заполяняются первые 17 строк, затем если кол-во выданных запиесей превыашет 17 строк, то это переносится на другую страницу, т.е. записи начинаются с 44 строки.
                 Это реализовано с помощью двух циклов, метки и оператора goto
                 */
                for (int r = 0; r <= ds.Tables[0].Rows.Count - 1; r++)
                {
                    if (r == 18)
                    {
                        goto T;
                    }
                    string naim = ds.Tables[0].Rows[r][3].ToString();
                    string edizm = ds.Tables[0].Rows[r][4].ToString();
                    string kolvo = ds.Tables[0].Rows[r][5].ToString();
                    string per_isp = ds.Tables[0].Rows[r][6].ToString();
                    string dt_vidach = ds.Tables[0].Rows[r][7].ToString().Remove(11);//Отсекаем время из даты
                    
                    ObjWorkSheet.Cells[r + 21, 2] = dataGridView1.CurrentRow.Cells[2].Value;
                    ObjWorkSheet.Cells[r + 21, 4] = naim;
                    ObjWorkSheet.Cells[r + 21, 7] = edizm;
                    ObjWorkSheet.Cells[r + 21, 8] = kolvo;
                    ObjWorkSheet.Cells[r + 21, 9] = dt_vidach;
                    ObjWorkSheet.Cells[r + 21, 10] = per_isp;                    
                                   
                }

            T: for (int t = 18; t <= ds.Tables[0].Rows.Count - 1; t++)
                {
                    string naim = ds.Tables[0].Rows[t][3].ToString();
                    string edizm = ds.Tables[0].Rows[t][4].ToString();
                    string kolvo = ds.Tables[0].Rows[t][5].ToString();
                    string per_isp = ds.Tables[0].Rows[t][6].ToString();
                    string dt_vidach = ds.Tables[0].Rows[t][7].ToString().Remove(11);//Отсекаем время из даты

                    ObjWorkSheet.Cells[t + 26, 2] = dataGridView1.CurrentRow.Cells[2].Value;
                    ObjWorkSheet.Cells[t + 26, 4] = naim;
                    ObjWorkSheet.Cells[t + 26, 7] = edizm;
                    ObjWorkSheet.Cells[t + 26, 8] = kolvo;
                    ObjWorkSheet.Cells[t + 26, 9] = dt_vidach;
                    ObjWorkSheet.Cells[t + 26, 10] = per_isp; 
                }
                

                    ObjExcel.Visible = true;

                GC.Collect();
            }
        }
    }
}

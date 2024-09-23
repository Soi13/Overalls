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
    public partial class view_zayavki : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        public static string gd;
        public static string[] arr_guid = new string[2];

        public view_zayavki()
        {
            InitializeComponent();
        }

        //Заполнение DataGridView наименованиями полей 
        public void fill_gridview()
        {
            dataGridView1.Columns["ID"].HeaderText = "ID";
            dataGridView1.Columns["ID"].Width = 40;
            dataGridView1.Columns["ID"].Visible = false;
            dataGridView1.Columns["GUID"].HeaderText = "GUID";
            dataGridView1.Columns["GUID"].Width = 20;
            dataGridView1.Columns["GUID"].Visible = false;
            dataGridView1.Columns["USER_ID"].HeaderText = "USER_ID";
            dataGridView1.Columns["USER_ID"].Width = 20;
            dataGridView1.Columns["USER_ID"].Visible = false;
            dataGridView1.Columns["KIND_ZAYAV"].HeaderText = "Тип заявки";
            dataGridView1.Columns["KIND_ZAYAV"].Width = 100;
            dataGridView1.Columns["PODR"].HeaderText = "Отдел/участок/цех";
            dataGridView1.Columns["PODR"].Width = 200;
            dataGridView1.Columns["PERIOD"].HeaderText = "Период заявки";
            dataGridView1.Columns["PERIOD"].Width = 150;
            dataGridView1.Columns["DATETIME_CREATE"].HeaderText = "DATETIME_CREATE";
            dataGridView1.Columns["DATETIME_CREATE"].Width = 20;
            dataGridView1.Columns["DATETIME_CREATE"].Visible = false;            
        }
        //////////////////////////////////////////////////////

        private void view_zayavki_Load(object sender, EventArgs e)
        {
            if ((Form2.val == 115) || (Form2.val == 3) || (Form2.val == 2)) //ПОказ всех сотрудников если заходит Скворцов ОИ или Петушков ИВ или Шепилева
            {
                SqlCommand command = new SqlCommand("select * from ZAYAVKA_MAIN", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_MAIN");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }
            else
            {
                SqlCommand command = new SqlCommand("select * from ZAYAVKA_MAIN where USER_ID=" + Form2.val, conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_MAIN");
                dataGridView1.DataSource = ds.Tables[0];

                statusStrip1.Items[0].Text = "Всего записей: " + Convert.ToString(ds.Tables[0].Rows.Count);

                fill_gridview();
            }
        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {
            gd = dataGridView1.CurrentRow.Cells[1].Value.ToString();

            zayav_detail zayav_detail = new zayav_detail();
            zayav_detail.ShowDialog();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Сформировать заявку?", "Вопрос", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                /////Создание объекта 
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(Environment.CurrentDirectory + @"\template\pril_6.xlsm", 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
                /////////

                ObjWorkSheet.Cells[6, 14] = DateTime.Today.ToShortDateString();
                ObjWorkSheet.Cells[11, 2] = dataGridView1.CurrentRow.Cells[5].Value+ " год";
                ObjWorkSheet.Cells[15, 3] = dataGridView1.CurrentRow.Cells[4].Value;

                progressBar1.Value = 0;

                SqlCommand command = new SqlCommand("select * from ZAYAVKA_BODY where GUID=@gd", conn);
                SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
                SqlCommandBuilder cb = new SqlCommandBuilder(da);
                DataSet ds = new DataSet();
                command.Parameters.AddWithValue("gd", dataGridView1.CurrentRow.Cells[1].Value);
                conn.Close();
                //Заполнение DataGridView наименованиями полей 
                da.Fill(ds, "ZAYAVKA_BODY");

                progressBar1.Maximum = ds.Tables[0].Rows.Count;
                
                //Добавление кол-ва строк, столько сколько выбрано из БД запросом
                for (int r = 1; r <= ds.Tables[0].Rows.Count; r++)
                {
                    //Выполняем макрос для вставки строки
                    ObjExcel.Run((object)"InsRow", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                }

                double ssum=0;
                int t;
                for (t = 24; t <= ds.Tables[0].Rows.Count + 23; t++)
                {
                    progressBar1.Value = progressBar1.Value + 1;
                    
                    ObjWorkSheet.Cells[t, 1] = ds.Tables[0].Rows[t - 24][2].ToString();
                    ObjWorkSheet.Cells[t, 2] = ds.Tables[0].Rows[t - 24][3].ToString();
                    ObjWorkSheet.Cells[t, 4] = ds.Tables[0].Rows[t - 24][5].ToString();
                    ObjWorkSheet.Cells[t, 5] = ds.Tables[0].Rows[t - 24][6].ToString();
                    ObjWorkSheet.Cells[t, 6] = dataGridView1.CurrentRow.Cells[4].Value;
                    ObjWorkSheet.Cells[t, 7] = ds.Tables[0].Rows[t - 24][7].ToString();
                    ObjWorkSheet.Cells[t, 8] = ds.Tables[0].Rows[t - 24][8].ToString();
                    ObjWorkSheet.Cells[t, 9] = ds.Tables[0].Rows[t - 24][9].ToString();
                    ObjWorkSheet.Cells[t, 10] = ds.Tables[0].Rows[t - 24][10];
                    ObjWorkSheet.Cells[t, 11] = ds.Tables[0].Rows[t - 24][11];
                    ObjWorkSheet.Cells[t, 20] = ds.Tables[0].Rows[t - 24][25].ToString();
                    ObjWorkSheet.Cells[t, 21] = ds.Tables[0].Rows[t - 24][12].ToString();
                    ObjWorkSheet.Cells[t, 22] = ds.Tables[0].Rows[t - 24][13].ToString();
                    ObjWorkSheet.Cells[t, 23] = ds.Tables[0].Rows[t - 24][14].ToString();
                    ObjWorkSheet.Cells[t, 24] = ds.Tables[0].Rows[t - 24][15].ToString();
                    ObjWorkSheet.Cells[t, 25] = ds.Tables[0].Rows[t - 24][16].ToString();
                    ObjWorkSheet.Cells[t, 26] = ds.Tables[0].Rows[t - 24][17].ToString();
                    ObjWorkSheet.Cells[t, 27] = ds.Tables[0].Rows[t - 24][18].ToString();
                    ObjWorkSheet.Cells[t, 28] = ds.Tables[0].Rows[t - 24][19].ToString();
                    ObjWorkSheet.Cells[t, 29] = ds.Tables[0].Rows[t - 24][20].ToString();
                    ObjWorkSheet.Cells[t, 30] = ds.Tables[0].Rows[t - 24][21].ToString();
                    ObjWorkSheet.Cells[t, 31] = ds.Tables[0].Rows[t - 24][22].ToString();
                    ObjWorkSheet.Cells[t, 32] = ds.Tables[0].Rows[t - 24][23].ToString();

                    ssum = ssum + Convert.ToDouble(ds.Tables[0].Rows[t - 24][9]);
                }

                ObjWorkSheet.Cells[t + 1, 8] = "Итого (общее кол-во):";
                ObjWorkSheet.Cells[t + 1, 9] = ssum;


                ObjExcel.Visible = true;

                GC.Collect();

            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            ////////
            SqlCommand command = new SqlCommand("select SUM(PRICE_KOMPLEKT_PLAN) from ZAYAVKA_BODY where GUID=@gd", conn);
            command.Parameters.AddWithValue("gd", dataGridView1.CurrentRow.Cells[1].Value);
            SqlDataAdapter da = new SqlDataAdapter(command);//Переменная объявлена как глобальная
            SqlCommandBuilder cb = new SqlCommandBuilder(da);
            DataSet ds = new DataSet();
            conn.Close();
            //Заполнение DataGridView наименованиями полей 
            da.Fill(ds, "ZAYAVKA_BODY");
            /////////////////

            label2.Text = "Сумма по заявке: " + Math.Round(Convert.ToDouble(ds.Tables[0].Rows[0][0].ToString()),2)+ " руб.";
            
            if (dataGridView1.SelectedRows.Count == 2)
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Вы уверены, что необходимо сравнить выбранные заявки?", "Вопрос", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                //Очищаем массив
                arr_guid[0] = "";
                arr_guid[1] = "";
                                
                //Заполняем массив GUIDами выделенных строк, чтобы по ним потом сделать выборку из тела заявки. В arr_guid[0] будет хроаниться первая выбранная заявка, в arr_guid[1] соответсвенно вторая.                
                int nn=0;
                foreach (DataGridViewRow dgvRow in dataGridView1.SelectedRows)
                {
                    arr_guid[nn] = dgvRow.Cells[1].Value.ToString();
                    nn++;                    
                }                    

                diff_between_zayav diff_between_zayav = new diff_between_zayav();
                diff_between_zayav.Owner = this;
                diff_between_zayav.ShowDialog();              
            }
        }
    }
}

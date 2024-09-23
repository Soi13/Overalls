using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Media;
using System.Data.SqlClient;

namespace COVERALLS
{
    public partial class load_itogi : Form
    {
        SqlConnection conn = new SqlConnection(@"Password=000;Persist Security Info=True;User ID=sa;Initial Catalog=COVERALLS;Data Source=T1212-W00079\MSSQLSERVER2012");

        public load_itogi()
        {
            InitializeComponent();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.FileName.Length == 0)
            {
                SystemSounds.Beep.Play();
                MessageBox.Show("ВЫ не выбрали файл для загрузки!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            /////Создание объекта итоги для загрузки
            Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
            ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];
            /////////

            progressBar1.Minimum = 0;
            progressBar1.Maximum = ObjWorkSheet.UsedRange.Rows.Count - 1;

            ////////////////////////////

            for (int i = 2; i <= ObjWorkSheet.UsedRange.Rows.Count - 1; i++)
            {
                string kod_guid = Convert.ToString(ObjWorkSheet.Cells[i, 1].value);
                string naim = Convert.ToString(ObjWorkSheet.Cells[i, 2].value);
                string dep = Convert.ToString(ObjWorkSheet.Cells[i, 4].value);
                string br = Convert.ToString(ObjWorkSheet.Cells[i, 5].value);
                string kolvo = Convert.ToString(ObjWorkSheet.Cells[i, 6].value);
                string offer = Convert.ToString(ObjWorkSheet.Cells[i, 9].value);
                string naim_ka = Convert.ToString(ObjWorkSheet.Cells[i, 10].value);
                string price = Convert.ToString(ObjWorkSheet.Cells[i, 11].value);

                if ((kod_guid.Length != 0) || (naim.Length != 0) || (dep.Length != 0) || (br.Length != 0) || (kolvo.Length != 0) || (offer.Length != 0) || (naim_ka.Length != 0) || (price.Length != 0))
                {

                    /////////Вставка данных
                    SqlCommand scmd = conn.CreateCommand();
                    scmd.CommandText = "update NEED_BODY set OFFERED_NAIM_COVERALL_KA=@offer, NAIM_KA=@naim_ka, PRICE_WITHOUT_NDS=@price where GUID=@kd_guid and NAIMENOVAN_COVERALLS=@naim and DEPARTMENT=@dp and BRANCH=@br and KOL_VO=@kolvo";
                    scmd.Parameters.AddWithValue("kd_guid", kod_guid);
                    scmd.Parameters.AddWithValue("naim", naim);
                    scmd.Parameters.AddWithValue("dp", dep);
                    scmd.Parameters.AddWithValue("br", br);
                    scmd.Parameters.AddWithValue("kolvo", kolvo);
                    scmd.Parameters.AddWithValue("offer", offer);
                    scmd.Parameters.AddWithValue("naim_ka", naim_ka);
                    scmd.Parameters.AddWithValue("price", Convert.ToDouble(price));

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

                    progressBar1.Value++;
                }
            }

            ObjWorkBook.Close();
            ObjExcel.Quit();
            progressBar1.Value = 0;

            SystemSounds.Beep.Play();
            MessageBox.Show("Загрузка прошла удачно!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            label1.Text = openFileDialog1.FileName;
        }
    }
}

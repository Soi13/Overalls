using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace COVERALLS
{
    public partial class Form11 : Form
    {
        public Form7 f7;
        string date_sled_vidachi;

        public Form11(Form7 form7)
        {
            InitializeComponent();
            f7 = form7;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            f7.dataGridView2.CurrentRow.Cells[6].Value = Convert.ToString(dateTimePicker1.Value);



            //Определение срока использования носки, если он меньше месяца, например 0.7, то используется функция AddDays, если более месяца (целое число), то AddMonth
            string srok = f7.dataGridView2.CurrentRow.Cells[5].Value.ToString();
            string decimal_sep = System.Globalization.NumberFormatInfo.CurrentInfo.CurrencyDecimalSeparator;
            srok = srok.Replace(".", decimal_sep);
            srok = srok.Replace(",", decimal_sep);
            double srok_p = Convert.ToDouble(srok);

            if (srok_p < 1)
            {
                double rs = srok_p * Convert.ToDouble(DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month));
                date_sled_vidachi = Convert.ToString(Convert.ToDateTime(f7.dataGridView2.CurrentRow.Cells[6].Value).AddDays(rs));//Вычисление даты следующей выдачи спецодежы                      
            }
            else
            {
                date_sled_vidachi = Convert.ToString(Convert.ToDateTime(f7.dataGridView2.CurrentRow.Cells[6].Value).AddMonths(Convert.ToInt16(f7.dataGridView2.CurrentRow.Cells[5].Value)));//Вычисление даты следующей выдачи спецодежы                
            }
            ////////////////////////////////////////

            f7.dataGridView2.CurrentRow.Cells[7].Value = Convert.ToDateTime(date_sled_vidachi);  //Вычисление даты следующей выдачи спецодежы                       
            this.Close();
        }
    }
}

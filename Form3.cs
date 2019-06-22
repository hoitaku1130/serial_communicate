using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace serial_com
{
    public partial class Form3 : Form
    {
        private Form1 super_class_ins;
        public Form3(Form1 __su)
        {
            InitializeComponent();
            this.ControlBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.super_class_ins = __su;

            this.button2.Click += exit_button_click;
            this.button1.Click += calc_button_click;
        }

        private void exit_button_click(object sender,EventArgs e) {
            this.Dispose();
        }
        private void calc_button_click(object sender,EventArgs e) {
            if (this.textBox1.Text != null && this.textBox2.Text != null && this.textBox3.Text != null)
            {
                try {
                    if (double.Parse(this.textBox1.Text) > double.Parse(this.textBox2.Text)) {
                        throw new Exception("start_secondの方が値が大きいです");
                    }
                    Dictionary<string, string> dc = this.super_class_ins.my_all_data;
                    List<double> sum = new List<double>();
                    foreach (var value in dc)
                    {
                        if (double.Parse(value.Key) >= double.Parse(textBox1.Text) && double.Parse(value.Key) <= double.Parse(textBox2.Text))
                        {
                            sum.Add(double.Parse(value.Value));
                        }
                    }
                    //sumリストを積分する
                    double integral_func(List<double> mylist) { //間が等間隔1sとすると
                        int length = mylist.Count();
                        double result = 0.0;
                        for (int i = 0; i < length; i++) {
                            if (i == 0 || i == length - 1) {
                                result += mylist[i];
                                continue;
                            }
                            result += 2.0 * mylist[i];
                        }
                        MessageBox.Show(length.ToString());
                        return result*0.5;
                    }
                    this.label4.Text = (double.Parse(this.textBox3.Text) * integral_func(sum)).ToString();
                } catch (Exception ex) {
                    MessageBox.Show("データ範囲外か\n" + ex.Message);
                }
                
            }
            else {
                return;
            }
        }
    }
}

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
    public partial class Form2 : Form
    {
        private Form1 super_class;
        public Form2(Form1 my_form1)
        {
            InitializeComponent();
            this.super_class = my_form1;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.button2.Click += Button_Cancel_Click;
            this.button1.Click += Button_OK_Click;
            
        }

        private void Button_Cancel_Click(object sender,EventArgs e) {
            this.Dispose();
        }
        private void Button_OK_Click(object sender,EventArgs e) {   //エラーチェックしていない
            if (textBox1.Text.Length != 0 && textBox2.Text.Length != 0) {
                try
                {

                    if (int.Parse(textBox2.Text) - int.Parse(textBox1.Text) < 0)
                    {
                        
                        throw new Exception("xmaxよりyminのほうが大きいです");
                    }
                    else {
                        this.super_class.class_plot_xmin = int.Parse(textBox1.Text);
                        this.super_class.class_plot_xmax = int.Parse(textBox2.Text);
                    }
                    
                    this.Dispose();
                }
                catch (Exception ex) {
                    MessageBox.Show(ex.Message);
                }
                
            }
            if (textBox3.Text.Length != 0 && textBox4.Text.Length != 0) //y軸のほう
            {
                try
                {

                    if (int.Parse(textBox4.Text) - int.Parse(textBox3.Text) < 0)
                    {

                        throw new Exception("ymaxよりyminのほうが大きいです");
                    }
                    else
                    {
                        this.super_class.class_plot_ymin = double.Parse(textBox3.Text);
                        this.super_class.class_plot_ymax = double.Parse(textBox4.Text);
                    }

                    this.Dispose();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }

            }
        }

        //objectの変換のため　更新したらこちらも更新必要

        
    }
}

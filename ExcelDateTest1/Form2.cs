using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Calculator
{
    
    public partial class Form2 : Form
    {


        //作者:陈继民
        //版本号：v2.0
        //时间：2015年12月11日



        public Form2()
        {
            InitializeComponent();
        }
        

        
        private void button1_Click(object sender, EventArgs e)//登陆验证
        {
            if (textBox1.Text=="" && textBox2.Text=="")
            {
                MessageBox.Show("请输入用户名和密码后重试");

                textBox1.Focus();
            }

            else if (textBox1.Text == "")
            {
                MessageBox.Show("用户名为空，请输入用户名");

                textBox1.Focus();
            }

            else if (textBox2.Text == "")
            {
                MessageBox.Show("密码为空，请输入密码");

                textBox2.Focus();
            }

            else if (textBox1.Text != "FYL" && textBox2.Text == "wasd")
            {
                MessageBox.Show("用户名错误，请重试");

                textBox1.Clear();
            }

            else if (textBox1.Text == "FYL" && textBox2.Text != "wasd")
            {
                MessageBox.Show("密码错误，请重试");

                textBox2.Clear();
            }

            else if (textBox1.Text == "FYL" && textBox2.Text == "wasd")//当账户和密码正确后激活Form1
            {
                this.Visible = false;

                Form1 f1 = new Form1();

                f1.ShowDialog();

              //  this.DialogResult = DialogResult.OK;
               
            }

            else
            {
                MessageBox.Show("用户名或密码错误，请重试");

                textBox1.Clear();

                textBox2.Clear();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }//退出程序
    }
}

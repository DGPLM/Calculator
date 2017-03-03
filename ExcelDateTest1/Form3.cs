using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.IO;

namespace ExcelDateTest1
{
    public partial class Form3 : Form
    {


        //作者:陈继民
        //版本号：v2.0
        //时间：2015年12月11日



        public Form3()
        {
            InitializeComponent();

        }
               

        
        private void button1_Click_1(object sender, EventArgs e)
        {
            System.Windows.Forms.TextBox[] tb1 = new System.Windows.Forms.TextBox[48];

            tb1[0] = textBox1;

            tb1[1] = textBox2;

            tb1[2] = textBox3;

            tb1[3] = textBox4;

            tb1[4] = textBox5;

            tb1[5] = textBox6;

            tb1[6] = textBox7;

            tb1[7] = textBox8;

            tb1[8] = textBox9;

            tb1[9] = textBox10;

            tb1[10] = textBox11;

            tb1[11] = textBox12;

            tb1[12] = textBox13;

            tb1[13] = textBox14;

            tb1[14] = textBox15;

            tb1[15] = textBox16;

            tb1[16] = textBox17;

            tb1[17] = textBox18;

            tb1[18] = textBox19;

            tb1[19] = textBox20;

            tb1[20] = textBox21;

            tb1[21] = textBox22;

            tb1[22] = textBox23;

            tb1[23] = textBox24;

            tb1[24] = textBox25;

            tb1[25] = textBox26;

            tb1[26] = textBox27;

            tb1[27] = textBox28;

            tb1[28] = textBox29;

            tb1[29] = textBox30;

            tb1[30] = textBox31;

            tb1[31] = textBox32;

            tb1[32] = textBox33;

            tb1[33] = textBox34;

            tb1[34] = textBox35;

            tb1[35] = textBox36;

            tb1[36] = textBox37;

            tb1[37] = textBox38;

            tb1[38] = textBox39;

            tb1[39] = textBox40;

            tb1[40] = textBox41;

            tb1[41] = textBox42;

            tb1[42] = textBox43;

            tb1[43] = textBox44;

            tb1[44] = textBox45;

            tb1[45] = textBox46;

            tb1[46] = textBox47;

            tb1[47] = textBox48;
            
            for (int i = 0; i < tb1.Length; i++)
            {
                if (tb1[i].Text == "")
                {
                    MessageBox.Show("数据缺少，请输入！");

                    tb1[i].Focus();

                    return;

                    if (tb1[i].Text != "")
                    {
                        continue;
                    }
                }
            }

            string a = tb1[45].Text;//创建abc三个变量存储路径名，以便后面更改

            string b = tb1[46].Text;

            string c = tb1[47].Text;

            if (a == "")
            {
                MessageBox.Show("班级名不能为空，请输入！");

                tb1[45].Focus();
            }

            if (a != "" && c == "")
            {
                MessageBox.Show("姓名不能为空，请输入！");

                tb1[47].Focus();
            }

            if (a != "" && b == "")
            {
                MessageBox.Show("学号不能为空，请输入！");

                tb1[46].Focus();
            }

            if (a != "" && b != "" && c != "")//当abc三个变量都存在时，生成数据文件
            {
               Calculate.InputData(tb1, a, b, c);
            }

            for (int i = 0; i < tb1.Length; i++)
            {
                tb1[i].Clear();
            }
            


        }



        #region 菜单代码
        private void fdfdToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Visible = false;

            Form1.path = null;

            Form1 f1 = new Form1();

            f1.ShowDialog();

            this.DialogResult = DialogResult.OK;


        }
        
        
        
        private void 计算ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form1.path = null;

            Form1 f1 = new Form1();

            this.Visible = false;
            f1.Show();            
        }



        private void 数据录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            return;
        }



        private void 退出程序ToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            this.Close();
        }



        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 f4 = new Form4();
            f4.Show();
        }
        #endregion 



    }
}

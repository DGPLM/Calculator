using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office;
using System.Data.Odbc;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace ExcelDateTest1
{

    public partial class Form1 : Form
    {


        //作者:陈继民
        //版本号：v2.0
        //时间：2015年12月11日



        #region 公共变量定义
        public static double[] Convert1 = new double[6];

        public static string[] temp = new string[6];

        public static string path;

        public static double[] Cal = new double[16];

        public static double[] iTheoryOutput = new double[6];


        public static double[] iRealityOutput = new double[6];

        public static double[] iBaifenwucha = new double[6];

        public static double[] fiveData = new double[5];
        #endregion



        public Form1()
        {
            InitializeComponent();
        }



        private void button1_Click(object sender, EventArgs e)//打开文件
        {

            OpenFileDialog dlg = new OpenFileDialog();

            dlg.Title = "请打开学生文件";

            dlg.ShowDialog();

            path = dlg.FileName;//定义路径名为文件所在路径

            if (path == "")//判断路径是否缺失（导入后未选择文件便点击取消的结果）
            {
                button2.Enabled = false;

                button3.Enabled = false;

                button4.Enabled = false;

                button5.Enabled = false;

                button6.Enabled = false;

                button7.Enabled = false;

                button8.Enabled = false;

                button9.Enabled = false;

                button10.Enabled = false;

                button11.Enabled = false;

                button12.Enabled = false;

                button13.Enabled = false;

                button14.Enabled = false;

                button15.Enabled = false;

            }

            else//路径不为空或者不缺失，则激活对应计算控件
            {
                button2.Enabled = true;

                button3.Enabled = true;

                button4.Enabled = true;

                button5.Enabled = true;

                button6.Enabled = true;

                button7.Enabled = true;
            }

        }




        public void button2_Click(object sender, EventArgs e) //周期T相关计算
        {
            Calculate.ExcelToString(path, 1, 4, 2, false);//读取对应数据，保存到公共变量中，后面同理

            Calculate.StringToDouble(temp);//将读取的string数组进行转换，后面同理

            Calculate.Avgn(Convert1, 20);//求平均值

            double Avg = Calculate.AvgDouble(Convert1);

            Cal[8] = Avg;//将周期T的平均值保留到Cal数组中，以便后面计算

            textBox1.Text = Convert.ToString(Calculate.A(Avg, 3));

            double Axiangdui1 = Calculate.Axiangdui(Convert1);//求A类不确定度

            textBox2.Text = Convert.ToString(Calculate.A(Axiangdui1, 3));

            double Bxiangdui1 = Calculate.Bxiangdui(0.002);//求B类不确定度

            textBox3.Text = Convert.ToString(Calculate.A(Bxiangdui1, 3));


            double Zongbuqueding1 = Calculate.Zongbuqueding(Axiangdui1, Bxiangdui1);//求总不确定度

            Cal[9] = Zongbuqueding1;//将周期T的总不确定度保留到Cal数组中，以便后面计算

            textBox4.Text = Convert.ToString(Calculate.A(Zongbuqueding1, 3));

            double Xiangduibuqueding = Calculate.Xiangduibuqueding1(Zongbuqueding1, Avg); //求相对不确定度

            textBox5.Text = Convert.ToString(Calculate.A(Xiangduibuqueding, 2)) + "%";

            fiveData[0] = 0;//定义一个新数组，用来控制后面暂时计算不了的计算按钮失效或者激活，后面同理

        }



        private void button3_Click(object sender, EventArgs e) //上孔a相关计算
        {
            Calculate.ExcelToString(path, 1, 10, 2, true);

            Calculate.StringToDouble(temp);

            double Avg = Calculate.AvgDouble(Convert1);//求平均值

            Cal[0] = Avg;//将上孔A的平均值保留到Cal数组中，以便后面计算

            textBox1.Text = Convert.ToString(Calculate.A(Avg, 2));

            double Axiangdui1 = Calculate.Axiangdui(Convert1);//求A类不确定度

            textBox2.Text = Convert.ToString(Calculate.A(Axiangdui1, 2));

            double Bxiangdui1 = Calculate.Bxiangdui(0.02);//求B类不确定度

            textBox3.Text = Convert.ToString(Calculate.A(Bxiangdui1, 2));

            double Zongbuqueding1 = Calculate.Zongbuqueding(Axiangdui1, Bxiangdui1);//求总不确定度

            Cal[3] = Zongbuqueding1;//将上孔A的总不确定度保留到Cal数组中，以便后面计算

            textBox4.Text = Convert.ToString(Calculate.A(Zongbuqueding1, 2));

            double Xiangduibuqueding = Calculate.Xiangduibuqueding1(Zongbuqueding1, Avg);//求相对不确定度

            textBox5.Text = Convert.ToString(Calculate.A(Xiangduibuqueding, 2)) + "%";

            fiveData[1] = 1;
        }



        private void button4_Click(object sender, EventArgs e)//下孔b相关计算
        {
            Calculate.ExcelToString(path, 1, 10, 3, true);

            Calculate.StringToDouble(temp);

            double Avg = Calculate.AvgDouble(Convert1);//求平均值

            Cal[1] = Avg;//将下孔B的平均值保留到Cal数组中，以便后面计算

            textBox1.Text = Convert.ToString(Calculate.A(Avg, 2));

            double Axiangdui1 = Calculate.Axiangdui(Convert1);//求A类不确定度

            textBox2.Text = Convert.ToString(Calculate.A(Axiangdui1, 2));

            double Bxiangdui1 = Calculate.Bxiangdui(0.02);//求B类不确定度

            textBox3.Text = Convert.ToString(Calculate.A(Bxiangdui1, 2));

            double Zongbuqueding1 = Calculate.Zongbuqueding(Axiangdui1, Bxiangdui1);//求总不确定度

            Cal[4] = Zongbuqueding1;//将下孔B的总不确定度保留到Cal数组中，以便后面计算

            textBox4.Text = Convert.ToString(Calculate.A(Zongbuqueding1, 2));

            double Xiangduibuqueding = Calculate.Xiangduibuqueding1(Zongbuqueding1, Avg);//求相对不确定度

            textBox5.Text = Convert.ToString(Calculate.A(Xiangduibuqueding, 2)) + "%";

            fiveData[2] = 2;
        }



        private void button5_Click(object sender, EventArgs e)//悬线长l相关计算
        {
            Calculate.ExcelToString(path, 1, 10, 4, true);

            Calculate.StringToDouble(temp);

            double Avg = Calculate.AvgDouble(Convert1);//求平均值

            Cal[2] = Avg;//将悬线长l的平均值保留到Cal数组中，以便后面计算

            textBox1.Text = Convert.ToString(Calculate.A(Avg, 2));

            double Axiangdui1 = Calculate.Axiangdui(Convert1);//求A类不确定度

            textBox2.Text = Convert.ToString(Calculate.A(Axiangdui1, 2));

            double Bxiangdui1 = Calculate.Bxiangdui(0.02);//求B类不确定度

            textBox3.Text = Convert.ToString(Calculate.A(Bxiangdui1, 2));

            double Zongbuqueding1 = Calculate.Zongbuqueding(Axiangdui1, Bxiangdui1); //求总不确定度

            Cal[5] = Zongbuqueding1;//将悬线长l的总不确定度保留到Cal数组中，以便后面计算

            textBox4.Text = Convert.ToString(Calculate.A(Zongbuqueding1, 2));

            double Xiangduibuqueding = Calculate.Xiangduibuqueding1(Zongbuqueding1, Avg); //求相对不确定度

            textBox5.Text = Convert.ToString(Calculate.A(Xiangduibuqueding, 2)) + "%";

            fiveData[3] = 3;
        }



        private void button6_Click(object sender, EventArgs e)//圆柱体半径r相关计算
        {
            Calculate.ExcelToString(path, 1, 10, 5, true);

            Calculate.StringToDouble(temp);

            Calculate.Avgn(Convert1, 2);

            double Avg = Calculate.AvgDouble(Convert1); //求平均值

            Cal[15] = Avg;//将半径r的平均值保留到Cal数组中，以便后面计算

            textBox1.Text = Convert.ToString(Calculate.A(Avg, 2));

            double Axiangdui1 = Calculate.Axiangdui(Convert1);//求A类不确定度

            textBox2.Text = Convert.ToString(Calculate.A(Axiangdui1, 2));

            double Bxiangdui1 = Calculate.Bxiangdui(0.02);//求B类不确定度

            textBox3.Text = Convert.ToString(Calculate.A(Bxiangdui1, 2));

            double Zongbuqueding1 = Calculate.Zongbuqueding(Axiangdui1, Bxiangdui1);//求总不确定度

            textBox4.Text = Convert.ToString(Calculate.A(Zongbuqueding1, 2));

            double Xiangduibuqueding = Calculate.Xiangduibuqueding1(Zongbuqueding1, Avg); //求相对不确定度

            textBox5.Text = Convert.ToString(Calculate.A(Xiangduibuqueding, 2)) + "%";

            fiveData[4] = 4;

        }



        private void Form1_Load(object sender, EventArgs e)
        {
            if (path == null)//当窗体载入时，如果路径为空则使计算按钮失效
            {
                button2.Enabled = false;

                button3.Enabled = false;

                button4.Enabled = false;

                button5.Enabled = false;

                button6.Enabled = false;

                button7.Enabled = false;

                button8.Enabled = false;

                button9.Enabled = false;

                button10.Enabled = false;

                button11.Enabled = false;

                button12.Enabled = false;

                button13.Enabled = false;

                button14.Enabled = false;

                button15.Enabled = false;

            }
        }//窗体载入



        private void button7_Click(object sender, EventArgs e)//间接测量量H的相关计算

        {
            for (int i = 0; i < fiveData.Length; i++)
            {
                if (fiveData[i] != i)
                {
                    MessageBox.Show("数据缺少,请返回回第一页将直接测量量计算完整！");//确保计算H所需数据全部计算完毕

                    return;
                }
            }

            double Hbuquedingdu;

            double Hxiangdui;

            Cal[6] = Calculate.Hcalculate(Cal[2], Cal[0], Cal[1]);//计算H的平均值

            Hbuquedingdu = Calculate.Hbuquedingdu(Cal);//计算H的总不确定度

            Cal[10] = Hbuquedingdu;

            Hxiangdui = Calculate.Xiangduibuqueding1(Hbuquedingdu, Cal[6]);

            textBox6.Text = Convert.ToString(Calculate.A(Cal[6], 2));

            textBox7.Text = Convert.ToString(Calculate.A(Hbuquedingdu, 2));

            textBox8.Text = Convert.ToString(Calculate.A(Hxiangdui, 2)) + "%";

            button8.Enabled = true;//H计算完毕，激活I的计算按钮
        }



        private void button8_Click(object sender, EventArgs e)//间接测量量转动惯量I的相关计算
        {
            Cal[7] = Convert.ToDouble(Calculate.ExcelToString(path, 1, 26, 2));

            double Icalculate = Calculate.Icalculate(Cal);//计算I的平均值

            Cal[14] = Icalculate;

            double Ixiangduibuquedingdu = Calculate.Ixiangduibuqueding(Cal);//计算I的相对不确定度

            double Ibuquedingdu = Calculate.A(Icalculate * Ixiangduibuquedingdu, 5);//计算I的总不确定度，以及保留小数（四舍六入五凑偶）

            textBox6.Text = Convert.ToString(Calculate.A(Icalculate, 5));

            textBox7.Text = Ibuquedingdu.ToString("0.00000");

            textBox8.Text = Convert.ToString(Calculate.A(Ixiangduibuquedingdu, 2)) + "%";

            button15.Enabled = true;
        }



        private void button15_Click(object sender, EventArgs e)//验证平行轴定理
        {
            Cal[11] = Convert.ToDouble(Calculate.ExcelToString(path, 1, 25, 2));//下圆盘直径

            Cal[12] = Convert.ToDouble(Calculate.ExcelToString(path, 1, 26, 2));//下圆盘质量

            Cal[13] = Convert.ToDouble(Calculate.ExcelToString(path, 1, 27, 2));//圆柱体质量

            //计算转动惯量理论值
            Calculate.ExcelToString(path, 1, 22, 2, false);

            Calculate.StringToDouble(temp);

            Calculate.Itheory(Cal[13], Cal[15], Convert1);

            //计算转动惯量实验值
            Calculate.ExcelToString(path, 1, 23, 2, false);

            Calculate.StringToDouble(temp);

            Calculate.Ireality(Cal, Convert1);

            //计算百分误差
            Calculate.Ibaifenwucha(iTheoryOutput, iRealityOutput);

            //激活对应按钮控件
            button9.Enabled = true;

            button10.Enabled = true;

            button11.Enabled = true;

            button12.Enabled = true;

            button13.Enabled = true;

            button14.Enabled = true;
        }



        #region 点击对应按钮显示对应六组数据的计算结果
        private void button9_Click(object sender, EventArgs e)//第一组
        {
            textBox9.Text = Convert.ToString(Calculate.A(iTheoryOutput[0], 5));

            textBox10.Text = Convert.ToString(Calculate.A(iRealityOutput[0], 5));

            textBox11.Text = Convert.ToString(Calculate.A(iBaifenwucha[0], 2)) + "%";
        }

        private void button10_Click(object sender, EventArgs e)//第二组
        {
            textBox9.Text = Convert.ToString(Calculate.A(iTheoryOutput[1], 5));

            textBox10.Text = Convert.ToString(Calculate.A(iRealityOutput[1], 5));

            textBox11.Text = Convert.ToString(Calculate.A(iBaifenwucha[1], 2)) + "%";
        }

        private void button11_Click(object sender, EventArgs e)//第三组
        {
            textBox9.Text = Convert.ToString(Calculate.A(iTheoryOutput[2], 5));

            textBox10.Text = Convert.ToString(Calculate.A(iRealityOutput[2], 5));

            textBox11.Text = Convert.ToString(Calculate.A(iBaifenwucha[2], 2)) + "%";
        }

        private void button12_Click(object sender, EventArgs e)//第四组
        {
            textBox9.Text = Convert.ToString(Calculate.A(iTheoryOutput[3], 5));

            textBox10.Text = Convert.ToString(Calculate.A(iRealityOutput[3], 5));

            textBox11.Text = Convert.ToString(Calculate.A(iBaifenwucha[3], 2)) + "%";
        }

        private void button13_Click(object sender, EventArgs e)//第五组
        {
            textBox9.Text = Convert.ToString(Calculate.A(iTheoryOutput[4], 5));

            textBox10.Text = Convert.ToString(Calculate.A(iRealityOutput[4], 5));

            textBox11.Text = Convert.ToString(Calculate.A(iBaifenwucha[4], 2)) + "%";
        }

        private void button14_Click(object sender, EventArgs e)//第六组
        {
            textBox9.Text = Convert.ToString(Calculate.A(iTheoryOutput[5], 5));

            textBox10.Text = Convert.ToString(Calculate.A(iRealityOutput[5], 5));

            textBox11.Text = Convert.ToString(Calculate.A(iBaifenwucha[5], 2)) + "%";
        }
        #endregion



        #region 菜单代码
        private void 计算ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            return;
        }

        private void 数据录入ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 f3 = new Form3();

            this.Visible = false;

            f3.Show();
        }

        private void 关于ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form4 f4 = new Form4();

            f4.Show();
        }

        private void 退出程序ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        #endregion



    }
}

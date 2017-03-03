using System;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.IO;
using System.Windows.Forms;

namespace Calculator
{
    //作者:陈继民
    //版本号：v2.0
    //时间：2015年12月11日

    class Calculate
    {      
        


        /// <summary>
     /// 求A类不确定度
     /// </summary>
     /// <param name="Convert1">参数为一个double数组</param>
     /// <returns>返回一个double类型的A类不确定度</returns>
        public static double Axiangdui(double[] Convert1)
        {
            double temp = 0;//创建一个double值作为累加的容器

            decimal Axiangdui = 0;//转换为decimal以便保留小数

            double Axiangdui1 = 0;

            double Axiangdui2 = 0;

            for (int i = 0; i < Convert1.Length; i++)
            {
                temp += Math.Pow((Convert1[i] - AvgDouble(Convert1)), 2);

            }

            Axiangdui = Convert.ToDecimal(temp) / (Convert1.Length - 1);

            Axiangdui1 = Convert.ToDouble(Axiangdui);

            Axiangdui2 = Math.Sqrt(Axiangdui1);

            return Axiangdui2;
        }
        


        /// <summary>
        /// 求B类不确定度
        /// </summary>
        /// <param name="sanjiao">参数为一个double的标准差</param>
        /// <returns>返回double类型的B类不确定度</returns>
        public static double Bxiangdui(double sanjiao)
        {
            decimal temp = Convert.ToDecimal(Math.Sqrt(3));//转换为decimal以便保留小数

            decimal temp1 = Convert.ToDecimal(sanjiao);

            return Convert.ToDouble(temp1 / temp);
        }
               
        
        
        
        /// <summary>
        /// 求总不确定度
        /// </summary>
        /// <param name="Axiangdui1">一个double的A类不确定度</param>
        /// <param name="Bxiangdui1">一个double的B类不确定度</param>
        /// <returns>返回double的总不确定度</returns>
        public static double Zongbuqueding(double Axiangdui1, double Bxiangdui1)
        {
            return Math.Sqrt(Math.Pow(Axiangdui1, 2) + Math.Pow(Bxiangdui1, 2));
        }
        


        /// <summary>
        /// 求相对不确定度
        /// </summary>
        /// <param name="Zongbuqueding">一个double的总不确定度</param>
        /// <param name="Avg">一个double的平均值</param>
        /// <returns></returns>
        public static double Xiangduibuqueding1(double Zongbuqueding, double Avg)
        {
            decimal Zongbuqueding1 = Convert.ToDecimal(Zongbuqueding);//转换为decimal以便保留小数

            decimal Avg1 = Convert.ToDecimal(Avg);

            return Convert.ToDouble(Zongbuqueding1 * 100 / Avg1);


        }
        
        

        /// <summary>
        /// 求一个double数组的n次平均
        /// </summary>
        /// <param name="Avg">一个double数组</param>
        /// <param name="j">平均次数</param>
        /// <returns>返回一个double数组</returns>
        public static double[] Avgn(double[] Avg, int j)
        {
            decimal[] temp1 = new decimal[6]; //转换为decimal以便保留小数

            for (int i = 0; i < Avg.Length; i++)
            {
                temp1[i] = Convert.ToDecimal(Avg[i]);

                Form1.Convert1[i] = Convert.ToDouble(temp1[i] / j);
            }

            return Form1.Convert1;
        }

        

        /// <summary>
        /// 换算单元（求平均数）
        /// </summary>
        /// <param name="Avg">参数为一个double值</param>
        /// <param name="i">换算倍率（求平均）</param>
        /// <returns>返回一个double值的平均值</returns>
        public static double Avgn(double Avg, int i)
        {
            decimal temp1 = Convert.ToDecimal(Avg);//转换为decimal以便保留小数

            return Convert.ToDouble(temp1 / i);
        }
        


        /// <summary>
        /// 调用excel的单元格数据
        /// </summary>
        /// <param name="path">一个string的路径参数</param>
        /// <param name="worksheet1">一个int的工作簿名称</param>
        /// <param name="row">一个int的行号（row）</param>
        /// <param name="col">一个int的列号（col）</param>
        /// <param name="bool1">一个bool值，true时从行开始加载，false时从列开始加载</param>
        /// <returns>一个string数组</returns>
        public static string[] ExcelToString(string path, int worksheetnumber, int row, int col, bool bool1)
        {


            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();//调用excel

            Workbook wbook = app.Workbooks.Open(path); //打开工作簿

            Worksheet worksheet = (Worksheet)wbook.Worksheets[worksheetnumber];

            //判断传入的bool值，以便更改读取方式，按行垂直读取或者按列垂直读取
            if (bool1 == true)
            {
                for (int i = 0; i < Form1.temp.Length; i++)
                {
                    Form1.temp[i] = ((Range)worksheet.Cells[row + i, col]).Text;
                }
            }

            else if (bool1 == false)
            {
                for (int i = 0; i < Form1.temp.Length; i++)
                {
                    Form1.temp[i] = ((Range)worksheet.Cells[row, col + i]).Text;
                }
            }

            Kill(app);

            return Form1.temp;

        }
        

        
        /// <summary>
        /// 调用excel的一个单元格数据
        /// </summary>
        /// <param name="path">一个string的路径参数</param>
        /// <param name="worksheetnumber">一个int的工作簿名称</param>
        /// <param name="row">一个int的行号（row）</param>
        /// <param name="col">一个int的列号（col）</param>
        /// <returns>返回一个string的数据</returns>
        public static string ExcelToString(string path, int worksheetnumber, int row, int col)
        {

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();//调用excel

            Workbook wbook = app.Workbooks.Open(path);//打开工作簿

            Worksheet worksheet = (Worksheet)wbook.Worksheets[worksheetnumber];

            string temp = ((Range)worksheet.Cells[row, col]).Text;//读取对应单元格数据

            Kill(app);

            return temp;
        }
        


        /// <summary>
        /// 将一个string数组转成double数组
        /// </summary>
        /// <param name="temp">参数为一个string数组</param>
        /// <returns>返回一个double数组</returns>
        public static double[] StringToDouble(string[] temp)
        {
            for (int i = 0; i < temp.Length; i++)
            {
                Form1.Convert1[i] = Convert.ToDouble(temp[i]);//遍历数组进行转换
            }

            return Form1.Convert1;

        }
        


        /// <summary>
        /// 求一组double数组的平均值
        /// </summary>
        /// <param name="Convert1">参数为一个double数组</param>
        /// <returns>返回一个double平均值</returns>
        public static double AvgDouble(double[] Convert1)
        {
            double Add = 0;//创建一个double值作为累加容器

            decimal Avg = 0;

            double Avg1 = 0;

            for (int i = 0; i < Convert1.Length; i++)//遍历数组进行累加
            {
                Add += Convert1[i];
            }
            Avg = Convert.ToDecimal(Add);//转换为decimal以便保留小数

            Avg1 = Convert.ToDouble(Avg / Convert1.Length);
            return Avg1;
        }
        


        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        public static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        public static void Kill(Microsoft.Office.Interop.Excel.Application excel)
        {
            IntPtr t = new IntPtr(excel.Hwnd);

            int k = 0;

            GetWindowThreadProcessId(t, out k);//找到对应句柄

            System.Diagnostics.Process p = System.Diagnostics.Process.GetProcessById(k);//根据句柄找到对应Excel进程

            p.Kill();//结束对应Excel进程

        }
        


        /// <summary>
        /// 计算间接测量量H
        /// </summary>
        /// <param name="l">悬线长l平均值</param>
        /// <param name="a">上孔a平均值</param>
        /// <param name="b">下孔b平均值</param>
        /// <returns>返回double类型的间接测量量H</returns>
        public static double Hcalculate(double l, double a, double b)
        {
            decimal a1 = Convert.ToDecimal(Math.Sqrt(3));//转换为decimal以便保留小数

            decimal a2 = a1 / 3;

            double e = Convert.ToDouble(a2);

            double h = Math.Sqrt(Math.Pow(l, 2) - Math.Pow((b * e - a * e), 2));//根据对应公式进行计算

            return h;
        }

        

        /// <summary>
        /// 计算间接测量量H的不确定度
        /// </summary>
        /// <param name="Cal">参数为一个double数组</param>
        /// <returns>返回一个double值</returns>
        public static double Hbuquedingdu(double[] Cal)
        {
            double H1 = Math.Sqrt(Math.Pow(3 * Cal[2] * Cal[5], 2) + ((Math.Pow(Cal[4], 2) + Math.Pow(Cal[5], 2)) * Math.Pow(Cal[0] - Cal[1], 2)));

            decimal H2 = Convert.ToDecimal(H1);

            decimal H3 = Convert.ToDecimal(3 * Cal[6]);

            return Convert.ToDouble(H2 / H3);//转换为decimal以便保留小数
        }
        


        #region 开发过程对测试的一种结束Excel进程的方法，失败，原因不明
        /*   

             public static void QuitExcel(ref Microsoft.Office.Interop.Excel.Application application)
             {
                 application.Quit();
                 try
                 {
                     System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
                 }
                 catch (System.Exception ex)
                 {
                     MessageBox.Show(ex.ToString());
                 }
                 finally
                 {
                     application = null;
                     GC.Collect();
                 }
             }*/

        #endregion



        /// <summary>
        /// 求间接测量量转动惯量I
        /// </summary>
        /// <param name="Cal">参数为一个double数组</param>
        /// <returns>返回一个double类型的转动惯量I</returns>
        public static double Icalculate(double[] Cal)
        {
            const double g = 9.8;//定义常量，重力G

            double Cal7 = Avgn(Cal[7], 1000);

            double Cal0 = Avgn(Cal[0], 100);

            double Cal1 = Avgn(Cal[1], 100);

            double Cal6 = Avgn(Cal[6], 100);

            decimal temp1 = Convert.ToDecimal(12 * Math.Pow(Math.PI, 2) * Cal6);

            decimal temp2 = Convert.ToDecimal(Cal7 * g * Cal0 * Cal1 * Math.Pow(Cal[8], 2));//根据公式计算间接测量量I

            double Itheory1 = Convert.ToDouble(temp2 / temp1);//转换为decimal以便保留小数

            return Itheory1;
        }
        


        /// <summary>
        /// 计算间接测量量I的相对不确定度
        /// </summary>
        /// <param name="Cal">参数是一个double数组，里面有对应的数据，具体参照公式</param>
        /// <returns>返回一个double类型的间接测量量I的相对不确定度</returns>
        public static double Ixiangduibuqueding(double[] Cal)
        {
            decimal[] temp = new decimal[11];

            for (int i = 0; i < temp.Length; i++)
            {
                temp[i] = Convert.ToDecimal(Cal[i]);
            }

            double a = Convert.ToDouble(temp[3] / temp[0]);

            double b = Convert.ToDouble(temp[4] / temp[1]);

            double m = Convert.ToDouble(1 / temp[7]);

            double t = Convert.ToDouble(2 * temp[9] / temp[8]);

            double h = Convert.ToDouble(temp[10] / temp[6]);

            //根据公式计算间接测量量I的相对不确定度
            double I = Math.Sqrt(Math.Pow(a, 2) + Math.Pow(b, 2) + Math.Pow(m, 2) + Math.Pow(t, 2) + Math.Pow(h, 2));

            return I;
        }
        


        /// <summary>
        /// 计算转动惯量I的理论值 
        /// </summary>
        /// <param name="m">圆柱体质量</param>
        /// <param name="r">圆柱体半径</param>
        /// <param name="d">圆柱体到下圆盘中心的距离</param>
        /// <returns>返回double数组类型的理论值</returns>
        public static double[] Itheory(double m, double r2, double[] d)
        {
            double[] d1 = new double[6];

            for (int i = 0; i < d.Length; i++)//遍历数组求得对应半径
            {
                d1[i] = Calculate.Avgn(d[i], 200);
            }

            double r1 = Calculate.Avgn(r2, 1000);

            double m1 = Calculate.Avgn(m, 1000);

            for (int i = 0; i < d.Length; i++)
            {
                Form1.iTheoryOutput[i] = Avgn(m1 * Math.Pow(r1, 2), 2) + m1 * Math.Pow(d1[i], 2);//根据公式计算理论值
            }

            return Form1.iTheoryOutput;
        }
        


        /// <summary>
        /// 计算转动惯量I的实验值
        /// </summary>
        /// <param name="Cal">参数为一个double数组，里面包含对应计算数据，具体参照公式</param>
        /// <param name="t">参数为一个double数组，里面包含对应计算数据，具体参照公式</param>
        /// <returns>返回double数组类型的实验值</returns>
        public static double[] Ireality(double[] Cal, double[] t)
        {
            const double g = 9.8;

            double m0_ = Avgn(Cal[12], 1000);

            double m2_ = Avgn(Cal[13], 1000);

            double a_ = Avgn(Cal[0], 100);

            double b_ = Avgn(Cal[1], 100);

            double h_ = Avgn(Cal[6], 100);

            double I_ = Avgn(Cal[14], 2);//转换数据，保证在国际单位制下进行计算

            decimal H_ = Convert.ToDecimal(24 * Math.Pow(Math.PI, 2) * h_);

            decimal[] con = new decimal[6];

            for (int i = 0; i < t.Length; i++)
            {
                con[i] = (Convert.ToDecimal((m0_ + 2 * m2_) * g * a_ * b_ * Math.Pow(t[i], 2))) / H_;

            }

            for (int i = 0; i < Form1.iRealityOutput.Length; i++)
            {
                Form1.iRealityOutput[i] = (Convert.ToDouble(con[i]) - I_);
            }

            //根据公式进行计算实验值
            return Form1.iRealityOutput;
        }
        


        /// <summary>
        /// 计算转动惯量I的百分误差
        /// </summary>
        /// <param name="iTheoryOutput">参数为一个double数组，里面包含对应转动惯量I的理论值</param>
        /// <param name="iReality">参数为一个double数组，里面包含对应转动惯量I的实验值，具体参照公式</param>
        /// <returns>返回一个double数组类型的转动惯量I的百分误差</returns>
        public static double[] Ibaifenwucha(double[] iTheoryOutput, double[] iReality)
        {
            decimal[] iTheoryTran = new decimal[6];//转换为decimal以便保留小数

            decimal[] iCollect = new decimal[6];

            for (int i = 0; i < iTheoryOutput.Length; i++)
            {
                iTheoryTran[i] = Convert.ToDecimal(iTheoryOutput[i]);

                iCollect[i] = Convert.ToDecimal(Math.Abs((iTheoryOutput[i] - iReality[i])) * 100);//根据公式进行计算

                Form1.iBaifenwucha[i] = Convert.ToDouble(iCollect[i] / iTheoryTran[i]);
            }

            return Form1.iBaifenwucha;
        }


        
        /// <summary>
        /// 四舍六入五凑偶
        /// </summary>
        /// <param name="a">double类型的原始数据</param>
        /// <param name="b">int类型的小数位数</param>
        /// <returns>double类型的，计算后的结果</returns>
        public static double A(double a, int b)
        {
            double aa = Math.Floor(a * Math.Pow(10, b + 1));//乘上相应次方，取得整数

            double bb = Math.Floor(a * Math.Pow(10, b)) * 10;

            double cc = aa - bb;//通过绝对值的方法，取得最后一位数

            string cc1 = Convert.ToString(cc);

            double aaa = Math.Floor((aa / 10)) - Math.Floor((aa / 100)) * 10;//通过绝对值的方法，取得倒数第二位数

            switch (cc1)
            {
                //当最后一位是小于等于4时，舍去最后一位
                case "0":

                case "1":

                case "2":

                case "3":

                case "4":

                    aa = aa - cc;

                    break;

                //当最后一位是五时，判断倒数第二位是奇数还是偶数
                case "5":

                    if (aaa % 2 == 0)
                    {
                        aa = aa - cc;
                    }

                    else
                    {
                        aa = aa + (10 - cc);
                    }
                    break;

                //当最后一位是大于等于6的，向前进一位
                case "6":

                case "7":

                case "8":

                case "9":

                    aa = aa + (10 - cc);

                    break;
            }

            double endresult = Convert.ToDouble((decimal)aa / (decimal)Math.Pow(10, b + 1));//转换结果，使之符合要求四舍六入五凑偶

            return endresult;
        }



        /// <summary>
        /// 生成学生实验数据文件
        /// </summary>
        /// <param name="tb">参数为一个文本框数组，包含从窗体读取的数据</param>
        /// <param name="Sclass">文件路径之一，班级名</param>
        /// <param name="Snumber">文件名之一，学号</param>
        /// <param name="Sname">文件名之一，姓名</param>
        public static void InputData(System.Windows.Forms.TextBox[] tb, string Sclass, string Snumber, string Sname)
        {
            string path = @"e:\三线摆实验学生数据\Sclass\Snumber.xls";//定义文件存储目录

            string path1 = path.Remove(19);//修改路径,方便下面创建对应文件夹，如果不创建文件夹会导致无法保存

            string path2 = path.Replace("Sclass", Sclass);//替换文件夹名为班级名

            string path3 = path1.Replace("Sclass", Sclass); //替换文件夹名为班级名

            string path4 = path2.Replace("Snumber", Snumber + Sname);//替换文件名为学号+姓名

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

            Workbook wbook = app.Workbooks.Add();

            Worksheet worksheet = (Worksheet)wbook.Worksheets[1];

            #region 生成数据

            //时间T1
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[4, 2 + i] = tb[i].Text;
            }

            //上孔a
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[10 + i, 2] = tb[6 + i].Text;
            }

            //下孔b
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[10 + i, 3] = tb[12 + i].Text;
            }

            //悬线长l
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[10 + i, 4] = tb[18 + i].Text;
            }

            //直径d
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[10 + i, 5] = tb[24 + i].Text;
            }

            //孔距2d
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[22, 2 + i] = tb[30 + i].Text;
            }

            //周期T1
            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[23, 2 + i] = tb[36 + i].Text;
            }

            //其他
            worksheet.Cells[25, 2] = tb[42].Text;//下圆盘直径

            worksheet.Cells[26, 2] = tb[43].Text;//下圆盘质量

            worksheet.Cells[27, 2] = tb[44].Text;//圆柱体质量

            //补充表格
            string[] tempRow1 = { "直尺仪器误差", "0.02", "cm", "质量仪器误差", "1", "g", "时间仪器误差", "0.002", "s" };

            for (int i = 0; i < 9; i++)
            {
                worksheet.Cells[1, i + 1] = tempRow1[i];
            }

            string[] tempCol1 = {"表一", " ", "时间", " ", " ", " ", "表2", " ", "1", "2",
            "3","4","5","6"," "," "," "," ","表三"," ",
            " 孔距","周期"," ","下圆盘直径","下圆盘质量","圆柱体质量" };

            for (int i = 0; i < 26; i++)
            {
                worksheet.Cells[2 + i, 1] = tempCol1[i];
            }

            string[] tempRow9 = { "上孔a", "下孔b", "悬线长l", "直径2r" };

            for (int i = 0; i < 4; i++)
            {
                worksheet.Cells[9, 2 + i] = tempRow9[i];
            }

            for (int i = 0; i < 6; i++)
            {
                worksheet.Cells[3, 2 + i] = i + 1;
                worksheet.Cells[21, 2 + i] = i + 1;
            }
            #endregion
            Directory.CreateDirectory(@"e:\三线摆实验学生数据");
            Directory.CreateDirectory(path3);//创建文件夹，如果不创建文件夹会导致无法保存成功
            worksheet.SaveAs(path4, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.DisplayAlerts = true;

            MessageBox.Show("学生实验数据生成成功1个,失败0个。文件目录为" + path4 + "下一位同学请准备输入");

            Kill(app);
        }



    }

}


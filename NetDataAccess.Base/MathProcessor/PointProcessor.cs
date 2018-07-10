using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.MathProcessor
{
    public class PointerProcessor
    { 
        #region PointRotate
        /// <summary> 
        /// 对一个坐标点按照一个中心进行旋转 
        /// </summary> 
        /// <param name="center">中心点</param> 
        /// <param name="p1">要旋转的点</param> 
        /// <param name="angle">旋转角度，笛卡尔直角坐标</param> 
        /// <returns></returns> 
        public static double[] PointRotate(double[] p1, double angle)
        {
            double[] center = new double[] { 0, 0 };
            double[] tmp = new double[2];
            double angleHude = angle * Math.PI / 180;/*角度变成弧度*/
            double x1 = (p1[0] - center[0]) * Math.Cos(angleHude) + (p1[1] - center[1]) * Math.Sin(angleHude) + center[0];
            double y1 = -(p1[0] - center[0]) * Math.Sin(angleHude) + (p1[1] - center[1]) * Math.Cos(angleHude) + center[1];
            tmp[0] = x1;
            tmp[1] = y1;
            return tmp;
        }
        #endregion
         
    }
}

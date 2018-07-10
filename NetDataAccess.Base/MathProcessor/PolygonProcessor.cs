using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.MathProcessor
{
    public class PolygonProcessor
    {
        #region CalculateArea
        /// <summary>  
        /// 计算多边形面积的函数  
        /// (以原点为基准点,分割为多个三角形)  
        /// 定理：任意多边形的面积可由任意一点与多边形上依次两点连线构成的三角形矢量面积求和得出。矢量面积=三角形两边矢量的叉乘。  
        /// </summary>  
        /// <param name="points"></param>  
        /// <returns></returns>  
        public static double CalculateArea(List<double[]> points)
        {
            int iCycle, iCount;
            iCycle = 0;
            double iArea = 0;
            iCount = points.Count;

            for (iCycle = 0; iCycle < iCount; iCycle++)
            {
                iArea = iArea + (points[iCycle][0] * points[(iCycle + 1) % iCount][1] - points[(iCycle + 1) % iCount][0] * points[iCycle][1]);
            }

            return (double)Math.Abs(0.5 * iArea);
        }
        #endregion

        #region PolygonRotate
        /// <summary> 
        /// 对一个坐标点按照一个中心进行旋转 
        /// </summary> 
        /// <param name="points">多边形</param>  
        /// <param name="angle">旋转角度</param> 
        /// <returns></returns> 
        public static List<double[]> PolygonRotate(List<double[]> points, double angle)
        {
            List<double[]> resultPoints = new List<double[]>();
            for (int i = 0; i < points.Count; i++)
            {
                double[] tempPoint = PointerProcessor.PointRotate(points[i], angle);
                resultPoints.Add(tempPoint);
            }
            return resultPoints;
        }
        #endregion

        #region GetMaxRectangle
        /// <summary> 
        /// 获取最大内接矩形
        /// </summary> 
        /// <param name="points">多边形</param>  
        /// <returns></returns> 
        public static List<double[]> GetMaxRectangle(List<double[]> boundaryPoints)
        {
            double maxRectArea = 0;
            List<double[]> maxRectPoints = null;
            for (int i = 0; i < 90; i = i + 5)
            {
                double rectArea = 0;
                List<double[]> transfromBoundaryPoints = PolygonRotate(boundaryPoints, i);
                List<double[]> rectPoints = GetMaxRightRectangle(transfromBoundaryPoints, ref rectArea);
                if (rectArea > maxRectArea)
                {
                    maxRectArea = rectArea;
                    maxRectPoints = PolygonRotate(rectPoints, -i);
                }
            }
            return maxRectPoints;
        }
        #endregion

        #region GetMaxRectangle
        /// <summary> 
        /// 获取最大内接矩形
        /// </summary> 
        /// <param name="points">多边形</param>  
        /// <returns></returns> 
        public static List<double[]> GetMaxRightRectangle(List<double[]> boundaryPoints, ref double maxRectArea)
        {
            int splitCount = 20;
            double[] minPoint = null;
            double[] maxPoint = null;
            GetMinMaxXY(boundaryPoints, ref minPoint, ref maxPoint);
            double xStep = (maxPoint[0] - minPoint[0]) / splitCount;
            double yStep = (maxPoint[1] - minPoint[1]) / splitCount;
            double blockArea = xStep * yStep;

            Dictionary<string, double[]> innerPoints = new Dictionary<string, double[]>();
            for (int x = 0; x < splitCount; x++)
            {
                for (int y = 0; y < splitCount; y++)
                {
                    double[] p = new double[] { minPoint[0] + (x + 0.5) * xStep, minPoint[1] + (y + 0.5) * yStep };
                    if (CheckPolygonContainsPoint(boundaryPoints, p))
                    {
                        innerPoints.Add(x.ToString() + "_" + y.ToString(), p);
                    }
                }
            }

            int maxBlockCount = 0;
            int[] maxAreaIndexes = new int[] { 0, 0, 0, 0 };
            for (int x = 0; x < splitCount; x++)
            {
                for (int i = x; i < splitCount; i++)
                {
                    for (int y = 0; y < splitCount; y++)
                    {
                        for (int j = y; j < splitCount; j++)
                        {
                            if (CheckAllBlockInRectangle(innerPoints, x, y, i, j))
                            {
                                int blockCount = (i - x + 1) * (j - y + 1);
                                if (blockCount > maxBlockCount)
                                {
                                    maxBlockCount = blockCount;
                                    maxAreaIndexes = new int[] { x, y, i, j };
                                }
                            }
                        }
                    }
                }
            }
            List<double[]> maxRectPoints = new List<double[]>();
            if (maxBlockCount > 0)
            {
                maxRectPoints.Add(innerPoints[maxAreaIndexes[0] + "_" + maxAreaIndexes[1]]);
                maxRectPoints.Add(innerPoints[maxAreaIndexes[2] + "_" + maxAreaIndexes[1]]);
                maxRectPoints.Add(innerPoints[maxAreaIndexes[2] + "_" + maxAreaIndexes[3]]);
                maxRectPoints.Add(innerPoints[maxAreaIndexes[0] + "_" + maxAreaIndexes[3]]);
                maxRectPoints.Add(innerPoints[maxAreaIndexes[0] + "_" + maxAreaIndexes[1]]);
            }
            maxRectArea = maxBlockCount * blockArea;
            return maxRectPoints;
        }

        private static bool CheckAllBlockInRectangle(Dictionary<string, double[]> innerPoints, int minX, int minY, int maxX, int maxY)
        {
            for (int i = minX; i <= maxX; i++)
            {
                for (int j = minY; j <= maxY; j++)
                {
                    if (!innerPoints.ContainsKey(i.ToString() + "_" + j.ToString()))
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        #endregion

        #region GetMinMaxXY
        /// <summary> 
        /// 获取多边形的最大最小坐标
        /// </summary> 
        /// <param name="points">多边形</param>  
        /// <returns></returns> 
        public static void GetMinMaxXY(List<double[]> points, ref double[] minPoint, ref double[] maxPoint)
        {
            minPoint = new double[] { double.MaxValue, double.MaxValue };
            maxPoint = new double[] { double.MinValue, double.MinValue };
            for (int i = 0; i < points.Count; i++)
            {
                double[] p = points[i];
                if (p[0] > maxPoint[0])
                {
                    maxPoint[0] = p[0];
                }

                if (p[1] > maxPoint[1])
                {
                    maxPoint[1] = p[1];
                }

                if (p[0] < minPoint[0])
                {
                    minPoint[0] = p[0];
                }

                if (p[1] < minPoint[1])
                {
                    minPoint[1] = p[1];
                }
            }
        }
        #endregion

        /** 
     * 返回一个点是否在一个多边形区域内 
     * 
     * @param mPoints 多边形坐标点列表 
     * @param point   待判断点 
     * @return true 多边形包含这个点,false 多边形未包含这个点。 
     */
        public static bool CheckPolygonContainsPoint(List<double[]> mPoints, double[] point)
        {
            int nCross = 0;
            for (int i = 0; i < mPoints.Count; i++)
            {
                double[] p1 = mPoints[i];
                double[] p2 = mPoints[(i + 1) % mPoints.Count];
                // 取多边形任意一个边,做点point的水平延长线,求解与当前边的交点个数  
                // p1p2是水平线段,要么没有交点,要么有无限个交点  
                if (p1[1] == p2[1])
                    continue;
                // point 在p1p2 底部 --> 无交点  
                if (point[1] < Math.Min(p1[1], p2[1]))
                    continue;
                // point 在p1p2 顶部 --> 无交点  
                if (point[1] >= Math.Max(p1[1], p2[1]))
                    continue;
                // 求解 point点水平线与当前p1p2边的交点的 X 坐标  
                double x = (point[1] - p1[1]) * (p2[0] - p1[0]) / (p2[1] - p1[1]) + p1[0];
                if (x > point[0]) // 当x=point.x时,说明point在p1p2线段上  
                    nCross++; // 只统计单边交点  
            }
            // 单边交点为偶数，点在多边形之外 ---  
            return (nCross % 2 == 1);
        }

        /** 
         * 返回一个点是否在一个多边形边界上 
         * 
         * @param mPoints 多边形坐标点列表 
         * @param point   待判断点 
         * @return true 点在多边形边上,false 点不在多边形边上。 
         */
        public static bool CheckPointInPolygonBoundary(List<double[]> mPoints, double[] point)
        {
            for (int i = 0; i < mPoints.Count; i++)
            {
                double[] p1 = mPoints[i];
                double[] p2 = mPoints[(i + 1) % mPoints.Count];
                // 取多边形任意一个边,做点point的水平延长线,求解与当前边的交点个数  

                // point 在p1p2 底部 --> 无交点  
                if (point[1] < Math.Min(p1[1], p2[1]))
                    continue;
                // point 在p1p2 顶部 --> 无交点  
                if (point[1] > Math.Max(p1[1], p2[1]))
                    continue;

                // p1p2是水平线段,要么没有交点,要么有无限个交点  
                if (p1[1] == p2[1])
                {
                    double minX = Math.Min(p1[0], p2[0]);
                    double maxX = Math.Max(p1[0], p2[0]);
                    // point在水平线段p1p2上,直接return true  
                    if ((point[1] == p1[1]) && (point[0] >= minX && point[0] <= maxX))
                    {
                        return true;
                    }
                }
                else
                { // 求解交点  
                    double x = (point[1] - p1[1]) * (p2[0] - p1[0]) / (p2[1] - p1[1]) + p1[0];
                    if (x == point[0]) // 当x=point.x时,说明point在p1p2线段上  
                        return true;
                }
            }
            return false;
        }  
    }
}

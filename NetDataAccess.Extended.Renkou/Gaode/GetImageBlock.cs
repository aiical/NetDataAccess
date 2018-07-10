using System;
using System.Collections.Generic;
using System.Text;
using NetDataAccess.Base.DLL;
using NetDataAccess.Base.Config;
using System.Threading;
using System.Windows.Forms;
using mshtml;
using NetDataAccess.Base.Definition;
using System.IO;
using NetDataAccess.Base.Common;
using NPOI.SS.UserModel;
using NetDataAccess.Base.EnumTypes;
using NetDataAccess.Base.UI;
using Newtonsoft.Json.Linq;
using HtmlAgilityPack;
using NetDataAccess.Base.Writer;
using NetDataAccess.Base.DB;
using System.Drawing;
using NetDataAccess.Base.Reader;

namespace NetDataAccess.Extended.Renkou.Gaode
{
    /// <summary>
    /// GetImageBlock
    /// </summary>
    public class GetImageBlock : CustomProgramBase
    {
        public bool Run(string parameters, IListSheet listSheet)
        {
            return this.GetImageBlockInfo(parameters, listSheet);
        }
        private bool GetImageBlockInfo(string parameters, IListSheet listSheet)
        { 
            bool succeed = true;
            string exportDir = this.RunPage.GetExportDir();
            string pageSourceDir = this.RunPage.GetDetailSourceFileDir();

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                GetBlockInfo(exportDir, listRow, pageSourceDir);
            }

            MergeInfoFile(exportDir, listSheet);

            return succeed;
        } 

        private double ToBaiduY(int gaodeY, int yp, int z)
        {
            switch (z)
            {
                case 18:
                    double y = ((double)gaodeY + (double)(yp - 128) / (double)256) * (-152.88324536) + 20039240.499163;
                    y = y * 180 / 20037508.34;
                    y = 180 * (2 * Math.Atan(Math.Exp(y * Math.PI / 180)) - Math.PI / 2) / Math.PI;
                    return y;
                default:
                    throw new Exception("无法处理坐标, z=" + z.ToString());
            }
        }

        private double ToBaiduX(int gaodeX, int xp, int z)
        {
            switch (z)
            {
                case 18:
                    double x = ((double)gaodeX + (double)(xp - 128) / (double)256) * 152.875047753 - 20036932.2651061;
                    x = x * 180 / 20037508.34;
                    return x;
                default:
                    throw new Exception("无法处理坐标, z=" + z.ToString());
            }
        }

        private void MergeInfoFile(string exportDir, IListSheet listSheet)
        {

            string allBlockInfoPath = Path.Combine(exportDir, "爬取结果.csv");
            StringBuilder ss = new StringBuilder();            

            Dictionary<string, int> allBlockInfoDic = new Dictionary<string, int>();
            allBlockInfoDic.Add("x", 0);
            allBlockInfoDic.Add("y", 1);
            allBlockInfoDic.Add("z", 2);
            allBlockInfoDic.Add("xp", 3);
            allBlockInfoDic.Add("yp", 4);
            allBlockInfoDic.Add("v", 5);
            allBlockInfoDic.Add("time", 6);
            CsvWriter allBlockInfoCW = new CsvWriter(allBlockInfoPath, allBlockInfoDic);

            for (int i = 0; i < listSheet.RowCount; i++)
            {
                Dictionary<string, string> listRow = listSheet.GetRow(i);
                string x = listRow["x"];
                string y = listRow["y"];
                string z = listRow["z"];
                string time = listRow["time"];
                string tempDir = Path.Combine(exportDir, "temp");
                string blockInfoPath = Path.Combine(tempDir, x + "_" + y + "_" + z + "_" + time + ".csv");
                CsvReader csvReader = new CsvReader(blockInfoPath);
                int pCount = csvReader.GetRowCount();
                if (pCount > 0)
                {
                    for (int pIndex = 0; pIndex < pCount; pIndex++)
                    {
                        Dictionary<string, string> pValues = csvReader.GetFieldValues(pIndex);
                        string xp = pValues["xp"];
                        string yp = pValues["yp"];
                        string v = pValues["v"];
                        allBlockInfoCW.AddRow(pValues);

                        if(ss.Length!=0)
                        {
                            ss.Append(",");
                            ss.AppendLine();
                        }
                        ss.Append("    {\"x\":" + x.ToString() + "." + xp + ",\"y\":" + y.ToString() + "." + yp + ", \"lng\":" + this.ToBaiduX(int.Parse(x), int.Parse(xp), int.Parse(z)) + ",\"lat\":" + this.ToBaiduY(int.Parse(y), int.Parse(yp), int.Parse(z)) + ",\"count\":" + double.Parse(v).ToString() + "}");

                    }
                }
            }
            allBlockInfoCW.SaveToDisk();
            
            string allBlockInfoTextPath = Path.Combine(exportDir, "爬取结果.txt");
            FileHelper.SaveTextToFile(ss.ToString(), allBlockInfoTextPath);
        }

        private void GetBlockInfo(string exportDir, Dictionary<string, string> listRow, string pageSourceDir)
        {
            string detailUrl = listRow["detailPageUrl"];
            string x = listRow["x"];
            string y = listRow["y"];
            string z = listRow["z"];
            string time = listRow["time"];

            string tempDir = Path.Combine(exportDir, "temp");
            string blockInfoPath = Path.Combine(tempDir, x + "_" + y + "_" + z + "_" + time + ".csv");
            if (!File.Exists(blockInfoPath))
            {

                string localFilePath = this.RunPage.GetFilePath(detailUrl, pageSourceDir);

                Dictionary<string, int> blockInfoDic = new Dictionary<string, int>();
                blockInfoDic.Add("x", 0);
                blockInfoDic.Add("y", 1);
                blockInfoDic.Add("z", 2);
                blockInfoDic.Add("xp", 3);
                blockInfoDic.Add("yp", 4);
                blockInfoDic.Add("v", 5);
                blockInfoDic.Add("time", 6);

                CsvWriter blockInfoCW = new CsvWriter(blockInfoPath, blockInfoDic);



                int blackSize = 16;
                Bitmap img = new Bitmap(localFilePath);
                List<Point> allPoints = new List<Point>();
                Dictionary<Point, float> p2hs = new Dictionary<Point, float>();

                for (int xx = 0; xx < img.Width; xx = xx + blackSize)
                {
                    for (int yy = 0; yy < img.Width; yy = yy + blackSize)
                    {
                        float sumH = 0;
                        int X = xx +blackSize/2;
                        int Y = yy +blackSize/2;
                        List<Point> sameHPoints = new List<Point>();

                        for (int i = 0; i < blackSize; i++)
                        {
                            if (xx + i < img.Width)
                            {
                                for (int j = 0; j < blackSize; j++)
                                {
                                    if (yy + j < img.Height)
                                    {
                                        Color c = img.GetPixel(xx + i, yy + j);
                                        float h = 360 - c.GetHue();
                                        if (h > 0 && h < 360)
                                        {
                                            sumH += h;
                                        }
                                        /*
                                        if (h > maxH && h != 360)
                                        {
                                            maxH = h;
                                            X = xx + i;
                                            Y = yy + j;
                                            sameHPoints.Clear();
                                            sameHPoints.Add(new Point(X, Y));
                                        }
                                        else if (h == maxH && h != 360)
                                        {
                                            sameHPoints.Add(new Point(xx + i, yy + j));
                                        }*/
                                    }
                                }
                            }
                        }
                        float avgH = sumH / (blackSize * blackSize);
                        if (avgH < 360 && avgH > 0)
                        { 
                            Point p = new Point(X, Y);
                            allPoints.Add(p);
                            p2hs.Add(p, avgH);
                        }
                    }
                }
                /*
                List<Point> remainPoints = new List<Point>();
                while (allPoints.Count > 0)
                {
                    Point maxHP = getMaxHPoint(allPoints, p2hs);
                    remainPoints.Add(maxHP);
                    allPoints.Remove(maxHP); 
                    List<Point> deletePoints = new List<Point>();
                    foreach (Point p in allPoints)
                    {
                        if ((maxHP.X - p.X) * (maxHP.X - p.X) + (maxHP.Y - p.Y) * (maxHP.Y - p.Y) < blackSize / 2 * blackSize / 2)
                        {
                            deletePoints.Add(p);
                        }
                    }
                    foreach (Point p in deletePoints)
                    {
                        allPoints.Remove(p);
                    } 
                }*/
                foreach (Point p in allPoints)
                {

                    float h = p2hs[p];
                    Dictionary<string, string> cityReport = new Dictionary<string, string>();
                    cityReport.Add("x", x);
                    cityReport.Add("y", y);
                    cityReport.Add("z", z);
                    cityReport.Add("xp", p.X.ToString());
                    cityReport.Add("yp", p.Y.ToString());
                    cityReport.Add("v", h.ToString());
                    cityReport.Add("time", time);
                    blockInfoCW.AddRow(cityReport);
                }

                blockInfoCW.SaveToDisk();
            }
        }

        private Point getMaxHPoint(List<Point> allPoints, Dictionary<Point, float> p2hs)
        {
            Point maxP = new Point();
            float maxH = 0;
            foreach (Point p in allPoints)
            {
                float h = p2hs[p];
                if (maxH < h)
                {
                    maxH = h;
                    maxP = p;
                }
            }
            return maxP;
        }
    }
}
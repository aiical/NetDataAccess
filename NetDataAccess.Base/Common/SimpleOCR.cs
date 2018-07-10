using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using Tesseract;

namespace NetDataAccess.Base.Common
{
    /// <summary>
    /// 简单的文字识别功能
    /// </summary>
    public class SimpleOCR
    {
        #region 识别一行文字
        /// <summary>
        /// 识别一行文字
        /// </summary>
        /// <param name="data"></param>
        /// <param name="tessractData"></param>
        /// <param name="language"></param>
        /// <param name="variables"></param>
        /// <returns></returns>
        public static string OCRSingleLine(byte[] data, string tessractData, string language, Dictionary<string, string> variables)
        {
            Bitmap bmp = null;
            try
            {
                bmp = To16Bmp(data);
                return OCRSingleLine(bmp, tessractData, language, variables);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (bmp != null)
                {
                    bmp.Dispose();
                }
            }
        }
        #endregion

        #region 识别一行文字
        public static string OCRSingleLine(string filePath, string tessractData, string language, Dictionary<string, string> variables)
        {
            Bitmap bmp = null;
            try
            {
                bmp = To16Bmp(filePath);
                return OCRSingleLine(bmp, tessractData, language, variables);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (bmp != null)
                {
                    bmp.Dispose();
                }
            }
        }
        #endregion

        #region 识别一行文字
        public static string OCRSingleLine(Bitmap bmp, string tessractData, string language, Dictionary<string, string> variables)
        {
            TesseractEngine test = new TesseractEngine(tessractData, language, EngineMode.Default);
            foreach (string key in variables.Keys)
            {
                string value = variables[key];
                test.SetVariable(key, value);
            }

            Graphics graph = Graphics.FromImage(bmp);
            Page page = test.Process(bmp, pageSegMode: PageSegMode.SingleLine);
            string txt = page.GetText();
            return txt;
        }
        #endregion

        #region 转化成图片
        public static Bitmap ToBmp(byte[] sourceData)
        {
            MemoryStream sourceMs = null; 
            try
            {
                sourceMs = new MemoryStream(sourceData);
                return new Bitmap(sourceMs); 
            }
            finally
            { 
            }
        }
        #endregion

        #region 转化成图片
        public static Bitmap To16Bmp(byte[] sourceData)
        {
            MemoryStream sourceMs = null;
            Bitmap bitmap = null;  
            try
            {
                sourceMs = new MemoryStream(sourceData);
                bitmap = new Bitmap(sourceMs); 
                return To16Bmp(bitmap);
            }
            finally
            {
                if (bitmap != null)
                {
                    bitmap.Dispose();
                }
                if (sourceMs != null)
                {
                    sourceMs.Dispose();
                } 
            }
        }
        #endregion

        #region 转化成图片
        public static Bitmap To16Bmp(string filePath)
        {
            Bitmap bitmap = null; 
            try
            {
                bitmap = new Bitmap(filePath);
                return To16Bmp(bitmap);
            }
            finally
            {
                if (bitmap != null)
                {
                    bitmap.Dispose();
                } 
            }
        }
        #endregion

        #region 转化成图片
        public static Bitmap To16Bmp(Bitmap sourceBmp)
        {
            Bitmap bitmap = sourceBmp;
            Bitmap bitmap2 = null;
            MemoryStream ms = null;
            try
            { 
                BitmapData data = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadOnly, PixelFormat.Format16bppArgb1555);
                bitmap2 = new Bitmap(bitmap.Width, bitmap.Height, data.Stride, PixelFormat.Format16bppArgb1555, data.Scan0);

                ms = new MemoryStream();
                bitmap2.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                Bitmap bmp = new Bitmap(ms);
                bitmap.UnlockBits(data);
                return bmp;
            }
            finally
            { 
                if (bitmap2 != null)
                {
                    bitmap2.Dispose();
                } 
            }
        }
        #endregion

        #region 图像灰度化
        /// <summary> 
        /// 图像灰度化 
        /// </summary> 
        /// <param name="bmp"></param> 
        /// <returns></returns> 
        public static Bitmap ToGray(Bitmap bmp)
        {
            for (int i = 0; i < bmp.Width; i++)
            {
                for (int j = 0; j < bmp.Height; j++)
                {
                    //获取该点的像素的RGB的颜色 
                    Color color = bmp.GetPixel(i, j);
                    //利用公式计算灰度值 
                    int gray = (int)(color.R * 0.3 + color.G * 0.59 + color.B * 0.11);
                    Color newColor = Color.FromArgb(gray, gray, gray);
                    bmp.SetPixel(i, j, newColor);
                }
            }
            return bmp;
        }
        #endregion

        #region 图像灰度化（不是标准的二值化算法，这里只是处理如果不是背景色，那么就是黑色）
        /// <summary> 
        /// 图像二值化（不是标准的二值化算法，这里只是处理如果不是背景色，那么就是黑色）
        /// </summary> 
        /// <param name="bmp"></param> 
        /// <returns></returns> 
        public static Bitmap ConvertTo2Color(Bitmap bmp, Color backColor)
        { 
            for (int i = 0; i < bmp.Width; i++)
            {
                for (int j = 0; j < bmp.Height; j++)
                {
                    //获取该点的像素的RGB的颜色 
                    Color color = bmp.GetPixel(i, j);
                    if (!(color.R == backColor.R && color.G == backColor.G && color.B == backColor.B))
                    {
                        bmp.SetPixel(i, j, Color.Black);
                    } 
                }
            }
            return bmp;
        }
        #endregion

        #region 去除噪音
        /// <summary>
        /// 去除噪音（不是标准的去噪算法，如果此点周围出现小于exceptCount个不是toColor的，那么就把此点替换成toColor）
        /// </summary>
        /// <param name="sourceImg"></param>
        /// <param name="squareSize"></param>
        /// <param name="toColor"></param>
        /// <returns></returns>
        public static Bitmap ReplaceByAround(Bitmap sourceImg, int squareSize, int enableExceptCount, Color toColor)
        {
            for (int i = 0; i < sourceImg.Width; i++)
            {
                for (int j = 0; j < sourceImg.Height; j++)
                {
                    //获取该点的像素的RGB的颜色 
                    Color color = sourceImg.GetPixel(i, j);
                    if (!(color.R == toColor.R && color.G == toColor.G && color.B == toColor.B))
                    {
                        int count = 0;
                        for (int x = i - squareSize; x <= i + squareSize; x++)
                        {
                            for (int y = j - squareSize; y <= j + squareSize; y++)
                            {
                                if (x != i || y != j)
                                {
                                    if (!CheckIsColor(sourceImg, x, y, toColor, true))
                                    {
                                        count++;
                                    }
                                }
                            }
                        }
                        if (count < enableExceptCount)
                        {
                            sourceImg.SetPixel(i, j, toColor);
                        }
                    }
                }
            }
            return sourceImg;
        }
        #endregion

        #region 判断是否为某个颜色
        private static bool CheckIsColor(Bitmap img, int x, int y, Color color, bool defaultValue)
        {
            if(x <0 || y<0||x>=img.Width||y>=img.Height)
            {
                return defaultValue;
            }
            else
            {
                Color c = img.GetPixel(x, y);
                return (c.R == color.R && c.G == color.G && c.B == color.B);
            }
        }
        #endregion 
    }
}

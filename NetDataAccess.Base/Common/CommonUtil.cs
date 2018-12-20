using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Windows.Forms; 

namespace NetDataAccess.Base.Common
{
    /// <summary>
    /// 通用方法类
    /// 提供文件、简单类型数据的通用处理方法
    /// </summary>
    public static class CommonUtil
    {
        #region 判断字符串是否为空（null或者长度为0的字符串）
        /// <summary>
        /// 判断字符串是否为空（null或者长度为0的字符串）
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsNullOrBlank(string value)
        {
            return value == ""
                || value == null
                || (object)value == DBNull.Value;
        }
        #endregion

        #region 判断是否为空（null或者DBNnull）
        /// <summary>
        /// 判断是否为空（null或者DBNnull）
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        public static bool IsNullOrDBNul(object value)
        {
            return value == null || value == DBNull.Value;
        }
        #endregion

        #region 提示
        /// <summary>
        /// 提示
        /// </summary>
        /// <param name="title"></param>
        /// <param name="msg"></param>
        public static void Alert(string title, string msg)
        {
            MessageBox.Show(msg, title);
        }
        #endregion

        #region 确认
        /// <summary>
        /// 确认
        /// </summary>
        /// <param name="title"></param>
        /// <param name="msg"></param>
        public static bool Confirm(string title, string msg)
        {
            return MessageBox.Show(msg, title, MessageBoxButtons.OKCancel) == DialogResult.OK;
        }
        #endregion

        #region 文件名字符替换
        /// <summary>
        /// 文件名字符替换
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="replaceStr"></param>
        /// <returns></returns>
        public static string ProcessFileName(string fileName, string replaceStr)
        {
            string[] strs = new string[] { ":", ";", "/", "\\", "|", ",", "*", "?", "\"", "<", ">", "\r", "\n" };
            foreach (string str in strs)
            {
                fileName = fileName.Replace(str, replaceStr);
            }
            return fileName;
        }
        #endregion 

        #region 获取异常信息
        /// <summary>
        /// 获取异常信息
        /// </summary>
        /// <param name="ex"></param>
        /// <returns></returns>
        public static string GetExceptionAllMessage(Exception ex)
        {
            StringBuilder errors = new StringBuilder();
            while (ex != null)
            {
                errors.AppendLine(ex.Message);
                ex = ex.InnerException;
            }
            return errors.ToString();
        }
        #endregion

        #region 递归创建路径中的文件夹
        /// <summary>
        /// 递归创建路径中的文件夹
        /// </summary>
        /// <param name="filePath"></param>
        public static void CreateFileDirectory(string filePath)
        {
            string tempPath = Path.GetDirectoryName(filePath);
            if (!Directory.Exists(tempPath))
            {
                CreateFileDirectory(tempPath);
                Directory.CreateDirectory(tempPath);
            }
        }
        #endregion

        #region 判断文件是否在用
        public static bool IsFileInUse(string filePath)
        {
            bool inUse = true;

            FileStream fs = null;
            try
            {

                fs = new FileStream(filePath, FileMode.Open, FileAccess.Read,

                FileShare.None);

                inUse = false;
            }
            catch
            {

            }
            finally
            {
                if (fs != null)

                    fs.Close();
            }
            return inUse;//true表示正在使用,false没有使用
        }
        #endregion

        #region InitStringIndexDic
        public static Dictionary<string, int> InitStringIndexDic(string[] strs)
        {
            Dictionary<string, int> dic = new Dictionary<string, int>();
            for (int i = 0; i < strs.Length; i++)
            {
                dic.Add(strs[i], i);
            }
            return dic;
        }
        #endregion

        #region InitStringIndexDic
        public static string StringArrayToString(string[] strs, string spliter)
        {
            StringBuilder ss = new StringBuilder();
            for (int i = 0; i < strs.Length; i++)
            {
                ss.Append(ss.Length == 0 ? "" : spliter);
                ss.Append(strs[i]);
            }
            return ss.ToString();
        }
        #endregion

        #region 将字符串中的ASCII码转义为字符
        /// <summary>
        /// ASCII码转字符
        /// </summary>
        /// <param name="asciiCode"></param>
        /// <returns></returns>
        public static string Chr(int asciiCode)
        {
            if (asciiCode >= 0 && asciiCode <= 255)
            {
                System.Text.ASCIIEncoding asciiEncoding = new System.Text.ASCIIEncoding();
                byte[] byteArray = new byte[] { (byte)asciiCode };
                string strCharacter = asciiEncoding.GetString(byteArray);
                return (strCharacter);
            }
            else
            {
                throw new Exception("ASCII Code is not valid.");
            }
        }

        private static Dictionary<string, string> _AsciiStrDic = null;
        private static Dictionary<string, string> AsciiStrDic
        {
            get
            {
                if (_AsciiStrDic == null)
                {
                    _AsciiStrDic = new Dictionary<string, string>();
                    for (int i = 32; i <= 126; i++)
                    {
                        _AsciiStrDic.Add(i.ToString(), Chr(i));
                    }
                }
                return _AsciiStrDic;
            }
        }
        public static string ReplaceAsciiByString(string sourceStr)
        {
            foreach (string ascii in AsciiStrDic.Keys)
            {
                sourceStr = Regex.Replace(sourceStr, "&#" + ascii + ";", AsciiStrDic[ascii]);
            }
            return sourceStr;
        }
        #endregion

        #region html转义
        /// <summary>
        /// HtmlDecode
        /// </summary>
        /// <param name="sourceStr"></param>
        /// <returns></returns>
        public static string HtmlDecode(string sourceStr)
        {
            return HttpUtility.HtmlDecode(sourceStr);
        }
        /// <summary>
        /// HtmlEncode
        /// </summary>
        /// <param name="sourceStr"></param>
        /// <returns></returns>
        public static string HtmlEncode(string sourceStr)
        {
            return HttpUtility.HtmlEncode(sourceStr);
        }
        #endregion

        #region url转义
        /// <summary>
        /// UrlDecode
        /// </summary>
        /// <param name="sourceStr"></param>
        /// <returns></returns>
        public static string UrlDecode(string sourceStr)
        {
            return HttpUtility.UrlDecode(sourceStr);
        }
        /// <summary>
        /// UrlEncode
        /// </summary>
        /// <param name="sourceStr"></param>
        /// <returns></returns>
        public static string UrlEncode(string sourceStr)
        {
            return HttpUtility.UrlEncode(sourceStr);
        }
        #endregion

        public static string StringToHexString(string s, Encoding encode)
        {
            byte[] b = encode.GetBytes(s);//按照指定编码将string编程字节数组
            string result = string.Empty;
            for (int i = 0; i < b.Length; i++)//逐字节变为16进制字符，以%隔开
            {
                result += "%" + Convert.ToString(b[i], 16);
            }
            return result;
        }

        #region 将字符串里的&amp;替换为&
        /// <summary>
        /// 将字符串里的&amp;替换为&
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string UrlDecodeSymbolAnd(string url)
        {
            return url == null ? null : url.Replace("&amp;", "&");
        }
        #endregion 

        #region MD5加密
        public static string MD5Crypto(string sourceString)
        {
            MD5 md5 = new MD5CryptoServiceProvider();
            byte[] result = Encoding.Default.GetBytes(sourceString);
            byte[] output = md5.ComputeHash(result);
            string destString = BitConverter.ToString(output);
            return destString;
        }
        #endregion

        #region 判断是否为汉字
        /// <summary>
        /// 用 正则表达式 判断字符是不是汉字
        /// </summary>
        /// <param name="text">待判断字符或字符串</param>
        /// <returns>真：是汉字；假：不是</returns>
        public static bool CheckStringChineseReg(string text)
        {
            return System.Text.RegularExpressions.Regex.IsMatch(text, @"[\u4e00-\u9fbb]+$");
        }
        #endregion
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace NetDataAccess.Base.Common
{
    public class FileHelper
    {
        #region 未知的文件类型
        public static string UnknownFileType = "unknown";
        #endregion

        #region FilesHeader
        private static Dictionary<string, object[]> _FilesHeader = null;
        public static Dictionary<string, object[]> FilesHeader
        {
            get
            {
                if (_FilesHeader == null)
                {
                    _FilesHeader = new Dictionary<string, object[]>();

                    #region 文件头说明
                    /*
                    JPEG (jpg)，文件头：FFD8FF  
                    PNG (png)，文件头：89504E47  
                    GIF (gif)，文件头：47494638  
                    TIFF (tif)，文件头：49492A00  
                    Windows Bitmap (bmp)，文件头：424D  
                    CAD (dwg)，文件头：41433130  
                    Adobe Photoshop (psd)，文件头：38425053  
                    Rich Text Format (rtf)，文件头：7B5C727466  
                    XML (xml)，文件头：3C 3F 78 6D 6C;  60,63,120,109,108
                    HTML (html)，文件头：68 74 6D 6C 3E; 104,116,109,108,62
                    Email [thorough only] (eml)，文件头：44656C69766572792D646174653A  
                    Outlook Express (dbx)，文件头：CFAD12FEC5FD746F  
                    Outlook (pst)，文件头：2142444E  
                    MS Word/Excel (xls.or.doc)，文件头：D0CF11E0  
                    MS Access (mdb)，文件头：5374616E64617264204A  
                    WordPerfect (wpd)，文件头：FF575043  
                    Postscript (eps.or.ps)，文件头：252150532D41646F6265  
                    Adobe Acrobat (pdf)，文件头：255044462D312E  
                    Quicken (qdf)，文件头：AC9EBD8F  
                    Windows Password (pwl)，文件头：E3828596  
                    ZIP Archive (zip)，文件头：504B0304  
                    RAR Archive (rar)，文件头：52617221  
                    Wave (wav)，文件头：57415645  
                    AVI (avi)，文件头：41564920  
                    Real Audio (ram)，文件头：2E7261FD  
                    Real Media (rm)，文件头：2E524D46  
                    MPEG (mpg)，文件头：000001BA  
                    MPEG (mpg)，文件头：000001B3  
                    Quicktime (mov)，文件头：6D6F6F76  
                    Windows Media (asf)，文件头：3026B2758E66CF11  
                    MIDI (mid)，文件头：4D546864 
                    */
                    #endregion

                    _FilesHeader.Add("pdf", new object[] { 
                        new byte[] { 37, 80, 68, 70, 45, 49, 46, 53 },
                        new byte[] { 37, 80, 68, 70, 45 }});
                    _FilesHeader.Add("docx", new object[] { new object[] { new byte[] { 80, 75, 3, 4, 20, 0, 6, 0, 8, 0, 0, 0, 33 }, new Regex(@"word/_rels/document\.xml\.rels", RegexOptions.IgnoreCase) } });
                    _FilesHeader.Add("xlsx", new object[] { new object[] { new byte[] { 80, 75, 3, 4, 20, 0, 6, 0, 8, 0, 0, 0, 33 }, new Regex(@"xl/_rels/workbook\.xml\.rels", RegexOptions.IgnoreCase) } });
                    _FilesHeader.Add("pptx", new object[] { new object[] { new byte[] { 80, 75, 3, 4, 20, 0, 6, 0, 8, 0, 0, 0, 33 }, new Regex(@"ppt/_rels/presentation\.xml\.rels", RegexOptions.IgnoreCase) } });
                    _FilesHeader.Add("doc", new object[] { new object[] { new byte[] { 208, 207, 17, 224, 161, 177, 26, 225 }, new Regex(@"microsoft( office)? word(?![\s\S]*?microsoft)", RegexOptions.IgnoreCase) } });
                    _FilesHeader.Add("xls", new object[] { new object[] { new byte[] { 208, 207, 17, 224, 161, 177, 26, 225 }, new Regex(@"microsoft( office)? excel(?![\s\S]*?microsoft)", RegexOptions.IgnoreCase) } });
                    _FilesHeader.Add("ppt", new object[] { new object[] { new byte[] { 208, 207, 17, 224, 161, 177, 26, 225 }, new Regex(@"c.u.r.r.e.n.t. .u.s.e.r(?![\s\S]*?[a-z])", RegexOptions.IgnoreCase) } });
                    _FilesHeader.Add("avi", new object[] { new byte[] { 65, 86, 73, 32 } });
                    _FilesHeader.Add("mpg", new object[] { new byte[] { 0, 0, 1, 0xBA } });
                    _FilesHeader.Add("mpeg", new object[] { new byte[] { 0, 0, 1, 0xB3 } });
                    _FilesHeader.Add("rar", new object[] { new byte[] { 82, 97, 114, 33, 26, 7 } });
                    _FilesHeader.Add("zip", new object[] { new byte[] { 80, 75, 3, 4 } });
                    _FilesHeader.Add("gif", new object[] { new byte[] { 71, 73, 70, 56, 57, 97 } });
                    _FilesHeader.Add("bmp", new object[] { new byte[] { 66, 77 } });
                    _FilesHeader.Add("jpg", new object[] { new byte[] { 255, 216, 255 } });
                    _FilesHeader.Add("png", new object[] { new byte[] { 137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82 } });
                    _FilesHeader.Add("xml", new object[] { new byte[] { 60, 63, 120, 109, 108 } });
                    _FilesHeader.Add("html", new object[] { 
                        new byte[] { 104, 116, 109, 108, 62 },
                        new byte[] { 60, 115, 99, 114, 105, 112, 116, 32 },
                        new byte[] { 60, 33, 68, 79, 67, 84, 89, 80, 69, 32, 104, 116, 109, 108 }});
                }
                return _FilesHeader;
            }
        }
        #endregion

        #region CheckFileType
        public static string CheckFileType(Stream str)
        {
            #region 使用文件头判断
            string FileExt = FileHelper.UnknownFileType;
            foreach (string ext in FilesHeader.Keys)
            {
                object[] headerObjects = FilesHeader[ext];
                for (int h = 0; h < headerObjects.Length; h++)
                {
                    byte[] header = headerObjects[h].GetType() == (new byte[] { }).GetType() ? (byte[])headerObjects[h] : (byte[])(((object[])headerObjects[h])[0]);
                    byte[] test = new byte[header.Length];
                    str.Position = 0;
                    str.Read(test, 0, test.Length);
                    bool same = true;
                    for (int i = 0; i < test.Length; i++)
                    {
                        if (test[i] != header[i])
                        {
                            same = false;
                            break;
                        }
                    }
                    if (headerObjects[h].GetType() != (new byte[] { }).GetType() && same)
                    {
                        object[] obj = (object[])headerObjects[h];
                        bool exists = false;
                        if (obj[1].GetType().ToString() == "System.Int32")
                        {
                            for (int ii = 2; ii < obj.Length; ii++)
                            {
                                if (str.Length >= (int)obj[1])
                                {
                                    str.Position = str.Length - (int)obj[1];
                                    byte[] more = (byte[])obj[ii];
                                    byte[] testmore = new byte[more.Length];
                                    str.Read(testmore, 0, testmore.Length);
                                    if (Encoding.GetEncoding(936).GetString(more) == Encoding.GetEncoding(936).GetString(testmore))
                                    {
                                        exists = true;
                                        break;
                                    }
                                }
                            }
                        }
                        else if (obj[1].GetType().ToString() == "System.Text.RegularExpressions.Regex")
                        {
                            Regex re = (Regex)obj[1];
                            str.Position = 0;
                            byte[] buffer = new byte[(int)str.Length];
                            str.Read(buffer, 0, buffer.Length);
                            string txt = Encoding.ASCII.GetString(buffer);
                            if (re.IsMatch(txt))
                            {
                                exists = true;
                            }
                        }
                        if (!exists)
                        {
                            same = false;
                        }
                    }
                    if (same)
                    {
                        FileExt = ext;
                        break;
                    }
                }
            }
            #endregion

            return FileExt;
        }
        #endregion

        #region 获取文本文件字符串
        public static string GetTextFromFile(string filePath)
        {
            return GetTextFromFile(filePath, Encoding.UTF8);
        }
        public static string GetTextFromFile(string filePath, Encoding encoding)
        {
            TextReader tr = null;
            try
            {
                tr = new StreamReader(filePath, encoding);
                string fileText = tr.ReadToEnd();
                return fileText;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (tr != null)
                {
                    tr.Close();
                    tr.Dispose();
                    tr = null;
                }
            }
        }
        #endregion

        #region 保存文本文件字符串
        public static void SaveTextToFile(string text, string filePath)
        {
            SaveTextToFile(text, filePath, Encoding.UTF8);
        }
        public static void SaveTextToFile(string text, string filePath, Encoding encoding)
        {
            TextWriter tw = null;
            try
            {
                tw = new StreamWriter(filePath, false, encoding);
                tw.Write(text);
                tw.Flush();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (tw != null)
                {
                    tw.Close();
                    tw.Dispose();
                    tw = null;
                }
            }
        }
        #endregion
    }
}

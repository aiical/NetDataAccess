using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class Html2Txt
    {
        public static bool Html2TxtByHtmlAgilityPack(string sourceFile, string destFile, bool deleteExistFile)
        {
            return Html2TxtByHtmlAgilityPack(sourceFile, destFile, deleteExistFile, "utf-8");
        }

        public static bool Html2TxtByHtmlAgilityPack(string sourceFile, string destFile, bool deleteExistFile, string encoding)
        {
            bool needConvert = true;
            if (File.Exists(destFile))
            {
                if (deleteExistFile)
                {
                    File.Delete(destFile);
                }
                else
                {
                    needConvert = false;
                }
            }

            if (needConvert)
            {
                string parentDir = Directory.GetParent(destFile).FullName;
                CreateDir(parentDir);
                StreamWriter htmlSW = null;
                try
                {

                    HtmlAgilityPack.HtmlDocument htmlDoc = new HtmlAgilityPack.HtmlDocument();
                    htmlDoc.Load(sourceFile, Encoding.GetEncoding(encoding));
                    string text = htmlDoc.DocumentNode.InnerText;
                    htmlSW = new StreamWriter(destFile, false);
                    htmlSW.Write(text);
                    return true;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    if (htmlSW != null)
                    {
                        htmlSW.Close();
                        htmlSW.Dispose();
                    }
                }
            }
            else
            {
                return true;
            }
        }

        private static void CreateDir(string dir)
        {
            if (!Directory.Exists(dir))
            {
                string parentDir = Directory.GetParent(dir).FullName;
                CreateDir(parentDir);
                Directory.CreateDirectory(dir);
            }
        }
    }
}

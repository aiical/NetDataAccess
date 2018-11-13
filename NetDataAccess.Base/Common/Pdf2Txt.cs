using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class Pdf2Txt
    {
        public static bool Pdf2TxtByITextSharp(string sourceFile, string destFile, bool deleteExistFile)
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
                StreamWriter swPdfChange = null;
                try
                {
                    StringBuilder s = new StringBuilder();

                    using (PdfReader reader = new PdfReader(sourceFile))
                    {
                        int pageCount = reader.NumberOfPages;
                        for (int i = 1; i <= pageCount; i++)
                        {
                            try
                            {
                                MyTexExStrat strategy = new MyTexExStrat();
                                PdfTextExtractor.GetTextFromPage(reader, i, strategy);
                                string content = strategy.GetFullText();
                                s.Append(content);
                            }
                            catch (Exception ex)
                            {
                                throw ex;
                            }
                        }
                    }

                    swPdfChange = new StreamWriter(destFile, false);
                    swPdfChange.Write(s.ToString());
                    return true;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    if (swPdfChange != null)
                    {
                        swPdfChange.Close();
                        swPdfChange.Dispose();
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

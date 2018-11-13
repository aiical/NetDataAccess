using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class PdfSpliter
    {
        public static string[] ExtractPages(string sourcePdfPath, string outputPdfDir)
        {
            PdfReader reader = null;
            Document sourceDocument = null;
            PdfCopy pdfCopyProvider = null;
            PdfImportedPage importedPage = null;
            List<string> filePaths = new List<string>();
            try
            {
                reader = new PdfReader(sourcePdfPath);
                int pageCount = reader.NumberOfPages;
                for (int i = 1; i <= pageCount; i++)
                {
                    sourceDocument = new Document(reader.GetPageSizeWithRotation(i));
                    string outputPdfPath = System.IO.Path.Combine(outputPdfDir, i.ToString() + ".pdf");

                    filePaths.Add(outputPdfPath);

                    pdfCopyProvider = new PdfCopy(sourceDocument, new System.IO.FileStream(outputPdfPath, System.IO.FileMode.Create));
                    sourceDocument.Open();
                    importedPage = pdfCopyProvider.GetImportedPage(reader, i);
                    pdfCopyProvider.AddPage(importedPage);
                    sourceDocument.Close();
                }
                reader.Close();
                return filePaths.ToArray();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}

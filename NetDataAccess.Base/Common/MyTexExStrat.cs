using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using iTextSharp.xmp.impl;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class MyTexExStrat : SimpleTextExtractionStrategy 
    {
        public MyTexExStrat()
        {

        }
        private StringBuilder TextBuilder = new StringBuilder();
        private Vector LastEndPoint = null;

        public string GetFullText()
        {
            return this.TextBuilder.ToString();
        }
        public override void BeginTextBlock()
        {
            base.BeginTextBlock();
        }

        public override void EndTextBlock()
        {
            base.EndTextBlock();
        }
         
        public override string GetResultantText()
        {
            return base.GetResultantText();
        }
         

        public override void RenderText(TextRenderInfo renderInfo)
        { 
            DocumentFont font = renderInfo.GetFont();
            byte[] bytes = font.ConvertToBytes(renderInfo.GetText());
            StringBuilder partStr = new StringBuilder();
            //PdfDictionary dict = font.FontDictionary;
            //PdfDictionary encoding = dict.GetAsDict(PdfName.ENCODING);
            //if (encoding != null)
            {
                string renderTxt = renderInfo.GetText(); 
                foreach (char c in renderTxt)
                {
                    string name = c.ToString();
                    if (CharMaps.ContainsKey(name))
                    {
                        string chat = CharMaps[name];
                        partStr.Append(chat);
                    }
                    else
                    {
                        /*
                        String s = name.ToString().Substring(4);
                        byte[] nameBytes = this.HexStringToBytes(s);
                        string text = Encoding.BigEndianUnicode.GetString(nameBytes);
                        partStr.Append(text);
                         */
                        partStr.Append(name);
                    }
                } 
                /*
                PdfArray diffs = encoding.GetAsArray(PdfName.DIFFERENCES); 
                StringBuilder builder = new StringBuilder();
                foreach (byte b in renderInfo.PdfString.GetBytes())
                { 
                    string name = "";
                    try
                    {
                        name = diffs.GetAsName((char)b).ToString();
                    }
                    catch (Exception ex)
                    {
                        //throw ex;
                    }
                    if (CharMaps.ContainsKey(name))
                    {
                        string chat = CharMaps[name];
                        //partStr.Append(chat);
                    }
                    else
                    {
                        //String s = name.ToString().Substring(4);
                        //byte[] nameBytes = this.HexStringToBytes(s);
                        //string text = Encoding.BigEndianUnicode.GetString(nameBytes);
                        //partStr.Append(text);
                    } 
                }
                 */
            }
            //else
            //{
            //    partStr.Append(renderInfo.GetText());
            //} 

            if (partStr.ToString().Trim().Length > 0)
            {
                LineSegment segment = renderInfo.GetAscentLine();
                Vector startPoint = segment.GetStartPoint();
                if (LastEndPoint != null)
                {
                    int charCount = partStr.Length;
                    float singleWidth = renderInfo.GetSingleSpaceWidth();// segment.GetEndPoint().Subtract(segment.GetStartPoint()).Length / (float)charCount;
                    string[] pStartStrs = startPoint.ToString().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    string[] pLastEndStrs = LastEndPoint.ToString().Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    float verticalSpan = Math.Abs(float.Parse(pStartStrs[1]) - float.Parse(pLastEndStrs[1]));
                    if (verticalSpan > renderInfo.GetSingleSpaceWidth())
                    {
                        if (verticalSpan > 4 * renderInfo.GetSingleSpaceWidth())
                        {
                            TextBuilder.Append("\r\n\r\n");
                        }
                        else
                        {
                            TextBuilder.Append("\r\n");
                        }
                    }
                    else
                    {
                        Vector spanVector = startPoint.Subtract(LastEndPoint);
                        float span = spanVector.Length;
                        if (span > singleWidth / 2f)
                        {
                            TextBuilder.Append(" ");
                        }
                    }
                }
                TextBuilder.Append(partStr.ToString());
                LastEndPoint = segment.GetEndPoint();
            }

            base.RenderText(renderInfo);
        }
        private static Dictionary<string, string> _CharMaps = null;
        private static Dictionary<string, string> CharMaps
        {
            get
            {
                if (_CharMaps == null)
                {
                    Dictionary<string, string> charMaps = new Dictionary<string, string>();
                    charMaps.Add("ﬀ", "ff");
                    charMaps.Add("ﬁ", "fi");
                    charMaps.Add("ﬂ", "fl");
                    charMaps.Add("ﬃ", "ffi");
                    charMaps.Add("ﬄ", "ffl");
                    charMaps.Add("ﬅ", "st");
                    charMaps.Add("ﬆ", "st");
                    charMaps.Add("Þ", "fi"); 
                    /*
                    charMaps.Add("/uniFB00", "ff");
                    charMaps.Add("/uniFB01", "fi");
                    charMaps.Add("/uniFB02", "fl");
                    charMaps.Add("/uniFB03", "ffi");
                    charMaps.Add("/uniFB04", "ffl");
                    charMaps.Add("/uniFB05", "st");
                    charMaps.Add("/uniFB06", "st");
                    charMaps.Add("/uni0020", " ");
                    charMaps.Add("/uni2002", " ");
                    charMaps.Add("/uni2003", " ");
                     */
                    _CharMaps = charMaps;
                }
                return _CharMaps;
            }
        }
        private byte[] HexStringToBytes(string hs)
        {
            string strTemp = "";
            byte[] b = new byte[hs.Length / 2];
            for (int i = 0; i < hs.Length / 2; i++)
            {
                strTemp = hs.Substring(i * 2, 2);
                b[i] = Convert.ToByte(strTemp, 16);
            }
            return b;
        }
    }
}

using Microsoft.International.Converters.TraditionalChineseToSimplifiedConverter;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NetDataAccess.Base.Common
{
    public class CharProcessor
    {
        public static string ConverterChinese(string text, ChineseConversionDirection direction)
        {
            return ChineseConverter.Convert(text, direction);
        }
    }
}

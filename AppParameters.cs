using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Hello
{
    public static class AppParameters
    {
        public static bool LoggingMode { get; set; } = true;

        static string _pathFileLog = @"C:\Windows\Temp\docnetLog.txt";
        public static string PathFileLog
        {
            get { return _pathFileLog; }
            set
            {
                if (Directory.GetParent(value).Exists) _pathFileLog = value;
                else TextError.Add("Отсутствует папка для логгироавния файлов по пути: " + value);
            }
        }

        public static List<string> TextError { get; set; } = new List<string>();
        public static List<string> _listInfoQualityImages = new List<string>();

        public static string GetErrors()
        {
            if (TextError.Count > 0)
                return string.Join(System.Environment.NewLine, TextError);
            else
                return "";
        }

        public static bool _isSuitableForOcr = true;

        public static string _warnColorTypeBad = "";
        public static string _warnResolutionImageBad = "";
    }
}

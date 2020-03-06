using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookExportTool
{
    public static class StringExtensions
    {
        public const int Indent = 4;
        public const int ConsoleWidth = 119;

        public static bool IsSentItems(this string name) => name.Contains("Sent Items");

        public static string FilterFormat(this DateTime date) => date.ToString("g");

        public static bool Contains(this string value, string[] items)
        {
            foreach (var item in items)
            {
                if (value.Contains(item))
                {
                    return true;
                }
            }

            return false;
        }

        public static void WriteIndented(this string message, int level) => Console.Write($"{GetIndentedMessage(message, level).PadRight(ConsoleWidth)}\r");

        public static void WriteLineIndented(this string message, int level) => Console.WriteLine(GetIndentedMessage(message, level).PadRight(ConsoleWidth));

        private static string GetIndentedMessage(string message, int level) => Truncate($"{GetIndent(level)}{message}");

        private static string GetIndent(int level) => string.Empty.PadLeft(level * Indent);

        private static string Truncate(string value, int length = ConsoleWidth) => value.Length < length ? value : value.Substring(0, length);
    }
}

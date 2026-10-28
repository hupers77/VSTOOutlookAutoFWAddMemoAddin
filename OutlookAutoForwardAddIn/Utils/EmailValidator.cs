using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;

namespace OutlookAutoForwardAddIn.Utils
{
    public static class EmailValidator
    {
        private static readonly char[] Separators = { ';', ',', '\r', '\n' };

        public static List<string> Split(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return new List<string>();

            return value.Split(Separators, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .Where(s => s.Length > 0)
                .ToList();
        }

        /// <summary>모든 주소가 유효하면 true. 비어 있으면 true(필수 여부는 호출 측에서 판단).</summary>
        public static bool TryGetInvalid(string value, out string invalidAddress)
        {
            invalidAddress = null;
            foreach (string address in Split(value))
            {
                if (!IsValid(address))
                {
                    invalidAddress = address;
                    return false;
                }
            }
            return true;
        }

        public static bool IsValid(string address)
        {
            try
            {
                var parsed = new MailAddress(address);
                return parsed.Address == address && parsed.Host.Contains(".");
            }
            catch (FormatException)
            {
                return false;
            }
        }

        /// <summary>Outlook 형식(세미콜론 구분)으로 정규화한다.</summary>
        public static string Normalize(string value)
        {
            return string.Join("; ", Split(value));
        }
    }
}

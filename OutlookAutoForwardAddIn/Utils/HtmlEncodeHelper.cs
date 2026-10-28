using System.Net;

namespace OutlookAutoForwardAddIn.Utils
{
    public static class HtmlEncodeHelper
    {
        public static string Encode(string value)
        {
            if (string.IsNullOrEmpty(value))
                return string.Empty;

            return WebUtility.HtmlEncode(value);
        }

        /// <summary>HTML 인코딩 후 줄바꿈을 br 태그로 변환한다.</summary>
        public static string EncodeWithLineBreaks(string value)
        {
            return Encode(value)
                .Replace("\r\n", "<br/>")
                .Replace("\n", "<br/>")
                .Replace("\r", "<br/>");
        }
    }
}

using System;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using OutlookAutoForwardAddIn.Models;
using OutlookAutoForwardAddIn.Utils;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAutoForwardAddIn.Services
{
    public static class MailForwardService
    {
        private static readonly Regex BodyOpenTag =
            new Regex(@"<body[^>]*>", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        public static void ForwardByRule(Outlook.MailItem originalMail, AutoForwardRule rule)
        {
            Outlook.MailItem forwardMail = null;
            try
            {
                forwardMail = originalMail.Forward();

                forwardMail.To = rule.ToRecipients;
                forwardMail.CC = rule.CcRecipients ?? string.Empty;
                forwardMail.BCC = rule.BccRecipients ?? string.Empty;

                InsertHeaderMessage(originalMail, forwardMail, rule.HeaderMessage);

                if (!rule.IncludeAttachments)
                    RemoveAttachments(forwardMail);

                // 이 Add-in이 보낸 메일임을 표시하는 헤더 (수신 측에서 재전달 루프 방지에 사용)
                forwardMail.PropertyAccessor.SetProperty(MailMonitorService.LoopHeaderSchema, "1");

                forwardMail.Send();
            }
            finally
            {
                ReleaseCom(forwardMail);
            }
        }

        private static void InsertHeaderMessage(Outlook.MailItem original, Outlook.MailItem forward, string message)
        {
            if (original.BodyFormat == Outlook.OlBodyFormat.olFormatHTML)
            {
                try
                {
                    string block =
                        "<div style='font-family:Malgun Gothic; font-size:10pt;'>" +
                        HtmlEncodeHelper.EncodeWithLineBreaks(message) +
                        "</div><hr/>";

                    string html = forward.HTMLBody ?? string.Empty;
                    Match body = BodyOpenTag.Match(html);
                    forward.HTMLBody = body.Success
                        ? html.Insert(body.Index + body.Length, block)
                        : block + html;
                    return;
                }
                catch (Exception ex)
                {
                    // HTML 처리 실패 시 일반 텍스트 방식으로 대체
                    System.Diagnostics.Debug.WriteLine("HTML 본문 처리 실패: " + ex.Message);
                }
            }

            forward.Body =
                message +
                Environment.NewLine +
                Environment.NewLine +
                "----- 원본 메일 -----" +
                Environment.NewLine +
                forward.Body;
        }

        private static void RemoveAttachments(Outlook.MailItem mail)
        {
            Outlook.Attachments attachments = mail.Attachments;
            try
            {
                for (int i = attachments.Count; i >= 1; i--)
                    attachments.Remove(i);
            }
            finally
            {
                ReleaseCom(attachments);
            }
        }

        public static void ReleaseCom(object comObject)
        {
            if (comObject != null && Marshal.IsComObject(comObject))
                Marshal.ReleaseComObject(comObject);
        }
    }
}

using System;
using System.Collections.Generic;
using OutlookAutoForwardAddIn.Models;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAutoForwardAddIn.Services
{
    /// <summary>Application.NewMailEx 이벤트로 신규 메일을 감지하고 규칙에 따라 전달한다.</summary>
    public class MailMonitorService
    {
        // 이 Add-in이 보낸 전달 메일에 심는 인터넷 헤더. 수신 시 이 헤더가 있으면 다시 전달하지 않는다.
        public const string LoopHeaderSchema =
            "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/X-OutlookAutoForward";
        private const string InternetMessageIdSchema = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
        private const string ProcessedPropertyName = "AutoForwardProcessed";

        private readonly Outlook.Application application;
        private readonly HashSet<string> processedKeys = new HashSet<string>();

        public MailMonitorService(Outlook.Application application)
        {
            this.application = application;
        }

        public void Start()
        {
            application.NewMailEx += Application_NewMailEx;
        }

        public void Stop()
        {
            application.NewMailEx -= Application_NewMailEx;
        }

        private void Application_NewMailEx(string entryIDCollection)
        {
            if (string.IsNullOrEmpty(entryIDCollection))
                return;

            // 여러 메일이 한 번에 수신되면 EntryID가 쉼표로 구분되어 전달된다.
            foreach (string entryId in entryIDCollection.Split(','))
            {
                string id = entryId.Trim();
                if (id.Length > 0)
                    ProcessMail(id);
            }
        }

        private void ProcessMail(string entryId)
        {
            Outlook.NameSpace session = null;
            object item = null;
            try
            {
                session = application.Session;
                item = session.GetItemFromID(entryId);

                var mail = item as Outlook.MailItem;
                if (mail == null || string.IsNullOrWhiteSpace(mail.Subject))
                    return;

                // 1) 이 Add-in이 보낸 전달 메일이면 처리하지 않는다 (전달 루프 방지)
                if (GetStringProperty(mail, LoopHeaderSchema) != null)
                    return;

                // 2) 이미 처리한 메일이면 처리하지 않는다 (메일당 1회 발송)
                //    EntryID는 메일 이동 시 바뀔 수 있으므로 InternetMessageId를 우선 키로 사용한다.
                string key = GetStringProperty(mail, InternetMessageIdSchema) ?? entryId;
                lock (processedKeys)
                {
                    if (!processedKeys.Add(key))
                        return;
                }
                if (mail.UserProperties.Find(ProcessedPropertyName) != null)
                    return;

                AutoForwardRule matchedRule = RuleMatchService.FindMatch(RuleRepository.LoadRules(), mail.Subject);
                if (matchedRule == null)
                    return;

                // 발송 전에 먼저 처리 표시를 남겨, 발송 중 재진입/재시작이 있어도 중복되지 않게 한다.
                MarkProcessed(mail);
                MailForwardService.ForwardByRule(mail, matchedRule);
            }
            catch (Exception ex)
            {
                // 로그 기능은 범위 제외. Outlook 동작에 영향이 없도록 예외를 삼킨다.
                System.Diagnostics.Debug.WriteLine("자동 전달 실패: " + ex);
            }
            finally
            {
                MailForwardService.ReleaseCom(item);
                MailForwardService.ReleaseCom(session);
            }
        }

        private static void MarkProcessed(Outlook.MailItem mail)
        {
            try
            {
                Outlook.UserProperty prop = mail.UserProperties.Add(
                    ProcessedPropertyName, Outlook.OlUserPropertyType.olYesNo, false);
                prop.Value = true;
                mail.Save();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("처리 표시 실패: " + ex.Message);
            }
        }

        private static string GetStringProperty(Outlook.MailItem mail, string schema)
        {
            try
            {
                object value = mail.PropertyAccessor.GetProperty(schema);
                string text = value as string;
                return string.IsNullOrEmpty(text) ? null : text;
            }
            catch
            {
                return null;
            }
        }
    }
}

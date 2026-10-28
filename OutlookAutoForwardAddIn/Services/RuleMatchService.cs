using System;
using System.Collections.Generic;
using System.Linq;
using OutlookAutoForwardAddIn.Models;
using OutlookAutoForwardAddIn.Utils;

namespace OutlookAutoForwardAddIn.Services
{
    public static class RuleMatchService
    {
        /// <summary>
        /// 사용 중인 규칙 중 제목 조건이 일치하고 제외 단어가 없는 첫 번째 규칙을 반환한다.
        /// </summary>
        public static AutoForwardRule FindMatch(IEnumerable<AutoForwardRule> rules, string subject)
        {
            if (string.IsNullOrWhiteSpace(subject) || rules == null)
                return null;

            return rules
                .Where(r => r.Enabled && !string.IsNullOrWhiteSpace(r.SubjectContains))
                .Where(r => Contains(subject, r.SubjectContains))
                .FirstOrDefault(r => !EmailValidator.Split(r.ExcludeSubjects).Any(word => Contains(subject, word)));
        }

        private static bool Contains(string text, string word)
        {
            return text.IndexOf(word, StringComparison.OrdinalIgnoreCase) >= 0;
        }
    }
}

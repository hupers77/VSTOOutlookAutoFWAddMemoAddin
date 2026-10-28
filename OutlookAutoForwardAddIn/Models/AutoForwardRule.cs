using System;
using System.Runtime.Serialization;

namespace OutlookAutoForwardAddIn.Models
{
    [DataContract]
    public class AutoForwardRule
    {
        [DataMember] public string RuleId { get; set; }
        [DataMember] public string RuleName { get; set; }
        [DataMember] public bool Enabled { get; set; }

        [DataMember] public string SubjectContains { get; set; }
        [DataMember] public string ExcludeSubjects { get; set; }
        [DataMember] public string HeaderMessage { get; set; }

        [DataMember] public string ToRecipients { get; set; }
        [DataMember] public string CcRecipients { get; set; }
        [DataMember] public string BccRecipients { get; set; }

        [DataMember] public bool IncludeAttachments { get; set; }

        [DataMember] public DateTime CreatedAt { get; set; }
        [DataMember] public DateTime UpdatedAt { get; set; }

        public AutoForwardRule Clone()
        {
            return (AutoForwardRule)MemberwiseClone();
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.Serialization.Json;
using System.Text;
using OutlookAutoForwardAddIn.Models;

namespace OutlookAutoForwardAddIn.Services
{
    public static class RuleRepository
    {
        private static readonly object SyncRoot = new object();

        private static readonly DataContractJsonSerializerSettings Settings = new DataContractJsonSerializerSettings
        {
            DateTimeFormat = new System.Runtime.Serialization.DateTimeFormat("yyyy-MM-ddTHH:mm:ss")
        };

        public static string RulesFilePath
        {
            get
            {
                string dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "OutlookAutoForwardAddIn");
                return Path.Combine(dir, "rules.json");
            }
        }

        /// <summary>규칙 로드. 파일이 없거나 손상된 경우 빈 목록을 반환한다.</summary>
        public static List<AutoForwardRule> LoadRules()
        {
            lock (SyncRoot)
            {
                try
                {
                    string path = RulesFilePath;
                    if (!File.Exists(path))
                    {
                        EnsureFile(path);
                        return new List<AutoForwardRule>();
                    }

                    byte[] bytes = File.ReadAllBytes(path);
                    if (bytes.Length == 0)
                        return new List<AutoForwardRule>();

                    var serializer = new DataContractJsonSerializer(typeof(List<AutoForwardRule>), Settings);
                    using (var stream = new MemoryStream(bytes))
                    {
                        var rules = serializer.ReadObject(stream) as List<AutoForwardRule>;
                        return rules ?? new List<AutoForwardRule>();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("규칙 로드 실패: " + ex.Message);
                    return new List<AutoForwardRule>();
                }
            }
        }

        /// <summary>규칙 저장. 임시 파일에 쓴 뒤 교체하여 파일 손상을 방지한다.</summary>
        public static void SaveRules(IEnumerable<AutoForwardRule> rules)
        {
            lock (SyncRoot)
            {
                string path = RulesFilePath;
                Directory.CreateDirectory(Path.GetDirectoryName(path));

                var serializer = new DataContractJsonSerializer(typeof(List<AutoForwardRule>), Settings);
                string tempPath = path + ".tmp";
                using (var stream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                using (var writer = JsonReaderWriterFactory.CreateJsonWriter(stream, Encoding.UTF8, false, true, "  "))
                {
                    serializer.WriteObject(writer, new List<AutoForwardRule>(rules));
                }

                if (File.Exists(path))
                    File.Delete(path);
                File.Move(tempPath, path);
            }
        }

        private static void EnsureFile(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, "[]", new UTF8Encoding(false));
        }

        public static string NewRuleId()
        {
            return "RULE-" + Guid.NewGuid().ToString("N").Substring(0, 8).ToUpperInvariant();
        }
    }
}

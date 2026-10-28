using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using OutlookAutoForwardAddIn.Forms;

namespace OutlookAutoForwardAddIn
{
    /// <summary>Outlook 홈 탭에 '자동 전달 규칙' 버튼을 추가한다.</summary>
    [ComVisible(true)]
    public class AutoForwardRibbon : IRibbonExtensibility
    {
        public string GetCustomUI(string ribbonID)
        {
            // Explorer(메인 창) 리본에만 적용
            if (ribbonID != "Microsoft.Outlook.Explorer")
                return null;

            return GetResourceText("OutlookAutoForwardAddIn.AutoForwardRibbon.xml");
        }

        public void OnOpenRules(IRibbonControl control)
        {
            using (var form = new RuleListForm())
            {
                form.ShowDialog();
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                    return null;
                using (var reader = new StreamReader(stream))
                    return reader.ReadToEnd();
            }
        }
    }
}

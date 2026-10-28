using Microsoft.Office.Core;
using OutlookAutoForwardAddIn.Services;

namespace OutlookAutoForwardAddIn
{
    public partial class ThisAddIn
    {
        private MailMonitorService mailMonitor;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            mailMonitor = new MailMonitorService(Application);
            mailMonitor.Start();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (mailMonitor != null)
                mailMonitor.Stop();
        }

        protected override IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new AutoForwardRibbon();
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            Startup += ThisAddIn_Startup;
            Shutdown += ThisAddIn_Shutdown;
        }

        #endregion
    }
}

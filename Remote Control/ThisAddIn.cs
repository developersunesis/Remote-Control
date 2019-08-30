using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using System.Diagnostics;

namespace Remote_Control
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Microsoft.Office.Interop.PowerPoint.Application app = this.Application;
            app.PresentationBeforeClose += beforeClose;

        }

        private void beforeClose(PowerPoint.Presentation Pres, ref bool Cancel)
        {
            if(Application.Presentations.Count <=1)
            {
                Marshal.ReleaseComObject(this.Application);
                Process[] processes = Process.GetProcessesByName("powerpnt");
                foreach(Process p in processes)
                {
                    p.Kill();
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}

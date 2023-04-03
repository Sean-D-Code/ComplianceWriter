using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Configuration;

namespace ComplianceWriter
{
    public partial class ThisAddIn
    {
        Outlook.Inspectors Inspectors { get; set; }
        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            Inspectors = this.Application.Inspectors;
            Inspectors.NewInspector += new InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
            this.Application.ItemSend += new ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);

            Settings.InitializeSettings();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        private void Inspectors_NewInspector(Inspector inspector)
        {
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if(mailItem != null)
            {
                if(mailItem.EntryID == null)
                {
                }
            }
        }
        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            string line = "________________________________________________________________________________";
            try
            {                
                if (Item != null && Item is Outlook.MailItem)
                {
                    MailItem mailItem = (Outlook.MailItem)Item;

                    mailItem.HTMLBody += Settings.DisclaimerTextHTML;

                }
                else if(Item != null && Item is Outlook.MeetingItem)
                {
                    MeetingItem meetingItem = (MeetingItem)Item;
                    meetingItem.Body += $"{line}\r\n\r\n" + Settings.DisclaimerText;
                }
            } 
            catch (System.Exception ex)
            {
                //Report Error
            }
            finally
            {
                
            }
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

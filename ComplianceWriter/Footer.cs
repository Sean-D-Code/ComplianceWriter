using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace ComplianceWriter
{
    partial class Footer
    {
        #region Form Region Factory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Note)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("ComplianceWriter.Footer")]
        public partial class FooterFactory
        {
            // Occurs before the form region is initialized.
            // To prevent the form region from appearing, set e.Cancel to true.
            // Use e.OutlookItem to get a reference to the current Outlook item.
            private void FooterFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
                
                if(Settings.DisableDisclaimerForm)
                {
                    //Set Overall Disable/Enable
                    e.Cancel = Settings.DisableDisclaimerForm;
                }
                else if (e.OutlookItem != null && e.OutlookItem is Outlook.MailItem)
                {
                    e.Cancel = Settings.DisableFormOnMailItem;
                }
                else if(e.OutlookItem != null && e.OutlookItem is Outlook.AppointmentItem)
                {
                    AppointmentItem appointmentItem = (AppointmentItem)e.OutlookItem;
                    if(appointmentItem != null && appointmentItem.MeetingStatus == 0)
                    {
                        e.Cancel = Settings.DisableFormOnAppointmentItem;
                    }
                    else if(appointmentItem != null && appointmentItem.MeetingStatus == 0)
                    {
                        e.Cancel = Settings.DisableFormOnMeetingItem;
                    }                    
                }
            }
        }

        #endregion

        // Occurs before the form region is displayed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void Footer_FormRegionShowing(object sender, System.EventArgs e)
        {
            //Resize Text
            if (this.OutlookFormRegion != null)
            {
                this.disclaimer.Width = this.OutlookFormRegion.Parent.Width - 20;
            }
        }

        // Occurs when the form region is closed.
        // Use this.OutlookItem to get a reference to the current Outlook item.
        // Use this.OutlookFormRegion to get a reference to the form region.
        private void Footer_FormRegionClosed(object sender, System.EventArgs e)
        {
        }

        private void Footer_Resize(object sender, System.EventArgs e)
        {
            //Resize Text
            if (this.OutlookFormRegion != null)
            {
                this.disclaimer.Width = this.OutlookFormRegion.Parent.Width - 20;
            }
        }
       
    }
}

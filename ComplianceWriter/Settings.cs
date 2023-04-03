using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ComplianceWriter
{
    internal class Settings
    {
        public static string DisclaimerText { get; set; }        
        public static string DisclaimerTextHTML { get; set; }
        public static string DisclaimerFormName { get; set; }
        public static bool DisableDisclaimerForm { get; set; }
        public static bool DisableFormOnMailItem { get; set; }
        public static bool DisableFormOnMeetingItem { get; set; }
        public static bool DisableFormOnAppointmentItem { get; set; }


        public static void InitializeSettings()
        {
            List<string> configKeys = ConfigurationManager.AppSettings.AllKeys.ToList();
            List<PropertyInfo> properties = typeof(Settings).GetProperties().ToList();


            foreach(PropertyInfo prop in properties)
            {
                try
                {
                    string configKey = configKeys.FirstOrDefault(x => x == prop.Name);

                    if (configKey != null)
                    {
                        if (prop.PropertyType == typeof(string))
                        {
                            prop.SetValue(prop, ConfigurationManager.AppSettings[configKey]);
                        }
                        else if (prop.PropertyType == typeof(bool))
                        {
                            prop.SetValue(prop, bool.Parse(ConfigurationManager.AppSettings[configKey]));
                        }
                        else if (prop.PropertyType == typeof(int))
                        {
                            prop.SetValue(prop, int.Parse(ConfigurationManager.AppSettings[configKey]));
                        }
                        else
                        {
                            //TODO: Add Logging
                        }
                    }
                }
                catch (Exception ex)
                {
                    //TODO: Add Logging
                }
            }

            DisclaimerHtmlMarkUp();
        }

        private static void DisclaimerHtmlMarkUp()
        {
            string startHtml = @"<p class=MsoNormal><span style='font-size:12.0pt'>	<o:p>&nbsp;</o:p></span></p><hr><p class=xmsonormal><span style='font-size:12.0pt'><i><br>";
            string endHtml = @"</span><o:p/></i></p>";
            DisclaimerTextHTML = startHtml + DisclaimerText + endHtml;
        }
        
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Configuration;
using Microsoft.InformationProtection;
using Microsoft.InformationProtection.File;
using Microsoft.InformationProtection.Protection;
using Microsoft.InformationProtection.Policy;

namespace RmsFileWatcher
{
    static class Program
    {
        private static readonly string clientId = ConfigurationManager.AppSettings["ida:ClientId"];
        private static readonly string appName = ConfigurationManager.AppSettings["app:Name"];
        private static readonly string appVersion = ConfigurationManager.AppSettings["app:Version"];


        [STAThread]
        static void Main()
        {
            // Create ApplicationInfo, setting the clientID from Azure AD App Registration as the ApplicationId
            // If any of these values are not set API throws BadInputException.
            ApplicationInfo appInfo = new ApplicationInfo()
            {
                // ApplicationId should ideally be set to the same ClientId found in the Azure AD App Registration.
                // This ensures that the clientID in AAD matches the AppId reported in AIP Analytics.
                ApplicationId = clientId,
                ApplicationName = appName,
                ApplicationVersion = appVersion
            };
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormRmsFileWatcher(appInfo));
        }
    }
}

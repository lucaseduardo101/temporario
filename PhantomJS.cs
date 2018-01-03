using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OpenQA.Selenium.PhantomJS;
using OpenQA.Selenium.Remote;
using System.IO;
using Automate.Utils;


namespace Automate.Services.Applications
{
    class PhantomJS : WebBrowser
    {
        protected override string BrowserWindow
        {
            get { return "PhantomJS"; }
        }
        protected override string PrintKeys
        {
            get { return "^+p"; }
        }
        protected override string[] DownloadExtensions
        {
            get { return new string[] { ".tmp" }; }
        }

        public override void Open(string url)
        {
            PhantomJSOptions options = new PhantomJSOptions();
           // DesiredCapabilities capabilities = DesiredCapabilities.PhantomJS();

            tempDownloadPath = Path.Combine(Manager.parameters.Get("$Caminho Download Final$").ToString(), "$Hera$");
            var downloadPrefs = new Dictionary<string, object>
            {
                {"default_directory", tempDownloadPath},
                {"directory_upgrade", true}
            };

            options.AddAdditionalCapability("phantomjs.page.settings.userAgent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/60.0.3112.113 Safari/537.36");

            PhantomJSDriverService service = PhantomJSDriverService.CreateDefaultService(AppDomain.CurrentDomain.BaseDirectory);
            service.SuppressInitialDiagnosticInformation = true;

            browser = new PhantomJSDriver(service, options);

            browserHandle = UIAutomation.FindWindow(BrowserWindow, UIAutomation.GetActiveWindowTitle());
            if (!(url.StartsWith("https://") || url.StartsWith("http://")))
            {
                url = "http://" + url;
            }

            browser.Navigate().GoToUrl(url);

        }
    }
}

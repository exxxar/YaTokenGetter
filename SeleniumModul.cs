using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;


namespace RegAccYandex
{
    class SeleniumModul
    {
        ChromeDriver driver;
        string url = "https://oauth.yandex.ru/authorize?response_type=code&client_id={0}";

        public SeleniumModul()
        {

        }

        public string run(string clientId, string clientSecret, string login = null, string pass = null)
        {
            this.driver = this.initDriver(clientId);
            if (login != null && pass != null)
            {
                Thread.Sleep(2000);
                IWebElement elem = driver.FindElement(By.CssSelector("[name='login']"));
                elem.SendKeys(login);
                Thread.Sleep(2000);
                elem = driver.FindElement(By.CssSelector("[name='passwd']"));
                elem.SendKeys(pass);
                Thread.Sleep(2000);
                elem = driver.FindElement(By.CssSelector(".passport-Button[type='submit']"));
                elem.SendKeys(Keys.Enter);
                Thread.Sleep(2000);
                try
                {
                    if (isSelectorExist(By.CssSelector(".passport-Domik-Form-Error.passport-Domik-Form-Error_active")))
                    {
                        driver.Close();
                        driver.Dispose();
                        return null;
                    }
                }
                catch { }

            }

            var task = Task.Run(() => (new CodeWaiter()).execute(this.driver, new Dictionary<string, string>
                {
                    { "client_id",clientId},
                    { "client_secret",clientSecret}
                }));


            var r = task.GetAwaiter();
            Task.WaitAll(new Task[] { task });

            driver.Close();
            driver.Dispose();

            return r.GetResult();



        }

        public bool isSelectorExist(By selector)
        {
            return driver.FindElements(selector).Count != 0;
        }

        private ChromeDriver initDriver(string clientId)
        {
            var options = new ChromeOptions();
            //options.AddArgument("no-sandbox");
            // options.AddUserProfilePreference("download.default_directory", Directory.GetCurrentDirectory());
            options.AddArguments("--disable-extensions");
            options.AddArgument("no-sandbox");
            options.AddArgument("--incognito");
            //options.AddArgument("--headless");
            options.AddArgument("--disable-gpu");  //--d


            ChromeDriver driver = new ChromeDriver(options);//открываем сам браузер
            driver.LocationContext.PhysicalLocation = new OpenQA.Selenium.Html5.Location(55.751244, 37.618423, 152);
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10); //время ожидания компонента страницы после загрузки страницы
            driver.Manage().Cookies.DeleteAllCookies();


            driver.Navigate().GoToUrl(String.Format(this.url, clientId));

            return driver;
        }
    }
}

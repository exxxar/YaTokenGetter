using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RegAccYandex
{
     class SeleniumUtils
    {
        protected ChromeDriver driver;

        protected bool isSelectorExist(By selector)
        {
            Boolean result = false;

            try
            {
                result = driver.FindElements(selector).Count != 0;
            }
            catch
            {

            }
            return result;
        }

        protected void UrlGo(string url)
        {
            this.driver.Navigate().GoToUrl(url);
        }

        protected IWebElement getElement(string selector)
        {
            return this.driver.FindElement(By.CssSelector(selector));
        }

        protected IWebElement[] getElements(string selector)
        {
            return this.driver.FindElements(By.CssSelector(selector)).ToArray();
        }


        protected string doJS(string path, string param = "")
        {
            StringBuilder s = new StringBuilder();
            File.ReadLines("js\\" + path)
                .ToList()
                .ForEach(l =>
                {
                    //if (param.Trim().Length > 0)
                    //    l = (new Regex("{{item}}")).Replace(l, param);
                    //Console.WriteLine(l);
                    s.Append(l);
                });

            return s.ToString();
        }

        protected ChromeDriver initDriver()
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


           

            return driver;
        }
    }
}

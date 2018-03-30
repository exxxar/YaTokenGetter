using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System.Threading;
using System.IO;
using System.Text.RegularExpressions;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;

namespace RegAccYandex
{
    class SeleniumModul : SeleniumUtils
    {
        string url = "https://oauth.yandex.ru/authorize?response_type=code&client_id={0}";


        public string takeToken(string clientId, string clientSecret, string login = null, string pass = null)
        {
            this.driver = this.initDriver();
            this.UrlGo(String.Format(this.url, clientId));

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


        public void findAndDo(string selector, params string[] args )
        {
            //new WebDriverWait(driver, TimeSpan.FromSeconds(15)).Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".captcha__container")));
            try
            {
                IWebElement elem = driver.FindElement(By.CssSelector(selector));
                Actions actions = new Actions(driver);
                actions.MoveToElement(elem).Click().Perform();
                if (args != null)
                {
                    foreach (var k in args)
                    {
                        Console.WriteLine(k);
                        elem.SendKeys(k);
                        Thread.Sleep(2000);
                    }
                }
                    
            }catch
            {

            }
        }

        public void doCaptchaAwaiter()
        {
            while (true)
            {
                if (!isSelectorExist(By.CssSelector(".captcha__container"))
                    && !isSelectorExist(By.CssSelector(".b-captcha-form")))
                    break;
            }
        }
        public User regYandexAcc()
        {


            User user = new User();
            user.name = Faker.Name.First();
            user.tname = Faker.Name.Last();
            user.login = Faker.Name.FullName().ToLower().Replace(' ', 'a')
                + Faker.RandomNumber.Next(10000, 1000000);
            user.pass = "testPass1234fxdxttrr";
            user.question = "Ваш любимый музыкант";
            user.answer = "Андрей!";
            user.date_reg = String.Format("{0}", DateTime.Now);


            this.driver = this.initDriver();
            this.UrlGo("https://passport.yandex.ru/registration");

            findAndDo("#firstname", user.name);
            findAndDo("#lastname", user.tname);
            findAndDo("#login", user.login);
            findAndDo("#password", user.pass);
            findAndDo("#password_confirm", user.pass);
            findAndDo(".link_has-no-phone");
            findAndDo("#hint_answer", user.answer);
            doCaptchaAwaiter();
            this.UrlGo("https://direct.yandex.ru/");
            findAndDo(".b-morda-hero__button", Keys.Enter);
            doCaptchaAwaiter();
            findAndDo("[name='email']", $"{user.login}@yandex.ru");
            findAndDo(".p-collect-emails__button");
            doCaptchaAwaiter();
            findAndDo(".b-choose-country-currency__cell .select button",  Keys.Down, Keys.Enter);
            findAndDo(".p-choose-interface__submit-button", Keys.Enter);
            doCaptchaAwaiter();
            findAndDo(".b-regions-selector__switcher",  Keys.Enter);
            findAndDo(".p-choose-interface__submit-button", Keys.Enter);
            doCaptchaAwaiter();
            findAndDo(".b-regions__quick-select_id_1", Keys.Enter);
            findAndDo(".b-edit-regions-popup__save", Keys.Enter);
            findAndDo(".b-campaign-settings__submit", Keys.Enter);
            doCaptchaAwaiter();
            findAndDo(".b-edit-group-header__name-input input", "Test company");
            findAndDo(".b-edit-banner2__title input", "Test company");
            findAndDo(".b-edit-banner2__body textarea", "Test data");
            findAndDo(".b-edit-banner2__href input", "mail.ru");
            findAndDo(".p-multiedit2__submit-button", Keys.Enter);
            doCaptchaAwaiter();
            this.UrlGo("https://direct.yandex.ru/registered/main.pl?cmd=apiSettings");
            doCaptchaAwaiter();
            findAndDo(".link_theme_direct", Keys.Enter);
            doCaptchaAwaiter();
            findAndDo(".b-oferta-accept__oferta-content .radiobox__radio:nth-of-type(1) input", Keys.Enter);
            findAndDo(".b-oferta-accept__submit-button", Keys.Enter);
            doCaptchaAwaiter();

            Thread.Sleep(3000);
            driver.Dispose();
            return user;
        }


    }
}

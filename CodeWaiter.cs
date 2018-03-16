using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;

namespace RegAccYandex
{
    class CodeWaiter
    {
        ChromeDriver driver;
        public async Task<string> execute(ChromeDriver driver, Dictionary<string, string> param)
        {
            this.driver = driver;
            while (true)
            {
                Thread.Sleep(2000);
                Console.WriteLine("ping {0}", DateTime.Now);

                if (isSelectorExist(By.CssSelector(".error-message")))
                {
                    var message = driver.FindElement(By.CssSelector(".error-message")).Text;
                    Console.WriteLine(message);
                    driver.Close();
                    driver.Dispose();
                    throw new Exception(message);
                }

                try
                {
                    if (isSelectorExist(By.CssSelector(".submit-once_allow")))
                    {
                        driver.FindElement(By.CssSelector("button:nth-of-type(1)")).SendKeys(Keys.Enter);
                    }
                    if (isSelectorExist(By.CssSelector(".verification-code-code")))
                        break;
                }
                catch
                {
                    return null;
                }
            }

            string code = driver.FindElement(By.CssSelector(".verification-code-code")).Text;
            Console.WriteLine("ДОЖДАЛИСЬ {0}", code);


            HttpClient client = new HttpClient();

            var query = new Dictionary<string, string>
            {
                {  "grant_type","authorization_code"},
                {  "code",code},
                {  "client_id",param["client_id"]},
                {  "client_secret",param["client_secret"]}
            };

            WebRequest request = WebRequest.Create("https://oauth.yandex.ru/token");

            request.Method = "POST";
            string postData = String.Format("grant_type={0}&code={1}&client_id={2}&client_secret={3}",
                   query["grant_type"],
                   query["code"],
                   query["client_id"],
                   query["client_secret"]

               );
            byte[] byteArray = Encoding.UTF8.GetBytes(postData);
            request.ContentType = "application/x-www-form-urlencoded";
            request.ContentLength = byteArray.Length;
            Stream dataStream = request.GetRequestStream();
            dataStream.Write(byteArray, 0, byteArray.Length);
            dataStream.Close();
            WebResponse response = request.GetResponse();
            dataStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(dataStream);
            string responseFromServer = reader.ReadToEnd();
            reader.Close();
            dataStream.Close();
            response.Close();


            return (new JavaScriptSerializer()).Deserialize<Dictionary<string, string>>(responseFromServer)["access_token"];
        }

        public bool isSelectorExist(By selector)
        {
            return driver.FindElements(selector).Count != 0;
        }
    }
}

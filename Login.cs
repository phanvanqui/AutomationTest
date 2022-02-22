using NUnit.Framework;
using RestSharp;
using RestSharp.Serialization.Json;
using SpreadsheetLight;
using System.Collections.Generic;
//using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using SeleniumExtras.WaitHelpers;
using System.IO;

namespace Tests_Data_Driven
{
    class Login
    {
        WebDriver driver;
        IWebElement textBoxUsername;
        IWebElement textBoxPassword;
        IWebElement buttonLogin;
        StreamWriter sw;

        [OneTimeSetUp]
        public void Setup()
        {
            driver = new ChromeDriver();

            sw = new StreamWriter(@"..\\..\\..\\Login_Report.csv", false);
            sw.Write("username,password,result,status");
            sw.Write(sw.NewLine);
            sw.Flush();
        }

        static IEnumerable<TestCaseData> LoadTestDataFromExcel()
        {
            using (var sheet = new SLDocument("../../../Login_Data.xlsx"))
            {
                int endRowIndex = sheet.GetWorksheetStatistics().EndRowIndex;
                for (int row = 2; row <= endRowIndex; row++)
                {
                    string username = sheet.GetCellValueAsString(row, 1);
                    string password = sheet.GetCellValueAsString(row, 2);
                    string status = sheet.GetCellValueAsString(row, 3);
                    yield return new TestCaseData(username, password, status);
                }
            }
        }

        [TestCaseSource("LoadTestDataFromExcel")]
        public void TestCalculatorWebApp(string username, string password, string expectedResult)
        {
            driver.Navigate().GoToUrl("http://www.supermarketnhom11.somee.com/Login.aspx");

            textBoxUsername = driver.FindElement(By.Id("UserName"));
            textBoxPassword = driver.FindElement(By.Id("Password"));
            buttonLogin = driver.FindElement(By.Id("LoginButton"));


            if (username != "")
                textBoxUsername.SendKeys(username);

            if (password != "")
                textBoxPassword.SendKeys(password);


            //new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("LoginButton"))));
            //buttonLogin.Click();
            buttonLogin.SendKeys(Keys.Enter);

            var actualResult = username;

            var classes = driver.FindElements(By.CssSelector("div.text-danger"));
            if (classes.Count > 0)
                actualResult = classes[0].Text;


            sw.Write(username);
            sw.Write(",");
            sw.Write(password);
            sw.Write(",");
            sw.Write(actualResult);
            sw.Write(",");
            sw.Write(expectedResult == actualResult ? "success" : "fail");
            sw.Write(sw.NewLine);
            sw.Flush();


            Assert.AreEqual(expectedResult, actualResult);
        }

        [OneTimeTearDown]
        public void TearDown()
        {
            sw.Flush();
            sw.Close();
            driver.Quit();
        }
    }
}

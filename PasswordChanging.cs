using NUnit.Framework;
using RestSharp;
using RestSharp.Serialization.Json;
using SpreadsheetLight;
using System.Collections.Generic;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.IO;

namespace Tests_Data_Driven
{
    // kiểm tra các trường hợp nhập thay đổi mật khẩu người dùng
    class PasswordChanging
    {
        WebDriver driver;
        IWebElement textCurrentPassword;
        IWebElement textNewPassword;
        IWebElement textConfirmNewPassword;
        IWebElement buttonConfirm;

        StreamWriter sw;

        [OneTimeSetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("http://www.supermarketnhom11.somee.com/Member/Account.aspx");
            driver.FindElement(By.Id("UserName")).SendKeys("pvqui0910");
            driver.FindElement(By.Id("Password")).SendKeys("naruto123@");
            driver.FindElement(By.Id("LoginButton")).SendKeys(Keys.Enter);

            //new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("LoginButton"))));


            sw = new StreamWriter(@"..\\..\\..\\PasswordChanging_Report.csv", false);
            sw.Write("Current Password,New Password,Confirm New Password,result,status");
            sw.Write(sw.NewLine);
            sw.Flush();


        }

        static IEnumerable<TestCaseData> LoadTestDataFromExcel()
        {
            using (var sheet = new SLDocument("../../../PasswordChanging_Data.xlsx"))
            {
                int endRowIndex = sheet.GetWorksheetStatistics().EndRowIndex;
                for (int row = 2; row <= endRowIndex; row++)
                {
                    string CurrentPassword = sheet.GetCellValueAsString(row, 1);
                    string NewPassword = sheet.GetCellValueAsString(row, 2);
                    string ConfirmNewPassword = sheet.GetCellValueAsString(row, 3);
                    string status = sheet.GetCellValueAsString(row, 4);
                    yield return new TestCaseData(CurrentPassword, NewPassword,ConfirmNewPassword, status);
                }
            }
        }

        [TestCaseSource("LoadTestDataFromExcel")]
        public void TestCalculatorWebApp(string CurrentPassword, string NewPassword, string ConfirmNewPassword, string expectedResult)
        {
            driver.Navigate().GoToUrl("http://www.supermarketnhom11.somee.com/Member/Account.aspx");

            textCurrentPassword = driver.FindElement(By.Id("CurrentPassword"));
            textNewPassword = driver.FindElement(By.Id("NewPassword"));
            textConfirmNewPassword = driver.FindElement(By.Id("ConfirmNewPassword"));
            buttonConfirm = driver.FindElement(By.Id("ChangePasswordPushButton"));


            if (CurrentPassword != "")
                textCurrentPassword.SendKeys(CurrentPassword);

            if (NewPassword != "")
                textNewPassword.SendKeys(NewPassword);

            if (ConfirmNewPassword != "")
                textConfirmNewPassword.SendKeys(ConfirmNewPassword);


            //new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementExists((By.Id("LoginButton"))));
            //buttonLogin.Click();
            buttonConfirm.SendKeys(Keys.Enter);

            string content = "";

            var classes = driver.FindElements(By.CssSelector("div.text-danger"));
            if (classes.Count > 0)
                content = classes[0].Text;
            else
                content = driver.FindElement(By.XPath("//table[@id='ChangePasswordMember']/tbody/tr/td/h4")).Text;

            var actualResult = content;


            sw.Write(CurrentPassword);
            sw.Write(",");
            sw.Write(NewPassword);
            sw.Write(",");
            sw.Write(ConfirmNewPassword);
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

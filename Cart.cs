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
using System.Threading;

namespace Tests_Data_Driven
{
    // add sản phẩm vào giỏ hàng sau đó check số lượng sản phẩm trên giỏ hang
    class Cart
    {
        WebDriver driver;

        StreamWriter sw;

        [OneTimeSetUp]
        public void Setup()
        {
            //driver = new ChromeDriver();

            //driver.Navigate().GoToUrl("http://www.supermarketnhom11.somee.com/Member/Account.aspx");
            //driver.FindElement(By.Id("UserName")).SendKeys("pvqui0910");
            //driver.FindElement(By.Id("Password")).SendKeys("naruto123@");
            //driver.FindElement(By.Id("LoginButton")).SendKeys(Keys.Enter);

            sw = new StreamWriter(@"..\\..\\..\\Cart_Report.csv", false);
            sw.Write("link,numSelect,result,status");
            sw.Write(sw.NewLine);
            sw.Flush();


        }

        static IEnumerable<TestCaseData> LoadTestDataFromExcel()
        {
            using (var sheet = new SLDocument("../../../Cart_Data.xlsx"))
            {
                int endRowIndex = sheet.GetWorksheetStatistics().EndRowIndex;
                for (int row = 2; row <= endRowIndex; row++)
                {
                    string link = sheet.GetCellValueAsString(row, 1);
                    int numSelect = sheet.GetCellValueAsInt32(row, 2);
                    int status = sheet.GetCellValueAsInt32(row, 3);
                    yield return new TestCaseData(link, numSelect, status);
                }
            }
        }

        [TestCaseSource("LoadTestDataFromExcel")]
        public void TestCalculatorWebApp(string link, int numSelect, int expectedResult)
        {
            driver = new ChromeDriver();

            driver.Navigate().GoToUrl("http://www.supermarketnhom11.somee.com/Login.aspx");
            driver.FindElement(By.Id("UserName")).SendKeys("pvqui0910");
            driver.FindElement(By.Id("Password")).SendKeys("naruto123@");
            driver.FindElement(By.Id("LoginButton")).SendKeys(Keys.Enter);



            //new WebDriverWait(driver, TimeSpan.FromSeconds(25)).Until(ExpectedConditions.ElementExists(By.CssSelector("input.button")));

            //var classes = driver.FindElements(By.CssSelector("products-right"));
            //var product = classes[0].FindElements(By.CssSelector("button"));
            //var product = driver.FindElements(By.CssSelector("input.button"));

            //Console.WriteLine(numSelect);
            //Console.WriteLine(product.Count);
            //Console.WriteLine(product.Count);
            for (int i = 0; i < numSelect; i++)
            {
                ////new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(ExpectedConditions.ElementToBeClickable(By.Id("MainContent_ItemList_ProductsListView_ctrl0_ButtonAddToCart_" + i)));
                //new WebDriverWait(driver, TimeSpan.FromSeconds(2));
                ////IWebElement prod = driver.FindElement(By.Id("MainContent_ItemList_ProductsListView_ctrl0_ButtonAddToCart_"+i));
                ////prod.Click();
                //Thread.Sleep(2000);
                //product[i].SendKeys(Keys.Enter);
                ////driver.ExecuteScript("arguments[0].click();", product[0]);
                ///

                driver.Navigate().GoToUrl(link);
                var product = driver.FindElements(By.CssSelector("input.button"));

                //var product = driver.FindElement(By.Id("MainContent_ItemList_ProductsListView_ctrl0_ButtonAddToCart_" + i));
                product[0].SendKeys(Keys.Enter);
            }

            var count = driver.FindElement(By.Id("Header_LabelItemCount"));
            int num = Int32.Parse(count.Text);


            sw.Write(link);
            sw.Write(",");
            sw.Write(numSelect);
            sw.Write(",");
            sw.Write(expectedResult);
            sw.Write(",");
            sw.Write(expectedResult == num ? "success" : "fail");
            sw.Write(sw.NewLine);
            sw.Flush();


            Assert.AreEqual(expectedResult, num);

            driver.Close();
            driver.Quit();
            driver.Dispose();
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

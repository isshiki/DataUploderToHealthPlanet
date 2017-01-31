using System;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using System.IO;
using System.Diagnostics;

namespace DataUploderToHealthPlanet
{
    class Program
    {
        static void Main(string[] args)
        {
            // InternetExplorer
            IWebDriver _webDriver = new InternetExplorerDriver();

            try
            {
                _webDriver.Url = "https://www.healthplanet.jp/innerscan.do";
                // return; // 最初にここで終了してIEで［ログインしたままにする］オプションをオンにしてログインしてください。以下ではログイン処理は書いていないため

                var lines = File.ReadAllLines(@"weight.csv");
                foreach (var line in lines)
                {
                    if (line.Trim().Length == 0) continue;

                    var cells = line.Split(',');
                    if (cells.Length != 4) { Debug.Assert(false); continue; }

                    _webDriver.Url = "https://www.healthplanet.jp/innerscan.do?date=" + cells[0];
                    WaitForLoad(_webDriver);

                    IWebElement measurementTimeHH = _webDriver.FindElement(By.Name("measurementTimeHH"));
                    if (measurementTimeHH.Size.Height == 0) continue; // 入力済み
                    measurementTimeHH.Click();
                    measurementTimeHH.SendKeys(Keys.Backspace);
                    measurementTimeHH.SendKeys(Keys.Backspace);
                    measurementTimeHH.SendKeys("08");

                    IWebElement measurementTimeMM = _webDriver.FindElement(By.Name("measurementTimeMM"));
                    measurementTimeMM.Click();
                    measurementTimeMM.SendKeys(Keys.Backspace);
                    measurementTimeMM.SendKeys(Keys.Backspace);
                    measurementTimeMM.SendKeys("00");

                    IWebElement innerscanBean0keyData = _webDriver.FindElement(By.Name("innerscanBean[0].keyData"));
                    innerscanBean0keyData.SendKeys(cells[1]);

                    IWebElement innerscanBean1keyData = _webDriver.FindElement(By.Name("innerscanBean[1].keyData"));
                    innerscanBean1keyData.SendKeys(cells[2]);

                    IWebElement innerscanBean2keyData = _webDriver.FindElement(By.Name("innerscanBean[2].keyData"));
                    innerscanBean2keyData.SendKeys(cells[3]);

                    IJavaScriptExecutor js = _webDriver as IJavaScriptExecutor;
                    js.ExecuteScript("goSubmit('/innerscan_confirm.do');");
                    WaitForLoad(_webDriver);
                    WaitForLoad(_webDriver);

                    try
                    {
                        js.ExecuteScript("goSubmit('/innerscan_complete.do');");
                    }
                    catch (Exception)
                    {
                        js.ExecuteScript("goSubmit('/innerscan_complete.do');");
                    }
                    WaitForLoad(_webDriver);

                    Console.WriteLine("Done - " + cells[0]);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private static void WaitForLoad(IWebDriver driver)
        {
            try
            {
                var wait = new OpenQA.Selenium.Support.UI.WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
                wait.Until(driver1 => ((IJavaScriptExecutor)driver).ExecuteScript("return document.readyState").Equals("complete"));
            }
            catch (Exception)
            {
            }
        }

    }
}

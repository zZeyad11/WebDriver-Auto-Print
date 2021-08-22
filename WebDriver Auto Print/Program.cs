using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.html.simpleparser;
using iTextSharp.tool.xml;
using System.IO;
using System.Web;
using Corsis.Xhtml;
using System.Xml;
using Sgml;
using Aspose.Html.Saving;
using System.Windows.Forms;
using WindowsInput.Native;
using WindowsInput;
using SmartBear.TestLeft;
using Microsoft.VisualBasic.Devices;

namespace WebDriver_Auto_Print
{
    class Program : System.Web.UI.Page
    {
        static void Main(string[] args)
        {
           
            var I = new Program();
            I.Main2();


            Console.ReadLine();

        }



        public string ProcessString(string strInputHtml)
        {
            string strOutputXhtml = String.Empty;
            SgmlReader reader = new SgmlReader();
            reader.DocType = "HTML";
            StringReader sr = new System.IO.StringReader(strInputHtml);
            reader.InputStream = sr;
            StringWriter sw = new StringWriter();
            XmlTextWriter w = new XmlTextWriter(sw);
            reader.Read();
            while (!reader.EOF)
            {
                w.WriteNode(reader, true);
            }
            w.Flush();
            w.Close();
            return sw.ToString();
        }



        public async void Main2()
        {
            
            var driver = ChromeDriverService.CreateDefaultService();
            driver.HideCommandPromptWindow = true;

            var Webdriver = new ChromeDriver(driver, new ChromeOptions());
            Webdriver.Navigate().GoToUrl("http://dgftebrc.nic.in:8100/BRCQueryTrade/index.jsp");
            Console.WriteLine("Hit Any Key To start");
            Console.ReadKey();

            var Main = Webdriver.CurrentWindowHandle;
            var Elements = Webdriver.FindElementsByTagName("input");

            List<IWebElement> PrintButtons = new List<IWebElement>();
            //Get Print Buttons
            foreach (var e in Elements)
            {
                if (e.GetAttribute("type") == "submit" && e.GetAttribute("value") == "Print")
                {
                    PrintButtons.Add(e);
                }

            }


        
            foreach(var printBtn in PrintButtons)
            {
                printBtn.Click();

                var windowHandles = Webdriver.WindowHandles;
                if (!(windowHandles.Count == 0))
                {

                    Webdriver.SwitchTo().Window((String)windowHandles[windowHandles.Count - 1]);

                }

                while (Webdriver.PageSource.Length < 5000) { }


                var document = new Aspose.Html.HTMLDocument(Webdriver.PageSource, "");
                
                // Create Instance of PDF Options 
                var options = new PdfSaveOptions();
                
                // save XHTML as a PDF 
                Aspose.Html.Converters.Converter.ConvertHTML(document, options, Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + $"\\output_{PrintButtons.IndexOf(printBtn)}.pdf");

                Webdriver.Close();
                Webdriver.SwitchTo().Window((String)Main);

               
            }

            Webdriver.Close();
            Console.WriteLine("All Is Done And Saved On Desktop");
            Console.WriteLine("Exit IN 3...");
            await Task.Delay(1000);
            Console.WriteLine("Exit IN 2...");
            await Task.Delay(1000);
            Console.WriteLine("Exit IN 1...");
            await Task.Delay(1000);
            Environment.Exit(0);

        }


       

    }
}


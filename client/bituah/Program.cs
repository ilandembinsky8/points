using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.IO;
using Newtonsoft.Json;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System.Net;
using OpenQA.Selenium.Interactions;
using DeathByCaptcha;

namespace bituah
{
    class Program
    {
        static readonly bool DEBUG_SKIP_CAPCHA = false; //skipping capcha for debug - should be false in regular usage
        static readonly int CAPCHA_TRIALS = 3;
        static readonly int MONTH_STOP = 6;
        //14- july19 tel aviv, zichron
        //15- july19 alef, small
        static readonly string CURR_RUN = "test0015";
        static Dictionary<string, string> data_values;
        static readonly string[] server_fields = {
            "area",
            "gush_helka",
            "iska_date",
            "CITY",
            "STREET",
            "NUMBER",
            "ENTERANCE",
            "APPARTMENT",
            "TPRICENIS",
            "TPRICEDOLLAR",
            "EPRICENIS",
            "EPRICEDOLLAR",
            "ARNONASURFCE",
            "ROOMS",
            "ROOF",
            "LISTEDSURFACE",
            "FLOOR",
            "STOREHOUSE",
            "BUILDYEAR",
            "FLOORS",
            "YARD",
            "PRICEPERROOM",
            "APARTMENTS",
            "FIELD",
            "PRICEPERMETER",
            "PARKING",
            "GALLERY",
            "ELEVATOR",
            "DEALTYPE",
            "BUILDINGFUNC",
            "UNITFUNC",
            "SHUMAPARTS",
            "GUSHMOFA",
            "TABA",
            "MAHUT"
        };
        static readonly string[] html_fields = {
            "ContentUsersPage_lblEzor",
            "ContentUsersPage_lblGush",
            "ContentUsersPage_lblTarIska",
            "ContentUsersPage_lblYeshuv",
            "ContentUsersPage_lblRechov",
            "ContentUsersPage_lblBayit",
            "ContentUsersPage_lblKnisa",
            "ContentUsersPage_lblDira",
            "ContentUsersPage_lblMcirMozhar",
            "ContentUsersPage_lblMcirMozharDlr",
            "ContentUsersPage_lblMcirMorach",
            "ContentUsersPage_lblMcirMorachDlr",
            "ContentUsersPage_lblShetachBruto",
            "ContentUsersPage_lblMisHadarim",
            "ContentUsersPage_lblGag",
            "ContentUsersPage_lblShetachNeto",
            "ContentUsersPage_lblKoma",
            "ContentUsersPage_lblMachsan",
            "ContentUsersPage_lblShnatBniya",
            "ContentUsersPage_lblMisKomot",
            "ContentUsersPage_lblHzer",
            "ContentUsersPage_lblMechirCheder",
            "ContentUsersPage_lblDirotBnyn",
            "ContentUsersPage_lblMigrash",
            "ContentUsersPage_lblMechirLmr",
            "ContentUsersPage_lblHanaya",
            "ContentUsersPage_lblGlrya",
            "ContentUsersPage_lblMalit",
            "ContentUsersPage_lblSugIska",
            "ContentUsersPage_lblTifkudBnyn",
            "ContentUsersPage_lblTifkudYchida",
            "ContentUsersPage_lblShumaHalakim",
            "ContentUsersPage_lblMofaGush",
            "ContentUsersPage_lblTava",
            "ContentUsersPage_lblMahutZchut"
        };

        public enum DatePagingMode
        {
            Daily,
            Monthly,
            TenDays
        }

        public struct Data
        {
            public string area;
            public string gush_helka;
            public string iska_date;
            public string CITY;
            public string STREET;
            public string NUMBER;
            public string ENTERANCE;
            public string APPARTMENT;
            public string TPRICENIS;
            public string TPRICEDOLLAR;
            public string EPRICENIS;
            public string EPRICEDOLLAR;
            public string ARNONASURFACE;
            public string ROOMS;
            public string ROOF;
            public string LISTEDSURFACE;
            public string FLOOR;
            public string STOREHOUSE;
            public string BUILDYEAR;
            public string FLOORS;
            public string YARD;
            public string PRICEPERROOM;
            public string APARTMENTS;
            public string FIELD;
            public string PRICEPERMETER;
            public string PARKING;
            public string GALLERY;
            public string ELEVATOR;
            public string DEALTYPE;
            public string BUILDINGFUNC;
            public string UNITFUNC;
            public string SHUMAPARTS;
            public string GUSHMOFA;
            public string TABA;
            public string MAHUT;
        }

        static void SendData(string value)
        {
            //Console.OutputEncoding = System.Text.Encoding.BigEndianUnicode;
            //Console.WriteLine(value);
            try
            {
                // Create a request using a URL that can receive a post.   
                WebRequest request = WebRequest.Create("http://www.yan-systems.co.il/points/processData.asp ");
                // Set the Method property of the request to POST.  
                request.Method = "POST";

                // Create POST data and convert it to a byte array.  
                value = value.Replace(" ", "%20");
                value = "json =" + value;
                byte[] byteArray = Encoding.UTF8.GetBytes(value);

                // Set the ContentType property of the WebRequest.  
                request.ContentType = "application/x-www-form-urlencoded";
                // Set the ContentLength property of the WebRequest.  
                request.ContentLength = byteArray.Length;

                // Get the request stream.  
                Stream dataStream = request.GetRequestStream();
                // Write the data to the request stream.  
                dataStream.Write(byteArray, 0, byteArray.Length);
                // Close the Stream object.  
                dataStream.Close();

                // Get the response.  
                WebResponse response = request.GetResponse();
                // Display the status.  
                Console.WriteLine(((HttpWebResponse)response).StatusDescription);

                // Get the stream containing content returned by the server.  
                // The using block ensures the stream is automatically closed.
                using (dataStream = response.GetResponseStream())
                {
                    // Open the stream using a StreamReader for easy access.  
                    StreamReader reader = new StreamReader(dataStream);
                    // Read the content.  
                    string responseFromServer = reader.ReadToEnd();
                    // Display the content.  
                    Console.WriteLine(responseFromServer);
                }

                // Close the response.  
                response.Close();

            }
            catch (System.Exception ex)
            {
                    Console.WriteLine(ex.Message);
            }
        }

        static void Main(string[] args)
        {
            //SendData("{ \r\n   \"area\":\"NNN\",\r\n   \"gush_helka\":\"NNN\",\r\n   \"iska_date\":\"YYYYMMDD\",\r\n   \"CITY\":\"NNN\",\r\n   \"STREET\":\"NNN\",\r\n   \"NUMBER\":\"999\",\r\n   \"ENTERANCE\":\"NNN\",\r\n   \"APPARTMENT\":\"999\",\r\n   \"TPRICENIS\":\"999\",\r\n   \"TPRICEDOLLAR\":\"999\",\r\n   \"EPRICENIS\":\"999\",\r\n   \"EPRICEDOLLAR\":\"999\",\r\n   \"ARNONASURFCE\":\"999\",\r\n   \"ROOMS\":\"999.99\",\r\n   \"ROOF\":\"999\",\r\n   \"LISTEDSURFACE\":\"999\",\r\n   \"FLOOR\":\"999\",\r\n   \"STOREHOUSE\":\"999\",\r\n   \"BUILDYEAR\":\"999\",\r\n   \"FLOORS\":\"999\",\r\n   \"YARD\":\"999\",\r\n   \"PRICEPERROOM\":\"999\",\r\n   \"APARTMENTS\":\"999\",\r\n   \"FIELD\":\"999\",\r\n   \"PRICEPERMETER\":\"999\",\r\n   \"PARKING\":\"NNN\",\r\n   \"GALLERY\":\"999\",\r\n   \"ELEVATOR\":\"NNN\",\r\n   \"DEALTYPE\":\"NNN\",\r\n   \"BUILDINGFUNC\":\"NNN\",\r\n   \"UNITFUNC\":\"NNN\",\r\n   \"SHUMAPARTS\":\"NNN\",\r\n   \"GUSHMOFA\":\"NNN\",\r\n   \"TABA\":\"NNN\",\r\n   \"MAAHUT\":\"NNN\"\r\n}");
            //SendData("{ \r\n   \"area\":\"50 - מק'-תל אביב\",\r\n   \"gush_helka\":\"724470051005\",\r\n   \"iska_date\":\"02/09/2019\",\r\n   \"CITY\":\"תל אביב -יפו\",\r\n   \"STREET\":\"מעפילי אגוז\",\r\n   \"NUMBER\":\"67\",\r\n   \"ENTERANCE\":\"--\",\r\n   \"APPARTMENT\":\"5\",\r\n   \"TPRICENIS\":\"1,570,999\",\r\n   \"TPRICEDOLLAR\":\"444,036\",\r\n   \"EPRICENIS\":\"1,570,999\",\r\n   \"EPRICEDOLLAR\":\"444,036\",\r\n   \"ARNONASURFCE\":\"67\",\r\n   \"ROOMS\":\"3.0\",\r\n   \"ROOF\":\"0\",\r\n   \"LISTEDSURFACE\":\"67\",\r\n   \"FLOOR\":\"3.0\",\r\n   \"STOREHOUSE\":\"0\",\r\n   \"BUILDYEAR\":\"1960\",\r\n   \"FLOORS\":\"4\",\r\n   \"YARD\":\"0\",\r\n   \"PRICEPERROOM\":\"0\",\r\n   \"APARTMENTS\":\"8\",\r\n   \"FIELD\":\"0\",\r\n   \"PRICEPERMETER\":\"0\",\r\n   \"PARKING\":\"0 רכבים\",\r\n   \"GALLERY\":\"0\",\r\n   \"ELEVATOR\":\"\",\r\n   \"DEALTYPE\":\"מכר\",\r\n   \"BUILDINGFUNC\":\"מגורים\",\r\n   \"UNITFUNC\":\"דירה בבית קומות\",\r\n   \"SHUMAPARTS\":\"1 / 1 ליחידה בשלמותה\",\r\n   \"GUSHMOFA\":\"2\",\r\n   \"TABA\":\"--\",\r\n   \"MAAHUT\":\"בעלות\"\r\n}");
            IWebDriver driver;
            string homeURL;
            string dbcUsername = "iland";
            string dbcPassword = "Dambi1965";
            DatePagingMode datePagingMode = DatePagingMode.Monthly;

            homeURL = "https://www.misim.gov.il/svinfonadlan2010/startpageNadlanNewDesign.aspx";
            driver = new ChromeDriver("C:\\");
            driver.Navigate().GoToUrl(homeURL);
            WebDriverWait wait = new WebDriverWait(driver, new TimeSpan(0, 0, 15));
            wait.Until(x => x.FindElement(By.Id("ContentUsersPage_RadCaptcha1_CaptchaImageUP")));
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2); //wait another 2 seconds

            //loop towns from file
            string[] towns = File.ReadAllLines("towns.txt");
            Dictionary<string, int> towns_capcha_counter = new Dictionary<string, int>();
            for (int town_count = 0; town_count < towns.Length; town_count++)
            {
                string town = towns[town_count];
                if (town != null && town.Length > 0)
                {
                    //Encoding iso = Encoding.GetEncoding("Windows-1252");
                    //Encoding utf8 = Encoding.UTF8;
                    //byte[] utfBytes = utf8.GetBytes(town);
                    //byte[] isoBytes = Encoding.Convert(iso, utf8, utfBytes);
                    //string town_ = iso.GetString(isoBytes);

                    Console.WriteLine("processing town " + town);
                    towns_capcha_counter.Add(town, 0);
                    //Console.WriteLine(town);
                    IWebElement txtTeshuv = driver.FindElement(By.Id("txtYeshuv"));
                    txtTeshuv.Click();
                    txtTeshuv.Clear();
                    for (int i = 0; i < town.Length; i++)
                    {
                        if (town[i] == '-')
                        {
                            break;
                        }
                        txtTeshuv.SendKeys(town[i].ToString());
                        System.Threading.Thread.Sleep(100);
                    }
                    System.Threading.Thread.Sleep(1000);
                    txtTeshuv.SendKeys(Keys.ArrowDown);
                    System.Threading.Thread.Sleep(1000);
                    txtTeshuv.SendKeys(Keys.Return);

                    //loop ContentUsersPage_DDLTypeNehes - values 1,2,5,6,9
                    string[] typesNehes = { "1", "2", "5", "6", "9" };
                    //SelectElement drpNehes = new SelectElement(driver.FindElement(By.Id("ContentUsersPage_DDLTypeNehes")));
                    //SelectElement drpDateType, drpMahutIska;
                    for (int typeNehes_count = 0; typeNehes_count < typesNehes.Length; typeNehes_count++)
                    {
                        try
                        {
                            

                            //loop months from now back
                            int curr_year = DateTime.Now.Year;
                            //int curr_month = DateTime.Now.Month;
                            int curr_month = 7;
                            //int curr_day_of_month = DateTime.Now.Day;
                            int curr_day_of_month = 31;

                            IWebElement fromdate;
                            IWebElement todate;
                            bool is_first_search = true; //for current month only - range is from 1st till today. for prev months range is from 1st to last day of month
                            int last_day_of_month;

                            while (curr_month > MONTH_STOP)
                            {
                                homeURL = "https://www.misim.gov.il/svinfonadlan2010/startpageNadlanNewDesign.aspx";
                                driver.Close();
                                driver = new ChromeDriver("C:\\");
                                driver.Navigate().GoToUrl(homeURL);
                                wait = new WebDriverWait(driver, new TimeSpan(0, 0, 15));
                                wait.Until(x => x.FindElement(By.Id("ContentUsersPage_RadCaptcha1_CaptchaImageUP")));
                                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2); //wait another 2 seconds
                                txtTeshuv = driver.FindElement(By.Id("txtYeshuv"));
                                txtTeshuv.Click();
                                txtTeshuv.Clear();
                                for (int i = 0; i < town.Length; i++)
                                {
                                    if (town[i] == '-')
                                    {
                                        break;
                                    }
                                    txtTeshuv.SendKeys(town[i].ToString());
                                    System.Threading.Thread.Sleep(100);
                                }
                                System.Threading.Thread.Sleep(1000);
                                txtTeshuv.SendKeys(Keys.ArrowDown);
                                System.Threading.Thread.Sleep(1000);
                                txtTeshuv.SendKeys(Keys.Return);

                                SelectElement drpNehes = new SelectElement(driver.FindElement(By.Id("ContentUsersPage_DDLTypeNehes")));
                                SelectElement drpDateType, drpMahutIska;


                                string typeNehes = typesNehes[typeNehes_count];
                                drpNehes.SelectByValue(typeNehes);
                                System.Threading.Thread.Sleep(200);

                                //loop dates - ContentUsersPage_DDLDateType
                                //selet date range cusotom ContentUsersPage_DDLDateType, value = 1 
                                drpDateType = new SelectElement(driver.FindElement(By.Id("ContentUsersPage_DDLDateType")));
                                drpDateType.SelectByValue("1");

                                //select all deals type ContentUsersPage_DDLMahutIska, value = 999
                                drpMahutIska = new SelectElement(driver.FindElement(By.Id("ContentUsersPage_DDLMahutIska")));
                                drpMahutIska.SelectByValue("999");

                                fromdate = driver.FindElement(By.Id("ctl00_ContentUsersPage_txtyomMechira_dateInput"));
                                todate = driver.FindElement(By.Id("ctl00_ContentUsersPage_txtadYomMechira_dateInput"));
                                if (datePagingMode == DatePagingMode.Monthly)
                                {
                                    //get last day of month
                                    last_day_of_month = DateTime.DaysInMonth(curr_year, curr_month);
                                    //enter from date ctl00_ContentUsersPage_txtyomMechira_dateInput
                                    if (fromdate != null)
                                    {
                                        fromdate.Click();
                                        fromdate.SendKeys("01");
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        if (curr_month < 10)
                                        {
                                            fromdate.SendKeys("0");
                                        }
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys(curr_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys(curr_year.ToString());
                                    }
                                    //enter to date ctl00_ContentUsersPage_txtadYomMechira_dateInput
                                    if (todate != null)
                                    {
                                        todate.Click();
                                        if (is_first_search)
                                        {
                                            if (curr_day_of_month < 10)
                                            {
                                                todate.SendKeys("0");
                                            }
                                            System.Threading.Thread.Sleep(100);
                                            todate.SendKeys(curr_day_of_month.ToString());
                                        }
                                        else
                                        {
                                            todate.SendKeys(last_day_of_month.ToString());
                                        }
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        if (curr_month < 10)
                                        {
                                            todate.SendKeys("0");
                                        }
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys(curr_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys(curr_year.ToString());
                                    }
                                }
                                else if (datePagingMode == DatePagingMode.Daily)
                                {
                                    //enter from date ctl00_ContentUsersPage_txtyomMechira_dateInput
                                    if (fromdate != null)
                                    {
                                        fromdate.Click();
                                        if (curr_day_of_month < 10)
                                        {
                                            fromdate.SendKeys("0");
                                        }
                                        fromdate.SendKeys(curr_day_of_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        if (curr_month < 10)
                                        {
                                            fromdate.SendKeys("0");
                                        }
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys(curr_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        fromdate.SendKeys(curr_year.ToString());
                                    }
                                    if (todate != null)
                                    {
                                        todate.Click();
                                        if (curr_day_of_month < 10)
                                        {
                                            todate.SendKeys("0");
                                        }
                                        todate.SendKeys(curr_day_of_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        if (curr_month < 10)
                                        {
                                            todate.SendKeys("0");
                                        }
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys(curr_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys(curr_year.ToString());
                                    }
                                }
                                else if (datePagingMode == DatePagingMode.TenDays)
                                {
                                    //enter to date ctl00_ContentUsersPage_txtadYomMechira_dateInput
                                    if (todate != null)
                                    {
                                        todate.Click();
                                        if (curr_day_of_month < 10)
                                        {
                                            todate.SendKeys("0");
                                        }
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys(curr_day_of_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        if (curr_month < 10)
                                        {
                                            todate.SendKeys("0");
                                        }
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys(curr_month.ToString());
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys("/");
                                        System.Threading.Thread.Sleep(100);
                                        todate.SendKeys(curr_year.ToString());

                                        if (curr_day_of_month - 10 > 0)
                                        {
                                            curr_day_of_month -= 10;
                                        }
                                        else
                                        {
                                            curr_day_of_month = 1;
                                        }
                                        //enter from date ctl00_ContentUsersPage_txtyomMechira_dateInput
                                        if (fromdate != null)
                                        {
                                            fromdate.Click();
                                            if (curr_day_of_month < 10)
                                            {
                                                fromdate.SendKeys("0");
                                            }
                                            fromdate.SendKeys(curr_day_of_month.ToString());
                                            System.Threading.Thread.Sleep(100);
                                            fromdate.SendKeys("/");
                                            System.Threading.Thread.Sleep(100);
                                            if (curr_month < 10)
                                            {
                                                fromdate.SendKeys("0");
                                            }
                                            System.Threading.Thread.Sleep(100);
                                            fromdate.SendKeys(curr_month.ToString());
                                            System.Threading.Thread.Sleep(100);
                                            fromdate.SendKeys("/");
                                            System.Threading.Thread.Sleep(100);
                                            fromdate.SendKeys(curr_year.ToString());
                                        }
                                    }
                                }
                                towns_capcha_counter[town]++;
                                //Console.WriteLine(town + ":" + towns_capcha_counter[town]);
                                if (!DEBUG_SKIP_CAPCHA) //on debug skip the capcha and just loop
                                {
                                    bool capcha_success = false;
                                    int curr_capcha_trial = 0;
                                    //save capcha image ContentUsersPage_RadCaptcha1_CaptchaImageUP
                                    while (!capcha_success && curr_capcha_trial < CAPCHA_TRIALS)
                                    {
                                        if (!driver.PageSource.Contains("ctl00$ContentUsersPage$txtYeshuv"))
                                        {
                                            driver.FindElement(By.Id("menuItem1")).Click();//back to query page
                                        }
                                        IWebElement element = driver.FindElement(By.Id("ContentUsersPage_RadCaptcha1_CaptchaImageUP"));
                                        string src = element.GetAttribute("src");
                                        System.Drawing.Point loc = element.Location;
                                        System.Drawing.Size siz = element.Size;

                                        //driver.save_screenshot(path)

                                        IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                                        string script = @"var c = document.createElement('canvas');
                                            var ctx = c.getContext('2d');
                                            var img = document.getElementById('ContentUsersPage_RadCaptcha1_CaptchaImageUP');
                                            c.height=img.height;
                                            c.width=img.width;
                                            ctx.drawImage(img, 0, 0,img.width, img.height);
                                            var base64String = c.toDataURL();
                                            return base64String;";


                                        string base64string = (string)js.ExecuteScript(script);
                                        var base64 = base64string.Split(',').Last();
                                        using (var stream = new MemoryStream(Convert.FromBase64String(base64)))
                                        {
                                            using (var bitmap = new System.Drawing.Bitmap(stream))
                                            {
                                                var filepath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "capcha_img.jpg");
                                                bitmap.Save(filepath, System.Drawing.Imaging.ImageFormat.Jpeg);
                                            }
                                        }
                                        Client client = (Client)new DeathByCaptcha.HttpClient(dbcUsername, dbcPassword);
                                        Console.WriteLine("Your balance is {0:F2} US cents", client.Balance);
                                        Captcha captcha = client.Decode("capcha_img.jpg", Client.DefaultTimeout);
                                        if (null != captcha)
                                        {
                                            Console.WriteLine("CAPTCHA {0:D} solved: {1}",
                                                              captcha.Id, captcha.Text);
                                            string capcha_code = captcha.Text.Replace("o", "0").Replace("O", "0");
                                            IWebElement capcha_input = driver.FindElement(By.Id("ContentUsersPage_RadCaptcha1_CaptchaTextBox"));

                                            capcha_input.Click();
                                            for (int j = 0; j < capcha_code.Length; j++)
                                            {
                                                capcha_input.SendKeys(capcha_code[j].ToString());
                                                System.Threading.Thread.Sleep(100);
                                            }
                                            System.Threading.Thread.Sleep(100);
                                            driver.FindElement(By.Id("ContentUsersPage_btnHipus")).Click();
                                            WebDriverWait wait_capcha = new WebDriverWait(driver, new TimeSpan(0, 0, 15));
                                            try
                                            {
                                                System.Threading.Thread.Sleep(1000);
                                                //check if message "לא נמצאו נתונים לחתך המבוקש" shows
                                                //ContentUsersPage_LblAlert, ContentUsersPage_lblerrDanger
                                                if (driver.PageSource.Contains("ContentUsersPage_LblAlert"))
                                                {
                                                    //do nothing, no data, go to next loop
                                                }
                                                else
                                                {
                                                    if (driver.PageSource.Contains("ContentUsersPage_RadCaptcha1_ctl00"))
                                                    {
                                                        IWebElement wrong_capcha = driver.FindElement(By.Id("ContentUsersPage_RadCaptcha1_ctl00"));
                                                        if (wrong_capcha != null && wrong_capcha.Text == "הקלדה שגויה")
                                                        {
                                                            capcha_success = false;
                                                            curr_capcha_trial++;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        //all is good.
                                                        int num_of_pages = 1;
                                                        wait_capcha.Until(x => x.FindElement(By.Id("ContentUsersPage_btnExportExcel")));
                                                        //loop over pages
                                                        int curr_page = 1;
                                                        List<string> pages_list = new List<string>();
                                                        pages_list.Add("first_page"); //always read data from first page
                                                                                      //check if there is paging - more than one page
                                                        if (driver.PageSource.Contains("table_title tabelPages")) //there's more than one page
                                                        {
                                                            int page_counter = 2;
                                                            while (driver.PageSource.Contains("__doPostBack('ctl00$ContentUsersPage$GridMultiD1','Page$" + page_counter.ToString() + "'"))
                                                            {
                                                                pages_list.Add("javascript:__doPostBack('ctl00$ContentUsersPage$GridMultiD1','Page$" + page_counter.ToString() + "')");
                                                                page_counter++;
                                                            }

                                                        }


                                                        //ContentUsersPage_GridMultiD1 - data grid
                                                        if (driver.PageSource.Contains("ContentUsersPage_GridMultiD1"))
                                                        {
                                                            while (pages_list.Count > 0)
                                                            {
                                                                if (pages_list[0].Contains("first_page"))
                                                                {
                                                                    pages_list.RemoveAt(0);
                                                                }
                                                                else
                                                                {
                                                                    driver.FindElement(By.CssSelector("[href*=\"" + pages_list[0] + "\"]")).Click();
                                                                    pages_list.RemoveAt(0);
                                                                }
                                                                IWebElement data_grid = driver.FindElement(By.Id("ContentUsersPage_GridMultiD1"));
                                                                //var odd_rows = data_grid.FindElements(By.ClassName("row1")); //get odd rows
                                                                //var even_rows = data_grid.FindElements(By.ClassName("BoxB")); //get even rows
                                                                int curr_odd_row = 0;
                                                                int curr_even_row = 0;
                                                                bool got_all_data = false;
                                                                while (!got_all_data)
                                                                {
                                                                    data_grid = driver.FindElement(By.Id("ContentUsersPage_GridMultiD1"));
                                                                    var odd_rows = data_grid.FindElements(By.ClassName("row1")); //get odd rows
                                                                    var even_rows = data_grid.FindElements(By.ClassName("BoxB")); //get even rows

                                                                    if (odd_rows != null && odd_rows.Count > curr_odd_row)
                                                                    {
                                                                        var details_url = odd_rows[curr_odd_row].FindElements(By.TagName("td"));
                                                                        curr_odd_row++;
                                                                        if (details_url != null)
                                                                        {
                                                                            var inner_data = details_url[1].FindElements(By.TagName("a"));
                                                                            if (inner_data != null && inner_data.Count > 0)
                                                                            {
                                                                                string data_json = "{\r\n";
                                                                                data_values = new Dictionary<string, string>();
                                                                                inner_data[0].Click();
                                                                                WebDriverWait data_wait = new WebDriverWait(driver, new TimeSpan(0, 0, 10));
                                                                                data_wait.Until(x => x.FindElement(By.Id("ContentUsersPage_btHazara")));
                                                                                for (int i = 0; i < html_fields.Length; i++)
                                                                                {
                                                                                    string curr_field = html_fields[i];
                                                                                    if (driver.PageSource.Contains(curr_field))
                                                                                    {
                                                                                        string curr_field_value = driver.FindElement(By.Id(curr_field)).Text;
                                                                                        curr_field_value = curr_field_value.Replace("\"", "\\\"");
                                                                                        //curr_field_value = curr_field_value.Replace("\\\"", "");
                                                                                        //curr_field_value = curr_field_value.Replace("\"", "");
                                                                                        if (curr_field_value != null)
                                                                                        {
                                                                                            data_values.Add(curr_field, curr_field_value);
                                                                                            data_json += "\"" + server_fields[i] + "\":\"" + curr_field_value + "\"";

                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            data_json += "\"" + server_fields[i] + "\":\"\"";
                                                                                        }
                                                                                        //we add the run as last field, so no need to remove comma from last field
                                                                                        //if (i + 1 == html_fields.Length)
                                                                                        //{
                                                                                        //    data_json += "\r\n";
                                                                                        //}
                                                                                        //else
                                                                                        //{
                                                                                        data_json += ",\r\n";
                                                                                        //}
                                                                                    }
                                                                                }
                                                                                data_json += "\"RUN\":\"" + CURR_RUN + "\"\r\n"; //add the run as last field
                                                                                data_json += "}";
                                                                                SendData(data_json);
                                                                                driver.FindElement(By.Id("ContentUsersPage_btHazara")).Click();//back to data grid table
                                                                            }
                                                                            else
                                                                            {
                                                                                //take data from current table
                                                                                int k = 8;
                                                                                //details_url[2].Text;
                                                                            }
                                                                        }
                                                                    }
                                                                    data_grid = driver.FindElement(By.Id("ContentUsersPage_GridMultiD1"));
                                                                    odd_rows = data_grid.FindElements(By.ClassName("row1")); //get odd rows
                                                                    even_rows = data_grid.FindElements(By.ClassName("BoxB")); //get even rows

                                                                    if (even_rows != null && even_rows.Count > curr_even_row)
                                                                    {
                                                                        var details_url = even_rows[curr_even_row].FindElements(By.TagName("td"));
                                                                        curr_even_row++;
                                                                        if (details_url != null)
                                                                        {
                                                                            var inner_data = details_url[1].FindElements(By.TagName("a"));
                                                                            if (inner_data != null && inner_data.Count > 0)
                                                                            {
                                                                                string data_json = "{\r\n";
                                                                                data_values = new Dictionary<string, string>();
                                                                                inner_data[0].Click();
                                                                                WebDriverWait data_wait = new WebDriverWait(driver, new TimeSpan(0, 0, 10));
                                                                                data_wait.Until(x => x.FindElement(By.Id("ContentUsersPage_btHazara")));
                                                                                for (int i = 0; i < html_fields.Length; i++)
                                                                                {
                                                                                    string curr_field = html_fields[i];
                                                                                    if (driver.PageSource.Contains(curr_field))
                                                                                    {
                                                                                        string curr_field_value = driver.FindElement(By.Id(curr_field)).Text;
                                                                                        curr_field_value = curr_field_value.Replace("\"", "\\\"");
                                                                                        //curr_field_value = curr_field_value.Replace("\\\"", "");
                                                                                        //curr_field_value = curr_field_value.Replace("\"", "");
                                                                                        if (curr_field_value != null)
                                                                                        {
                                                                                            data_values.Add(curr_field, curr_field_value);
                                                                                            data_json += "\"" + server_fields[i] + "\":\"" + curr_field_value + "\"";

                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            data_json += "\"" + server_fields[i] + "\":\"\"";
                                                                                        }
                                                                                        if (i + 1 == html_fields.Length)
                                                                                        {
                                                                                            data_json += "\r\n";
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            data_json += ",\r\n";
                                                                                        }
                                                                                    }
                                                                                }
                                                                                data_json += "}";
                                                                                SendData(data_json);
                                                                                driver.FindElement(By.Id("ContentUsersPage_btHazara")).Click();//back to data grid table
                                                                            }
                                                                            else
                                                                            {
                                                                                //take data from current table
                                                                                int k = 8;
                                                                                //details_url[2].Text;
                                                                            }
                                                                        }

                                                                    }
                                                                    if ((odd_rows == null && even_rows == null) || (curr_odd_row >= odd_rows.Count && curr_even_row >= even_rows.Count))
                                                                    {
                                                                        got_all_data = true;
                                                                        //menuItem1
                                                                        //driver.FindElement(By.Id("menuItem1")).Click();//back to query page
                                                                    }
                                                                }
                                                            }
                                                            driver.FindElement(By.Id("menuItem1")).Click();//back to query page
                                                        }

                                                        //driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2); //wait another 2 seconds
                                                        //driver.FindElement(By.Id("ContentUsersPage_btnExportExcel")).Click();
                                                        //driver.Navigate().Back();
                                                        //###
                                                        capcha_success = true;
                                                    }
                                                }

                                            }
                                            catch (System.Exception ex)
                                            {
                                                Console.WriteLine("Downloading data failed");
                                                curr_capcha_trial++;
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine("CAPTCHA was not solved");
                                        }
                                        System.Threading.Thread.Sleep(500);
                                        curr_capcha_trial++;
                                    }
                                }
                                if (datePagingMode == DatePagingMode.Monthly)
                                {
                                    curr_month--; //loop back for last month
                                    is_first_search = false;
                                }
                                else if (datePagingMode == DatePagingMode.Daily)
                                {
                                    if (curr_day_of_month > 1)
                                    {
                                        curr_day_of_month--;
                                    }
                                    else
                                    {
                                        curr_month--;
                                        //get last day of month
                                        last_day_of_month = DateTime.DaysInMonth(curr_year, curr_month);
                                        curr_day_of_month = last_day_of_month;
                                    }
                                }
                                else if (datePagingMode == DatePagingMode.TenDays)
                                {
                                    if (curr_day_of_month > 1)
                                    {
                                        curr_day_of_month--;
                                    }
                                    else
                                    {
                                        curr_month--;
                                        //get last day of month
                                        if (curr_month > 0)
                                        {
                                            last_day_of_month = DateTime.DaysInMonth(curr_year, curr_month);
                                            curr_day_of_month = last_day_of_month;
                                        }
                                    }
                                }
                            }
                            System.Threading.Thread.Sleep(500);
                        }
                        catch (System.Exception ex)
                        {
                            Console.WriteLine("Exception in type nehes " + typesNehes[typeNehes_count]);
                        }
                        
                    }



                }
            }
            Console.WriteLine("Done!");
            Console.WriteLine("Towns counters:");
            foreach (string town in towns)
            {
                Console.WriteLine(town + ":" + towns_capcha_counter[town]);
            }
        }

        /*

        */

    }
}

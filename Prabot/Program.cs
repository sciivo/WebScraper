using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace Prabot
{
    class Program
    {
        public static class GV
        {
            //Constants
            public static string desktop = Environment.GetFolderPath(System.Environment.SpecialFolder.Desktop);
            public static string workingDir = desktop + "\\Prabot";
            public static string configDir = workingDir + "\\config";
            public static string configFile = "\\settings.xlsx";
            public static string scrapeWindow;
            public static FirefoxProfile ffProfile = new FirefoxProfile(configDir + "\\ffProfile");
            public static IWebDriver driver = new FirefoxDriver(ffProfile);
            
            //Formatting
            public static char searchCurrency = '£';
            public static List<int> dupeColumnsStart = new List<int>();
            public static List<int> dupeColumnsCount = new List<int>();
            public static int spacerWidth = 71;

            //Excel
            public static Application excelApp = new Application();
            public static Workbook wbConfig;
            public static Workbook wbOutput = excelApp.Workbooks.Add();
            public static List<string> xpathSearchLabels = new List<string>();
            public static List<string> xpathSearchLocators = new List<string>();
            public static List<string> xpathScrapeLabels = new List<string>();
            public static List<string> xpathScrapeLocators = new List<string>();

            //Output variables
            public static int noRows;
            public static int noColumns;
            public static DateTime launchTime;
            public static List<string> outputScrapedData = new List<string>();
            public static List<DateTime> scrapedTimeSearched = new List<DateTime>();
            public static int outputHeadersCount;
        }

        static void Main(string[] args)
        {
            //Get main window ID
            GV.scrapeWindow = GV.driver.CurrentWindowHandle;

            //Read input settings
            GV.excelApp.Visible = true; //TEMP: Remove when done
            GV.wbConfig = GV.excelApp.Workbooks.Open(GV.configDir + GV.configFile);

            //Input variables
            string searchString = "";
            string searchChannel = GV.wbConfig.Worksheets["Settings"].Cells(2, 3).Value;
            string searchProperty = GV.wbConfig.Worksheets["Settings"].Cells(3, 3).Value;
            string searchTown = GV.wbConfig.Worksheets["Settings"].Cells(4, 3).Value;
            string searchCounty = GV.wbConfig.Worksheets["Settings"].Cells(5, 3).Value;
            string searchCountry = GV.wbConfig.Worksheets["Settings"].Cells(6, 3).Value;
            string searchType = GV.wbConfig.Worksheets["Settings"].Cells(7, 3).Value;
            int startDateDay = (int)GV.wbConfig.Worksheets["Settings"].Cells(8, 3).Value;
            int startDateMonth = (int)GV.wbConfig.Worksheets["Settings"].Cells(9, 3).Value;
            int startDateYear = (int)GV.wbConfig.Worksheets["Settings"].Cells(10, 3).Value;
            int lengthOfStay = (int)GV.wbConfig.Worksheets["Settings"].Cells(11, 3).Value;
            int noAdults = (int)GV.wbConfig.Worksheets["Settings"].Cells(12, 3).Value;
            int noChildren = (int)GV.wbConfig.Worksheets["Settings"].Cells(13, 3).Value;
            int noRooms = (int)GV.wbConfig.Worksheets["Settings"].Cells(14, 3).Value;
            string bookingPurpose = GV.wbConfig.Worksheets["Settings"].Cells(15, 3).Value;
            int maxResults = (int)GV.wbConfig.Worksheets["Settings"].Cells(19, 3).Value;
            DateTime startDate = new DateTime(startDateYear, startDateMonth, startDateDay);
            DateTime endDate = startDate.AddDays(lengthOfStay);
            GV.launchTime = DateTime.Now;            
            List<int> childrenAges = new List<int>();
            string sortLocator = "";

            //Add children ages
            for (int i = 1; i <= noChildren; i++)
            {
                childrenAges.Add(GV.wbConfig.Worksheets["Settings"].Cells((i + 15), 3).Value);
            }

            //Output file
            string ext = ".xlsx";
            string launchTimestamp = GV.launchTime.ToString("-yyMMdd-HHmmss");
            string launchFilename = "\\Prabot-Processing-" + launchTimestamp + ext;
            string outputTimestamp = GV.launchTime.ToString("-yy.MM.dd-HH.mm");
            string outputFilename = "\\Prabot" + outputTimestamp + ext;

            //Reset formatting
            GV.wbOutput.ActiveSheet.Columns.ClearFormats();
            GV.wbOutput.ActiveSheet.Rows.ClearFormats();

            //Create List of way to sort and scrape data            
            List<string> sortTypes = new List<string>();
            int lastSortTypeRow = GV.wbConfig.Worksheets[searchChannel].Cells(1, 5).End(XlDirection.xlDown).Row();
            for (int i = 1; i < lastSortTypeRow; i++)
            {
                sortTypes.Add(GV.wbConfig.Worksheets[searchChannel].Cells((i + 1), 5).Value());
            }

            //Create List for input Xpath Locators
            int lastChannelInputRow = GV.wbConfig.Worksheets[searchChannel].Cells(1, 1).End(XlDirection.xlDown).Row();
            for (int i = 1; i < lastChannelInputRow; i++)
            {
                GV.xpathSearchLabels.Add(GV.wbConfig.Worksheets[searchChannel].Cells((i + 1), 1).Value());
                GV.xpathSearchLocators.Add(GV.wbConfig.Worksheets[searchChannel].Cells((i + 1), 2).Value());
            }
            
            //Set up workbook
            if (GV.wbOutput.Worksheets.Count > sortTypes.Count)
            {
                while (GV.wbOutput.Worksheets.Count > sortTypes.Count)
                {
                    GV.wbOutput.Worksheets[GV.wbOutput.Worksheets.Count].Delete();
                }
            }
            else if (GV.wbOutput.Worksheets.Count < sortTypes.Count)
            {
                while (GV.wbOutput.Worksheets.Count < sortTypes.Count)
                {
                    GV.wbOutput.Worksheets.Add();
                }
            }
            for (int i = 0; i < sortTypes.Count; i++)
            {
                GV.wbOutput.Worksheets[(i + 1)].Name = sortTypes[i];
            }

            //Determine Channel
            switch (searchChannel)
            {
                default:
                {
                    Console.WriteLine("Unsupported channel selected.");
                    break;
                }
                case "Booking.com":
                {
                    //Set search string
                    if (searchType == "Location")
                    {
                        searchProperty = "";
                    }

                    searchString = searchProperty + ", " + searchTown + ", " + searchCounty + ", " + searchCountry;
                    searchStart(searchString, searchChannel, noAdults, noChildren, noRooms, bookingPurpose, childrenAges, sortTypes, startDate, endDate);

                    break;
                }
                case "Laterooms":
                {
                    //Set search string
                    if (searchType == "Location")
                    {
                        searchProperty = "";
                    }

                    searchString = searchProperty; // +", " + GV.searchTown; //+ ", " + GV.searchCounty + ", " + GV.searchCountry;
                    searchStart(searchString, searchChannel, noAdults, noChildren, noRooms, bookingPurpose, childrenAges, sortTypes, startDate, endDate);
                    
                    break;
                }
                case "Expedia":
                {
                    //Set search string
                    if (searchType == "Location")
                    {
                        searchProperty = "";
                    }

                    searchString = searchProperty; // +", " + GV.searchTown; //+ ", " + GV.searchCounty + ", " + GV.searchCountry;
                    searchStart(searchString, searchChannel, noAdults, noChildren, noRooms, bookingPurpose, childrenAges, sortTypes, startDate, endDate);

                    break;
                }
            }

             
            
            //Start scraping
            for (int i = 0; i < sortTypes.Count; i++)
            {
                //Set sort locator
                for (int j = 0; j < GV.xpathSearchLabels.Count(); j++)
                {
                    if (GV.xpathSearchLabels[j] == "sortLocator")
                    {
                        sortLocator = GV.xpathSearchLocators[j].Replace("REPLACE-ME", sortTypes[i].ToLower().Trim());
                    }
                }

                //Output formatting
                if (i != 0)
                {
                    Console.WriteLine("\n\n");
                }

                Console.WriteLine(spacerGenerator(1, 1, false));
                Console.WriteLine("Scraping: " + searchChannel + "   Search Type: " + searchType + "   Sort: " + GV.driver.FindElement(By.XPath(sortLocator)).Text);
                Console.WriteLine(spacerGenerator(1, 1, false) + "\n");

                //Sort results
                GV.driver.FindElement(By.XPath(sortLocator)).Click();
                System.Threading.Thread.Sleep(2000);

                scrapeStart(searchChannel, searchType, maxResults);
                writeExcel(sortTypes[i], startDate, endDate);
            }

            //Excel variables
            GV.wbOutput.Worksheets[1].Activate();
            GV.wbOutput.SaveAs(GV.workingDir + launchFilename, XlFileFormat.xlOpenXMLWorkbook);
            
            //Save Excel
            GV.wbOutput.Save();
            string savedTime = DateTime.Now.ToLongTimeString();
            //Quit Excel
            GV.excelApp.DisplayAlerts = false;
            GV.wbConfig.Close();
            GV.excelApp.DisplayAlerts = true;
            GV.excelApp.Quit();
            //Rename workbook
            TimeSpan lastWrite = Convert.ToDateTime(DateTime.Now.ToLongTimeString()) - Convert.ToDateTime(savedTime);
            while (lastWrite.Seconds < 1)
            {
                //Console.WriteLine("Waiting to file to unlock...");
                lastWrite = Convert.ToDateTime(DateTime.Now.ToLongTimeString()) - Convert.ToDateTime(savedTime);
                System.Threading.Thread.Sleep(250);
            }
            File.Move(GV.workingDir + launchFilename, GV.workingDir + outputFilename);
            
            /*//Wait to confirm exit
            Console.WriteLine("Please press any key to exit.");
            Console.ReadKey();*/
            GV.driver.Close();
        }
        
        static void searchStart(string searchString, string searchChannel, int noAdults, int noChildren, int noRooms, string bookingPurpose, List<int> childrenAges, List<string> sortTypes, DateTime startDate, DateTime endDate)
        {
            //Declare variables
            string searchURL = "";
            int ageIteration = 0;
            SelectElement comboBoxSelection;
            IWebElement searchQuery = null;
            IWebElement searchButton = null;
            IWebElement searchNoRooms = null;
            IWebElement searchNoGuests = null;
            IWebElement searchNoAdults = null;
            IWebElement searchNoChildren = null;
            IWebElement searchChildrenAge = null;
            string searchCheckInDate_Input = null;
            IWebElement searchCheckInDate_Select = null;
            IWebElement searchCheckInDay_Select = null;
            IWebElement searchCheckInYearMonth_Select = null;
            string searchCheckOutDate_Input = null;
            IWebElement searchCheckOutDate_Select = null;
            IWebElement searchCheckOutDay_Select = null;
            IWebElement searchCheckOutYearMonth_Select = null;
            IWebElement searchBookingPurposeLeisure = null;
            IWebElement searchBookingPurposeBusiness = null;

            //Load channel site
            searchURL = GV.wbConfig.Worksheets[searchChannel].Cells(2, 2).Value;
            GV.driver.Navigate().GoToUrl(searchURL);
            
            //Initial channel specifc steps
            switch (searchChannel)
            {
                case "Booking.com":
                {
                    //Nothing
                    break;
                }
                case "Laterooms":
                {
                    //Nothing
                    break;
                }
                case "Expedia":
                {
                    //Click Hotel only tab
                    GV.driver.FindElement(By.XPath(GV.xpathSearchLocators[GV.xpathSearchLabels.IndexOf("hotelTab")])).Click();
                    break;
                }
            }
            
            

            //Define XPath locaters
            for (int i = 0; i < GV.xpathSearchLocators.Count(); i++)
            {
                switch (GV.xpathSearchLabels[i])
                {
                    default:
                        {
                            break;
                        }
                    case "SearchBox":
                        {
                            searchQuery = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "SearchButton":
                        {
                            searchButton = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "NoRooms":
                        {
                            searchNoRooms = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "NoGuests":
                        {
                            searchNoGuests = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "NoAdults":
                        {
                            searchNoAdults = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "NoChildren":
                        {
                            searchNoChildren = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "ChildrenAge":
                        {
                            ageIteration = i;
                            //Not assigned until element is generated. Iteration count stored to access later.
                            //searchChildrenAge = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "CheckInDate_Input":
                        {
                            searchCheckInDate_Input = GV.xpathSearchLocators[i];
                            break;
                        }
                    case "CheckInDate_Select":
                        {
                            searchCheckInDate_Select = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "CheckInDay_Select":
                        {
                            searchCheckInDay_Select = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "CheckInYearMonth_Select":
                        {
                            searchCheckInYearMonth_Select = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "CheckOutDate_Input":
                        {
                            searchCheckOutDate_Input = GV.xpathSearchLocators[i];
                            break;
                        }
                    case "CheckOutDate_Select":
                        {
                            searchCheckOutDate_Select = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "CheckOutDay_Select":
                        {
                            searchCheckOutDay_Select = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "CheckOutYearMonth_Select":
                        {
                            searchCheckOutYearMonth_Select = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "BookingPurposeLeisure":
                        {
                            searchBookingPurposeLeisure = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                    case "BookingPurposeBusiness":
                        {
                            searchBookingPurposeBusiness = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[i])));
                            break;
                        }
                }
            }
            
            //Create List for scrape Xpath Locators and Headers
            GV.outputScrapedData.Add("Channel");
            GV.outputScrapedData.Add("Search");
            int lastChannelScrapeRow = GV.wbConfig.Worksheets[searchChannel].Cells(1, 3).End(XlDirection.xlDown).Row();
            for (int i = 1; i < lastChannelScrapeRow; i++)
            {
                GV.outputScrapedData.Add(formatHeaders(GV.wbConfig.Worksheets[searchChannel].Cells((i + 1), 3).Value()));
                GV.xpathScrapeLabels.Add(GV.wbConfig.Worksheets[searchChannel].Cells((i + 1), 3).Value());
                GV.xpathScrapeLocators.Add(GV.wbConfig.Worksheets[searchChannel].Cells((i + 1), 4).Value());
                GV.outputHeadersCount = GV.outputScrapedData.Count();
            }

            //Search            
            //Name and location
            searchQuery.SendKeys(searchString);
            System.Threading.Thread.Sleep(1000);
            searchQuery.SendKeys(Keys.Down);
            System.Threading.Thread.Sleep(1000);
            searchQuery.SendKeys(Keys.Tab);

            //Dates
                //Check in
                if (searchCheckInDate_Input != null)
                {
                    GV.driver.FindElement(By.XPath(searchCheckInDate_Input)).SendKeys(startDate.ToShortDateString());
                }
                if (searchCheckInDate_Select != null)
                {
                    comboBoxSelection = new SelectElement(searchCheckInDate_Select);
                    comboBoxSelection.SelectByValue(startDate.Year.ToString() + prependZero(startDate.Month, 2) + startDate.Day.ToString());
                }
                if (searchCheckInDay_Select != null)
                {
                    comboBoxSelection = new SelectElement(searchCheckInDay_Select);
                    comboBoxSelection.SelectByValue(startDate.Day.ToString());
                }
                if (searchCheckInYearMonth_Select != null)
                {
                    comboBoxSelection = new SelectElement(searchCheckInYearMonth_Select);
                    comboBoxSelection.SelectByValue(startDate.Year.ToString() + "-" + startDate.Month.ToString());
                }
                //Check Out
                if (searchCheckOutDate_Input != null)
                {
                    GV.driver.FindElement(By.XPath(searchCheckOutDate_Input)).Clear();
                    //System.Threading.Thread.Sleep(10000);
                    GV.driver.FindElement(By.XPath(searchCheckOutDate_Input)).SendKeys(endDate.ToShortDateString());
                }
                if (searchCheckOutDate_Select != null)
                {
                    comboBoxSelection = new SelectElement(searchCheckOutDate_Select);
                    comboBoxSelection.SelectByValue(endDate.Year.ToString() + endDate.Month.ToString() + endDate.Day.ToString());
                }
                if (searchCheckOutDay_Select != null)
                {
                    comboBoxSelection = new SelectElement(searchCheckOutDay_Select);
                    comboBoxSelection.SelectByValue(endDate.Day.ToString());
                }
                if (searchCheckOutYearMonth_Select != null)
                {
                    comboBoxSelection = new SelectElement(searchCheckOutYearMonth_Select);
                    comboBoxSelection.SelectByValue(endDate.Year.ToString() + "-" + endDate.Month.ToString());
                }

            //Booking purpose
            if (bookingPurpose == "Leisure" && (searchBookingPurposeLeisure != null || searchBookingPurposeLeisure != null))
            {
                searchBookingPurposeLeisure.Click();
            }
            else if (bookingPurpose == "Business" && (searchBookingPurposeLeisure != null || searchBookingPurposeLeisure != null))
            {
                searchBookingPurposeBusiness.Click();
            }

            //TODO: Check if Laterooms has different adults/children amount implemented. AND EXPEDIA

            //Guests - Number of Adults and Children
            if (searchChannel == "Booking.com" && searchNoGuests != null)
            {
                if (noAdults <= 2 && noChildren == 0 && noRooms == 0)
                {
                    comboBoxSelection = new SelectElement(searchNoGuests);
                    comboBoxSelection.SelectByValue(noAdults.ToString());
                }
                else
                {
                    comboBoxSelection = new SelectElement(searchNoGuests);
                    comboBoxSelection.SelectByValue("3");
                    comboBoxSelection = new SelectElement(searchNoRooms);
                    comboBoxSelection.SelectByValue(noRooms.ToString());
                    comboBoxSelection = new SelectElement(searchNoAdults);
                    comboBoxSelection.SelectByValue(noAdults.ToString());
                    comboBoxSelection = new SelectElement(searchNoChildren);
                    comboBoxSelection.SelectByValue(noChildren.ToString());

                    if (noChildren > 0)
                    {
                        for (int i = 0; i < noChildren; i++)
                        {
                            searchChildrenAge = GV.driver.FindElement(By.XPath((GV.xpathSearchLocators[ageIteration])));
                            comboBoxSelection = new SelectElement(searchChildrenAge);
                            comboBoxSelection.SelectByValue(childrenAges[i].ToString());
                        }
                    }
                }
            }
            
            //Go!
            searchButton.Click();
            System.Threading.Thread.Sleep(2000);
        }

        static void scrapeStart(string searchChannel, string searchType, int maxResults)
        {            
            IList<IWebElement> propertyRows;
            string propertyRowLocator = "";
            int noOfProperties = 0;
            //int propertyCount = 0;
            OpenQA.Selenium.Interactions.Actions actions = new OpenQA.Selenium.Interactions.Actions(GV.driver);

            for (int i = 0; i < GV.xpathSearchLabels.Count(); i++)
            {
                if (GV.xpathSearchLabels[i] == "propertyLocator")
                {
                    propertyRowLocator = GV.xpathSearchLocators[i];
                }
            }

            propertyRows = GV.driver.FindElements(By.XPath(propertyRowLocator));
            noOfProperties = propertyRows.Count();

            for (int i = 0; i < GV.outputScrapedData.Count(); i++)
            {
                if (GV.outputScrapedData.IndexOf(GV.outputScrapedData[i]) != GV.outputScrapedData.LastIndexOf(GV.outputScrapedData[i]))
                {
                    //Only add headers once
                    if (i > 0 && GV.outputScrapedData[i-1] != GV.outputScrapedData[i])
                    {
                        GV.dupeColumnsStart.Add(i + 1);
                        GV.dupeColumnsCount.Add(GV.outputScrapedData.LastIndexOf(GV.outputScrapedData[i]) - GV.outputScrapedData.IndexOf(GV.outputScrapedData[i]));
                    }
                }
            }

            
            if (propertyRows.Count() < maxResults)
            {
                maxResults = propertyRows.Count();
            }

            for (int i = 0; i < maxResults; i++)
            {

                //Scroll page
                actions.MoveToElement(propertyRows[i]);
                actions.Perform();
                System.Threading.Thread.Sleep(250); //0.25 seconds
                                
                //Output formatting
                Console.WriteLine("");
                Console.WriteLine(spacerGenerator((i + 1), maxResults, true));
                
                //Add channel and search type to scraped data
                GV.outputScrapedData.Add(searchChannel);
                GV.outputScrapedData.Add(searchType);

                //Scrape
                for (int j = 0; j < GV.xpathScrapeLocators.Count(); j++)
                {
                    //The '.' is prepended to xpathScrapeLocator so it only searched child elements
                    GV.outputScrapedData.Add(findScrapeFormat(propertyRows[i], "." + GV.xpathScrapeLocators[j], GV.xpathScrapeLabels[j], searchChannel));
                    Console.WriteLine(GV.outputScrapedData[GV.outputScrapedData.Count() - 1]);
                }

                //Count rows and columns
                GV.noRows = i + 2; //+2 for headers and counter 0 start
                GV.noColumns = GV.xpathScrapeLocators.Count() + 2; //+2 for the channel abd searcg type columns

                //Log Current time and date
                GV.scrapedTimeSearched.Add(DateTime.Now);

                //Output formatting
                Console.WriteLine(spacerGenerator(maxResults, noOfProperties, false));              
            }
        }

        private static string findScrapeFormat(IWebElement areaToSearch, string xpathLocator, string scrapeLabel, string searchChannel)
        {
            string returnString;
            
            returnString = "";

            if (elementExists(areaToSearch, xpathLocator) == true)
            {
                returnString = areaToSearch.FindElement(By.XPath(xpathLocator)).Text;
            }
            else
            {
                returnString = "MISSING";
            }

            //Add total to ratings
            if (scrapeLabel.Contains("Rating"))
            {
                if (searchChannel == "Booking.com")
                {
                    returnString = returnString + "/10";
                }
                else if (searchChannel == "Laterooms")
                {
                    returnString = GV.driver.FindElements(By.XPath(xpathLocator)).Count() + ""; //TODO: Fix this
                }
            }

            //Remove unwwated line breaks
            if (scrapeLabel.Contains("Room") && returnString.LastIndexOf("\r\n") > 0)
            {
                returnString = returnString.Substring(returnString.LastIndexOf("\r\n") + 2);
            }

            //Remove everything before currency symbol
            if (scrapeLabel.Contains("Price"))
            {
                if (returnString.Contains("\r\n"))
                {
                    returnString = returnString.Substring(returnString.IndexOf("\r\n") + 2);
                }
                if (returnString.IndexOf(GV.searchCurrency) > 0)
                {
                    returnString = returnString.Substring(returnString.IndexOf(GV.searchCurrency) + 1);
                }
                if (returnString.IndexOf(GV.searchCurrency) == 0)
                {
                    returnString = returnString.Substring(returnString.IndexOf(GV.searchCurrency) + 1);
                }
            }

            //Remove extra text from stars field
            if (scrapeLabel.Contains("Stars"))
            {
                if (returnString.IndexOf("-star hotel") > 0 || returnString.IndexOf(" stars") > 0 || returnString.IndexOf("\r\nstars") > 0)
                {
                    returnString = returnString.Substring(0, 1);
                }
                else if (returnString.IndexOf("out of 5") > 0)
                {
                    returnString = returnString.Substring(0, 3);
                }
                else
                {
                    returnString = "0";
                }
            }

            //Determine preferred partner status
            if (scrapeLabel.Contains("PreferredPartner"))
            {
                if (returnString.Contains("Preferred Property"))
                {
                    returnString = "Yes";
                }
                else
                {
                    returnString = "No";
                }
            }
            return returnString.Trim();
        }

        static void writeExcel(string sheetName, DateTime startDate, DateTime endDate)
        {
            int outputCount = 0;
            int priceCol = 0;

            //Replace nulls with string for sorting
            for (int i = 0; i < GV.outputScrapedData.Count(); i++)
            {
                if (GV.outputScrapedData[i] == null || GV.outputScrapedData[i] == "")
                {
                    GV.outputScrapedData[i] = "MISSING";
                }
            }

            //Write scraped data to sheet
            GV.wbOutput.Worksheets[sheetName].Activate();
            for (int i = 0; i < GV.noRows; i++)
            {
                for (int j = 0; j < GV.noColumns; j++)
                {
                    GV.wbOutput.Worksheets[sheetName].Cells(i + 1, j + 1).Value = GV.outputScrapedData[outputCount];
                    outputCount = (outputCount + 1);
                }
            }
            
            //Write times to sheet
            List<string> headerTimings = new List<string>();
            headerTimings.Add("Check In");
            headerTimings.Add("Check Out");
            headerTimings.Add("Launched");
            headerTimings.Add("Scraped");
            int totalNoColumns = GV.noColumns + headerTimings.Count();

            //Add timing headers
            for (int i = 1; i <= headerTimings.Count(); i++) //Columns (Timings only)
            {
                GV.wbOutput.Worksheets[sheetName].Cells(1, GV.noColumns + i).Value = headerTimings[i - 1];
            }

            for (int i = 2; i < (GV.noRows + 1); i++) //Rows
            {
                //Add timings
                GV.wbOutput.Worksheets[sheetName].Cells(i, (GV.noColumns + 1)).Value = startDate;
                GV.wbOutput.Worksheets[sheetName].Cells(i, (GV.noColumns + 2)).Value = endDate;
                GV.wbOutput.Worksheets[sheetName].Cells(i, (GV.noColumns + 3)).Value = GV.launchTime;
                GV.wbOutput.Worksheets[sheetName].Cells(i, (GV.noColumns + 4)).Value = GV.scrapedTimeSearched[i - 2];

                //Dedupe missing columns
                for (int j = 1; j <= GV.noColumns; j++) //Each column
                {
                    if (GV.dupeColumnsStart.IndexOf(j) != -1) //Start col of group of dupe cols
                    {
                        for (int k = 1; k <= GV.dupeColumnsCount[GV.dupeColumnsStart.IndexOf(j)]; k++) //Iterates through each group of dupe cols
                        {
                            if (GV.wbOutput.Worksheets[sheetName].Cells(i, (j + k)).Value.ToString() != "MISSING")
                            {
                                GV.wbOutput.Worksheets[sheetName].Cells(i, j).Value = GV.wbOutput.Worksheets[sheetName].Cells(i, (j + k)).Value;
                                GV.wbOutput.Worksheets[sheetName].Cells(i, (j + k)).Value = "MISSING";
                            }
                        }
                    }
                }

                //Delete uneeded columns
                //Similar to above, but loop runs backwards
                if (i == GV.noRows)
                {
                    for (int j = GV.noColumns; j > 1 ; j--) 
                    {
                        if (GV.dupeColumnsStart.IndexOf(j) != -1) 
                        {
                            for (int k = GV.dupeColumnsCount[GV.dupeColumnsStart.IndexOf(j)]; k >= 1; k--)
                            {
                                GV.wbOutput.Worksheets[sheetName].Cells(1, j + k).EntireColumn.Delete();
                                GV.noColumns--;
                            }
                        }
                    }
                }
            }

            //Reset global variables
            GV.scrapedTimeSearched.Clear();

            //Determine price column
            for (int i = 1; i <= GV.noColumns; i++)
            {
                if (GV.wbOutput.Worksheets[sheetName].Cells(1, i).value.Contains("Price"))
                {
                    priceCol = i;
                }
            }

            //Number formatting
            GV.wbOutput.Worksheets[sheetName].Range(GV.wbOutput.ActiveSheet.Cells(1, 1), GV.wbOutput.ActiveSheet.Cells(1, totalNoColumns)).Font.Bold = true;
            GV.wbOutput.Worksheets[sheetName].Range(GV.wbOutput.ActiveSheet.Cells(2, priceCol), GV.wbOutput.ActiveSheet.Cells(GV.noRows, priceCol)).NumberFormat = "_-" + GV.searchCurrency + "* #,##0.00_-;-" + GV.searchCurrency + "* #,##0.00_-;_-" + GV.searchCurrency + "* \"-\"??_-;_-@_-";
            GV.wbOutput.Worksheets[sheetName].Range(GV.wbOutput.ActiveSheet.Cells(2, GV.noColumns + 3), GV.wbOutput.ActiveSheet.Cells(GV.noRows, GV.noColumns + 3)).NumberFormat = "dd/mm/yyyy hh:mm:ss";
            GV.wbOutput.Worksheets[sheetName].Range(GV.wbOutput.ActiveSheet.Cells(2, GV.noColumns + 4), GV.wbOutput.ActiveSheet.Cells(GV.noRows, GV.noColumns + 4)).NumberFormat = "hh:mm:ss";
            GV.wbOutput.Worksheets[sheetName].Columns.AutoFit();
            
            //Remove all scraped data except headers
            GV.outputScrapedData.RemoveRange(GV.outputHeadersCount, (GV.outputScrapedData.Count() - GV.outputHeadersCount));
        }

        private static Boolean elementExists(IWebElement areaToSearch, String xpath)
        {
            try
            {
                areaToSearch.FindElement(By.XPath(xpath));
            }
            catch (NoSuchElementException e)
            {
                return false;
            }
            return true;
        }

        private static string formatHeaders(string header)
        {
            switch (header)
            {
                default:
                    {
                        break;
                    }
                case "scrapePropertyName":
                    {
                        header = "Property Name";
                        break;
                    }
                case "scrapePropertyRating":
                    {
                        header = "Property Rating";
                        break;
                    }
                case "scrapeRoomSingle":
                    {
                        header = "Room Type";
                        break;
                    }
                case "scrapeRoomMultiple":
                    {
                        header = "Room Type";
                        break;
                    }
                case "scrapePriceStandard":
                    {
                        header = "Room Price";
                        break;
                    }
                case "scrapePriceDiscount":
                    {
                        header = "Room Price";
                        break;
                    }
                case "scrapePriceMultiple":
                    {
                        header = "Room Price";
                        break;
                    }
                case "scrapePropertyArea":
                    {
                        header = "Area";
                        break;
                    }
                case "scrapePropertyStars":
                    {
                        header = "Stars";
                        break;
                    }
                case "scrapePreferredPartner":
                    {
                        header = "Preferred Partner";
                        break;
                    }
            }
            return header;
        }

        private static string reverseString(string textString)
        {
            if (textString == null) return null;
            char[] array = textString.ToCharArray();
            Array.Reverse(array);
            return new String(array);
        }

        private static string spacerGenerator(int number, int totalNumber, bool openingSpacer)
        {
            string spacerString;
            string spacerPad = "";
            char spacerChar = '=';
            int openingSpacerLength;
            int spacerPadding;
            int prependCount = (int)totalNumber.ToString().Length;
            
            if (GV.spacerWidth % 2 == 0)
            {
                GV.spacerWidth = (GV.spacerWidth + 1);
            }

            spacerPadding = ((GV.spacerWidth - 1) / 2);

            for (int i = (0 + prependCount); i < spacerPadding; i++)
            {
                spacerPad = spacerPad + spacerChar;
            }

            spacerString = spacerPad + prependZero(number, prependCount) + "/" + totalNumber + spacerPad;
            openingSpacerLength = spacerString.Length;

            if (openingSpacer != true)
            {
                spacerString = "";

                for (int i = 0; i < openingSpacerLength; i++)
                {
                    spacerString = spacerString + spacerChar;
                }                
            }

            return spacerString;
        }

        private static string prependZero(int number, int noDigits)
        {
            string numberString = number.ToString();

            if (number.ToString().Length < noDigits)
            {
                for (int i = 0; i < noDigits - number.ToString().Length; i++)
                {
                    numberString = "0" + numberString;
                }                
            }
            return numberString;
        }
    }
}
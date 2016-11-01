using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using System.Net;
using HtmlAgilityPack;
using System.Xml;
using System.Xml.Linq;
using System.Configuration;
using System.Threading;
using System.Collections;


namespace WinScrapping
{
    public partial class Form1 : Form
    {
        private String msgText = String.Empty;

        public string CatfilePath
        {
            get { return ConfigurationManager.AppSettings["filePath"].ToString() + "CatData_" + GetBrandNameFromURL(cbxCategory.SelectedItem.ToString()) + ".xlsx"; }
        }
        public string ImageFolderPath
        {
            get
            {
                if (!Directory.Exists(ConfigurationManager.AppSettings["filePath"].ToString() + "Image Folder_" + GetBrandNameFromURL(cbxCategory.SelectedItem.ToString())))
                {
                    Directory.CreateDirectory(ConfigurationManager.AppSettings["filePath"].ToString() + "Image Folder_" + GetBrandNameFromURL(cbxCategory.SelectedItem.ToString()));
                }
                return ConfigurationManager.AppSettings["filePath"].ToString() + "Image Folder_" + GetBrandNameFromURL(cbxCategory.SelectedItem.ToString());
            }
        }
        public string ProdLinkfilePath
        {
            get { return ConfigurationManager.AppSettings["filePath"].ToString() + "ProdLinkData_" + GetBrandNameFromURL(cbxCategory.SelectedItem.ToString()) + ".xlsx"; }
        }
        public string ProdfilePath
        {
            get { return ConfigurationManager.AppSettings["filePath"].ToString() + "ProdData_" + GetBrandNameFromURL(cbxCategory.SelectedItem.ToString()) + ".xlsx"; }
        }
        //public string RelatedProdfilePath
        //{
        //    get { return cbxCategory.SelectedText; }
        //}
        public string htmlPath
        {
            get { return Convert.ToString(cbxCategory.SelectedItem); }
        }
        // string website;

        public string website
        {
            get { return GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com" ? ConfigurationManager.AppSettings["website"].ToString() : ConfigurationManager.AppSettings["website1"].ToString(); }

        }


        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            IEnumerable<string> hrefList;

            List<string> prodhrefList = new List<string>();

            #region Category

            hrefList = GetCategoryFromFile(CatfilePath);
            if (hrefList.Count() == 0)
            {
                if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                {
                    hrefList = GetCategoryFromWeb(htmlPath);
                    SaveCategoryToExcel(CatfilePath, hrefList);
                }
                else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                {

                    hrefList = GetCategoryFromWebNew(htmlPath);
                    SaveCategoryToExcelNew(CatfilePath, hrefList);
                }
            }
            #endregion

            #region Product
            int prodcount = 0;
            int pagecount = 48;
            int pageindex = 1;
            //website = "http://www.lightingdirect.com";

            int rowcount = GetProductTotal(ProdLinkfilePath);

            foreach (string cathref in hrefList)
            {
                prodhrefList.Clear();

                if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                {
                    hrefList = GetProductFromFile(ProdLinkfilePath, cathref.Split(',')[0]);
                    pagecount = 30;
                    prodcount = Convert.ToInt32(cathref.Split(',')[1]);
                }
                else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                {

                    hrefList = GetProductFromFile(ProdLinkfilePath, cathref);
                    prodcount = Convert.ToInt32(GetProductCount(website + cathref.Split(',')[0].Replace(website, "")));
                }


                pageindex = 1;
                if (hrefList.Count() == 0)
                {
                    //<a href="/brand/elk-lighting.html?pageIndex=2&amp;cat=6" id="pr_show_more" url="/brand/elk-lighting.html" np="5" tp="6" type="brand" style="display: block;">
                    //    SHOW MORE RESULTS</a>



                    while (prodcount > 0)
                    {
                        if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                        {
                            prodhrefList.AddRange(GetProductFromWeb(website + cathref.Split(',')[0].Replace(website, "") + "&pageIndex=" + pageindex));
                        }
                        else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                        {
                            prodhrefList.AddRange(GetProductFromWebNew(website + cathref.Split(',')[0].Replace(website, "") + "?p=" + pageindex));
                        }


                        prodcount = prodcount - pagecount;
                        pageindex++;
                    }

                    SaveProductToExcel(ProdLinkfilePath, prodhrefList, cathref.Split(',')[0], ref rowcount);
                }
            }

            #endregion

            #region "Product Data"
            string url;

            //website = "http://www.lightingdirect.com";


            hrefList = GetProductFromFile(ProdLinkfilePath);
            int prodrowno = GetProductTotal(ProdfilePath);
            if (prodrowno == 0)
            {
                prodrowno++;
                if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                {
                    addHeader(prodrowno, ProdfilePath);
                }
                else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                {
                    addHeaderNew(prodrowno, ProdfilePath);
                }

            }

            if (hrefList.Count() > 0)
            {
                foreach (string cathref in hrefList)
                {
                    prodrowno++;
                    url = website + cathref.Split(',')[0];
                    try
                    {
                        if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                        {
                            GetProductDetailsFromWeb(url, ProdfilePath, prodrowno, cathref.Split(',')[1]);
                        }
                        else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                        {
                            GetProductDetailsFromWebNew(url, ProdfilePath, prodrowno, cathref.Split(',')[1]);
                        }

                    }
                    catch (Exception ex)
                    {
                        using (StreamWriter tw = new StreamWriter("Error.txt", true))
                        {
                            tw.WriteLine("=======================================" + DateTime.Now.ToString() + "=============================================");
                            tw.WriteLine(url);
                            tw.WriteLine();
                            tw.WriteLine(ex.StackTrace);
                            tw.WriteLine("=======================================================================================================");
                            tw.WriteLine();
                            tw.WriteLine();
                        }

                    }
                }
            }
            #endregion

        }


        #region "GetPage"
        /// <summary>
        /// A method to return html of  the dansukker.dk page.
        /// </summary>
        /// <param name="requestURL">url of the requested page</param>
        /// <param name="msgText">Error message text, if any.</param>
        /// <returns>Html of the dansukker.dk page.</returns>
        public string GetPage(string href, out string msgText)
        {
            string strHtml;
            msgText = string.Empty;
            try
            {
                strHtml = GetPageRequest(href, out msgText);

            }
            catch (Exception ex)
            {
                msgText = ex.Message;
                return null;
            }

            return strHtml;
        }

        public string PostPage(string href, out string msgText)
        {
            string strHtml;
            msgText = string.Empty;
            try
            {
                strHtml = PostPageRequest(href, out msgText);

            }
            catch (Exception ex)
            {
                msgText = ex.Message;
                return null;
            }

            return strHtml;
        }
        #endregion

        #region "GetPageRequest"
        /// <summary>
        /// A method to send a request to the dansukker.dk page.
        /// </summary>
        /// <param name="requestURL">Request URL.</param>
        /// <param name="msgText">Error message text, if any.</param>
        /// <returns>The output of the request to the dansukker.dk page.</returns>
        private string GetPageRequest(string requestURL, out string msgText)
        {
            string pageHtml = null;
            msgText = string.Empty;

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(requestURL);

                request.Method = "GET";
                // request.ContentType = "application/x-www-form-urlencoded";
                request.AllowAutoRedirect = true;
                //request.Timeout = 4000;
                request.CookieContainer = new CookieContainer();

                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                using (StreamReader sr = new StreamReader(response.GetResponseStream()))
                    pageHtml = sr.ReadToEnd();
            }
            catch (TimeoutException ex)
            {
                msgText = ex.Message;
                throw ex;
            }
            catch (Exception ex)
            {
                msgText = ex.Message;
                throw ex;
            }

            return pageHtml;
        }

        private string PostPageRequest(string requestURL, out string msgText)
        {
            string pageHtml = null;
            msgText = string.Empty;

            try
            {
                string url = requestURL;
                string quary = "";

                int index = requestURL.IndexOf('?');
                if (index > 0)
                {
                    url = requestURL.Substring(0, index);
                    quary = requestURL.Substring(index + 1, requestURL.Length - index - 1);
                }
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);

                request.Method = "Post";
                byte[] byteArray = Encoding.UTF8.GetBytes(quary);
                // Set the ContentType property of the WebRequest.
                request.ContentType = "application/x-www-form-urlencoded";
                // Set the ContentLength property of the WebRequest.
                request.ContentLength = byteArray.Length;

                Stream dataStream = request.GetRequestStream();
                // Write the data to the request stream.
                dataStream.Write(byteArray, 0, byteArray.Length);
                // Close the Stream object.
                dataStream.Close();

                WebResponse response = request.GetResponse();
                // Display the status.
                Console.WriteLine(((HttpWebResponse)response).StatusDescription);
                // Get the stream containing content returned by the server.
                dataStream = response.GetResponseStream();
                // Open the stream using a StreamReader for easy access.
                StreamReader reader = new StreamReader(dataStream);
                // Read the content.
                pageHtml = reader.ReadToEnd();
                // Display the content.
                //Console.WriteLine(responseFromServer);
                // Clean up the streams.
                reader.Close();
                dataStream.Close();
                response.Close();


            }
            catch (TimeoutException ex)
            {
                msgText = ex.Message;
                throw ex;
            }
            catch (Exception ex)
            {
                msgText = ex.Message;
                throw ex;
            }

            return pageHtml;
        }
        #endregion

        //private IEnumerable<string> SaveCategory(string filePath, string htmlPath)
        //{

        //    //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";
        //    FileInfo newfile = new FileInfo(filePath);

        //    string html = "";
        //    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
        //    html = GetPage(htmlPath, out msgText);
        //    if (html.Length > 0 || html != null)
        //    {
        //        doc.LoadHtml(html);
        //        /* For title this is working very fine */
        //        var res = doc.DocumentNode.SelectSingleNode("//div[@class='brand_land_box']");

        //        doc.LoadHtml(res.InnerHtml.ToString());

        //        //brand_land_box
        //        using (ExcelPackage xlPackage = new ExcelPackage(newfile))
        //        {
        //            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Category"];

        //            if (worksheet == null)
        //                worksheet = xlPackage.Workbook.Worksheets.Add("Category");

        //            /* set column in excel */
        //            int i = 1;
        //            foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a"))
        //            {
        //                /* project title geted here */
        //                worksheet.Cell(i, 1).Value = link.Attributes["href"].Value;
        //                i++;
        //            }
        //            xlPackage.Save();
        //        }
        //    }
        //}

        #region Category href

        private IEnumerable<string> GetCategoryFromFile(string filePath)
        {
            List<string> hrefList = new List<string>();
            FileInfo newfile = new FileInfo(filePath);


            //brand_land_box
            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Category"];

                if (worksheet != null)
                {
                    /* set column in excel */
                    int i = 1;
                    while (true)
                    {
                        if (worksheet.Cell(i, 1).Value == "")
                            break;

                        hrefList.Add(worksheet.Cell(i, 1).Value + "," + worksheet.Cell(i, 2).Value);
                        i++;
                    }
                }
            }

            return hrefList.Distinct();
        }


        private IEnumerable<string> GetCategoryFromWeb(string htmlPath)
        {
            List<string> hrefList = new List<string>();
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";

            string txt = "";
            string html = "";
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            html = GetPage(htmlPath, out msgText);
            if (html.Length > 0 || html != null)
            {
                doc.LoadHtml(html);
                /* For title this is working very fine */
                var res = doc.DocumentNode.SelectSingleNode("//div[@class='brand_land_box']");

                doc.LoadHtml(res.InnerHtml.ToString());

                //brand_land_box
                /* set column in excel */
                //<span class="num_of_prods">161 products</span>

                //<div class="brand_cat">
                //    <a href="http://www.lightingnewyork.com/brand/elk-lighting.html?cat=61">
                //    <div class="brand_cat_img">
                //    <img src="./Category_files/1092.jpg" width="150" border="0">
                //    </div>
                //    </a>

                //        <a href="http://www.lightingnewyork.com/brand/elk-lighting.html?cat=61">ELK Light Bulbs</a>
                //        <br>
                //        <span class="num_of_prods">1 products</span>
                //        <br clear="all">
                //</div>
                foreach (var divcat in doc.DocumentNode.SelectNodes("//div[@class='brand_cat']"))
                {
                    HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                    doc2.LoadHtml(divcat.InnerHtml.ToString());
                    txt = doc2.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value;

                    txt += "," + doc2.DocumentNode.SelectSingleNode("//span[@class='num_of_prods']").InnerHtml.ToString().Replace(" products", "");

                    hrefList.Add(txt);


                    //foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a"))
                    //{
                    //    /* project title geted here */
                    //    hrefList.Add(link.Attributes["href"].Value);
                    //}
                }

            }
            return hrefList.Distinct();
        }

        private IEnumerable<string> GetCategoryFromWebNew(string htmlPath)
        {
            List<string> hrefList = new List<string>();
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";

            string txt = "";
            string html = "";
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            html = GetPage(htmlPath, out msgText);
            if (html.Length > 0 || html != null)
            {
                doc.LoadHtml(html);
                /* For title this is working very fine */

                var res = doc.DocumentNode.SelectSingleNode("//table[@id='liTable']");

                doc.LoadHtml(res.InnerHtml.ToString());




                foreach (var divcat in doc.DocumentNode.SelectNodes("//td"))
                {
                    HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                    doc2.LoadHtml(divcat.InnerHtml.ToString());
                    txt = doc2.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value;



                    if (doc2.DocumentNode.SelectSingleNode("//span").InnerText.ToLower() != "shop all")
                        hrefList.Add(txt);




                }

            }
            return hrefList.Distinct();
        }


        private string GetProductCount(string htmlPath)
        {

            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";

            string txt = "0";
            string html = "";
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            html = GetPage(htmlPath, out msgText);
            if (html.Length > 0 || html != null)
            {
                doc.LoadHtml(html);
                /* For title this is working very fine */
                //var res = doc.DocumentNode.SelectSingleNode("//div[@class='brand_land_box']");

                if (doc.DocumentNode.SelectSingleNode("//div[@id='numberOfItemsDisplayed']") != null)
                {
                    var res = doc.DocumentNode.SelectSingleNode("//div[@id='numberOfItemsDisplayed']");

                    doc.LoadHtml(res.InnerHtml.ToString());






                    // txt = doc2.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value;

                    txt = doc.DocumentNode.SelectSingleNode("//span[@id='search_results_count']").InnerHtml.ToString().Replace(" products", "").Replace("We have ", "");




                    //foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a"))
                    //{
                    //    /* project title geted here */
                    //    hrefList.Add(link.Attributes["href"].Value);
                    //}

                }
            }
            return txt;
        }

        private void SaveCategoryToExcel(string filePath, IEnumerable<string> hrefList)
        {
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";
            FileInfo newfile = new FileInfo(filePath);



            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Category"];

                if (worksheet == null)
                    worksheet = xlPackage.Workbook.Worksheets.Add("Category");

                /* set column in excel */
                int i = 1;
                foreach (string link in hrefList)
                {
                    /* project title geted here */
                    string[] links = link.Split(',');

                    worksheet.Cell(i, 1).Value = links[0];
                    worksheet.Cell(i, 2).Value = links[1];
                    i++;
                }
                xlPackage.Save();
            }

        }

        private void SaveCategoryToExcelNew(string filePath, IEnumerable<string> hrefList)
        {
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";
            FileInfo newfile = new FileInfo(filePath);



            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Category"];

                if (worksheet == null)
                    worksheet = xlPackage.Workbook.Worksheets.Add("Category");

                /* set column in excel */
                int i = 1;
                foreach (string link in hrefList)
                {
                    /* project title geted here */
                    //string[] links = link.Split(',');

                    //worksheet.Cell(i, 1).Value = links[0];
                    //worksheet.Cell(i, 2).Value = links[1];

                    worksheet.Cell(i, 1).Value = link;
                    i++;
                }
                xlPackage.Save();
            }

        }

        #endregion

        #region Product href

        private IEnumerable<string> GetProductFromFile(string filePath, string cathref)
        {
            List<string> hrefList = new List<string>();
            FileInfo newfile = new FileInfo(filePath);


            //brand_land_box
            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet != null)
                {
                    /* set column in excel */
                    int i = 1;
                    while (true)
                    {
                        if (worksheet.Cell(i, 1).Value == "")
                            break;
                        if (cathref == worksheet.Cell(i, 2).Value)
                            hrefList.Add(worksheet.Cell(i, 1).Value);
                        i++;
                    }
                }
            }

            return hrefList.Distinct();
        }

        private IEnumerable<string> GetProductFromFile(string filePath)
        {
            List<string> hrefList = new List<string>();
            FileInfo newfile = new FileInfo(filePath);


            //brand_land_box
            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet != null)
                {
                    /* set column in excel */
                    int i = 1;
                    while (true)
                    {
                        if (worksheet.Cell(i, 1).Value == "")
                            break;

                        hrefList.Add(worksheet.Cell(i, 1).Value + "," + worksheet.Cell(i, 2).Value);
                        i++;
                    }
                }
            }

            return hrefList.Distinct();
        }


        private int GetProductTotal(string filePath)
        {
            int rowtotal = 0;
            FileInfo newfile = new FileInfo(filePath);


            //brand_land_box
            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet != null)
                {
                    /* set column in excel */

                    while (true)
                    {
                        rowtotal++;

                        if (worksheet.Cell(rowtotal, 1).Value == "")
                        {
                            rowtotal--;
                            break;
                        }


                    }
                }
            }

            return rowtotal;
        }

        private IEnumerable<string> GetProductFromWeb(string htmlPath)
        {
            List<string> hrefList = new List<string>();
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";


            string html = "", txt;
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            html = GetPage(htmlPath, out msgText);
            if (html.Length > 0 || html != null)
            {
                doc.LoadHtml(html);
                /* For title this is working very fine */

                var res = doc.DocumentNode.SelectSingleNode("//div[@id='products_container']"); // gallery -- products_container

                doc.LoadHtml(res.InnerHtml.ToString());

                /* set column in excel */

                //foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a"))
                //{
                //    /* project title geted here */
                //    hrefList.Add(link.Attributes["href"].Value.Replace("%", "-"));
                //}
                foreach (var divcat in doc.DocumentNode.SelectNodes("//div[@class='one_product']"))
                {
                    HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                    doc2.LoadHtml(divcat.InnerHtml.ToString());
                    txt = doc2.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value.Replace("%2D", "-");


                    hrefList.Add(txt);
                }

            }
            return hrefList.Distinct();
        }
        private IEnumerable<string> GetProductFromWebNew(string htmlPath)
        {
            List<string> hrefList = new List<string>();
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";


            string html = "", txt;
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            html = GetPage(htmlPath, out msgText);
            if (html.Length > 0 || html != null)
            {
                doc.LoadHtml(html);
                /* For title this is working very fine */

                // var res = doc.DocumentNode.SelectSingleNode("//table[@id='searchResultsItems']"); // gallery -- products_container

                // doc.LoadHtml(res.InnerHtml.ToString());

                /* set column in excel */

                //foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//a"))
                //{
                //    /* project title geted here */
                //    hrefList.Add(link.Attributes["href"].Value.Replace("%", "-"));
                //}

                if (doc.DocumentNode.SelectNodes("//table[@id='searchResultsItems']") != null)
                {
                    foreach (var divcat in doc.DocumentNode.SelectNodes("//table[@id='searchResultsItems']"))
                    {
                        HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                        doc2.LoadHtml(divcat.InnerHtml.ToString());

                        foreach (var divProdLink in doc2.DocumentNode.SelectNodes("//td[@class='resultBox']"))
                        {
                            HtmlAgilityPack.HtmlDocument doc3 = new HtmlAgilityPack.HtmlDocument();
                            doc3.LoadHtml(divProdLink.InnerHtml.ToString());
                            if (doc3.DocumentNode.SelectSingleNode("//a") != null)
                            {
                                txt = doc3.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value.Replace("%2D", "-");
                                hrefList.Add(txt);
                            }
                        }




                    }
                }
                else
                {
                    foreach (var divcat in doc.DocumentNode.SelectNodes("//table[@class='listView']"))
                    {
                        HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                        doc2.LoadHtml(divcat.InnerHtml.ToString());

                        if (doc2.DocumentNode.SelectSingleNode("//a") != null)
                        {
                            txt = doc2.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value.Replace("%2D", "-");
                            hrefList.Add(txt);
                        }

                    }
                }

            }
            return hrefList.Distinct();
        }

        private void SaveProductToExcel(string filePath, List<string> hrefList, string CatPath, ref int rowcount)
        {
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";
            FileInfo newfile = new FileInfo(filePath);

            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet == null)
                    worksheet = xlPackage.Workbook.Worksheets.Add("Product");

                /* set column in excel */

                foreach (string link in hrefList)
                {
                    rowcount++;
                    /* project title geted here */
                    worksheet.Cell(rowcount, 1).Value = link;
                    worksheet.Cell(rowcount, 2).Value = CatPath;

                }
                if (hrefList.Count > 0)
                    xlPackage.Save();
            }

        }

        #endregion

        private void GetProductDetailsFromWeb(string htmlPath, string filePath, int prodrowno, string caturl)
        {

            string productTitle = "";
            string productLink = "";
            string productSku = "";
            string strinStock = "";
            string strCurrentPrice = "";
            string strOldPrice = "";
            string description = "";

            string Manufacturer = "";
            string Collection = "";
            string SKU = "";
            string UPC = "";

            string finishType = "N/A";
            string shadeType = "N/A";
            string bulbType = "N/A";

            string Category = "";
            string Finish = "";
            string Glass = "";
            string Material = "";
            //string Dimensions= "";

            string length = "";
            string width = "";
            string height = "";
            string ext = "";
            string bcWidth = "";
            string bcHeight = "";

            string Weight = "";

            string BulbTypePrimary = "";
            string NumberofBulbsPrimary = "";
            string MaxWattagePrimary = "";

            string BulbTypeSecondary = "";
            string NumberofBulbsSecondary = "";
            string MaxWattageSecondary = "";

            string BulbsIncluded = "";

            string ShipsVia = "";
            //string ShipDimensions = "";

            string Shiplength = "";
            string Shipwidth = "";
            string Shipheight = "";

            string ShipWeight = "";

            string OutdoorRating = "";
            string ULRating = "";

            string relativelist = "";

            string imagepath = "";

            string wtd = "rel", pr_id, pr_collection, pr_vendor;
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";


            string html = "", txt;
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            html = GetPage(htmlPath, out msgText);
            if (html != null)
            {
                if (html.Length > 0)
                {
                    doc.LoadHtml(html);
                    /* For title this is working very fine */
                    var res = doc.DocumentNode.SelectSingleNode("//div[@id='pr_title']");

                    Int32 pos = res.InnerHtml.IndexOf(@"<div id=""pr_title_txt"">");
                    string strRemstring = res.InnerHtml.Substring(pos);
                    int lastpos = strRemstring.IndexOf("</div>");

                    pr_id = doc.DocumentNode.SelectSingleNode("//span[@id='pr_id']").InnerHtml;
                    pr_collection = doc.DocumentNode.SelectSingleNode("//span[@id='pr_collection']").InnerHtml;
                    pr_vendor = doc.DocumentNode.SelectSingleNode("//span[@id='pr_vendor']").InnerHtml;

                    strRemstring = res.InnerHtml.Substring(pos, lastpos);

                    doc.LoadHtml(strRemstring);

                    productTitle = GetText(doc.DocumentNode.SelectNodes("//h1[@class='pr_name']").FirstOrDefault());

                    /*foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//h1[@class='pr_name']"))
                    {
                        productTitle = link.InnerText;
                    }*/

                    foreach (HtmlNode link in doc.DocumentNode.SelectNodes("//meta"))
                    {
                        switch (link.Attributes["itemprop"].Value.ToLower())
                        {
                            case "url":
                                productLink = link.Attributes["content"].Value.ToString();
                                break;
                            case "sku":
                                productSku = "Sku # " + link.Attributes["content"].Value.ToString();
                                break;
                        }
                    }
                    /* upto here */
                    doc.LoadHtml(html);

                    try
                    {
                        //res = doc.DocumentNode.SelectSingleNode("//div[@id='badges']");
                        //doc.LoadHtml(res.InnerHtml);
                        res = doc.DocumentNode.SelectSingleNode("//div[@class='badge stock']");
                        doc.LoadHtml(res.InnerHtml);
                        res = doc.DocumentNode.SelectSingleNode("//span[@style='font-size: 24px;']");
                        strinStock = res.InnerText.Trim();
                    }
                    catch { }


                    /* We are going to take price,available quantity, */
                    doc.LoadHtml(html);

                    res = doc.DocumentNode.SelectSingleNode("//div[@id='price_tbl_stf']");
                    doc.LoadHtml(res.InnerHtml);
                    HtmlNodeCollection tables = doc.DocumentNode.SelectNodes("//table");

                    //HtmlNodeCollection blubTable = doc.DocumentNode.SelectNodes("//table[@class='detail_tbl bulb_tbl']");

                    HtmlNodeCollection rows = tables[0].SelectNodes(".//tr");


                    if (rows != null)
                    {
                        for (int i = 0; i < rows.Count; ++i)
                        {
                            // Iterate all columns in this row
                            HtmlNodeCollection cols = rows[i].SelectNodes(".//td");
                            if (cols != null)
                            {
                                if (i == 0)
                                {
                                    /* current price here */
                                    strCurrentPrice = GetText(cols[0]).Replace("per each", "").Replace("$", "").Trim();
                                }
                                if (i == 2)
                                {
                                    /* old price here */
                                    strOldPrice = GetText(cols[1]).Replace("$", "").Trim();
                                }

                            }
                        }
                    }

                    doc.LoadHtml(html);
                    res = doc.DocumentNode.SelectSingleNode("//div[@id='general_info']");
                    doc.LoadHtml(res.InnerHtml);

                    description = GetText(doc.DocumentNode.SelectNodes("//p").FirstOrDefault()).ToString().Trim();
                    HtmlNodeCollection lis;

                    //Manufacturer Information
                    var resprodd = doc.DocumentNode.SelectSingleNode("//div[@id='info_prod']");
                    if (resprodd != null)
                    {
                        lis = resprodd.SelectNodes(".//li");
                        if (lis != null)
                        {

                            for (int t = 0; t < lis.Count; ++t)
                            {
                                if (lis[t].InnerText.Trim().Contains("Brand"))
                                {
                                    Manufacturer = GetTextOnly(lis[t]).Replace("Brand", "").Replace(":", "").Trim();
                                }

                                if (lis[t].InnerText.Trim().Contains("Collection"))
                                {
                                    Collection = GetTextOnly(lis[t]).Replace("Collection", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("SKU"))
                                {
                                    SKU = GetTextOnly(lis[t]).Replace("SKU", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("UPC"))
                                {
                                    UPC = GetTextOnly(lis[t]).Replace("UPC", "").Replace(":", "").Trim();
                                }
                            }
                        }
                    }
                    //Design Information
                    resprodd = doc.DocumentNode.SelectSingleNode("//div[@id='info_dsgn']");
                    if (resprodd != null)
                    {
                        lis = resprodd.SelectNodes(".//li");
                        if (lis != null)
                        {

                            for (int t = 0; t < lis.Count; ++t)
                            {
                                if (lis[t].InnerText.Trim().Contains("Category"))
                                {
                                    Category = GetTextOnly(lis[t]).Replace("Category", "").Replace(":", "").Trim();
                                }

                                if (lis[t].InnerText.Trim().Contains("Finish"))
                                {
                                    Finish = GetTextOnly(lis[t]).Replace("Finish", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Glass"))
                                {
                                    Glass = GetTextOnly(lis[t]).Replace("Glass", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Material"))
                                {
                                    Material = GetTextOnly(lis[t]).Replace("Material", "").Replace(":", "").Trim();
                                }
                            }
                        }
                    }
                    //Dimensions and Weight (inches and pounds)
                    resprodd = doc.DocumentNode.SelectSingleNode("//div[@id='info_dims']");
                    if (resprodd != null)
                    {
                        lis = resprodd.SelectNodes(".//li");
                        if (lis != null)
                        {

                            for (int t = 0; t < lis.Count; ++t)
                            {
                                if (lis[t].InnerText.Trim().Contains("Width") && !lis[t].InnerText.Trim().Contains("Backplate/Canopy"))
                                {
                                    width = GetTextOnly(lis[t]).Replace("Width", "").Replace(":", "").Trim();
                                }

                                if (lis[t].InnerText.Trim().Contains("Height") && !lis[t].InnerText.Trim().Contains("Backplate/Canopy"))
                                {
                                    height = GetTextOnly(lis[t]).Replace("Height", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Backplate/Canopy Width"))
                                {
                                    bcWidth = GetTextOnly(lis[t]).Replace("Backplate/Canopy Width", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Backplate/Canopy Height"))
                                {
                                    bcHeight = GetTextOnly(lis[t]).Replace("Backplate/Canopy Height", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Weight"))
                                {
                                    Weight = GetTextOnly(lis[t]).Replace("Weight", "").Replace(":", "").Trim();
                                }
                            }

                        }
                    }
                    //Bulb Information
                    resprodd = doc.DocumentNode.SelectSingleNode("//div[@id='info_bulb']");
                    if (resprodd != null)
                    {
                        lis = resprodd.SelectNodes(".//li");

                        if (lis != null)
                        {

                            for (int t = 0; t < lis.Count; ++t)
                            {
                                if (lis[t].InnerText.Trim().Contains("Bulbs Included"))
                                {
                                    BulbsIncluded = GetTextOnly(lis[t]).Replace("Bulbs Included", "").Replace(":", "").Trim();
                                }

                                if (lis[t].InnerText.Trim().Contains("Bulb Type"))
                                {
                                    BulbTypePrimary = GetTextOnly(lis[t]).Replace("Bulb Type", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Number of Bulbs"))
                                {
                                    NumberofBulbsPrimary = GetTextOnly(lis[t]).Replace("Number of Bulbs", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Max Wattage per Bulb"))
                                {
                                    MaxWattagePrimary = GetTextOnly(lis[t]).Replace("Max Wattage per Bulb", "").Replace(":", "").Trim();
                                }

                            }
                        }
                    }
                    //Shipping Information
                    resprodd = doc.DocumentNode.SelectSingleNode("//div[@id='info_dims']");
                    if (resprodd != null)
                    {
                        lis = resprodd.SelectNodes(".//li");
                        if (lis != null)
                        {

                            for (int t = 0; t < lis.Count; ++t)
                            {
                                if (lis[t].InnerText.Trim().Contains("Ships Via"))
                                {
                                    ShipsVia = GetTextOnly(lis[t]).Replace("Ships Via", "").Replace(":", "").Trim();
                                }

                                if (lis[t].InnerText.Trim().Contains("Ship Length"))
                                {
                                    Shiplength = GetTextOnly(lis[t]).Replace("Ship Length", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Ship Width"))
                                {
                                    Shipwidth = GetTextOnly(lis[t]).Replace("Ship Width", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Ship Height"))
                                {
                                    Shipheight = GetTextOnly(lis[t]).Replace("Ship Height", "").Replace(":", "").Trim();
                                }
                                if (lis[t].InnerText.Trim().Contains("Ship Weight"))
                                {
                                    ShipWeight = GetTextOnly(lis[t]).Replace("Ship Weight", "").Replace(":", "").Trim();
                                }
                            }
                        }
                    }
                    //Product Rating
                    resprodd = doc.DocumentNode.SelectSingleNode("//div[@id='info_rate']");
                    if (resprodd != null)
                    {
                        lis = resprodd.SelectNodes(".//li");
                        if (lis != null)
                        {

                            for (int t = 0; t < lis.Count; ++t)
                            {
                                if (lis[t].InnerText.Trim().Contains("Voltage"))
                                {
                                    OutdoorRating = GetTextOnly(lis[t]).Replace("Voltage", "").Replace(":", "").Trim();
                                }

                                if (lis[t].InnerText.Trim().Contains("UL Rating"))
                                {
                                    ULRating = GetTextOnly(lis[t]).Replace("UL Rating", "").Replace(":", "").Trim();
                                }

                            }
                        }
                    }

                    /* New Data, */
                    doc.LoadHtml(html);

                    res = doc.DocumentNode.SelectSingleNode("//div[@id='price_tbl_stf']");
                    doc.LoadHtml(res.InnerHtml);
                    tables = doc.DocumentNode.SelectNodes("//table");

                    //HtmlNodeCollection blubTable = doc.DocumentNode.SelectNodes("//table[@class='detail_tbl bulb_tbl']");
                    if (tables.Count() > 1)
                    {
                        if (tables[1] != null)
                        {
                            rows = tables[1].SelectNodes(".//tr");
                            if (rows != null)
                            {
                                for (int i = 0; i < rows.Count - 1; i++)
                                {
                                    try
                                    {
                                        if (rows[i].Id != "add_skus")
                                        {
                                            doc.LoadHtml(rows[i].InnerHtml);
                                            if (doc.DocumentNode.SelectSingleNode("//div[@type='Finish']") != null)
                                            {
                                                resprodd = doc.DocumentNode.SelectSingleNode("//div[@type='Finish']");
                                                lis = resprodd.SelectNodes("//div[@class='name']");
                                                for (int t = 0; t < lis.Count; t++)
                                                {
                                                    if (finishType == "N/A")
                                                    {
                                                        finishType = string.Empty;
                                                    }
                                                    finishType = finishType + GetText(lis[t]) + ",";
                                                }
                                            }
                                            else
                                                if (doc.DocumentNode.SelectSingleNode("//div[@type='Shade']") != null)
                                            {
                                                resprodd = doc.DocumentNode.SelectSingleNode("//div[@type='Shade']");
                                                lis = resprodd.SelectNodes("//div[@class='name']");
                                                for (int t = 0; t < lis.Count; t++)
                                                {
                                                    if (shadeType == "N/A")
                                                    {
                                                        shadeType = string.Empty;
                                                    }
                                                    shadeType = shadeType + GetText(lis[t]) + ",";
                                                }
                                            }
                                            else
                                                    if (doc.DocumentNode.SelectSingleNode("//div[@type='Bulb_Type']") != null)
                                            {
                                                resprodd = doc.DocumentNode.SelectSingleNode("//div[@type='Bulb_Type']");
                                                lis = resprodd.SelectNodes("//div[@class='name']");
                                                for (int t = 0; t < lis.Count; t++)
                                                {
                                                    if (bulbType == "N/A")
                                                    {
                                                        bulbType = string.Empty;
                                                    }
                                                    bulbType = bulbType + GetText(lis[t]) + ",";
                                                }
                                            }
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }

                    //image Path
                    doc.LoadHtml(html);
                    resprodd = doc.DocumentNode.SelectSingleNode("//meta[@property='og:image']");
                    if (resprodd != null)
                    {
                        imagepath = resprodd.Attributes["content"].Value;
                        try
                        {
                            if (imagepath.Length > 0)
                            {
                                string saveLocation = ImageFolderPath + "\\" + Path.GetFileName(imagepath);
                                byte[] imageBytes;
                                HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imagepath);
                                WebResponse imageResponse = imageRequest.GetResponse();

                                Stream responseStream = imageResponse.GetResponseStream();

                                using (BinaryReader br = new BinaryReader(responseStream))
                                {
                                    imageBytes = br.ReadBytes(500000);
                                    br.Close();
                                }
                                responseStream.Close();
                                imageResponse.Close();

                                FileStream fs = new FileStream(saveLocation, FileMode.Create);
                                BinaryWriter bw = new BinaryWriter(fs);
                                try
                                {
                                    bw.Write(imageBytes);

                                }
                                finally
                                {
                                    fs.Close();
                                    bw.Close();
                                }
                            }
                        }
                        catch { }

                    }

                    #region code
                    //resprodd = doc.DocumentNode.SelectSingleNode("//div[@id='pr_image_ie']");
                    //if (resprodd != null)
                    //{

                    //    HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                    //    doc2.LoadHtml(resprodd.InnerHtml.ToString());
                    //    imagepath = doc2.DocumentNode.SelectSingleNode("//img").Attributes["src"].Value;
                    //    try
                    //    {                            
                    //        if (imagepath.Length > 0)
                    //        {
                    //            string saveLocation = ImageFolderPath + "\\" + Path.GetFileName(imagepath);
                    //            byte[] imageBytes;
                    //            HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imagepath);
                    //            WebResponse imageResponse = imageRequest.GetResponse();

                    //            Stream responseStream = imageResponse.GetResponseStream();

                    //            using (BinaryReader br = new BinaryReader(responseStream))
                    //            {
                    //                imageBytes = br.ReadBytes(500000);
                    //                br.Close();
                    //            }
                    //            responseStream.Close();
                    //            imageResponse.Close();

                    //            FileStream fs = new FileStream(saveLocation, FileMode.Create);
                    //            BinaryWriter bw = new BinaryWriter(fs);
                    //            try
                    //            {
                    //                bw.Write(imageBytes);

                    //            }
                    //            finally
                    //            {
                    //                fs.Close();
                    //                bw.Close();
                    //            }
                    //        }
                    //    }
                    //    catch { }

                    //} 






                    //tables = doc.DocumentNode.SelectNodes("//table");
                    //if (tables != null)
                    //{

                    //    for (int t = 0; t < tables.Count; ++t)
                    //    {
                    //        rows = tables[t].SelectNodes(".//tr");
                    //        if (rows != null)
                    //        {
                    //            for (int i = 0; i < rows.Count; ++i)
                    //            {
                    //                // Iterate all columns in this row
                    //                HtmlNodeCollection cols = rows[i].SelectNodes(".//td");
                    //                if (cols != null)
                    //                {
                    //                    ////Manufacturer Information
                    //                    //if (cols[0].InnerText.Trim() == "Manufacturer")
                    //                    //{
                    //                    //    Manufacturer = GetText(cols[1]).Trim();
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Collection")
                    //                    //{
                    //                    //    Collection = GetText(cols[1]).Trim();
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "SKU")
                    //                    //{
                    //                    //    SKU = GetText(cols[1]).Trim();
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "UPC")
                    //                    //{
                    //                    //    UPC = GetText(cols[1]).Trim();
                    //                    //}

                    //                    ////Design Information
                    //                    //if (cols[0].InnerText.Trim() == "Category")
                    //                    //{
                    //                    //    Category = GetText(cols[1]).Trim();
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Finish")
                    //                    //{
                    //                    //    Finish = GetText(cols[1]).Trim();
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Glass")
                    //                    //{
                    //                    //    Glass = GetText(cols[1]).Trim();
                    //                    //}

                    //                    ////Dimensions and Weight (inches and pounds)
                    //                    //if (cols[0].InnerText.Trim() == "Dimensions")
                    //                    //{
                    //                    //    HtmlNode Dimensions = cols[1];
                    //                    //    HtmlAgilityPack.HtmlDocument docdim = new HtmlAgilityPack.HtmlDocument();
                    //                    //    docdim.LoadHtml(Dimensions.InnerHtml);
                    //                    //    HtmlNodeCollection tablesDim = docdim.DocumentNode.SelectNodes("//table");
                    //                    //    if (tablesDim != null)
                    //                    //    {
                    //                    //        HtmlNodeCollection rowsDim = tablesDim[0].SelectNodes(".//tr");
                    //                    //        if (rowsDim != null)
                    //                    //        {
                    //                    //            // Iterate all columns in this row
                    //                    //            HtmlNodeCollection colsDim = rowsDim[1].SelectNodes(".//td");
                    //                    //            if (colsDim != null)
                    //                    //            {
                    //                    //                try
                    //                    //                {
                    //                    //                    length = GetText(colsDim[0]).Trim();
                    //                    //                    width = GetText(colsDim[1]).Trim();
                    //                    //                    height = GetText(colsDim[2]).Trim();
                    //                    //                    ext = GetText(colsDim[3]).Trim();
                    //                    //                }
                    //                    //                catch { }
                    //                    //            }

                    //                    //        }

                    //                    //    }
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Weight")
                    //                    //{
                    //                    //    Weight = GetText(cols[1]);
                    //                    //}

                    //                    ////Bulb Information
                    //                    //if (cols[0].InnerText.Trim() == "Bulb Type")
                    //                    //{
                    //                    //    BulbTypePrimary = GetText(cols[1]);
                    //                    //    BulbTypeSecondary = GetText(cols[2]);
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Number of Bulbs")
                    //                    //{
                    //                    //    NumberofBulbsPrimary = GetText(cols[1]);
                    //                    //    NumberofBulbsSecondary = GetText(cols[2]);
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Max Wattage")
                    //                    //{
                    //                    //    MaxWattagePrimary = GetText(cols[1]);
                    //                    //    MaxWattageSecondary = GetText(cols[2]);
                    //                    //}


                    //                    ////Shipping Information
                    //                    //if (cols[0].InnerText.Trim() == "Ships Via")
                    //                    //{
                    //                    //    ShipsVia = GetText(cols[1]);
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Ship Dimensions")
                    //                    //{
                    //                    //    HtmlNode ShipDimensions = cols[1];
                    //                    //    HtmlAgilityPack.HtmlDocument docdim = new HtmlAgilityPack.HtmlDocument();
                    //                    //    docdim.LoadHtml(ShipDimensions.InnerHtml);
                    //                    //    HtmlNodeCollection tablesDim = docdim.DocumentNode.SelectNodes("//table");
                    //                    //    if (tablesDim != null)
                    //                    //    {
                    //                    //        HtmlNodeCollection rowsDim = tablesDim[0].SelectNodes(".//tr");
                    //                    //        if (rowsDim != null)
                    //                    //        {
                    //                    //            // Iterate all columns in this row
                    //                    //            HtmlNodeCollection colsDim = rowsDim[1].SelectNodes(".//td");
                    //                    //            if (colsDim != null)
                    //                    //            {
                    //                    //                try
                    //                    //                {
                    //                    //                    Shiplength = GetText(colsDim[0]);
                    //                    //                    Shipwidth = GetText(colsDim[1]);
                    //                    //                    Shipheight = GetText(colsDim[2]);
                    //                    //                }
                    //                    //                catch { }
                    //                    //            }

                    //                    //        }

                    //                    //    }
                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "Ship Weight")
                    //                    //{
                    //                    //    ShipWeight = GetText(cols[1]);
                    //                    //}

                    //                    ////Product Rating
                    //                    //if (cols[0].InnerText.Trim() == "Outdoor Rating")
                    //                    //{
                    //                    //    OutdoorRating = GetText(cols[1]);

                    //                    //}
                    //                    //if (cols[0].InnerText.Trim() == "UL Rating")
                    //                    //{
                    //                    //    ULRating = GetText(cols[1]);
                    //                    //}
                    //                }
                    //            }
                    //        }
                    //    }



                    //}

                    #endregion

                    //htmlPath = "http://www.lightingnewyork.com/post/product.cfm?wtd=rel&pr_id=5097&pr_collection=Elysburg&pr_vendor=2";
                   string _path = "http://www.lightingnewyork.com/post/product.cfm?wtd=" + wtd + "&pr_id=" + pr_id + "&pr_collection=" + pr_collection + "&pr_vendor=" + pr_vendor + "";

                    html = PostPage(_path, out msgText);
                    if (html != null)
                    {
                        if (html.Length > 0)
                        {
                            doc.LoadHtml(html);

                            //res = doc.DocumentNode.SelectSingleNode("//div[@id='relative_products']");
                            //if (res != null)
                            //{
                            //doc.LoadHtml(res.InnerHtml);

                            foreach (var divcat in doc.DocumentNode.SelectNodes("//div[@class='rel_pr']"))
                            {
                                HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                                doc2.LoadHtml(divcat.InnerHtml.ToString());
                                relativelist += doc2.DocumentNode.SelectSingleNode("//a").Attributes["href"].Value + ",";
                            }
                            //}
                        }
                    }





                    //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";
                    FileInfo newfile = new FileInfo(filePath);



                    using (ExcelPackage xlPackage = new ExcelPackage(newfile))
                    {
                        ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                        if (worksheet == null)
                            worksheet = xlPackage.Workbook.Worksheets.Add("Product");

                        /* set column in excel */
                        worksheet.Cell(prodrowno, 1).Value = SKU;
                        worksheet.Cell(prodrowno, 2).Value = productTitle;
                        worksheet.Cell(prodrowno, 3).Value = productLink;
                        worksheet.Cell(prodrowno, 4).Value = productSku;
                        worksheet.Cell(prodrowno, 5).Value = strinStock;
                        worksheet.Cell(prodrowno, 6).Value = strCurrentPrice;
                        worksheet.Cell(prodrowno, 7).Value = strOldPrice;
                        worksheet.Cell(prodrowno, 8).Value = description;

                        worksheet.Cell(prodrowno, 9).Value = Manufacturer;
                        worksheet.Cell(prodrowno, 10).Value = Collection;

                        worksheet.Cell(prodrowno, 11).Value = UPC;

                        worksheet.Cell(prodrowno, 12).Value = Category;
                        worksheet.Cell(prodrowno, 13).Value = Finish;
                        worksheet.Cell(prodrowno, 14).Value = Glass;
                        worksheet.Cell(prodrowno, 15).Value = Material;


                        //worksheet.Cell(prodrowno, 15).Value = length;
                        worksheet.Cell(prodrowno, 16).Value = width;
                        worksheet.Cell(prodrowno, 17).Value = height;
                        worksheet.Cell(prodrowno, 18).Value = bcWidth;
                        worksheet.Cell(prodrowno, 19).Value = bcHeight;


                        worksheet.Cell(prodrowno, 20).Value = Weight;

                        worksheet.Cell(prodrowno, 21).Value = BulbsIncluded;
                        worksheet.Cell(prodrowno, 22).Value = BulbTypePrimary;
                        worksheet.Cell(prodrowno, 23).Value = NumberofBulbsPrimary;
                        worksheet.Cell(prodrowno, 24).Value = MaxWattagePrimary;

                        //worksheet.Cell(prodrowno, 23).Value = BulbTypeSecondary;
                        //worksheet.Cell(prodrowno, 24).Value = NumberofBulbsSecondary;
                        //worksheet.Cell(prodrowno, 25).Value = MaxWattageSecondary;


                        worksheet.Cell(prodrowno, 25).Value = ShipsVia;


                        worksheet.Cell(prodrowno, 26).Value = Shiplength;
                        worksheet.Cell(prodrowno, 27).Value = Shipwidth;
                        worksheet.Cell(prodrowno, 28).Value = Shipheight;

                        worksheet.Cell(prodrowno, 29).Value = ShipWeight;

                        worksheet.Cell(prodrowno, 30).Value = OutdoorRating;
                        worksheet.Cell(prodrowno, 31).Value = ULRating;
                        worksheet.Cell(prodrowno, 32).Value = caturl;
                        worksheet.Cell(prodrowno, 33).Value = relativelist;
                        worksheet.Cell(prodrowno, 34).Value = imagepath;
                        worksheet.Cell(prodrowno, 35).Value = finishType;
                        worksheet.Cell(prodrowno, 36).Value = shadeType;
                        worksheet.Cell(prodrowno, 37).Value = bulbType;

                        xlPackage.Save();
                        WriteLog("scrapping complete");
                    }

                }
            }
            else
            {
                WriteLog("html null");
            }
        }

        void WriteLog(string msg)
        {
            using (StreamWriter tw = new StreamWriter("Log.txt", true))
            {
                tw.WriteLine("===========================" + msg + "=========================");
                tw.WriteLine(htmlPath);
            }
        }

        private int GetProductDetailsFromWebNew(string htmlPath, string filePath, int prodrowno, string caturl)
        {

            string ItemCode = "";
            string ProductTitle = "";
            string productLink = "";
            string productSku = "";
            string strinStock = "";
            string CurrentPrice = "";
            string OldPrice = "";
            string Description = "";

            string ADA = "";
            string BackplateDiameter = "";
            string BackplateHeight = "";
            string BackplateWidth = "";
            string BaseColor = "";
            string BulbBase = "";
            string BulbIncluded = "";
            string BulbType = "";
            string ChandelierType = "";
            string Chainlength = "";
            string Depth = "";
            string MaximumHeight = "";
            string Length = "";
            string Collection = "";
            string DownRodSize = "";
            string DownRodIncluded = "";
            string EnergyStar = "";
            string Extension = "";
            string FinishApplication = "";
            string FullBackplate = "";
            string Genre = "";
            string Height = "";
            string InstallationAvailable = "";
            string LightDirection = "";
            string Material = "";
            string NumberofBulbs = "";
            string NumberofTiers = "";
            string Series = "";
            string Shade = "";
            string ShadeColor = "";
            string ShadeMaterial = "";
            string ShadeShape = "";
            string ShadeType = "";
            string Style = "";
            string SuggestedRoomFit = "";
            string Theme = "";
            string ULlisted = "";
            string ULRating = "";
            string Wattage = "";
            string WattsperBulb = "";
            string Width = "";
            string WireLength = "";

            string CanopyDiameter = "";
            string CanopyHeight = "";
            string Dimmable = "";
            string Voltage = "";
            string Characteristics = "";
            string productImagePath = "";
            string finish = "";



            // string wtd = "rel", pr_id, pr_collection, pr_vendor;
            //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";


            string html = "", txt;
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            html = GetPage(htmlPath, out msgText);
            if (html != null)
            {
                if (html.Length > 0)
                {
                    doc.LoadHtml(html);

                    HtmlNode.ElementsFlags.Remove("option");
                    ArrayList finList = new ArrayList();
                    ArrayList imgList = new ArrayList();

                    foreach (HtmlNode node in doc.DocumentNode.SelectNodes("//select[@id='thisfinish']//option"))
                    {
                        imgList.Add(node.Attributes["data-image"].Value);
                        finList.Add(node.NextSibling.InnerText.Replace("\n", "").Replace("\t", "").Replace("\r", ""));
                    }
                    /* For title this is working very fine */
                    var res = doc.DocumentNode.SelectSingleNode("//h1[@id='title']");
                    int count = 0;



                    // doc.LoadHtml(strRemstring);
                    ItemCode = GetText(doc.DocumentNode.SelectNodes("//div[@id='itemId']").FirstOrDefault()).Replace("Item #:", "");
                    ProductTitle = GetText(doc.DocumentNode.SelectNodes("//span[@itemprop='manufacturer']").FirstOrDefault()) + GetText(doc.DocumentNode.SelectNodes("//span[@itemprop='model']").FirstOrDefault()) + GetText(doc.DocumentNode.SelectNodes("//span[@itemprop='description']").FirstOrDefault());



                    doc.LoadHtml(html);
                    productLink = doc.DocumentNode.SelectSingleNode("//link[@rel='canonical']").Attributes["href"].Value;
                    CurrentPrice = GetText(doc.DocumentNode.SelectSingleNode("//div[@itemprop='price']")).Replace("$", "");
                    OldPrice = GetText(doc.DocumentNode.SelectSingleNode("//div[@class='savings']")).Split(',')[0].Replace("Originally", "").Replace("$", "");
                    strinStock = GetText(doc.DocumentNode.SelectSingleNode("//div[@id='stockCount']")).Replace("In Stock", "");
                    productSku = GetText(doc.DocumentNode.SelectSingleNode("//strong[@class='oursku']")).Replace("Our SKU:", "");
                    // productImagePath = doc.DocumentNode.SelectSingleNode("//img[@id='productImage']").Attributes["src"].Value;

                    /* upto here */





                    /* We are going to take price,available quantity, */
                    doc.LoadHtml(html);

                    res = doc.DocumentNode.SelectSingleNode("//div[@id='techspecs']");
                    doc.LoadHtml(res.InnerHtml);
                    foreach (var spec in doc.DocumentNode.SelectNodes("//table"))
                    {
                        HtmlAgilityPack.HtmlDocument doc2 = new HtmlAgilityPack.HtmlDocument();

                        doc2.LoadHtml(spec.InnerHtml.ToString());

                        foreach (var divProdLink in doc2.DocumentNode.SelectNodes("//tr"))
                        {
                            HtmlAgilityPack.HtmlDocument doc3 = new HtmlAgilityPack.HtmlDocument();
                            doc3.LoadHtml(divProdLink.InnerHtml.ToString());
                            string strTechSpecName = string.Empty;

                            if (doc3.DocumentNode.SelectSingleNode("//a") != null)
                                strTechSpecName = GetText(doc3.DocumentNode.SelectSingleNode("//a"));
                            else
                                strTechSpecName = GetText(doc3.DocumentNode.SelectSingleNode("//th"));

                            switch (strTechSpecName.ToLower())
                            {
                                case "ada":
                                    ADA = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "backplate diameter":
                                    BackplateDiameter = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "backplate height":
                                    BackplateHeight = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "backplate width":
                                    BackplateWidth = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "base color":
                                    BaseColor = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "bulb base":
                                    BulbBase = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "bulb included":
                                    BulbIncluded = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "bulb type":
                                    BulbType = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "length":
                                    Length = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "collection":
                                    Collection = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "energy star":
                                    EnergyStar = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "extension":
                                    Extension = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "finish application":
                                    FinishApplication = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "full backplate":
                                    FullBackplate = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "genre":
                                    Genre = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "height":
                                    Height = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "installation available":
                                    InstallationAvailable = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "light direction":
                                    LightDirection = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "material":
                                    Material = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "number of bulbs":
                                    NumberofBulbs = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "shade":
                                    Shade = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "shade color":
                                    ShadeColor = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "shade material":
                                    ShadeMaterial = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "shade shape":
                                    ShadeShape = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "shade type":
                                    ShadeType = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "style":
                                    Style = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "suggested room fit":
                                    SuggestedRoomFit = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "theme":
                                    Theme = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "ul listed":
                                    ULlisted = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "ul rating":
                                    ULRating = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "wattage":
                                    Wattage = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "watts per bulb":
                                    WattsperBulb = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "width":
                                    Width = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "chandelier type":
                                    ChandelierType = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "number of tiers":
                                    NumberofTiers = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "series":
                                    Series = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "wire length":
                                    WireLength = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "maximum height":
                                    MaximumHeight = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "chain length":
                                    Chainlength = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "depth":
                                    Depth = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "downrod size(s)":
                                    DownRodSize = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "downrod(s) included":
                                    DownRodIncluded = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "canopy diameter":
                                    CanopyDiameter = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "canopy height":
                                    CanopyHeight = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "dimmable":
                                    Dimmable = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "voltage":
                                    Voltage = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                case "characteristics":
                                    Characteristics = GetText(doc3.DocumentNode.SelectSingleNode("//td"));
                                    break;
                                default: break;
                            }
                        }
                    }

                    FileInfo newfile = new FileInfo(filePath);

                    foreach (string fin in finList)
                    {


                        finish = fin;
                        productImagePath = imgList[count].ToString();

                        //string filePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\Data.xlsx";




                        using (ExcelPackage xlPackage = new ExcelPackage(newfile))
                        {
                            ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                            if (worksheet == null)
                                worksheet = xlPackage.Workbook.Worksheets.Add("Product");

                            /* set column in excel */
                            worksheet.Cell(prodrowno, 1).Value = ItemCode;
                            worksheet.Cell(prodrowno, 2).Value = ProductTitle;
                            worksheet.Cell(prodrowno, 3).Value = productLink;
                            worksheet.Cell(prodrowno, 4).Value = productSku;
                            worksheet.Cell(prodrowno, 5).Value = strinStock;
                            worksheet.Cell(prodrowno, 6).Value = CurrentPrice;
                            worksheet.Cell(prodrowno, 7).Value = OldPrice;
                            worksheet.Cell(prodrowno, 8).Value = Description;

                            worksheet.Cell(prodrowno, 9).Value = ADA;
                            worksheet.Cell(prodrowno, 10).Value = BackplateDiameter;

                            worksheet.Cell(prodrowno, 11).Value = BackplateHeight;

                            worksheet.Cell(prodrowno, 12).Value = BackplateWidth;
                            worksheet.Cell(prodrowno, 13).Value = BaseColor;
                            worksheet.Cell(prodrowno, 14).Value = BulbBase;
                            worksheet.Cell(prodrowno, 15).Value = BulbIncluded;
                            worksheet.Cell(prodrowno, 16).Value = BulbType;


                            //worksheet.Cell(prodrowno, 15).Value = Length;
                            worksheet.Cell(prodrowno, 17).Value = Collection;
                            worksheet.Cell(prodrowno, 18).Value = EnergyStar;
                            worksheet.Cell(prodrowno, 19).Value = Extension;
                            worksheet.Cell(prodrowno, 20).Value = FinishApplication;
                            worksheet.Cell(prodrowno, 21).Value = FullBackplate;
                            worksheet.Cell(prodrowno, 22).Value = Genre;
                            worksheet.Cell(prodrowno, 23).Value = Height;
                            worksheet.Cell(prodrowno, 24).Value = InstallationAvailable;
                            worksheet.Cell(prodrowno, 25).Value = LightDirection;
                            worksheet.Cell(prodrowno, 26).Value = Material;
                            worksheet.Cell(prodrowno, 27).Value = NumberofBulbs;
                            worksheet.Cell(prodrowno, 28).Value = Shade;
                            worksheet.Cell(prodrowno, 29).Value = ShadeColor;
                            worksheet.Cell(prodrowno, 30).Value = ShadeMaterial;
                            worksheet.Cell(prodrowno, 31).Value = ShadeShape;
                            worksheet.Cell(prodrowno, 32).Value = ShadeType;
                            worksheet.Cell(prodrowno, 33).Value = Style;
                            worksheet.Cell(prodrowno, 34).Value = SuggestedRoomFit;
                            worksheet.Cell(prodrowno, 35).Value = Theme;
                            worksheet.Cell(prodrowno, 36).Value = ULlisted;
                            worksheet.Cell(prodrowno, 37).Value = ULRating;
                            worksheet.Cell(prodrowno, 38).Value = Wattage;
                            worksheet.Cell(prodrowno, 39).Value = WattsperBulb;
                            worksheet.Cell(prodrowno, 40).Value = Width;
                            worksheet.Cell(prodrowno, 41).Value = ChandelierType;
                            worksheet.Cell(prodrowno, 42).Value = Series;
                            worksheet.Cell(prodrowno, 43).Value = WireLength;
                            worksheet.Cell(prodrowno, 44).Value = MaximumHeight;
                            worksheet.Cell(prodrowno, 45).Value = Chainlength;
                            worksheet.Cell(prodrowno, 46).Value = Depth;
                            worksheet.Cell(prodrowno, 47).Value = DownRodSize;
                            worksheet.Cell(prodrowno, 48).Value = DownRodIncluded;
                            worksheet.Cell(prodrowno, 49).Value = CanopyDiameter;
                            worksheet.Cell(prodrowno, 50).Value = CanopyHeight;
                            worksheet.Cell(prodrowno, 51).Value = Dimmable;
                            worksheet.Cell(prodrowno, 52).Value = Voltage;
                            worksheet.Cell(prodrowno, 53).Value = Characteristics;
                            worksheet.Cell(prodrowno, 54).Value = NumberofTiers;
                            worksheet.Cell(prodrowno, 55).Value = productImagePath;
                            worksheet.Cell(prodrowno, 56).Value = finish;

                            xlPackage.Save();
                        }

                        count++;
                        prodrowno++;
                    }

                }
            }

            return prodrowno;
        }

        private void addHeader(int prodrowno, string filePath)
        {
            FileInfo newfile = new FileInfo(filePath);



            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet == null)
                    worksheet = xlPackage.Workbook.Worksheets.Add("Product");

                /* set column in excel */
                worksheet.Cell(prodrowno, 1).Value = "SKU";
                worksheet.Cell(prodrowno, 2).Value = "Product Title";
                worksheet.Cell(prodrowno, 3).Value = "product Link";
                worksheet.Cell(prodrowno, 4).Value = "product Sku";
                worksheet.Cell(prodrowno, 5).Value = "strin Stock";
                worksheet.Cell(prodrowno, 6).Value = "Current Price";
                worksheet.Cell(prodrowno, 7).Value = "Old Price";
                worksheet.Cell(prodrowno, 8).Value = "Description";

                worksheet.Cell(prodrowno, 9).Value = "Manufacturer";
                worksheet.Cell(prodrowno, 10).Value = "Collection";

                worksheet.Cell(prodrowno, 11).Value = "UPC";

                worksheet.Cell(prodrowno, 12).Value = "Category";
                worksheet.Cell(prodrowno, 13).Value = "Finish";
                worksheet.Cell(prodrowno, 14).Value = "Glass";
                worksheet.Cell(prodrowno, 15).Value = "Material";


                //worksheet.Cell(prodrowno, 15).Value = "Length";
                worksheet.Cell(prodrowno, 16).Value = "Width";
                worksheet.Cell(prodrowno, 17).Value = "Height";
                worksheet.Cell(prodrowno, 18).Value = "Backplate/Canopy Width";
                worksheet.Cell(prodrowno, 19).Value = "Backplate/Canopy Height";

                worksheet.Cell(prodrowno, 20).Value = "Weight";

                worksheet.Cell(prodrowno, 21).Value = "Bulbs Included";
                worksheet.Cell(prodrowno, 22).Value = "Bulb Type Primary";
                worksheet.Cell(prodrowno, 23).Value = "Number of Bulbs Primary";
                worksheet.Cell(prodrowno, 24).Value = "Max Wattage Primary";

                //worksheet.Cell(prodrowno, 23).Value = "Bulb Type Secondary";
                //worksheet.Cell(prodrowno, 24).Value = "Number of Bulbs Secondary";
                //worksheet.Cell(prodrowno, 25).Value = "Max Wattage Secondary";


                worksheet.Cell(prodrowno, 25).Value = "Ships Via";


                worksheet.Cell(prodrowno, 26).Value = "Ship length";
                worksheet.Cell(prodrowno, 27).Value = "Ship width";
                worksheet.Cell(prodrowno, 28).Value = "Ship height";

                worksheet.Cell(prodrowno, 29).Value = "Ship Weight";

                worksheet.Cell(prodrowno, 30).Value = "Outdoor Rating";
                worksheet.Cell(prodrowno, 31).Value = "UL Rating";
                worksheet.Cell(prodrowno, 32).Value = "Category URL";

                worksheet.Cell(prodrowno, 33).Value = "Relative List";
                worksheet.Cell(prodrowno, 34).Value = "Image Path";
                worksheet.Cell(prodrowno, 35).Value = "FinishType";

                worksheet.Cell(prodrowno, 36).Value = "ShadeType";
                worksheet.Cell(prodrowno, 37).Value = "BulbType";
                xlPackage.Save();
            }

        }

        private void addHeaderNew(int prodrowno, string filePath)
        {
            FileInfo newfile = new FileInfo(filePath);



            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet == null)
                    worksheet = xlPackage.Workbook.Worksheets.Add("Product");

                /* set column in excel */
                worksheet.Cell(prodrowno, 1).Value = "Item Code";
                worksheet.Cell(prodrowno, 2).Value = "Product Title";
                worksheet.Cell(prodrowno, 3).Value = "product Link";
                worksheet.Cell(prodrowno, 4).Value = "product Sku";
                worksheet.Cell(prodrowno, 5).Value = "strin Stock";
                worksheet.Cell(prodrowno, 6).Value = "Current Price";
                worksheet.Cell(prodrowno, 7).Value = "Old Price";
                worksheet.Cell(prodrowno, 8).Value = "Description";

                worksheet.Cell(prodrowno, 9).Value = "ADA";
                worksheet.Cell(prodrowno, 10).Value = "Backplate Diameter";

                worksheet.Cell(prodrowno, 11).Value = "Backplate Height";

                worksheet.Cell(prodrowno, 12).Value = "Backplate Width";
                worksheet.Cell(prodrowno, 13).Value = "Base Color";
                worksheet.Cell(prodrowno, 14).Value = "Bulb Base";
                worksheet.Cell(prodrowno, 15).Value = "Bulb Included";
                worksheet.Cell(prodrowno, 16).Value = "Bulb Type";


                //worksheet.Cell(prodrowno, 15).Value = "Length";
                worksheet.Cell(prodrowno, 17).Value = "Collection";
                worksheet.Cell(prodrowno, 18).Value = "Energy Star";
                worksheet.Cell(prodrowno, 19).Value = "Extension";
                worksheet.Cell(prodrowno, 20).Value = "Finish Application";
                worksheet.Cell(prodrowno, 21).Value = "Full Backplate";
                worksheet.Cell(prodrowno, 22).Value = "Genre";
                worksheet.Cell(prodrowno, 23).Value = "Height";
                worksheet.Cell(prodrowno, 24).Value = "Installation Available";
                worksheet.Cell(prodrowno, 25).Value = "Light Direction";
                worksheet.Cell(prodrowno, 26).Value = "Material";
                worksheet.Cell(prodrowno, 27).Value = "Number of Bulbs";
                worksheet.Cell(prodrowno, 28).Value = "Shade";
                worksheet.Cell(prodrowno, 29).Value = "Shade Color";
                worksheet.Cell(prodrowno, 30).Value = "Shade Material";
                worksheet.Cell(prodrowno, 31).Value = "Shade Shape";
                worksheet.Cell(prodrowno, 32).Value = "Shade Type";
                worksheet.Cell(prodrowno, 33).Value = "Style";
                worksheet.Cell(prodrowno, 34).Value = "Suggested Room Fit";
                worksheet.Cell(prodrowno, 35).Value = "Theme";
                worksheet.Cell(prodrowno, 36).Value = "UL listed";
                worksheet.Cell(prodrowno, 37).Value = "UL Rating";
                worksheet.Cell(prodrowno, 38).Value = "Wattage";
                worksheet.Cell(prodrowno, 39).Value = "Watts per Bulb";
                worksheet.Cell(prodrowno, 40).Value = "Width";
                worksheet.Cell(prodrowno, 41).Value = "Chandelier Type";
                worksheet.Cell(prodrowno, 42).Value = "Series";
                worksheet.Cell(prodrowno, 43).Value = "Wire Length";
                worksheet.Cell(prodrowno, 44).Value = "Maximum Height";
                worksheet.Cell(prodrowno, 45).Value = "Chain length ";
                worksheet.Cell(prodrowno, 46).Value = "Depth";
                worksheet.Cell(prodrowno, 47).Value = "DownRod Size";
                worksheet.Cell(prodrowno, 48).Value = "DownRod Included ";
                worksheet.Cell(prodrowno, 49).Value = "Canopy Diameter";
                worksheet.Cell(prodrowno, 50).Value = "Canopy Height ";
                worksheet.Cell(prodrowno, 51).Value = "Dimmable";
                worksheet.Cell(prodrowno, 52).Value = "Voltage";
                worksheet.Cell(prodrowno, 53).Value = "Characteristics";
                worksheet.Cell(prodrowno, 54).Value = "Number of Tiers";
                worksheet.Cell(prodrowno, 55).Value = "Product Image Path";
                worksheet.Cell(prodrowno, 56).Value = "Finish";

                xlPackage.Save();
            }




        }

        private string GetText(HtmlNode cols)
        {
            if (cols != null)
                return cols.InnerText.Trim();

            return "";
        }

        private string GetTextOnly(HtmlNode cols)
        {
            if (cols != null)
            {
                var atxt = cols.SelectSingleNode("//a");
                if (atxt == null)
                    return atxt.InnerText.Trim();
                else
                    return cols.InnerText.Trim();
            }

            return "";
        }

        private string GetBrandNameFromURL(string url)
        {
            string brand = string.Empty;

            String[] arrUrl = url.Split(':');

            String[] arrVal = arrUrl[1].Split('/');

            if (arrVal[2] == "www.lightingnewyork.com")
            {
                brand = arrVal[4].Split('.')[0];
            }
            else if (arrVal[2] == "www.lightingdirect.com")
                brand = arrVal[3];


            return brand;
        }

        private string GetWebsiteFromURL(string url)
        {


            String[] arrUrl = url.Split(':');

            String[] arrVal = arrUrl[1].Split('/');



            return arrVal[2];
        }

        private void btnGetCategory_Click(object sender, EventArgs e)
        {
            try
            {
                //string CatfilePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\CatData.xlsx";

                //string htmlPath = @"http://localhost:58755/Web/Category/Category.htm";


                IEnumerable<string> hrefList;

                List<string> prodhrefList = new List<string>();

                #region Category

                hrefList = GetCategoryFromFile(CatfilePath);
                if (hrefList.Count() == 0)
                {
                    if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                    {
                        hrefList = GetCategoryFromWeb(htmlPath);
                        SaveCategoryToExcel(CatfilePath, hrefList);
                    }
                    else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                    {

                        hrefList = GetCategoryFromWebNew(htmlPath);
                        SaveCategoryToExcelNew(CatfilePath, hrefList);
                    }

                }
                MessageBox.Show("All category are saved in excel.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            #endregion



        }

        private void btnProductLink_Click(object sender, EventArgs e)
        {

            //string ProdLinkfilePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\ProdLinkData.xlsx";



            IEnumerable<string> hrefList = GetCategoryFromFile(CatfilePath);
            hrefList = GetCategoryFromFile(CatfilePath);

            List<string> prodhrefList = new List<string>();


            #region Product
            int prodcount = 0;
            int pagecount = 48;
            int pageindex = 1;
            //website = "http://www.lightingdirect.com";

            int rowcount = GetProductTotal(ProdLinkfilePath);

            foreach (string cathref in hrefList)
            {
                prodhrefList.Clear();

                if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                {
                    hrefList = GetProductFromFile(ProdLinkfilePath, cathref.Split(',')[0]);
                    pagecount = 30;
                    prodcount = Convert.ToInt32(cathref.Split(',')[1]);
                }
                else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                {

                    hrefList = GetProductFromFile(ProdLinkfilePath, cathref);
                    prodcount = Convert.ToInt32(GetProductCount(website + cathref.Split(',')[0].Replace(website, "")));
                }


                pageindex = 1;
                if (hrefList.Count() == 0)
                {
                    //<a href="/brand/elk-lighting.html?pageIndex=2&amp;cat=6" id="pr_show_more" url="/brand/elk-lighting.html" np="5" tp="6" type="brand" style="display: block;">
                    //    SHOW MORE RESULTS</a>



                    while (prodcount > 0)
                    {
                        if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                        {
                            prodhrefList.AddRange(GetProductFromWeb(website + cathref.Split(',')[0].Replace(website, "") + "&pageIndex=" + pageindex));
                        }
                        else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                        {
                            prodhrefList.AddRange(GetProductFromWebNew(website + cathref.Split(',')[0].Replace(website, "") + "?p=" + pageindex));
                        }


                        prodcount = prodcount - pagecount;
                        pageindex++;
                    }

                    SaveProductToExcel(ProdLinkfilePath, prodhrefList, cathref.Split(',')[0], ref rowcount);
                }
            }
            MessageBox.Show("All Products links are saved in excel.");
            #endregion


        }

        private void btnProductData_Click(object sender, EventArgs e)
        {

            //string ProdLinkfilePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\ProdLinkData.xlsx";
            //string ProdfilePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\ProdData.xlsx";

            //string website = @"http://www.lightingnewyork.com";

            IEnumerable<string> hrefList;

            List<string> prodhrefList = new List<string>();

            string url;

            //website = "http://www.lightingdirect.com";


            hrefList = GetProductFromFile(ProdLinkfilePath);
            int prodrowno = GetProductTotal(ProdfilePath);
            if (prodrowno == 0)
            {
                prodrowno++;
                if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                {
                    addHeader(prodrowno, ProdfilePath);
                }
                else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                {
                    addHeaderNew(prodrowno, ProdfilePath);
                }

            }

            if (hrefList.Count() > 0)
            {
                foreach (string cathref in hrefList)
                {
                    prodrowno++;
                    url = website + cathref.Split(',')[0];
                    try
                    {
                        if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
                        {
                            GetProductDetailsFromWeb(url, ProdfilePath, prodrowno, cathref.Split(',')[1]);
                        }
                        else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
                        {
                            int totalRows = GetProductDetailsFromWebNew(url, ProdfilePath, prodrowno, cathref.Split(',')[1]);
                            prodrowno = totalRows - 1;
                        }

                    }
                    catch (Exception ex)
                    {
                        using (StreamWriter tw = new StreamWriter("Error.txt", true))
                        {
                            tw.WriteLine("=======================================" + DateTime.Now.ToString() + "=============================================");
                            tw.WriteLine(url);
                            tw.WriteLine();
                            tw.WriteLine(ex.StackTrace);
                            tw.WriteLine("=======================================================================================================");
                            tw.WriteLine();
                            tw.WriteLine();
                        }

                    }
                }
            }
            MessageBox.Show("All Products data are saved in excel.");
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {

                //string CatfilePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\CatData.xlsx";
                //string ProdLinkfilePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\ProdLinkData.xlsx";
                //string ProdfilePath = @"D:\My Project 2010\Hussen\WinScrapping\WinScrapping\Excel\ProdData.xlsx";               
                //string htmlPath =@"http://www.lightingnewyork.com/brand/elk-lighting.html";
                //string website = @"http://www.lightingnewyork.com";

                //CatfilePath = ConfigurationManager.AppSettings["CatfilePath"].ToString();
                //ProdLinkfilePath = ConfigurationManager.AppSettings["ProdLinkfilePath"].ToString();
                //ProdfilePath = ConfigurationManager.AppSettings["ProdfilePath"].ToString();
                //RelatedProdfilePath = ConfigurationManager.AppSettings["RelatedProdfilePath"].ToString();
                //htmlPath = ConfigurationManager.AppSettings["htmlPath"].ToString();

                cbxWebsite.Items.Add("http://www.lightingdirect.com");
                cbxWebsite.Items.Add("http://www.lightingnewyork.com");
                cbxCategory.SelectedIndex = 0;
                //MessageBox.Show(CatfilePath);
                /*
                
                string filePath = Application.StartupPath + "App.Config";
                XmlTextReader myReader = new XmlTextReader(filePath);
                XDocument configdoc = XDocument.Load(myReader);

                var q = from p in configdoc.Descendants("config") select p;

                CatfilePath = Convert.ToString(q.First().Element("CatfilePath").Value);
                ProdLinkfilePath = Convert.ToString(q.First().Element("ProdLinkfilePath").Value);
                ProdfilePath = Convert.ToString(q.First().Element("ProdfilePath").Value);
                htmlPath = Convert.ToString(q.First().Element("htmlPath").Value);
                website = Convert.ToString(q.First().Element("website").Value);
                */

            }
            catch (Exception ex)
            {

            }
        }

        private void btnImageDownload_Click(object sender, EventArgs e)
        {
            IEnumerable<string> hrefList = null;

            List<string> prodhrefList = new List<string>();

            string url;



            if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingnewyork.com")
            {
                hrefList = GetImagesProductFile(ProdfilePath);
            }
            else if (GetWebsiteFromURL(cbxCategory.SelectedItem.ToString()) == "www.lightingdirect.com")
            {
                hrefList = GetImagesProductFileNew(ProdfilePath);
            }
            //int count=0;
            if (hrefList.Count() > 0)
            {
                foreach (string imagepath in hrefList)
                {
                    //count++;
                    //if (imghref.Length > 0)
                    //{

                    //    string saveLocation = ImageFolderPath + "\\" + Path.GetFileName(imghref);
                    //    //Worker objwor = new Worker(saveLoc, imghref);
                    //    //new Thread(new ThreadStart(objwor.DoWork), 20).Start();

                    //    //Thread t = new Thread(new ParameterizedThreadStart(DoWork),5);
                    //    //t.Start(new Worker() { saveLocation = saveLoc, imagepath = imghref });
                    //}
                    try
                    {
                        if (imagepath.Length > 0)
                        {
                            string saveLocation = ImageFolderPath + "\\" + Path.GetFileName(imagepath);
                            byte[] imageBytes;
                            HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imagepath);
                            WebResponse imageResponse = imageRequest.GetResponse();

                            Stream responseStream = imageResponse.GetResponseStream();

                            using (BinaryReader br = new BinaryReader(responseStream))
                            {
                                imageBytes = br.ReadBytes(500000);
                                br.Close();
                            }
                            responseStream.Close();
                            imageResponse.Close();

                            FileStream fs = new FileStream(saveLocation, FileMode.Create);
                            BinaryWriter bw = new BinaryWriter(fs);
                            try
                            {
                                bw.Write(imageBytes);

                            }
                            finally
                            {
                                fs.Close();
                                bw.Close();
                            }
                        }
                    }
                    catch { }
                }

            }
            MessageBox.Show("All Products Images are saved in excel.");
        }

        public static void DoWork(string saveLocation, string imagepath)
        {
            try
            {
                if (imagepath.Length > 0)
                {
                    //string saveLocation = ImageFolderPath + "\\" + Path.GetFileName(imagepath);
                    byte[] imageBytes;
                    HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imagepath);
                    WebResponse imageResponse = imageRequest.GetResponse();

                    Stream responseStream = imageResponse.GetResponseStream();

                    using (BinaryReader br = new BinaryReader(responseStream))
                    {
                        imageBytes = br.ReadBytes(500000);
                        br.Close();
                    }
                    responseStream.Close();
                    imageResponse.Close();

                    FileStream fs = new FileStream(saveLocation, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    try
                    {
                        bw.Write(imageBytes);

                    }
                    finally
                    {
                        fs.Close();
                        bw.Close();
                    }
                }
            }
            catch { }
            //while (!_shouldStop)
            //{
            //    Console.WriteLine("worker thread: working...");
            //}
            //Console.WriteLine("worker thread: terminating gracefully.");
        }

        private IEnumerable<string> GetImagesProductFile(string filePath)
        {
            List<string> hrefList = new List<string>();
            FileInfo newfile = new FileInfo(filePath);


            //brand_land_box
            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet != null)
                {
                    /* set column in excel */
                    int i = 1;
                    while (true)
                    {
                        if (worksheet.Cell(i, 1).Value == "")
                            break;

                        hrefList.Add(worksheet.Cell(i, 34).Value);
                        i++;
                    }
                }
            }

            return hrefList.Distinct();
        }

        private IEnumerable<string> GetImagesProductFileNew(string filePath)
        {
            List<string> hrefList = new List<string>();
            FileInfo newfile = new FileInfo(filePath);


            //brand_land_box
            using (ExcelPackage xlPackage = new ExcelPackage(newfile))
            {
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets["Product"];

                if (worksheet != null)
                {
                    /* set column in excel */
                    int i = 1;
                    while (true)
                    {
                        if (worksheet.Cell(i, 1).Value == "")
                            break;

                        hrefList.Add(worksheet.Cell(i, 55).Value);
                        i++;
                    }
                }
            }

            return hrefList.Distinct();
        }

        private void cbxWebsite_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbxCategory.Items.Clear();
            if (cbxWebsite.SelectedItem.ToString() == "http://www.lightingdirect.com")
            {
                cbxCategory.Items.Add("http://www.lightingdirect.com/livex-lighting/c14882");
                cbxCategory.Items.Add("http://www.lightingdirect.com/crystorama-lighting-group/c6461");
                cbxCategory.Items.Add("http://www.lightingdirect.com/kichler-lighting/c28");
            }
            else if (cbxWebsite.SelectedItem.ToString() == "http://www.lightingnewyork.com")
            {
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/elk-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/elegant-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/crystorama.html");

                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/hinkley-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/fredrick-ramond-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/lite-source.html");

                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/fine-art-lamps.html");

                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/kalco-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/allegri.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/seagull-lighting.html");

                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/crystorama.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/elk-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/maxim-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/et2-lighting.html");
                cbxCategory.Items.Add("http://www.lightingnewyork.com/brand/quoizel-lighting.html");

            }





            cbxCategory.SelectedIndex = 0;
        }


    }

    public class Worker
    {
        // This method will be called when the thread is started. 
        private string saveLocation;
        private string imagepath;
        public Worker(string saveLoc, string imgpath)
        {
            saveLocation = saveLoc;
            imagepath = imgpath;
        }

        public void DoWork()
        {
            try
            {
                if (imagepath.Length > 0)
                {
                    //string saveLocation = ImageFolderPath + "\\" + Path.GetFileName(imagepath);
                    byte[] imageBytes;
                    HttpWebRequest imageRequest = (HttpWebRequest)WebRequest.Create(imagepath);
                    WebResponse imageResponse = imageRequest.GetResponse();

                    Stream responseStream = imageResponse.GetResponseStream();

                    using (BinaryReader br = new BinaryReader(responseStream))
                    {
                        imageBytes = br.ReadBytes(500000);
                        br.Close();
                    }
                    responseStream.Close();
                    imageResponse.Close();

                    FileStream fs = new FileStream(saveLocation, FileMode.Create);
                    BinaryWriter bw = new BinaryWriter(fs);
                    try
                    {
                        bw.Write(imageBytes);

                    }
                    finally
                    {
                        fs.Close();
                        bw.Close();
                    }
                }
            }
            catch { }
            //while (!_shouldStop)
            //{
            //    Console.WriteLine("worker thread: working...");
            //}
            //Console.WriteLine("worker thread: terminating gracefully.");
        }


        public void RequestStop()
        {
            _shouldStop = true;
        }
        // Volatile is used as hint to the compiler that this data 
        // member will be accessed by multiple threads. 
        private volatile bool _shouldStop;
    }
}

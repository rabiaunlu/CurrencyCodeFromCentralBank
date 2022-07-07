using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Xml;
using System.Xml.Linq;


namespace ExchangeRate.ClientMethods
{
    public class TCMBClient
    {
        private string BaseUrl = ConfigurationManager.AppSettings["Url"].ToString();
        public void TCMBMetod(string year, string month, string day)
        {
            string NameUSD = ""; string BuyingUSD = ""; string SellingUSD = "";
            string NameEUR = ""; string BuyingEUR = ""; string SellingEUR = "";
            string date = "";
            DataTable dt = new DataTable();
            dt.Columns.Add("Date");
            dt.Columns.Add("CurrencyName");
            dt.Columns.Add("Buying");
            dt.Columns.Add("Selling");
            string url = "";
            for (int i = 1; i <= 29; i++)
            {
                if (i == 1 || i == 2 || i == 3 || i == 4 || i == 5 || i == 6 || i == 7 || i == 8 || i == 9)
                {
                    url = BaseUrl + "kurlar/" + year + month + "/" + "0" + i + month + year + ".xml";
                    date = year + month + "0" + i;
                }
                else
                {
                    url = BaseUrl + "kurlar/" + year + month + "/" + i + month + year + ".xml";
                    date = year + month + i;
                }

                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Accept.Add(new
                    MediaTypeWithQualityHeaderValue("application/xml"));
                    HttpResponseMessage response = client.GetAsync(url).Result;
                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        XDocument xdoc = XDocument.Parse(response.Content.ReadAsStringAsync().Result);
                        var CurrencyDataUSD = from Rates in xdoc.Descendants("Currency")
                                              where Rates.Element("Isim").Value.ToString() == "ABD DOLARI"
                                              select new
                                              {
                                                  CurrencyNameUSD = Rates.Element("CurrencyName").Value.ToString(),
                                                  ForexBuyingUSD = Rates.Element("ForexBuying").Value.ToString(),
                                                  ForexSellingUSD = Rates.Element("ForexSelling").Value.ToString()
                                              };

                        var CurrencyDataEUR = from Rates in xdoc.Descendants("Currency")
                                              where Rates.Element("Isim").Value.ToString() == "EURO"
                                              select new
                                              {
                                                  CurrencyNameEUR = Rates.Element("CurrencyName").Value,
                                                  ForexBuyingEUR = Rates.Element("ForexBuying").Value,
                                                  ForexSellingEUR = Rates.Element("ForexSelling").Value
                                              };
                        foreach (var dovizler in CurrencyDataUSD)
                        {
                            NameUSD = "USD";
                            BuyingUSD = dovizler.ForexBuyingUSD.ToString();
                            SellingUSD = dovizler.ForexSellingUSD.ToString();
                            dt.Rows.Add(NameUSD, BuyingUSD, SellingUSD, date);
                        }
                        foreach (var dovizler in CurrencyDataEUR)
                        {
                            NameEUR = dovizler.CurrencyNameEUR.ToString();
                            BuyingEUR = dovizler.ForexBuyingEUR.ToString();
                            SellingEUR = dovizler.ForexSellingEUR.ToString();
                            dt.Rows.Add(NameEUR, BuyingEUR, SellingEUR, date);
                        }
                    }
                }
            }
            //create excel file
            string fileName = "";
            fileName = "C:\\dosyaYolu\\doviz" + "_" + year + "_" + month + ".xlsx";
            File.Create(fileName).Dispose();
            //writing Data table to excel 
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workbook = new ExcelFile();
            var worksheet = workbook.Worksheets.Add("Data");
            // Insert DataTable to an Excel worksheet.
            worksheet.InsertDataTable(dt,
                new InsertDataTableOptions()
                {
                    ColumnHeaders = true,

                });

            workbook.Save(fileName);

        }
    }
}
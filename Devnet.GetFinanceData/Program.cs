using System;
using System.Net.Http;
using System.Threading.Tasks;
using RestSharp;
using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.IO;
using ClosedXML.Excel;

namespace Devnet.GetFinanceData
{
    class Program
    {
        static async Task Main(string[] args)
        {
            var sirketler = await GetSirketList();
            foreach (var sirket in sirketler.data)
            {
                // DİLEDİĞİNİZ TARİH ARALIĞINI VEREBİLİRSİNİZ.
                await GetData(sirket.kod, "01-01-2017", "27-06-2024");
            }
        }

        static async Task GetData(string sirket, string startDate, string endDate)
        {
            var url = "https://www.isyatirim.com.tr/_layouts/15/Isyatirim.Website/Common/Data.aspx/HisseTekil?hisse=" + sirket + "&startdate=" + startDate + "&enddate=" + endDate;

            var client = new RestClient(url);
            var request = new RestRequest(url, Method.Get);
            RestResponse response = client.Execute(request);
            FinanceData financeData = JsonConvert.DeserializeObject<FinanceData>(response.Content);
            Console.WriteLine("ŞİRKET:" + sirket + " İÇİN VERİLER ÇEKİLDİ.");

            await SaveFile(financeData.value, sirket);
        }

        static async Task<Sirketler> GetSirketList()
        {
            var url = "https://bigpara.hurriyet.com.tr/api/v1/hisse/list";
            var client = new RestClient(url);
            var request = new RestRequest(url, Method.Get);
            RestResponse response = client.Execute(request);
            Sirketler sirketler = JsonConvert.DeserializeObject<Sirketler>(response.Content);
            return sirketler;
        }

        static async Task SaveFile(List<FinanceValue> financeValues, string sirketKodu)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sheet1");
                var range = worksheet.Cell(1, 1).InsertTable(financeValues);
                worksheet.Columns().AdjustToContents();
                
                // DOSYA YOLUNU DEĞİŞTİRMEYİ UNUTMAYIN!
                string filePath = @"/Users/emin.goroglu/Documents/exceldata/" + sirketKodu + ".xlsx";
                
                
                workbook.SaveAs(filePath);
            }
            
            Console.WriteLine(sirketKodu + " için excel dosyası oluşturuldu.");
        }
    }
    
    public class Sirket
    {
        public int id { get; set; }
        public string kod { get; set; }
        public string ad { get; set; }
        public string tip { get; set; }
    }

    public class Sirketler
    {
        public string code { get; set; }
        public List<Sirket> data { get; set; }
    }

    public class FinanceData
    {
        public bool ok { get; set; }
        public object errorCode { get; set; }
        public object errorDescription { get; set; }
        public string transactionId { get; set; }
        public List<FinanceValue> value { get; set; }
    }

    public class FinanceValue
    {
        public string? HGDG_HS_KODU { get; set; }
        public string? HGDG_TARIH { get; set; }
        public double? HGDG_KAPANIS { get; set; }
        public double? HGDG_AOF { get; set; }
        public double? HGDG_MIN { get; set; }
        public double? HGDG_MAX { get; set; }
        public double? HGDG_HACIM { get; set; }
        public string? END_ENDEKS_KODU { get; set; }
        public long? END_TARIH { get; set; }
        public int? END_SEANS { get; set; }
        public double? END_DEGER { get; set; }
        public string? DD_DOVIZ_KODU { get; set; }
        public string? DD_DT_KODU { get; set; }
        public long? DD_TARIH { get; set; }
        public double? DD_DEGER { get; set; }
        public double? DOLAR_BAZLI_FIYAT { get; set; }
        public double? ENDEKS_BAZLI_FIYAT { get; set; }
        public double? DOLAR_HACIM { get; set; }
        public double? SERMAYE { get; set; }
        public double? HG_KAPANIS { get; set; }
        public double? HG_AOF { get; set; }
        public double? HG_MIN { get; set; }
        public double? HG_MAX { get; set; }
        public double? PD { get; set; }
        public double? PD_USD { get; set; }
        public double? HAO_PD { get; set; }
        public double? HAO_PD_USD { get; set; }
        public double? HG_HACIM { get; set; }
        public double? DOLAR_BAZLI_MIN { get; set; }
        public double? DOLAR_BAZLI_MAX { get; set; }
        public double? DOLAR_BAZLI_AOF { get; set; }
    }
}
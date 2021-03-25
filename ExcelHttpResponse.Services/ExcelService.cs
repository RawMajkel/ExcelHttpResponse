using ExcelHttpResponse.Common;
using OfficeOpenXml;
using System.IO;
using System.Net;

namespace ExcelHttpResponse.Services
{
    public class ExcelService
    {
        private static readonly int ExcelMaxRowNumber = 1048576;

        public ExcelService() 
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public void GetAndSaveResultsToFile(string url)
        {
            var fileInfo = new FileInfo(url);

            using var excelPackage = new ExcelPackage(fileInfo);
            var firstWorksheet = excelPackage.Workbook.Worksheets[0];

            for (int i = 1; i <= ExcelMaxRowNumber; i++)
            {
                var value = firstWorksheet.GetValue<string>(i, 1);

                if (value != null)
                {
                    var response = GetResultFromUrl(value);
                    firstWorksheet.Cells[i, 2].Value = response.ToString();
                }
            }

            excelPackage.Save();
        }

        public MajkelWorkerResponse GetResultFromUrl(string url, bool allowAutoRedirect = false)
        {
            try
            {
                var request = (HttpWebRequest)WebRequest.Create(url);
                request.AllowAutoRedirect = allowAutoRedirect;
                request.AutomaticDecompression = DecompressionMethods.Deflate | DecompressionMethods.GZip;
                var response = (HttpWebResponse)request.GetResponse();

                var workerResponse = new MajkelWorkerResponse()
                {
                    StatusCode = response.StatusCode
                };

                if (response.StatusCode == HttpStatusCode.Moved || response.StatusCode == HttpStatusCode.MovedPermanently)
                {
                    var redirectionResponse = GetResponseUri(url);
                    workerResponse.AddRedirectionUrl(redirectionResponse);
                } 

                return workerResponse;
            }
            catch (WebException e)
            {
                var errorResult = (HttpWebResponse)e.Response;

                return new MajkelWorkerResponse()
                {
                    StatusCode = errorResult.StatusCode
                };
            }
        }
        private string GetResponseUri(string url)
        {
            var request = (HttpWebRequest)WebRequest.Create(url);
            var response = (HttpWebResponse)request.GetResponse();

            return response.ResponseUri.AbsoluteUri;
        }
    }
}

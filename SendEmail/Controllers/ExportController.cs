using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using Dapper.Contrib.Extensions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using SendEmail.Helpers;
using SendEmail.Models;

namespace SendEmail.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExportController : ControllerBase
    {
        private readonly ConnectionString _connectionString;

        public ExportController(ConnectionString connectionString)
        {
            _connectionString = connectionString;
        }

        [HttpGet]
        public async Task<IActionResult> Export()
        {
            var sql = SQLHelper.SelectUser;
            var getall = await _connectionString.Connection.QueryAsync<User>(sql);
            var columnheader = ParamHelper.columnHeaderExcel;

            byte[] result;

            using (var package = new ExcelPackage())
            {
                // add a new worksheet to the empty workbook
                var worksheet = package.Workbook.Worksheets.Add("Data User"); //Worksheet name
                using (var cells = worksheet.Cells[1, 1, 1, 3]) //(1,1) (1,3)
                {
                    cells.Style.Font.Bold = true;
                }
                //First add the headers
                for (var i = 0; i < columnheader.Count(); i++)
                {
                    worksheet.Cells[1, i + 1].Value = columnheader[i];
                    worksheet.Cells[1, i + 1].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Cells[1, i + 1].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                }
                //Add values
                var j = 2;
                var num = 0;
                foreach (var tdl in getall)
                {
                    worksheet.Cells["A" + j].Value = num + 1;
                    worksheet.Cells["A" + j].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Cells["A" + j].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    worksheet.Cells["B" + j].Value = tdl.Nama;
                    worksheet.Cells["B" + j].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Cells["B" + j].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    worksheet.Cells["C" + j].Value = tdl.No_Telp.ToString();
                    worksheet.Cells["C" + j].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    worksheet.Cells["C" + j].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                    j++;
                    num++;
                }
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();
                result = package.GetAsByteArray();
            }

            return File(result, ParamHelper.ExcelType, $"Excel.xlsx");
            //var hasilFile = File(result, ParamHelper.ExcelType, $"Excel.xls");

            ////Read the FileName and convert it to Byte array.
            //string fileName = Path.GetFileName(hasilFile.FileDownloadName);
            //try
            //{
            //    string ftp = "ftp://10.10.8.183/";
            //    string ftpFolder = "Uploads/";
            //    //Create FTP Request.
            //    FtpWebRequest request = (FtpWebRequest)WebRequest.Create(ftp + ftpFolder + fileName);
            //    request.Method = WebRequestMethods.Ftp.UploadFile;

            //    //Enter FTP Server credentials.
            //    request.Credentials = new NetworkCredential("marzotadwi@gmail.com", "25082016jovi");
            //    request.ContentLength = result.Length;
            //    request.UsePassive = true;
            //    request.UseBinary = true;
            //    request.KeepAlive = false;
            //    request.ServicePoint.ConnectionLimit = result.Length;
            //    request.EnableSsl = false;

            //    using (Stream requestStream = request.GetRequestStream())
            //    {
            //        requestStream.Write(result, 0, result.Length);
            //        requestStream.Close();
            //    }

            //    FtpWebResponse response = (FtpWebResponse)request.GetResponse();
            //    response.Close();
            //}
            //catch (WebException ex)
            //{
            //    throw new Exception((ex.Response as FtpWebResponse).StatusDescription);
            //}

            //return Redirect("/export");
        }
    }
}
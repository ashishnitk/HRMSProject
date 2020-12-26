using ClosedXML.Excel;
using HRReporting.Services;
using HRMS.Model;
using HRMS.Utility;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using ExcelDataReader;
using HRReporting.Model;
using HRReports.Utility.Excel;
using HRReports.Model;
using Microsoft.Azure.Cosmos;
using SelectPdf;

namespace HRMS.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportsController : ControllerBase
    {
        private readonly ILogger<ReportsController> _logger;
        private readonly ICosmosDbService _cosmosDbService;

        /// <summary>
        /// Generate APIs
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="cosmosDbService"></param>
        public ReportsController(ILogger<ReportsController> logger, ICosmosDbService cosmosDbService)
        {
            _logger = logger;
            _cosmosDbService = cosmosDbService;
        }

        /// <summary>
        /// Provident Fund statement of an employee
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("PFStatement")]

        public async Task<ActionResult> PFStatement(Month Month, int Year)
        {
            // DataTable dtProduct = ToDataTable<EmployeeData>(listEmp);
            try
            {
                QueryDefinition query = new QueryDefinition("select * from c where c.Month = @month").WithParameter("@month", string.Format("{0}{1}", Month, Year));

                List<Employee> listEmp = await _cosmosDbService.GetItemsAsync(query);

                ESICDataTable dt = Excel.GetPFDataTable(listEmp);

                using (XLWorkbook workBook = new XLWorkbook())
                {
                    workBook.Worksheets.Add(dt.DataTable, "ESI");
                    // workBook.Table. = false;
                    workBook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workBook.Style.Font.Bold = true;
                    var ws = workBook.Worksheet(1);
                    ws.Columns().AdjustToContents();
                    var rngHeaders = ws.Range("A" + dt.RowToBeBold + ":G" + dt.RowToBeBold);
                    // rngHeaders.Style.Fill.BackgroundColor = XLColor.VividViolet;
                    rngHeaders.Style.Font.Bold = true;

                    using (MemoryStream stream = new MemoryStream())
                    {
                        workBook.SaveAs(stream);
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("PF_Statement_{0}{1}.xlsx", Month, Year));
                    }
                }
            }
            catch (Exception e)
            {

                throw;
            }
        }

        /// <summary>
        /// Employees' State Insurance Corporation Report
        /// </summary>
        /// <param name="Month">Select the Month</param>
        /// <param name="Year">Four Digit Year</param>
        /// <returns></returns>
        /// <response code="200">Ok</response>
        /// <response code="404">Employee Data Not Found</response>
        /// <response code="500">Internal Server error</response>
        [HttpGet]
        // [ResponseCache(Duration = 60, Location = ResponseCacheLocation.Any, VaryByQueryKeys = new[] { "impactlevel", "pii" })]
        [Route("ESICStatement")]
        public async Task<ActionResult> ESICStatement(Month Month, int Year)
        {
            try
            {
                QueryDefinition query = new QueryDefinition("select * from c where c.Month = @month").WithParameter("@month", string.Format("{0}{1}", Month, Year));

                List<Employee> listEmp = await _cosmosDbService.GetItemsAsync(query);

                if (listEmp.Count > 0)
                {
                    ESICDataTable dt = Excel.GetESICDataTable(listEmp);

                    using (XLWorkbook workBook = new XLWorkbook())
                    {
                        workBook.Worksheets.Add(dt.DataTable, "ESI");
                        // workBook.Table. = false;
                        workBook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        workBook.Style.Font.Bold = true;
                        var ws = workBook.Worksheet(1);
                        ws.Columns().AdjustToContents();
                        var rngHeaders = ws.Range("A" + dt.RowToBeBold + ":G" + dt.RowToBeBold);
                        // rngHeaders.Style.Fill.BackgroundColor = XLColor.VividViolet;
                        rngHeaders.Style.Font.Bold = true;

                        using (MemoryStream stream = new MemoryStream())
                        {
                            workBook.SaveAs(stream);
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("ESIC_Statement_{0}{1}.xlsx", Month, Year));
                        }
                    }
                }
                else
                    return NotFound(string.Format("Employee data not found for the Month {0} and Year {1}", Month, Year));
            }
            catch (Exception e)
            {
                throw;
            }
        }

        /// <summary>
        /// Tax Deducted at Source statement
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("TDSStatement")]

        public async Task<ActionResult> TDSStatement(Month Month, int Year)
        {
            // DataTable dtProduct = ToDataTable<EmployeeData>(listEmp);
            try
            {
                QueryDefinition query = new QueryDefinition("select * from c where c.Month = @month").WithParameter("@month", string.Format("{0}{1}", Month, Year));

                List<Employee> listEmp = await _cosmosDbService.GetItemsAsync(query);

                ESICDataTable dt = Excel.GetPFDataTable(listEmp);

                using (XLWorkbook workBook = new XLWorkbook())
                {
                    workBook.Worksheets.Add(dt.DataTable, "ESI");
                    // workBook.Table. = false;
                    workBook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workBook.Style.Font.Bold = true;
                    var ws = workBook.Worksheet(1);
                    ws.Columns().AdjustToContents();
                    var rngHeaders = ws.Range("A" + dt.RowToBeBold + ":G" + dt.RowToBeBold);
                    // rngHeaders.Style.Fill.BackgroundColor = XLColor.VividViolet;
                    rngHeaders.Style.Font.Bold = true;

                    using (MemoryStream stream = new MemoryStream())
                    {
                        workBook.SaveAs(stream);
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", string.Format("PF_Statement_{0}{1}.xlsx", Month, Year));
                    }
                }
            }
            catch (Exception e)
            {

                throw;
            }
        }

        /// <summary>
        /// Get the Salary report of all Employee
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("GetAllEmployeeSalary")]
        //  [ApiExplorerSettings(IgnoreApi = true)]
        public async Task<ActionResult> GetAllEmployeeSalary()
        {
            try
            {
                QueryDefinition query = new QueryDefinition("select * from c");

                List<Employee> listEmp = await _cosmosDbService.GetItemsAsync(query);

                string cssFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Assets");

                string pdf_page_size = "A4";
                PdfPageSize pageSize = (PdfPageSize)Enum.Parse(typeof(PdfPageSize), pdf_page_size, true);

                string pdf_orientation = "Portrait";
                PdfPageOrientation pdfOrientation = (PdfPageOrientation)Enum.Parse(typeof(PdfPageOrientation), pdf_orientation, true);

                int webPageWidth = 1024;
                int webPageHeight = 0;

                HtmlToPdf converter = new HtmlToPdf();

                converter.Options.PdfPageSize = pageSize;
                converter.Options.PdfPageOrientation = pdfOrientation;
                converter.Options.WebPageWidth = webPageWidth;
                converter.Options.WebPageHeight = webPageHeight;

                PdfDocument doc = converter.ConvertHtmlString(PDF.GetSalariesHTML(listEmp), cssFilePath);


                // save pdf
                byte[] pdf = doc.Save();
                doc.Close();

                return new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = "Document.pdf"
                };
            }
            catch (Exception e)
            {

                throw;
            }
        }


        /// <summary>
        /// Get Payslip of an employee
        /// </summary>
        /// <param name="EmplID">Employee Code. e.g. MT152</param>
        /// <param name="Month">Select the Month</param>
        /// <param name="Year">Four Digit year. e.g. 2021</param>
        /// <returns></returns>
        [HttpGet]
        [Route("PaySlip")]
        //  [ApiExplorerSettings(IgnoreApi = true)]
        public async Task<ActionResult> PaySlip(string EmplID, Month Month, int Year)
        {
            try
            {
                QueryDefinition query = new QueryDefinition("select * from c where c.Month = @month and c.EmplId = @emplid");
                query.WithParameter("@month", string.Format("{0}{1}", Month, Year));
                query.WithParameter("@emplid", EmplID);
                List<Employee> listEmp = await _cosmosDbService.GetItemsAsync(query);

                string cssFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Assets");

                string pdf_page_size = "A4";
                PdfPageSize pageSize = (PdfPageSize)Enum.Parse(typeof(PdfPageSize), pdf_page_size, true);

                string pdf_orientation = "Portrait";
                PdfPageOrientation pdfOrientation = (PdfPageOrientation)Enum.Parse(typeof(PdfPageOrientation), pdf_orientation, true);

                int webPageWidth = 1024;
                int webPageHeight = 0;

                HtmlToPdf converter = new HtmlToPdf();

                converter.Options.PdfPageSize = pageSize;
                converter.Options.PdfPageOrientation = pdfOrientation;
                converter.Options.WebPageWidth = webPageWidth;
                converter.Options.WebPageHeight = webPageHeight;

                PdfDocument doc = converter.ConvertHtmlString(PDF.PaySlip(listEmp.Where(a => a.EmplId == EmplID).FirstOrDefault()), cssFilePath);


                // save pdf
                byte[] pdf = doc.Save();
                doc.Close();

                return new FileContentResult(pdf, "application/pdf")
                {
                    FileDownloadName = "Document.pdf"
                };
            }
            catch (Exception e)
            {

                throw;
            }
        }
    }
}

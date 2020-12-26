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

    public class ReportController : ControllerBase
    {
        public static List<Employee> listEmp;
        private readonly ILogger<ReportController> _logger;
        private readonly ICosmosDbService _cosmosDbService;

        public ReportController(ILogger<ReportController> logger, ICosmosDbService cosmosDbService)
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
        /// Upload the Salary Register
        /// </summary>
        /// <param name="file">Input File in xls format</param>
        /// <returns></returns>
        [HttpPost("SalaryRegister")]
        public async Task<IActionResult> Upload(IFormFile file)
        {
            try
            {
                if (file == null || file.Length == 0)
                    return Content("File Not Selected");

                string fileExtension = Path.GetExtension(file.FileName);
                if (fileExtension != ".xls" && fileExtension != ".xlsx")
                    return Content("Invalid file format, Please upload .xls file");

                List<Employee> listOfEmployees = Excel.ParseAllEmployees(file);

                await _cosmosDbService.createBulkItemAsync(listOfEmployees);
                return Ok();
            }
            catch (Exception e)
            {
                return Content(string.Format("Upload File Thrown Exception {0}", e.Message));
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

                // instantiate converter object
                SelectPdf.HtmlToPdf converter = new SelectPdf.HtmlToPdf();

                // set converter options
                converter.Options.WebPageWidth = 1024;
                converter.Options.WebPageHeight = 0;

                converter.Options.PdfPageSize = PdfPageSize.A4;
                converter.Options.PdfPageOrientation = PdfPageOrientation.Portrait;

                SelectPdf.PdfDocument doc;
                string url = "https://selectpdf.com/community-edition/";
                // convert url or html string to pdf
                // doc = converter.ConvertUrl(url);
                doc = converter.ConvertHtmlString(PDF.GetSalariesHTML(listEmp));
                //if (!string.IsNullOrEmpty(url))
                //{
                //}
                //else
                //{
                //    // doc = converter.ConvertHtmlString(html, base_url);
                //}

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

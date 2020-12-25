using ClosedXML.Excel;
using DinkToPdf;
using DinkToPdf.Contracts;
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

namespace HRMS.Controllers
{
    [ApiController]
    [Route("[controller]")]

    public class ReportController : ControllerBase
    {
        public static List<Employee> listEmp;
        private IConverter _converter;
        private readonly ILogger<ReportController> _logger;
        private readonly ICosmosDbService _cosmosDbService;

        public ReportController(ILogger<ReportController> logger, IConverter converter, ICosmosDbService cosmosDbService)
        {
            _logger = logger;
            _converter = converter;
            _cosmosDbService = cosmosDbService;
        }

        /// <summary>
        /// Provident Fund statement of an employee
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        [Route("PFStatement")]

        public async Task<ActionResult> PFStatement()
        {
            // DataTable dtProduct = ToDataTable<EmployeeData>(listEmp);
            try
            {
                QueryDefinition query = new QueryDefinition("select * from c where c.Month = @month").WithParameter("@month", string.Format("{0}{1}", "March", 2020));

                List<Employee> listEmp = await _cosmosDbService.GetItemsAsync(query);
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
                        return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ESIC_Statement.xlsx");
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
                            return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ESIC_Statement.xlsx");
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
        [ApiExplorerSettings(IgnoreApi = true)]
        public ActionResult GetAllEmployeeSalary()
        {
            try
            {
                var globalSettings = new GlobalSettings
                {
                    ColorMode = ColorMode.Color,
                    Orientation = Orientation.Portrait,
                    PaperSize = PaperKind.A4,
                    Margins = new MarginSettings { Top = 10 },
                    DocumentTitle = "Salary Report"
                };
                var objectSettings = new ObjectSettings
                {
                    PagesCount = true,
                    HtmlContent = TemplateGenerator.GetHTMLString(),
                    WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "Assets", "styles.css") },
                    HeaderSettings = { FontName = "Arial", FontSize = 9, Right = "Page [page] of [toPage]", Line = true },
                    FooterSettings = { FontName = "Arial", FontSize = 9, Line = true, Center = "Report Footer" }
                };
                var pdf = new HtmlToPdfDocument()
                {
                    GlobalSettings = globalSettings,
                    Objects = { objectSettings }
                };
                var file = _converter.Convert(pdf);

                var res = File(file, "application/pdf");
                // Response.AddHeader("Content-Disposition", "attachment; filename=receipt.pdf");

                return File(file, "application/pdf", "MyRenamedFile.pdf");
            }
            catch (Exception e)
            {

                throw;
            }
        }
    }
}

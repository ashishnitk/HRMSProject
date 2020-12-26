using HRMS.Model;
using HRReporting.Services;
using HRReports.Utility.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace HRMS.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DataFeedController : ControllerBase
    {

        private readonly ILogger<DataFeedController> _logger;
        private readonly ICosmosDbService _cosmosDbService;

        /// <summary>
        /// Generate APIs
        /// </summary>
        /// <param name="logger"></param>
        /// <param name="cosmosDbService"></param>
        public DataFeedController(ILogger<DataFeedController> logger, ICosmosDbService cosmosDbService)
        {
            _logger = logger;
            _cosmosDbService = cosmosDbService;
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

    }
}

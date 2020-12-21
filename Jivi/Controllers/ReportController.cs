using ClosedXML.Excel;
using DinkToPdf;
using DinkToPdf.Contracts;
using Jivi.Model;
using Jivi.Utility;
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

namespace Jivi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportController : ControllerBase
    {
        public static List<EmployeeData> listEmp;
        private IConverter _converter;
        private readonly ILogger<ReportController> _logger;

        public ReportController(ILogger<ReportController> logger, IConverter converter)
        {
            _logger = logger;
            _converter = converter;

            listEmp = new List<EmployeeData>();
            string con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Temp\SalaryData\Salary Register.xlsx; Extended Properties='Excel 8.0;HDR=Yes;'";
            using (OleDbConnection connection = new OleDbConnection(con))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select * from [Active$]", connection);
                using (OleDbDataReader rdr = command.ExecuteReader())
                {
                    while (rdr.Read())
                    {
                        if (rdr[0] != null && int.TryParse(rdr[0].ToString(), out int res))
                        {
                            EmployeeData ed = new EmployeeData()
                            {
                                SerialNumber = res,
                                Code = rdr[1].ToString(),
                                Name = rdr[2].ToString(),
                                DOJ = rdr[3].ToString(),
                                DOL = rdr[4].ToString(),
                                DOB = rdr[5].ToString(),
                                BankAcNo = rdr[6].ToString(),
                                Bankname = rdr[7].ToString(),
                                IFSCCode = rdr[8].ToString(),
                                PANNo = rdr[9].ToString(),
                                PFNo = rdr[10].ToString(),
                                UANNo = rdr[11].ToString(),
                                InsuranceNo = rdr[12].ToString(),
                                Location = rdr[13].ToString(),
                                PTLocation = rdr[14].ToString(),
                                Designation = rdr[15].ToString(),
                                Department = rdr[16].ToString(),
                                OriginalCTC = Convert.ToInt32(rdr[18]),
                                CTCPA = Convert.ToInt32(rdr[19]),
                                CTCPM = Convert.ToInt32(rdr[20]),
                                MasterBasic = Convert.ToInt32(rdr[21]),
                                MasterHRA = Convert.ToInt32(rdr[22]),
                                MasterStatutoryBonus = Convert.ToInt32(rdr[23]),
                                MasterLTA = Convert.ToInt32(rdr[24]),
                                MasterTelephoneReimbursement = Convert.ToInt32(rdr[25]),
                                MasterAttireAllowance = Convert.ToInt32(rdr[26]),
                                MasterFuelReimbursment = Convert.ToInt32(rdr[27]),
                                ERPF = Convert.ToInt32(rdr[28]),
                                ERESIC = Convert.ToInt32(rdr[29]),
                                MasterProjectAllowance = Convert.ToInt32(rdr[30]),
                                TotalCTC = Convert.ToInt32(rdr[31]),
                                CTCPA2 = Convert.ToInt32(rdr[32]),
                                CalculatedGross = Convert.ToInt32(rdr[33]),
                                DaysInMonth = Convert.ToInt32(rdr[34]),
                                Payout = Convert.ToInt32(rdr[35]),
                                EmpWorkeddays = Convert.ToInt32(rdr[36]),
                                LOPDays = Convert.ToInt32(rdr[37]),
                                LeaveEncashment = Convert.ToInt32(rdr[39]),
                                EffectiveWorkDays = Convert.ToInt32(rdr[40]),
                                BASIC = Convert.ToInt32(rdr[41]),
                                HRA = Convert.ToInt32(rdr[42]),
                                StatutoryBonus = Convert.ToInt32(rdr[43]),
                                LTA = Convert.ToInt32(rdr[44]),
                                TelephoneReimbursement = Convert.ToInt32(rdr[45]),
                                AttireAllowance = Convert.ToInt32(rdr[46]),
                                FuelReimbursment = Convert.ToInt32(rdr[47]),
                                ProjectAllowance = Convert.ToInt32(rdr[48]),
                                OtherEarnings = Convert.ToInt32(rdr[49]),
                                TotalEarnings = Convert.ToInt32(rdr[50]),
                                PF = Convert.ToInt32(rdr[51]),
                                ESIC = Convert.ToInt32(rdr[52]),
                                PT = Convert.ToInt32(rdr[53]),
                                IncomeTax = Convert.ToInt32(rdr[54]),
                                MedicalInsurance = Convert.ToInt32(rdr[55]),
                                OtherRecovery = Convert.ToInt32(rdr[56]),
                                TotalDeductions = Convert.ToInt32(rdr[57]),
                                NETPAY = Convert.ToInt32(rdr[58]),
                                Remarks = rdr[59].ToString(),
                                Status = rdr[60].ToString(),
                                GrossWages = Convert.ToInt32(rdr[61]),
                                BasicDA_Regular = Convert.ToInt32(rdr[62]),
                                BasicDA_Arrear = Convert.ToInt32(rdr[63]),
                                PF_Regular = Convert.ToInt32(rdr[64]),
                                PF_Arrear = Convert.ToInt32(rdr[65]),
                                VPF = Convert.ToInt32(rdr[66]),
                                PF2_Regular = Convert.ToInt32(rdr[67]),
                                PF2_Arrear = Convert.ToInt32(rdr[68]),
                                EPS_Regular = Convert.ToInt32(rdr[69]),
                                EPS_Arrear = Convert.ToInt32(rdr[70]),
                                Total = Convert.ToInt32(rdr[71]),
                                TaxRegime = rdr[72].ToString(),
                                TaxableIncome = Convert.ToInt32(rdr[73]),
                                IncomeTax2 = Convert.ToInt32(rdr[74]),
                                Surcharge = Convert.ToInt32(rdr[75]),
                                cess = Convert.ToInt32(rdr[76]),
                                TotalTax = Convert.ToInt32(rdr[77]),
                                ESICGross = Convert.ToInt32(rdr[78]),
                                EmployeesContribution = Convert.ToInt32(rdr[79])
                            };
                            listEmp.Add(ed);
                        }

                    }
                }
            }
        }

        [HttpGet]
        [Route("PFStatement")]
        public ActionResult PFStatement()
        {
            // DataTable dtProduct = ToDataTable<EmployeeData>(listEmp);
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.AddRange(new DataColumn[18] {
                    new DataColumn("Sl No",typeof(int)),
                    new DataColumn("Employee Number",typeof(string)),
                    new DataColumn("Name",typeof(Int64)),
                    new DataColumn("Date Of Joining",typeof(string)),
                    new DataColumn("Date of Leaving",typeof(int)),
                    new DataColumn("PF No",typeof(int)),
                    new DataColumn("UAN No",typeof(int)),
                    new DataColumn("Gross Wages",typeof(int)),
                    new DataColumn("Regular",typeof(int)),
                    new DataColumn("Arrear",typeof(int)),
                    new DataColumn("Regular",typeof(int)),
                    new DataColumn("Arrear",typeof(int)), new DataColumn("Regular",typeof(int)), new DataColumn("Arrear",typeof(int)), new DataColumn("Regular",typeof(int)), new DataColumn("Arrear",typeof(int)), new DataColumn("UAN No",typeof(int)), new DataColumn("UAN No",typeof(int))});

                var listESI = listEmp.Where(a => !string.IsNullOrEmpty(a.InsuranceNo));

                int TotalESI = 0;
                int TotalEmplContri = 0;
                int i = 2;
                foreach (var item in listESI)
                {
                    dt.Rows.Add(item.SerialNumber,
                           item.Code,
                           item.InsuranceNo,
                           item.Name,
                           item.EmpWorkeddays - item.LOPDays, // == "NULL" ? "" : item.agentResults,
                           item.ESICGross,
                           item.EmployeesContribution
                        );
                    TotalESI = TotalESI + item.ESICGross;
                    TotalEmplContri = TotalEmplContri + item.EmployeesContribution;
                    i++;
                }
                dt.Rows.Add(null, null, null, "Grand Total", null, TotalESI, TotalEmplContri);
                using (XLWorkbook workBook = new XLWorkbook())
                {
                    workBook.Worksheets.Add(dt, "ESI");
                    // workBook.Table. = false;
                    workBook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workBook.Style.Font.Bold = true;
                    var ws = workBook.Worksheet(1);
                    ws.Columns().AdjustToContents();
                    var rngHeaders = ws.Range("A" + i + ":G" + i);
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

        [HttpGet]
        [Route("ESICStatement")]
        public ActionResult ESICStatement()
        {
            try
            {
                DataTable dt = new DataTable();
                dt.Columns.AddRange(new DataColumn[7] {
                    new DataColumn("SI No",typeof(int)),
                    new DataColumn("Employee Number",typeof(string)),
                    new DataColumn("Insurance Number",typeof(Int64)),
                    new DataColumn("Name of Insured Person",typeof(string)),
                    new DataColumn("Days Worked",typeof(int)),
                    new DataColumn("ESI Gross",typeof(int)),
                    new DataColumn("Employee's Contribution",typeof(int))
                });

                var listESI = listEmp.Where(a => !string.IsNullOrEmpty(a.InsuranceNo));

                int TotalESI = 0;
                int TotalEmplContri = 0;
                int i = 2;
                foreach (var item in listESI)
                {
                    dt.Rows.Add(item.SerialNumber,
                           item.Code,
                           item.InsuranceNo,
                           item.Name,
                           item.EmpWorkeddays - item.LOPDays, // == "NULL" ? "" : item.agentResults,
                           item.ESICGross,
                           item.EmployeesContribution
                        );
                    TotalESI = TotalESI + item.ESICGross;
                    TotalEmplContri = TotalEmplContri + item.EmployeesContribution;
                    i++;
                }
                dt.Rows.Add(null, null, null, "Grand Total", null, TotalESI, TotalEmplContri);
                using (XLWorkbook workBook = new XLWorkbook())
                {
                    workBook.Worksheets.Add(dt, "ESI");
                    // workBook.Table. = false;
                    workBook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workBook.Style.Font.Bold = true;
                    var ws = workBook.Worksheet(1);
                    ws.Columns().AdjustToContents();
                    var rngHeaders = ws.Range("A" + i + ":G" + i);
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

        [HttpGet]
        [Route("GetAllEmployeeSalary")]
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
                    WebSettings = { DefaultEncoding = "utf-8", UserStyleSheet = Path.Combine(Directory.GetCurrentDirectory(), "assets", "styles.css") },
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

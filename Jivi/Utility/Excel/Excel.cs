using ExcelDataReader;
using HRMS.Model;
using HRReports.Model;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace HRReports.Utility.Excel
{
    public static class Excel
    {


        public static ESICDataTable GetPFDataTable(List<Employee> listEmployees)
        {
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[18] {
                    new DataColumn("SI No",typeof(int)),
                    new DataColumn("Employee Number",typeof(string)),
                    new DataColumn("Name",typeof(string)),
                    new DataColumn("Date Of Joining",typeof(string)),
                    new DataColumn("Date Of Leaving",typeof(string)),
                    new DataColumn("PF Number",typeof(string)),
                    new DataColumn("UAN Number",typeof(string)),
                    new DataColumn("Gross Wadges",typeof(int)),
                    new DataColumn("BasicDA Regular",typeof(int)),
                    new DataColumn("BasicDA Arrear",typeof(int)),
                    new DataColumn("PF Regular",typeof(int)),
                    new DataColumn("PF Arrear",typeof(int)),
                    new DataColumn("VPF",typeof(int)),
                    new DataColumn("ER PF Regular",typeof(int)),
                    new DataColumn("ER PF Arrear",typeof(int)),
                    new DataColumn("EPS Regular",typeof(int)),
                    new DataColumn("EPS Arrear",typeof(int)),
                    new DataColumn("Total",typeof(int))
                });

            int TotalESI = 0;
            int TotalEmplContri = 0;
            int i = 2;
            foreach (var item in listEmployees)
            {
                dt.Rows.Add(
                    item.SerialNumber,
                       item.EmplId,
                       item.Name,
                       item.DOJ,
                       item.DOL,
                       item.PFNo,
                       item.UANNo,
                       item.GrossWages,
                       item.BasicDA.Regular,
                       item.BasicDA.Arrear,
                       item.PF1.Regular,
                       item.PF1.Arrear,
                       item.VPF,
                       item.PF2.Regular,
                       item.PF2.Arrear,
                       item.EPS.Regular,
                       item.EPS.Arrear,
                       item.BasicDA.Regular + item.BasicDA.Arrear + item.PF1.Regular + item.PF1.Arrear + item.VPF + item.PF2.Regular + item.PF2.Arrear + item.EPS.Regular + item.EPS.Arrear
                    );
                TotalESI = TotalESI + item.ESICGross;
                TotalEmplContri = TotalEmplContri + item.EmployeesContribution;
                i++;
            }
            dt.Rows.Add(null, null, null, "Grand Total", null, TotalESI, TotalEmplContri);
            return new ESICDataTable()
            {
                DataTable = dt,
                RowToBeBold = i
            };
        }



        public static ESICDataTable GetESICDataTable(List<Employee> listEmployees)
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

            var listESI = listEmployees.Where(a => !string.IsNullOrEmpty(a.InsuranceNo));

            int TotalESI = 0;
            int TotalEmplContri = 0;
            int i = 2;
            foreach (var item in listESI)
            {
                dt.Rows.Add(item.SerialNumber,
                       item.EmplId,
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
            return new ESICDataTable()
            {
                DataTable = dt,
                RowToBeBold = i
            };
        }

        public static List<Employee> ParseAllEmployees(IFormFile file)
        {
            List<Employee> listEmpl = new List<Employee>();

            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = new MemoryStream())
            {
                file.CopyTo(stream);
                stream.Position = 0;
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var conf = new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = a => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };

                    DataSet dataSet = reader.AsDataSet(conf);
                    DataRowCollection row = dataSet.Tables["Active"].Rows;
                    List<object> rowDataList = null;
                    List<object> allRowsList = new List<object>();
                    string Month = string.Empty;
                    int i = 0;
                    foreach (DataRow rdr in row)
                    {
                        rowDataList = rdr.ItemArray.ToList(); //list of each rows
                        allRowsList.Add(rowDataList); //adding the above list of each row to another list
                        if (i == 0)
                        {
                            Month = rdr[0]?.ToString()?.Split(":")[1].Trim().Replace("-", "");
                        }
                        if (rdr[0] != null && int.TryParse(rdr[0].ToString(), out int slNo))
                        {
                            Employee ed = new Employee()
                            {
                                Id = Guid.NewGuid().ToString(),
                                SerialNumber = slNo,
                                Month = Month,
                                EmplId = rdr[1].ToString(),
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

                                BasicDA = new RA() { Regular = Convert.ToInt32(rdr[62]), Arrear = Convert.ToInt32(rdr[63]) },
                                PF1 = new RA() { Arrear = Convert.ToInt32(rdr[65]), Regular = Convert.ToInt32(rdr[64]) },

                                VPF = Convert.ToInt32(rdr[66]),

                                PF2 = new RA() { Arrear = Convert.ToInt32(rdr[68]), Regular = Convert.ToInt32(rdr[67]) },

                                EPS = new RA() { Arrear = Convert.ToInt32(rdr[70]), Regular = Convert.ToInt32(rdr[69]) },

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
                            listEmpl.Add(ed);
                        }
                        i++;
                    }
                }
            }
            return listEmpl;
        }
    }
}

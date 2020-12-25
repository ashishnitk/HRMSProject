using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace HRMS.Model
{
    public class Employee
    {
        [JsonProperty(PropertyName = "id")]
        public string Id { get; set; }
        public string Month { get; set; }
        public int SerialNumber { get; set; }
        public string EmplId { get; set; }
        public string Name { get; set; }
        public string DOJ { get; set; }
        public string DOL { get; set; }
        public string DOB { get; set; }
        public string BankAcNo { get; set; }
        public string Bankname { get; set; }
        public string IFSCCode { get; set; }
        public string PANNo { get; set; }
        public string PFNo { get; set; }
        public string UANNo { get; set; }
        public string InsuranceNo { get; set; }
        public string Location { get; set; }
        public string PTLocation { get; set; }
        public string Designation { get; set; }
        public string Department { get; set; }
        public int OriginalCTC { get; set; }
        public int CTCPA { get; set; }
        public int CTCPM { get; set; }
        public int MasterBasic { get; set; }
        public int MasterHRA { get; set; }
        public int MasterStatutoryBonus { get; set; }
        public int MasterLTA { get; set; }
        public int MasterTelephoneReimbursement { get; set; }
        public int MasterAttireAllowance { get; set; }
        public int MasterFuelReimbursment { get; set; }
        public int ERPF { get; set; }
        public int ERESIC { get; set; }
        public int MasterProjectAllowance { get; set; }
        public int TotalCTC { get; set; }
        public int CTCPA2 { get; set; }
        public int CalculatedGross { get; set; }
        public int DaysInMonth { get; set; }
        public int Payout { get; set; }
        public int EmpWorkeddays { get; set; }
        public int LOPDays { get; set; }
        public int LeaveEncashment { get; set; }
        public int EffectiveWorkDays { get; set; }
        public int BASIC { get; set; }
        public int HRA { get; set; }
        public int StatutoryBonus { get; set; }
        public int LTA { get; set; }
        public int TelephoneReimbursement { get; set; }
        public int AttireAllowance { get; set; }
        public int FuelReimbursment { get; set; }
        public int ProjectAllowance { get; set; }
        public int OtherEarnings { get; set; }
        public int TotalEarnings { get; set; }
        public int PF { get; set; }
        public int ESIC { get; set; }
        public int PT { get; set; }
        public int IncomeTax { get; set; }
        public int MedicalInsurance { get; set; }
        public int OtherRecovery { get; set; }
        public int TotalDeductions { get; set; }
        public int NETPAY { get; set; }
        public string Remarks { get; set; }
        public string Status { get; set; }
        public int GrossWages { get; set; }
        public RA BasicDA { get; set; }
        public RA PF1 { get; set; }
        public int VPF { get; set; }
        public RA PF2 { get; set; }

        public RA EPS { get; set; }

        public int Total { get; set; }
        public string TaxRegime { get; set; }
        public int TaxableIncome { get; set; }
        public int IncomeTax2 { get; set; }
        public int Surcharge { get; set; }
        public int cess { get; set; }
        public int TotalTax { get; set; }
        public int ESICGross { get; set; }
        public int EmployeesContribution { get; set; }

    }

    public class RA
    {
        public int Regular { get; set; }
        public int Arrear { get; set; }
    }

}

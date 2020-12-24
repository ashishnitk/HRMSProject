using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace HRMS.Model
{
    public class ESIReportModel
    {
        public string SerialNumber { get; set; }
        public string EmployeeNo { get; set; }
        public string InsuranceNo { get; set; }
        public string NameOfInsuredPerson { get; set; }
        public string DaysWorked { get; set; }
        public string ESIGross { get; set; }
        public string EmployeesContribution { get; set; }
    }
}

using HRMS.Controllers;
using HRMS.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HRMS.Utility
{
    public static class PDF
    {
        public static string GetSalariesHTML(List<Employee> listEmployee)
        {
            var sb = new StringBuilder();
            sb.Append(@"
                        <html>
                            <head>
.header {
    text-align: center;
    color: green;
    padding-bottom: 35px;
}

table {
    width: 80%;
    border-collapse: collapse;
}

td, th {
    border: 1px solid gray;
    padding: 15px;
    font-size: 22px;
    text-align: center;
}

table th {
    background-color: green;
    color: white;
}

                            </head>
                            <body>
                                <div class='header'><h1>This is the generated Salary report!!!</h1></div>
                                <table align='center'>
                                    <tr>
                                        <th>Name</th>
                                        <th>Department</th>
                                        <th>Designation</th>
                                        <th>CalculatedGross</th>
                                    </tr>");
            foreach (var emp in listEmployee)
            {
                sb.AppendFormat(@"<tr>
                                    <td>{0}</td>
                                    <td>{1}</td>
                                    <td>{2}</td>
                                    <td>{3}</td>
                                  </tr>", emp.Name, emp.Department, emp.Designation, emp.CalculatedGross);
            }
            sb.Append(@"
                                </table>
                            </body>
                        </html>");
            return sb.ToString();
        }


        public static string PaySlip(Employee emp)
        {
            var sb = new StringBuilder();
            sb.Append(@"<!DOCTYPE html>
<html>
<head>
    <style type='text / css'>
        table, th, td {
            border: 1px solid black;
                border - collapse: collapse;
            }

        .title {
            border: none;
            }

            th, td {
            padding: 5px;
            }
    </style >
</head > ");


            sb.AppendFormat(@"<body>
    <table>
        <tbody>
            <tr>
                <td>
                    <table class='title' style='border:hidden'>
                        <tbody style='border:hidden'>
                            <tr style='border:hidden'>
                                <td>Name</td>
                                <td>{0} {1}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>Join Date:</td>
                                <td>{2}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>Designation:</td>
                                <td>{3}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>Department:</td>
                                <td>{4}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>Location:</td>
                                <td>{5}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>Effective Work Days:</td>
                                <td>{6}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>Days In Month:</td>
                                <td>{7}</td>
                            </tr>
                        </tbody>
                    </table>
                </td>
                <td>
                    <table style='border:hidden'>
                        <tbody style='border:hidden'>
                            <tr style='border:hidden' >
                                <td>Bank Name:</td>
                                <td>{8}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>Bank Account No.:</td>
                                <td>{9}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>PF No.:</td>
                                <td>{10}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>PF UAN:</td>
                                <td>{11}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>ESI No.:</td>
                                <td>{12}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>PAN No.:</td>
                                <td>{13}</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>LOP:</td>
                                <td>{14}</td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>



            <tr>
                <td>
                    <table style='border:hidden'>
                        <tbody>
                            <tr style='font-weight:800'>
                                <td>Earnings</td>
                                <td>Full</td>
                                <td>Actual</td>
                            </tr>
                            <tr>
                                <td>BASIC</td>
                                <td>89,100.00</td>
                                <td>67,544.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>HRA</td>
                                <td>35,640.00</td>
                                <td>27,017.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>TELEPHONE ALLOWANCE</td>
                                <td>1,500.00</td>
                                <td>1,137.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>LTA</td>
                                <td>7,425.00</td>
                                <td>5,629.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>ATTIRE ALLOWANCE</td>
                                <td>3,000.00</td>
                                <td>2,274.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>FUEL REIMBURSEMENT</td>
                                <td>10,000.00</td>
                                <td>7,581.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>PROJECT ALLOWANCE</td>
                                <td>31,535.00</td>
                                <td>23,906.00</td>
                            </tr>
                        </tbody>
                    </table>
                </td>
                <td>
                    <table style='border:hidden'>
                        <tbody style='border:hidden'>
                            <tr style='font-weight:800'>
                                <td>Deductions</td>
                                <td>Actual</td>
                            </tr>
                            <tr>
                                <td>PF</td>
                                <td>1,800.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>PROF TAX</td>
                                <td>200.00</td>
                            </tr>
                            <tr style='border:hidden'>
                                <td>INCOME TAX</td>
                                <td>10,000.00</td>
                            </tr>
                        </tbody>
                    </table>
                </td>
            </tr>
            <tr>
                <td>
                </td>
                <td>
                    This cell contains a table:

                </td>
            </tr>
        </tbody>
    </table>

</body>
</html>", emp.Name, emp.EmplId,
emp.DOJ,
emp.Designation,
emp.Department,
emp.Location,
emp.EffectiveWorkDays,
emp.DaysInMonth,
emp.Bankname,
emp.BankAcNo,
emp.PFNo,
emp.UANNo,
emp.ESIC,
emp.PANNo,
emp.LOPDays);



            return sb.ToString();
        }

    }
}

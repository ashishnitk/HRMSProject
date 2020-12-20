using Jivi.Controllers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jivi.Utility
{
    public static class TemplateGenerator
    {
        public static string GetHTMLString()
        {
            var employees = ReportController.listEmp;
            var sb = new StringBuilder();
            sb.Append(@"
                        <html>
                            <head>
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
            foreach (var emp in employees)
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
    }
}

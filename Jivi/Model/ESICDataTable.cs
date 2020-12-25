using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;

namespace HRReports.Model
{
    public class ESICDataTable
    {
        public DataTable DataTable { get; set; }
        public int RowToBeBold { get; set; }
    }
}

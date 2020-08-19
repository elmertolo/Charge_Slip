using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Charge_Slip.Models
{
    class BankModel
    {
        public string IpAddress { get; set; }
        public string tableName { get; set; }
        public string ChkType { get; set; }
        public Int64 LastNo { get; set; }
        public string BranchCode { get; set; }
        public string Department { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Charge_Slip.Models
{
    class OrderModel
    {
        public string BankCode { get; set; }
  
        public string BankName { get; set; }
        public string ChkType { get; set; }
        public string Department { get; set; }
        public string BranchCode { get; set; }
        public string ChequeName { get; set; }
        public int Quantity { get; set; }
        public string StartingSerial { get; set; }
        public string EndingSerial { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.Entity
{
    public class OJDT
    {
        public int index{ get; set; }
        public string RefDate { get; set; }
        public string AcctCode{ get; set; }
        public string AcctName { get; set; }
        public string PrcCode { get; set; }
        public string PrcName { get; set; }
        public string Project { get; set; }
        public double Debit { get; set; }
        public double Credit { get; set; }
        public double TotalML { get; set; }
        public double TotalMS{ get; set; }
    }
}

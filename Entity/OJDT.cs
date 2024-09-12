using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddOnRclsGastos.Entity
{
    public class OJDT
    {
        [JsonProperty("ReferenceDate")] public string ReferenceDate { get; set; }
        [JsonProperty("TransactionCode")] public string TransactionCode { get; set; }
        [JsonProperty("TaxDate")] public string TaxDate { get; set; }
        [JsonProperty("Memo")] public string LineMemo { get; set; }
        [JsonProperty("JournalEntryLines")] public List<JDT1> Details { get; set; }
    }

    public class JDT1
    {
        public string TransId { get; set; }
        public string Line_ID { get; set; }
        [JsonProperty("BPLID")] public int? BPLID { get; set; }
        [JsonProperty("AccountCode")] public string AccountCode { get; set; }
        public string FormatCode { get; set; }
        public string AccountName { get; set; }
        [JsonProperty("CostingCode")] public string CostingCode { get; set; }
        public string CostingCodeName { get; set; }
        [JsonProperty("CostingCode2")] public string CostingCode2 { get; set; }
        public string CostingCode2Name { get; set; }
        [JsonProperty("CostingCode3")] public string CostingCode3 { get; set; }
        public string CostingCode3Name { get; set; }
        [JsonProperty("CostingCode4")] public string CostingCode4 { get; set; }
         public string CostingCode4Name { get; set; }
        [JsonProperty("CostingCode5")] public string CostingCode5 { get; set; }
        public string CostingCode5Name { get; set; }
        [JsonProperty("ProjectCode")] public string ProjectCode { get; set; }
        [JsonProperty("Debit")] public double Debit { get; set; }
        [JsonProperty("Credit")] public double Credit { get; set; }
        [JsonProperty("FCDebit")] public double? FCDebit { get; set; }
        [JsonProperty("FCCredit")] public double? FCCredit { get; set; }
        [JsonProperty("DebitSys")] public double DebitSys { get; set; }
        [JsonProperty("CreditSys")] public double CreditSys { get; set; }
        public double? TotalML { get; set; }
        public double? TotalME { get; set; }
        public double? TotalMS { get; set; }
        [JsonProperty("FCCurrency")] public string FCCurrency { get; set; }
    }
}

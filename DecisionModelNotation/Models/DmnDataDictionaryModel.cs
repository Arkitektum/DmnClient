using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace DecisionModelNotation.Models
{
    public class DmnDataDictionaryModel
    {
        public string FilNavn  { get; set; }
        public string DmnId { get; set; }
        public string DmnNavn { get; set; }
        public string VariabelId { get; set; }
        public string VariabelNavn { get; set; }
        public string VariabelType { get; set; }
        public string Type { get; set; }
        public string Kilde { get; set; }
    }
}

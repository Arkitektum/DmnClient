using System;
using System.Collections.Generic;
using System.Text;

namespace DecisionModelNotation.Models
{
   public class DataDictionaryModel
    {
        public BpmnDataDictionaryModel BpmnData { get; set; }
        public DmnDataDictionaryModel DmnData { get; set; }
    }
}

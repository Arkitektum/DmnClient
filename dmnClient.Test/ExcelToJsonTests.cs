using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Xml.Serialization;
using DecisionModelNotation;
using DecisionModelNotation.Shema;
using Excel;
using FluentAssertions;
using Newtonsoft.Json;
using OfficeOpenXml;
using Xunit;

namespace dmnClient.Test
{
   public class ExcelToJsonTests
    {
        [Fact(DisplayName = "Integration Test")]
        public void Test1()
        {
            var resourcePath = "dmnClient.Test.TestData.DataDictionaryFromModels.xlsx";
            var assembly = Assembly.GetExecutingAssembly();
            ExcelPackage ep;
            var name = string.Empty;
            using (Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath))
            {

                ep = new ExcelPackage(resourceAsStream);
            }

            var jsonList = new Dictionary<string,string>();

            var dmnTEK = new ExcelServices().ExcelToJsonObject(ep,"DmnTEK");
            if (!string.IsNullOrEmpty(dmnTEK))
                jsonList.Add("JsonDmn2TEK", dmnTEK);

            var variables = new ExcelServices().ExcelToJsonObject(ep, "Variables");
            if (!string.IsNullOrEmpty(variables))
                jsonList.Add("JsonDmnVariablesNames", variables);

            var dmnPlusVariables = new ExcelServices().ExcelToJsonObject(ep, "Dmn+Variables");
            if (!string.IsNullOrEmpty(dmnPlusVariables))
                jsonList.Add("JsonTable2Variables", dmnPlusVariables);

            foreach (var json in jsonList)
            {
                var jsonFile = string.Concat(@"c:\temp\", json.Key,".json");
                File.WriteAllText(jsonFile,json.Value,Encoding.UTF8);
                File.Exists(jsonFile).Should().BeTrue();
            }


            //var dmnFile = string.Concat(@"c:\temp\", name, "_", ".dmn");
            //XmlSerializer xs = new XmlSerializer(typeof(tDefinitions));
            //TextWriter tw = new StreamWriter(dmnFile);
            //xs.Serialize(tw, newDmn);

            //File.Exists(dmnFile).Should().BeTrue();

        }
    }
}

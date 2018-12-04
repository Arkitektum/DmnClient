using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using DecisionModelNotation;
using DecisionModelNotation.Shema;
using DecisionModelNotation.Models;
using Excel;
using FluentAssertions;
using Newtonsoft.Json;
using OfficeOpenXml;
using Xunit;

namespace dmnClient.Test
{
    public class DmnToDataDictionaryTests
    {
        [Fact(DisplayName = "INTEGRATION test - Create Data Dictionary From DMN & BPMN models")]
        public void Test1()
        {

            var file = "dmnTest1.dmn";
            var file2 = "dmnTest2.dmn";
            var file3 = "BpmnTest01.bpmn";
            var dmns = new List<tDefinitions>();
            var filePath1 = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file);
            var filePath2 = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file2);
            var bpmn = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file3);

            XDocument bpmnXml = XDocument.Load(bpmn);

            using (Stream dmnStream = File.Open(filePath1, FileMode.Open))
            {
                dmns.Add(new DmnServices().DeserializeStreamDmnFile(dmnStream));
            }
            using (Stream dmnStream = File.Open(filePath2, FileMode.Open))
            {
                dmns.Add(new DmnServices().DeserializeStreamDmnFile(dmnStream));
            }

            var dmnDataDictionaryModels = new List<DmnDataDictionaryModel>();


            var excelPkg = new ExcelPackage();
            foreach (var tdefinitions in dmns)
            {
                var Items = tdefinitions.Items;
                var decision = Items.Where(t => t.GetType() == typeof(tDecision));

                foreach (tDecision tdecision in decision)
                {
                    tDecisionTable decisionTable = null;
                    try
                    {
                        DmnServices.GetDecisionsVariables(tdecision, Path.GetFileNameWithoutExtension(filePath1),
                            ref dmnDataDictionaryModels);
                    }
                    catch
                    {
                        //
                    }
                }

            }

            var bpmnDataDictionary = new List<BpmnDataDictionaryModel>();
            DmnServices.GetDmnInfoFromBpmnModel(bpmnXml, ref bpmnDataDictionary);

            List<DataDictionaryModel> dataDictionaryModels = new List<DataDictionaryModel>();
            foreach (var dmnData in dmnDataDictionaryModels)
            {
                var submodel = new BpmnDataDictionaryModel();
                try
                {

                    var value = dmnData.GetType();
                    var property = value.GetProperty("DmnId");
                    String name = (String)(property.GetValue(dmnData, null));

                    submodel = bpmnDataDictionary.Single(b => b.DmnId == "sdsds");

                }
                catch 
                {
                }
                dataDictionaryModels.Add(new DataDictionaryModel()
                {
                    BpmnData = submodel,
                    DmnData = dmnData
                });

            }

            ExcelWorksheet wsSheet = excelPkg.Workbook.Worksheets.Add("DmnTEK");
            var dmnIds = dmnDataDictionaryModels.GroupBy(x => x.DmnId).Select(y => y.First());
            var objectPropertyNames = new[] { "DmnId", "DmnNavn", "TekKapitel", "TekLedd", "TekTabell", "TekForskriften", "TekWebLink" };
            ExcelServices.CreateDmnExcelTableDataDictionary(dmnIds, wsSheet, "dmnTek", objectPropertyNames);

            ExcelWorksheet wsSheet1 = excelPkg.Workbook.Worksheets.Add("Variables");
            var dmnVariablesIds = dmnDataDictionaryModels.GroupBy(x => x.VariabelId).Select(y => y.First());
            var dmnVariablesIdstPropertyNames = new[] { "VariabelId", "VariabelNavn", "VariabelBeskrivelse" };
            ExcelServices.CreateDmnExcelTableDataDictionary(dmnVariablesIds, wsSheet1, "Variables", dmnVariablesIdstPropertyNames);

            ExcelWorksheet wsSheet2 = excelPkg.Workbook.Worksheets.Add("Dmn+Variables");
            var objectPropertyNames1 = new[] { "DmnId", "VariabelId", "Type" };
            ExcelServices.CreateDmnExcelTableDataDictionary(dmnDataDictionaryModels, wsSheet2, "Dmn+Variables", objectPropertyNames1);


            ExcelWorksheet wsSheet3 = excelPkg.Workbook.Worksheets.Add("summary");
            var summaryPropertyNames = new[] { "DmnData.FilNavn", "BpmnData.BpmnId", "DmnData.DmnId", "DmnData.VariabelId", "DmnData.VariabelType", "DmnData.Type", "DmnData.Kilde" };
            ExcelServices.CreateSummaryExcelTableDataDictionary(dataDictionaryModels, wsSheet3, "summary", summaryPropertyNames);

            var path = string.Concat(@"c:\temp\");
            Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, string.Concat("dataDictionary", ".xlsx"));
            excelPkg?.SaveAs(new FileInfo(filePath));

            File.Exists(filePath).Should().BeTrue();

        }

        [Fact(DisplayName = "deserialize bpmn")]
        public void Test2()
        {

            var file3 = "BpmnTest01.bpmn";
            var dmns = new List<tDefinitions>();
            var bpmn = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file3);
            XmlDocument doc = new XmlDocument();
            doc.Load(bpmn);

            XDocument doc1 = XDocument.Load(bpmn); // Or whatever
            var ns = doc.NamespaceURI;
            var matchingElements2 = doc1.Descendants()
                .Where(x => x.Name.ToString().Contains("businessRuleTask"));

            //var bpmnDataDictionary = new DmnServices().GetDmnInfoFromBpmnModel(doc1);

            foreach (XElement element in matchingElements2)
            {
                //var noko = element.Attribute(@"camunda:resultVariable");
                var nn1 = element.Attributes();
                var Name = element.Attribute("name");
                var resultVariable = element.Attributes().Single(a => a.Name.ToString().Contains("resultVariable"))?.Value;
                var decisionRef = element.Attributes().Single(a => a.Name.ToString().Contains("decisionRef"))?.Value;
            }
        }
    }
}

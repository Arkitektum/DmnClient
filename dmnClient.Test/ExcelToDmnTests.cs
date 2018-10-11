using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Xml.Serialization;
using DecisionModelNotation;
using DecisionModelNotation.Shema;
using Excel;
using Xunit;
using FluentAssertions;
using OfficeOpenXml;

namespace dmnClient.Test
{
    public class ExcelToDmnTests
    {
        [Fact(DisplayName = "Integration Test")]
        public void Test1()
        {
            var resourcePath = "dmnClient.Test.TestData.dmnTest1.xlsx";
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);

            var file = "dmnTest1.xlsx";
            var savedFilePath = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file);
            //var savedFilePath = @"C:\Users\MatiasGonzalezTognon\Dropbox (Arkitektum AS)\Ark_prosjekter\DIBK\DigiTek\Brannteknisk prosjektering\DMN\Trinn1\ExcelFiles\13_OverflateKledning.xlsx";
            var name = Path.GetFileNameWithoutExtension(savedFilePath);
            var fi = new FileInfo(savedFilePath);
            ExcelPackage ep = new ExcelPackage(new FileInfo(savedFilePath));

            ExcelPackage ep1 = new ExcelPackage(resourceAsStream);

            //ExcelPackage ep = new ExcelPackage(new FileInfo(resourcePath));
            ExcelWorksheet workSheet = ep.Workbook.Worksheets.FirstOrDefault();
            var table = ep1.Workbook.Worksheets.FirstOrDefault().Tables.FirstOrDefault();

            var columnsDictionary = ExcelServices.GetColumnRagngeInLeters(table, 4, 3);
            columnsDictionary.TryGetValue("outputsIndex", out var outputsIndex);
            columnsDictionary.TryGetValue("inputsIndex", out var inputsIndex);


            var outputsRulesDictionary = new ExcelServices().GetRulesFromExcel(workSheet, outputsIndex, true);
            var inputsRulesDictionary = new ExcelServices().GetRulesFromExcel(workSheet, inputsIndex, true);
            var outputsRulesTypes = new ExcelServices().GetRulesTypes(workSheet, outputsIndex, true);
            var inputsRulesTypes = new ExcelServices().GetRulesTypes(workSheet, inputsIndex, true);
            var inputsDictionary = new ExcelServices().GetExcelHeaderName(workSheet, inputsIndex, true);
            var outputsDictionary = new ExcelServices().GetExcelHeaderName(workSheet, outputsIndex, true);
            var dmnInfo = new ExcelServices().GetDmnInfo(workSheet).FirstOrDefault();
            var dmnName = dmnInfo.Value;
            var dmnId = dmnInfo.Key;


            var newDmn = new DmnV1Builder()
                .AddDefinitionsInfo("Excel2Dmn_" + DateTime.Now.ToString("dd-mm-yy"), name)
                .AddDecision(dmnId, dmnName, "decisionTable")
                .AddInputsToDecisionTable(inputsDictionary, inputsRulesTypes)
                .AddOutputsToDecisionTable(outputsDictionary, outputsRulesTypes)
                .AddDecisionRules(inputsRulesDictionary, outputsRulesDictionary)
                .Build();

            var dmnFile = string.Concat(@"c:\temp\", name, "_", ".dmn");
            XmlSerializer xs = new XmlSerializer(typeof(tDefinitions));
            TextWriter tw = new StreamWriter(dmnFile);
            xs.Serialize(tw, newDmn);

            File.Exists(dmnFile).Should().BeTrue();

        }
    }
}

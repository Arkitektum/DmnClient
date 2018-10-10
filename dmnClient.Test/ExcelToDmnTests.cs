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
        [Fact]
        public void Test1()
        {
            var resourcePath = "dmnClient.Test.TestData.dmnTest1.xlsx";
            var assembly = Assembly.GetExecutingAssembly();
            Stream resourceAsStream = assembly.GetManifestResourceStream(resourcePath);

            var file = "dmnTest1.xlsx";
            var savedFilePath = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file);
            var name = Path.GetFileNameWithoutExtension(savedFilePath);
            var fi = new FileInfo(savedFilePath);
           
            ExcelPackage ep = new ExcelPackage(new FileInfo(savedFilePath));
            ExcelPackage ep1 = new ExcelPackage(resourceAsStream);

            //ExcelPackage ep = new ExcelPackage(new FileInfo(resourcePath));
            ExcelWorksheet workSheet = ep.Workbook.Worksheets.FirstOrDefault();
            ExcelWorksheet workSheet1 = ep1.Workbook.Worksheets.FirstOrDefault();
            var inputIndex = new[] { "B", "D", "C", "E" };
            var outputIndex = new[] { "F", "G", "H" };

            var outputsRulesDictionary = new ExcelServices().GetRulesFromExcel(workSheet, outputIndex, true);
            var inputsRulesDictionary = new ExcelServices().GetRulesFromExcel(workSheet, inputIndex, true);
            var outputsRulesTypes = new ExcelServices().GetRulesTypes(workSheet, outputIndex, true);
            var inputsRulesTypes = new ExcelServices().GetRulesTypes(workSheet, inputIndex, true);
            var inputsDictionary = new ExcelServices().GetExcelHeaderName(workSheet, inputIndex, true);
            var outputsDictionary = new ExcelServices().GetExcelHeaderName(workSheet, outputIndex, true);


            var newDmn = new DmnV1Builder()
                .AddDefinitionsInfo("Excel2Dmn_" + DateTime.Now.ToString("dd-mm-yy"), name)
                .AddDecision("KonsekvensBrannklassifisering", "Konsekvens brannklassifisering", "decisionTable")
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

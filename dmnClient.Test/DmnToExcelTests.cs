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
    public class DmnToExcelTests
    {

        [Fact(DisplayName = "Integration Test")]
        public void Test1()
        {
            var file = "dmnTest1.dmn";
            string ifcDataFile = Path.Combine(Directory.GetCurrentDirectory() + @"..\..\..\..\TestData\", file);
            tDefinitions dmn;
            using (Stream dmnStream = File.Open(ifcDataFile, FileMode.Open))
            {
                dmn = new DmnServices().DeserializeStreamDmnFile(dmnStream);
            }

            var Items = dmn.Items;
            var decision = Items.Where(t => t.GetType() == typeof(tDecision));

            var excelPkg = new ExcelPackage();
            foreach (var tdecision in decision)
            {
                tDecisionTable decisionTable = null;
                try
                {
                    var dt = ((tDecision)tdecision).Item;
                    decisionTable = (tDecisionTable)Convert.ChangeType(dt, typeof(tDecisionTable));
                    ExcelWorksheet wsSheet = excelPkg.Workbook.Worksheets.Add(tdecision.id);
                    //Add Table Title
                    ExcelServices.AddTableTitle(tdecision.name, wsSheet, decisionTable, tdecision.id);
                    // Add "input" and "output" headet to Excel table
                    ExcelServices.AddTableInputOutputTitle(wsSheet, decisionTable);
                    //Add DMN Table to excel Sheet
                    ExcelServices.CreateExcelTableFromDecisionTable(decisionTable, wsSheet, tdecision.id);

                }
                catch
                {
                    //
                }
            }

            var filename = Path.GetFileNameWithoutExtension(ifcDataFile);
            var path = string.Concat(@"c:\temp\");
            Directory.CreateDirectory(path);
            var filePath = Path.Combine(path, string.Concat(filename, ".xlsx"));
            excelPkg?.SaveAs(new FileInfo(filePath));

            File.Exists(filePath).Should().BeTrue();

        }
    }

}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Xml.Linq;
using System.Xml.Serialization;
using DecisionModelNotation;
using DecisionModelNotation.Models;
using DecisionModelNotation.Shema;
using Excel;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace dmnClient.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DmnController : ControllerBase
    {
        [HttpGet]
        public ActionResult<IEnumerable<string>> Get()
        {
            return new string[] { "value1", "value2" };
        }

        [HttpPost, Route("excelToDmn2")]
        public ActionResult<string> PostFromBody([FromBody] string value)
        {
            var noko = value;
            var nn = new JsonResult("");
            return nn;
        }


        [HttpPost, Route("excelToDmn/{inputs}/{outputs}/{haveId}")]
        public ActionResult<string> Post(string inputs, string outputs, bool haveId)
        {


            var httpRequest = HttpContext.Request;
            var responsDictionary = new Dictionary<string, string>();
            HttpResponseMessage response = null;

            if (httpRequest.Form.Files.Count != 1)
                return BadRequest(new Dictionary<string, string>() { { "Error:", "Not file fount" } });
            var file = httpRequest.Form.Files[0];
            var file1 = httpRequest.Form.Files.FirstOrDefault();
            ExcelPackage ep = null;
            ExcelWorksheet workSheet = null;
            ExcelTable table = null;
            if (file != null)
            {
                try
                {
                    //Open Excel file
                    using (Stream excelFile = file.OpenReadStream())
                    {
                        ep = new ExcelPackage(excelFile);
                    }

                    workSheet = ep.Workbook.Worksheets.FirstOrDefault();
                    table = workSheet.Tables.FirstOrDefault();

                }
                catch (Exception e)
                {
                    return BadRequest(new Dictionary<string, string>() { { "Error:", "Can't Open Excel File" } });
                }

                if (table != null)
                {
                    Dictionary<int, Dictionary<string, object>> outputsRulesFromExcel = null;
                    Dictionary<int, Dictionary<string, object>> inputsRulsFromExcel = null;
                    Dictionary<int, string> annotationsRulesDictionary = null;

                    Dictionary<int, string> outputsRulesTypes = null;
                    Dictionary<int, string> inputsRulesTypes = null;
                    Dictionary<string, Dictionary<string, string>> inputsDictionary = null;
                    Dictionary<string, Dictionary<string, string>> outputsDictionary = null;
                    var dmnName = string.Empty;
                    var dmnId = string.Empty;

                    var columnsDictionary = GetTablesIndex(table, inputs, outputs);
                    string[] outputsIndex = null;
                    string[] inputsIndex = null;
                    if (columnsDictionary != null)
                    {
                        columnsDictionary.TryGetValue("outputsIndex", out outputsIndex);
                        columnsDictionary.TryGetValue("inputsIndex", out inputsIndex);
                    }

                    if (inputsIndex == null && outputsIndex == null)
                    {
                        return BadRequest(new Dictionary<string, string>() { { "Error", "Can't get inputs/output rows" } });
                    }

                    try
                    {
                        outputsRulesFromExcel = new ExcelServices().GetRulesFromExcel(workSheet, outputsIndex, haveId);
                        inputsRulsFromExcel = new ExcelServices().GetRulesFromExcel(workSheet, inputsIndex, haveId);
                        annotationsRulesDictionary =new ExcelServices().GetAnnotationsRulesFromExcel(workSheet, inputsIndex, outputsIndex,haveId);
                        outputsRulesTypes = new ExcelServices().GetRulesTypes(workSheet, outputsIndex, haveId);
                        inputsRulesTypes = new ExcelServices().GetRulesTypes(workSheet, inputsIndex, haveId);
                        inputsDictionary = new ExcelServices().GetExcelHeaderName(workSheet, inputsIndex, haveId);
                        outputsDictionary = new ExcelServices().GetExcelHeaderName(workSheet, outputsIndex, haveId);
                        var dmnInfo = new ExcelServices().GetDmnInfo(workSheet).FirstOrDefault();
                        dmnName = dmnInfo.Value;
                        dmnId = dmnInfo.Key;
                        if (!outputsRulesFromExcel.Any() || !inputsRulsFromExcel.Any() ||
                            !outputsRulesTypes.Any() || !inputsRulesTypes.Any() || !inputsDictionary.Any()
                            || !outputsDictionary.Any())
                        {
                            return BadRequest(new Dictionary<string, string>() { { "Error:", "Wrong information to create DMN from Excel" } });
                        }
                    }
                    catch (Exception)
                    {
                        return BadRequest(new Dictionary<string, string>() { { "Error:", "Can't Get Excel info" } });
                    }

                    var filename = Path.GetFileNameWithoutExtension(file.FileName);
                    var newDmn = new DmnV1Builder()
                        .AddDefinitionsInfo("Excel2Dmn_" + DateTime.Now.ToString("dd-mm-yy"), filename)
                        .AddDecision(dmnId, dmnName, "decisionTable")
                        .AddInputsToDecisionTable(inputsDictionary, inputsRulesTypes)
                        .AddOutputsToDecisionTable(outputsDictionary, outputsRulesTypes)
                        .AddDecisionRules(inputsRulsFromExcel, outputsRulesFromExcel, annotationsRulesDictionary)
                        .Build();
                    // Save DMN 
                    try
                    {

                        var path = Path.Combine(@"C:\", "ExcelToDmn");

                        Directory.CreateDirectory(path);


                        //var dmnFile = string.Concat(path, filename, "_Exc2Dmn", ".dmn");
                        XmlSerializer xs = new XmlSerializer(typeof(DecisionModelNotation.Shema.tDefinitions));
                        var combine = Path.Combine(path, string.Concat(filename, ".dmn"));
                        using (TextWriter tw = new StreamWriter(combine))
                        {
                            xs.Serialize(tw, newDmn);
                        }

                        return Ok(new Dictionary<string, string>() { { filename + ".dmn", "Created" }, { "Path", combine } });
                    }
                    catch (Exception e)
                    {
                        return BadRequest(new Dictionary<string, string>() { { filename + ".dmn", "Can't be safe" } });

                    }
                }
                return BadRequest(new Dictionary<string, string>() { { file.FileName, "Excel file don't have table" } });

            }
            return Ok(responsDictionary);
        }

        [HttpPost, Route("dmnToExcel")]
        public IActionResult Pos()
        {
            var httpRequest = HttpContext.Request;
            HttpResponseMessage response = null;

            string okResponsText = null;
            var httpFiles = httpRequest.Form.Files;
            var okDictionary = new Dictionary<string, string>();
            var ErrorDictionary = new Dictionary<string, string>();

            if (httpFiles == null && !httpFiles.Any())
                return NotFound("Can't find any file");

            for (var i = 0; i < httpFiles.Count; i++)
            {
                string errorResponsText = null;
                string errorTemp = string.Empty;
                var file = httpFiles[i];
                tDefinitions dmn = null;

                //Deserialize DMN file
                if (file != null)
                {
                    using (Stream dmnfile = httpFiles[i].OpenReadStream())
                    {
                        dmn = new DmnServices().DeserializeStreamDmnFile(dmnfile);
                    }
                }
                if (dmn == null)
                {
                    ErrorDictionary.Add(file.FileName, "Can't validate Shema");
                    continue;
                }
                // check if DMN have desicion table

                var items = dmn.Items;
                var decision = items.Where(t => t.GetType() == typeof(tDecision));
                var tDrgElements = decision as tDRGElement[] ?? decision.ToArray();
                if (!tDrgElements.Any())
                {
                    ErrorDictionary.Add(file.FileName, "Dmn file have non decision");
                    continue;
                }

                // create Excel Package
                ExcelPackage excelPkg = null;
                try
                {
                    excelPkg = new ExcelPackage();
                    foreach (var tdecision in tDrgElements)
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
                            ErrorDictionary.Add(file.FileName,string.Concat("Dmn: ",tdecision.name, " Can't be create"));
                        }
                    }
                }
                catch
                {
                    ErrorDictionary.Add(file.FileName, "Can't create Excel file");
                    continue;
                }
                // Save Excel Package
                try
                {
                    var filename = Path.GetFileNameWithoutExtension(file.FileName);
                    var path = Path.Combine(@"C:\", "DmnToExcel");

                    Directory.CreateDirectory(path);
                    excelPkg.SaveAs(new FileInfo(Path.Combine(path, string.Concat(filename, ".xlsx"))));

                    okDictionary.Add(file.FileName, "Created in:"+path);
                }
                catch
                {

                    ErrorDictionary.Add(file.FileName, "Can't be saved");
                }

            }

            if (ErrorDictionary.Any())
            {
                if (okDictionary.Any())
                {
                    List<Dictionary<string, string>> dictionaries = new List<Dictionary<string, string>>();
                    dictionaries.Add(okDictionary);
                    dictionaries.Add(ErrorDictionary);
                    var result = dictionaries.SelectMany(dict => dict)
                        .ToLookup(pair => pair.Key, pair => pair.Value)
                        .ToDictionary(group => group.Key, group => group.First());
                    return Ok(result);
                }
                return BadRequest(ErrorDictionary);

            }
            return Ok(okDictionary);
        }

        //Data DIctionary
        [HttpPost, Route("GetModelDataDictionaryToExcel")]
        public IActionResult Post()
        {
            var httpRequest = HttpContext.Request;
            HttpResponseMessage response = null;

            string okResponsText = null;
            var httpFiles = httpRequest.Form.Files;
            var okDictionary = new Dictionary<string, string>();
            var ErrorDictionary = new Dictionary<string, string>();
            var dmnDataDictionaryModels = new List<DmnDataDictionaryModel>();
            var bpmnDataDictionaryModels = new List<BpmnDataDictionaryModel>();
            var dataDictionaryModels = new List<DataDictionaryModel>();

            if (httpFiles == null && !httpFiles.Any())
                return NotFound("Can't find any file");

            for (var i = 0; i < httpFiles.Count; i++)
            {
                string errorResponsText = null;
                string errorTemp = string.Empty;
                var file = httpFiles[i];
                tDefinitions dmn = null;
                var fileExtention = Path.GetExtension(file.FileName);
                if (fileExtention==".dmn")
                {
                    //Deserialize DMN file
                    if (file != null)
                    {
                        using (Stream dmnfile = httpFiles[i].OpenReadStream())
                        {
                            dmn = new DmnServices().DeserializeStreamDmnFile(dmnfile);
                        }
                    }
                    if (dmn == null)
                    {
                        ErrorDictionary.Add(file.FileName, "Can't validate Shema");
                        continue;
                    }
                    // check if DMN have desicion table

                    var items = dmn.Items;
                    var decision = items.Where(t => t.GetType() == typeof(tDecision));
                    var tDrgElements = decision as tDRGElement[] ?? decision.ToArray();
                    if (!tDrgElements.Any())
                    {
                        ErrorDictionary.Add(file.FileName, "Dmn file have non decision");
                        continue;
                    }
                    foreach (tDecision tdecision in decision)
                    {
                        tDecisionTable decisionTable = null;
                        try
                        {
                            DmnServices.GetDecisionsVariables(tdecision, Path.GetFileNameWithoutExtension(file.FileName),
                                ref dmnDataDictionaryModels);
                        }
                        catch
                        {
                            ErrorDictionary.Add(file.FileName, "Can't add serialize info from DMN");
                        }
                    }
                }

                if (fileExtention ==".bpmn")
                {
                    XDocument bpmnXml = null;
                    try
                    {
                        using (Stream dmnfile = httpFiles[i].OpenReadStream())
                        {
                            bpmnXml = XDocument.Load(dmnfile);
                        }
                    }
                    catch
                    {
                        ErrorDictionary.Add(file.FileName, "Can't add serialize bpmn to xml");
                    }

                    if (bpmnXml!= null)
                    {
                        try
                        {
                            DmnServices.GetDmnInfoFromBpmnModel(bpmnXml, ref bpmnDataDictionaryModels);
                        }
                        catch
                        {
                            ErrorDictionary.Add(file.FileName, "Can't add serialize bpmn to Data Model Dictionary");
                        }
                    }
                }
            }

            foreach (var dmnDataInfo in dmnDataDictionaryModels)
            { 
                var submodel = new BpmnDataDictionaryModel();
                try
                {
                    submodel = bpmnDataDictionaryModels.Single(b => b.DmnId == dmnDataInfo.DmnId);
                }
                catch
                {
                }
                dataDictionaryModels.Add(new DataDictionaryModel()
                {
                    BpmnData = submodel,
                    DmnData = dmnDataInfo
                });

            }




            // create Excel Package
            ExcelPackage excelPkg = null;
            var fileName = "DataDictionaryFromModels";
            try
            {
                excelPkg = new ExcelPackage();
                ExcelWorksheet wsSheet = excelPkg.Workbook.Worksheets.Add("DmnTEK");
                var dmnIds = dmnDataDictionaryModels.GroupBy(x => x.DmnId).Select(y => y.First());
                var objectPropertyNames = new[] { "DmnId", "DmnNavn", "TekKapitel", "TekLedd", "TekTabell", "TekForskriften", "TekWebLink" };
                ExcelServices.CreateDmnExcelTableDataDictionary(dmnIds, wsSheet, "dmnTek", objectPropertyNames);

                ExcelWorksheet wsSheet1 = excelPkg.Workbook.Worksheets.Add("Variables");
                var dmnVariablesIds = dmnDataDictionaryModels.GroupBy(x => x.VariabelId).Select(y => y.First());
                var dmnVariablesIdstPropertyNames = new[] { "VariabelId", "VariabelNavn", "VariabelBeskrivelse" };
                ExcelServices.CreateDmnExcelTableDataDictionary(dmnVariablesIds, wsSheet1, "Variables", dmnVariablesIdstPropertyNames);

                ExcelWorksheet wsSheet2 = excelPkg.Workbook.Worksheets.Add("Dmn+Variables");
                var objectPropertyNames1 = new[] { "DmnId", "VariabelId", "VariabelType" };
                ExcelServices.CreateDmnExcelTableDataDictionary(dmnVariablesIds, wsSheet2, "Dmn+Variables", objectPropertyNames1);

                ExcelWorksheet wsSheet3 = excelPkg.Workbook.Worksheets.Add("summary");
                var summaryPropertyNames = new[] { "DmnData.FilNavn", "BpmnData.BpmnId", "DmnData.DmnId", "DmnData.VariabelId", "DmnData.VariabelType", "DmnData.Type", "DmnData.Kilde" };
                ExcelServices.CreateSummaryExcelTableDataDictionary(dataDictionaryModels, wsSheet3, "summary", summaryPropertyNames);
            }
            catch
            {
                ErrorDictionary.Add("Error","Can't create Excel file");
            }
            // Save Excel Package
            try
            {
                var path = Path.Combine(@"C:\", "DmnToExcel");
                Directory.CreateDirectory(path);
                excelPkg.SaveAs(new FileInfo(Path.Combine(path, string.Concat(fileName, ".xlsx"))));
                okDictionary.Add(fileName, "Created in:" + path);
            }
            catch
            {
                ErrorDictionary.Add(fileName, "Can't be saved");
            }

            if (ErrorDictionary.Any())
            {
                if (okDictionary.Any())
                {
                    List<Dictionary<string, string>> dictionaries = new List<Dictionary<string, string>>();
                    dictionaries.Add(okDictionary);
                    dictionaries.Add(ErrorDictionary);
                    var result = dictionaries.SelectMany(dict => dict)
                        .ToLookup(pair => pair.Key, pair => pair.Value)
                        .ToDictionary(group => group.Key, group => group.First());
                    return Ok(result);
                }
                return BadRequest(ErrorDictionary);

            }
            return Ok(okDictionary);
        }





        private static Dictionary<string, string[]> GetTablesIndex(ExcelTable table, string inputs, string outputs)
        {
            var inputsIsNumber = int.TryParse(inputs, out var inputsColumnsCount);
            var outputsIsNumber = int.TryParse(outputs, out var outputsColumnsCount);
            Dictionary<string, string[]> dictionary = null;

            if (inputsIsNumber && outputsIsNumber)
            {
                dictionary = ExcelServices.GetColumnRagngeInLeters(table, inputsColumnsCount, outputsColumnsCount);
            }
            else
            {
                var inputsIndex = inputs.Split(',').ToArray();
                var outputsIndex = outputs.Split(',').ToArray();
                if (inputsIndex.Any() && outputsIndex.Any())
                {
                    dictionary = new Dictionary<string, string[]>()
                    {
                        {"inputsIndex",inputsIndex },
                        {"outputsIndex",outputsIndex }
                    };
                }

            }
            return dictionary;
        }
    }
}
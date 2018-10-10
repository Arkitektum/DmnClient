﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Xml.Serialization;
using DecisionModelNotation;
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
            var inputsIndex = inputs.Split(',').ToArray();
            var outputsIndex = outputs.Split(',').ToArray();

            if (!inputsIndex.Any())
                return BadRequest(new Dictionary<string, string>() { { "Inpust Row", "not specify" } });
            if (!inputsIndex.Any())
                return BadRequest(new Dictionary<string, string>() { { "Outpust Row", "not specify" } });

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
                    return BadRequest(new Dictionary<string, string>() { { "Error:", "Can Not Open Excel File" } });
                }

                if (table != null)
                {
                    Dictionary<int, Dictionary<string, object>> outputsRulesFromExcel = null;
                    Dictionary<int, Dictionary<string, object>> inputsRulsFromExcel = null;
                    Dictionary<int, string> outputsRulesTypes = null;
                    Dictionary<int, string> inputsRulesTypes = null;
                    Dictionary<string, Dictionary<string, string>> inputsDictionary = null;
                    Dictionary<string, Dictionary<string, string>> outputsDictionary = null;
                    var dmnName = string.Empty;
                    var dmnId = string.Empty;

                    try
                    {
                        outputsRulesFromExcel = new ExcelServices().GetRulesFromExcel(workSheet, outputsIndex, haveId);
                        inputsRulsFromExcel = new ExcelServices().GetRulesFromExcel(workSheet, inputsIndex, haveId);
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
                        return BadRequest(new Dictionary<string, string>() { { "Error:", "Can not Get Excel info" } });
                    }

                    var filename = Path.GetFileNameWithoutExtension(file.FileName);
                    var newDmn = new DmnV1Builder()
                        .AddDefinitionsInfo("Excel2Dmn_" + DateTime.Now.ToString("dd-mm-yy"), filename)
                        .AddDecision(dmnId, dmnName, "decisionTable")
                        .AddInputsToDecisionTable(inputsDictionary, inputsRulesTypes)
                        .AddOutputsToDecisionTable(outputsDictionary, outputsRulesTypes)
                        .AddDecisionRules(inputsRulsFromExcel, outputsRulesFromExcel)
                        .Build();
                    // Save DMN 
                    try
                    {

                        var path = Path.Combine(@"C:\", "ExcelToDmn");

                        Directory.CreateDirectory(path);


                        //var dmnFile = string.Concat(path, filename, "_Exc2Dmn", ".dmn");
                        XmlSerializer xs = new XmlSerializer(typeof(DecisionModelNotation.Shema.tDefinitions));
                        var combine = Path.Combine(path, string.Concat(filename,".dmn"));
                        using (TextWriter tw = new StreamWriter(combine))
                        {
                            xs.Serialize(tw, newDmn);
                        }

                        return Ok(new Dictionary<string, string>() { { filename+".dmn","Created" } });
                    }
                    catch (Exception e)
                    {
                        return BadRequest(new Dictionary<string, string>() { { filename + ".dmn", "Can not be safe" } });

                    }
                }
                    return BadRequest(new Dictionary<string, string>() { { file.FileName, "Excel don't have Table" } });

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
                return NotFound("Kan ikke finne noen fil");

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
                    ErrorDictionary.Add(file.FileName, " Can not validate Shema");
                    continue;
                }
                // check if DMN have desicion table

                var items = dmn.Items;
                var decision = items.Where(t => t.GetType() == typeof(tDecision));
                var tDrgElements = decision as tDRGElement[] ?? decision.ToArray();
                if (!tDrgElements.Any())
                {
                    ErrorDictionary.Add(file.FileName, " Dmn file have now decision");
                    continue;
                }

                // create Excel Package
                ExcelPackage excelPkg = null;
                try
                {
                    excelPkg = new ExcelPackage();
                    foreach (var tdecision in tDrgElements)
                    {
                        try
                        {
                            var dt = ((tDecision)tdecision).Item;
                            var decisionTable = (tDecisionTable)Convert.ChangeType(dt, typeof(tDecisionTable));
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
                            ErrorDictionary.Add(file.FileName, " DMN Can not be create");
                        }
                    }
                }
                catch
                {
                    ErrorDictionary.Add(file.FileName, " Can not create Excel");
                    continue;
                }
                // Save Excel Package
                try
                {
                    var filename = Path.GetFileNameWithoutExtension(file.FileName);
                    var path = Path.Combine(@"C:\", "DmnToExcel");

                    Directory.CreateDirectory(path);
                    excelPkg.SaveAs(new FileInfo(Path.Combine(path, string.Concat(filename, ".xlsx"))));

                    var temp = string.Concat("* ", file.FileName, ":", " created");
                    okResponsText = string.IsNullOrEmpty(okResponsText)
                        ? temp
                        : okResponsText + temp;
                    okDictionary.Add(file.FileName, "Created");
                }
                catch
                {

                    ErrorDictionary.Add(file.FileName, " Can not be saved");
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


    }
}
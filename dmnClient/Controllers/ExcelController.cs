using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Xml.Serialization;
using DecisionModelNotation;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Table;

namespace dmnClient.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
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
            var responseText = string.Empty;
            HttpResponseMessage response = null;
            string[] inputsIndex = new string[] { };
            string[] outputsIndex = new string[] { };
            try
            {
                inputsIndex = inputs.Split(',').ToArray();
                outputsIndex = outputs.Split(',').ToArray();
            }
            catch (Exception)
            {
                var ErrorText = string.Concat("*", "Can not serialize inputs or outputs");
                responseText = string.IsNullOrEmpty(responseText)
                    ? ErrorText
                    : responseText + ErrorText;
            }

            if (httpRequest.Form.Files.Count != 1)
            {

                var ErrorText = httpRequest.Form.Files.Count == 0 ? "Not file fount" : "Only one Excel File can be upload";
                responseText = string.IsNullOrEmpty(responseText)
                    ? ErrorText
                    : responseText + ErrorText;
            }
            else
            {
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
                        var ErrorText = string.Concat("*", "Can Not Open Excel File");
                        responseText = string.IsNullOrEmpty(responseText)
                            ? ErrorText
                            : responseText + ErrorText;
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
                            outputsRulesFromExcel = new Excel.ExcelServices().GetRulesFromExcel(workSheet, outputsIndex, haveId);
                            inputsRulsFromExcel = new Excel.ExcelServices().GetRulesFromExcel(workSheet, inputsIndex, haveId);
                            outputsRulesTypes = new Excel.ExcelServices().GetRulesTypes(workSheet, outputsIndex, haveId);
                            inputsRulesTypes = new Excel.ExcelServices().GetRulesTypes(workSheet, inputsIndex, haveId);
                            inputsDictionary = new Excel.ExcelServices().GetExcelHeaderName(workSheet, inputsIndex, haveId);
                            outputsDictionary = new Excel.ExcelServices().GetExcelHeaderName(workSheet, outputsIndex, haveId);
                            var dmnInfo = new Excel.ExcelServices().GetDmnInfo(workSheet).FirstOrDefault();
                            dmnName = dmnInfo.Value;
                            dmnId = dmnInfo.Key;
                            if (!outputsRulesFromExcel.Any() || !inputsRulsFromExcel.Any() ||
                                !outputsRulesTypes.Any() || !inputsRulesTypes.Any() || !inputsDictionary.Any()
                                || !outputsDictionary.Any())
                            {
                                var ErrorText = string.Concat("*", "Wrong information to get DMN from Excel");
                                responseText = string.IsNullOrEmpty(responseText)
                                    ? ErrorText
                                    : responseText + ErrorText;
                            }
                        }
                        catch (Exception)
                        {
                            var ErrorText = string.Concat("*", "Can Not Get Excel info");
                            responseText = string.IsNullOrEmpty(responseText)
                                ? ErrorText
                                : responseText + ErrorText;
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
                            var combine = Path.Combine(path, string.Concat(filename, "_Exc2Dmn", ".dmn"));
                            using (TextWriter tw = new StreamWriter(combine))
                            {
                                xs.Serialize(tw, newDmn);
                            }


                            responseText = "dmn Created";
                        }
                        catch (Exception e)
                        {
                            var ErrorText = string.Concat("*", file.FileName, ":", " Can not save dmn");
                            responseText = string.IsNullOrEmpty(responseText)
                                ? ErrorText
                                : responseText + ErrorText;
                        }
                    }
                    else
                    {
                        var ErrorText = string.Concat("*", file.FileName, ":", " Excel don't have Table");
                        responseText = string.IsNullOrEmpty(responseText)
                            ? ErrorText
                            : responseText + ErrorText;
                    }
                }
            }

            // Create error response with all the files errors
            if (!string.IsNullOrEmpty(responseText))
            {
                //response = Request.CreateErrorResponse(HttpStatusCode.PartialContent, responseText);
                return BadRequest(responseText);
            }

            //response = Request.CreateResponse(HttpStatusCode.Accepted, "", "application/json");
            return Ok(new JsonResult(responseText));
        }

        [HttpPost, Route("dmnToExcel")]
        public IActionResult Pos()
        {
            var httpRequest = HttpContext.Request;
            HttpResponseMessage response = null;
            string responseText = null;
            var httpFiles = httpRequest.Form.Files;

            if (httpFiles.Count == 0)
                return NotFound("Kan ikke finne noen fil");

            for (var i = 0; i < httpFiles.Count; i++)
            {
                responseText = i.ToString();
            }


            return Ok(responseText);
        }


    }
}
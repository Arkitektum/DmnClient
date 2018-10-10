﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml;
using System.Xml.Serialization;
using DecisionModelNotation.Shema;

namespace DecisionModelNotation
{
    public class DmnServices
    {
        public tDefinitions SerializeDictionariesToDmn(Dictionary<string, object> outputsDictionary, Dictionary<int, object> rulesDictionary, string fileName)
        {
            var tDecisionTable = new tDecisionTable();
            tDecisionTable.input = new tInputClause[] { };


            var tExpression = tDecisionTable;
            var tdecision = new tDecision();
            tdecision.Item = tExpression;
            var tDefinitions = new tDefinitions();
            tDefinitions.id = fileName;
            tDefinitions.Items = new tDRGElement[] { tdecision };

            return tDefinitions;
        }

        public tInputClause[] CreateDmnInputs(Dictionary<string, Dictionary<string, string>> inputsDictionary)
        {
            var inputs = new List<tInputClause>();
            foreach (var entry in inputsDictionary)
            {
                foreach (var inputValue in entry.Value)
                {
                    var inputId = inputValue.Value;
                    var inputLable = inputValue.Key;

                    if (String.IsNullOrEmpty(inputId))
                    {
                        inputId = Regex.Replace(inputLable, @"\s+", "");
                        inputId = inputId.Length <= 10 ? inputId : inputId.Substring(0, 10);
                    }
                    var input = new tInputClause()
                    {
                        id = inputId,
                        label = inputLable,

                        inputExpression = new tLiteralExpression()
                        {
                            id = String.Concat("exp_", inputId),
                            label = String.Concat("label_", inputId),
                            Item = inputId
                        }
                    };

                    inputs.Add(input);
                }
            }
            return inputs.ToArray();
        }
        public tOutputClause[] CreateDmnOutpus(Dictionary<string, Dictionary<string, string>> inputsDictionary)
        {
            var outputs = new List<tOutputClause>();

            foreach (var entry in inputsDictionary)
            {
                foreach (var output in entry.Value)
                {
                    var outputId = output.Value;
                    var outputLabel = output.Key;

                    if (String.IsNullOrEmpty(outputId))
                    {
                        outputId = Regex.Replace(outputLabel, @"\s+", "");
                        outputId = outputId.Length <= 10 ? outputId : outputId.Substring(0, 10);
                    }
                    var dmnOutputClause = new tOutputClause()
                    {
                        id = String.Concat(outputId, "_Id"),
                        label = outputLabel,
                        name = outputId,
                    };
                    outputs.Add(dmnOutputClause);
                }
            }
            return outputs.ToArray();
        }
        public tDefinitions DeserializeStreamDmnFile(Stream fileStream)
        {
            tDefinitions resultinMessage;
            try
            {
                var serializer = new XmlSerializer(typeof(tDefinitions));
                resultinMessage = (tDefinitions)serializer.Deserialize(new XmlTextReader(fileStream));
            }
            catch
            {

                resultinMessage = null;
            }
            return resultinMessage;
        }

        public static string GetComparisonNumber(string cellValue)
        {
            var regex = Regex.Match(cellValue, @"^[<,>][=]?\s?(?<number>\d+[\.]?(\d+)?)$");
            return regex.Success ? regex.Groups["number"].Value : null;
        }

        public static string[] GetRangeNumber(string cellValue)
        {
            var regex = Regex.Match(cellValue, @"^[\[,\],]\s?(?<range1>\d+(\.\d+)?).{2}?(?<range2>\d+(\.\d+)?)[\[,\]]$");
            return regex.Success ? new[] { regex.Groups["range1"].Value, regex.Groups["range2"].Value } : null;
        }
    }
}
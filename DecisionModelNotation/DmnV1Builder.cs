using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml;
using DecisionModelNotation.Shema;

namespace DecisionModelNotation
{
    public class DmnV1Builder
    {
        private readonly tDefinitions _dmn;

        public DmnV1Builder()
        {
            _dmn = new tDefinitions();
        }

        public tDefinitions Build()
        {
            return _dmn;
        }

        public DmnV1Builder AddDefinitionsInfo(string id, string name)
        {
            _dmn.id = id;
            _dmn.name = name;
            return this;
        }
        public DmnV1Builder AddDecision(string decisionId, string decisionName, string decisitonTableId, string hitPolicy = "FIRST")
        {

            _dmn.Items = new[]
            {
                new tDecision()
                {
                    id = decisionId,
                    name = decisionName,
                    Item = new tDecisionTable()
                    {
                        id = decisitonTableId,
                        hitPolicy = GethitPolicy(hitPolicy)
                    },
                },
            };
            return this;
        }

        private tHitPolicy GethitPolicy(string hitPolicy)
        {
            switch (hitPolicy)
            {
                case "FIRST":return tHitPolicy.FIRST;
                case "UNIQUE":return tHitPolicy.UNIQUE;
                case "PRIORITY":return tHitPolicy.PRIORITY;
                case "ANY":return tHitPolicy.ANY;
                case "COLLECT": return tHitPolicy.UNIQUE;
                default: return tHitPolicy.FIRST;
            }
        }

        public DmnV1Builder AddInputsToDecisionTable(Dictionary<string, Dictionary<string, string>> inputsDictionary, Dictionary<int, string> inputsTypes)
        {
            tDecisionTable decisionTable = null;
            if (_dmn.Items == null)
            {
                _dmn.Items = new[] {new tDecision()
                {
                    Item = new tDecisionTable(),
                },

                };
            }
            var decision = _dmn.Items.FirstOrDefault(i => i is tDecision);
            decisionTable = (tDecisionTable)((tDecision)decision)?.Item;

            if (decisionTable != null)
                decisionTable.input = CreateDmnInputs(inputsDictionary, inputsTypes);

            return this;
        }

        public DmnV1Builder AddOutputsToDecisionTable(Dictionary<string, Dictionary<string, string>> outputsDictionary, Dictionary<int, string> outputsType)
        {
            tDecisionTable decisionTable = null;
            if (_dmn.Items != null)
            {
                var decision = _dmn.Items.FirstOrDefault(i => i is tDecision);
                decisionTable = (tDecisionTable)((tDecision)decision)?.Item;
            }

            if (decisionTable != null)
                decisionTable.output = CreateDmnOutpus(outputsDictionary, outputsType);

            return this;
        }

        public DmnV1Builder AddDecisionRules(Dictionary<int, Dictionary<string, object>> inputsRulesDictionary, Dictionary<int, Dictionary<string, object>> outputsRulesDictionary)
        {
            var decision = _dmn.Items.FirstOrDefault(i => i is tDecision);
            var decisionTable = (tDecisionTable)((tDecision)decision)?.Item;


            var startKey = inputsRulesDictionary.FirstOrDefault().Key;
            var endKey = inputsRulesDictionary.Last().Key;

            var decisionRules = new List<tDecisionRule>();

            for (int i = startKey; i <= endKey; i++)
            {
                var inputRowValue = inputsRulesDictionary.TryGetValue(i, out Dictionary<string, object> inputsRowRuleDictionary);
                var outputRowValue = outputsRulesDictionary.TryGetValue(i, out Dictionary<string, object> outputsRowRuleDictionary);
                if (!inputRowValue || !outputRowValue)
                    break;
                var decisionRule = new tDecisionRule()
                {
                    id = string.Concat("rowRule_", i),
                    inputEntry =
                        CreateInputsRules(inputsRowRuleDictionary),
                    outputEntry =
                        CreateOutputsRules(outputsRowRuleDictionary)
                };
                decisionRules.Add(decisionRule);
            }

            decisionTable.rule = decisionRules.ToArray();
            return this;
        }

        private tLiteralExpression[] CreateOutputsRules(Dictionary<string, object> value)
        {
            var outputsRules = new List<tLiteralExpression>();
            foreach (var output in value)
            {
                var rule = new tLiteralExpression()
                {
                    id = string.Concat("LiteralExpression_", output.Key),
                };

                ////rule.Item = output.Value is double ? output.Value.ToString() : output.Value;
                //if (output.Value is Boolean || output.Value is double)
                //{
                //    rule.Item = output.Value.ToString();
                //}
                //else
                //{
                //    rule.Item = output.Value;
                //}
                rule.Item = GetValueParse(output.Value.ToString());
                outputsRules.Add(rule);
            }

            return outputsRules.ToArray();
        }

        private tUnaryTests[] CreateInputsRules(Dictionary<string, object> inputsRowRuleDictionary)
        {
            var inputsRules = new List<tUnaryTests>();
            foreach (var input in inputsRowRuleDictionary)
            {
                var value = string.IsNullOrEmpty(input.Value?.ToString()) ? "" : input.Value.ToString();

                var rule = new tUnaryTests()
                {
                    id = input.Key,
                    text = GetValueParse(value)
                };

                inputsRules.Add(rule);
            }
            return inputsRules.ToArray();
        }

        private string GetValueParse(string cellValue)
        {
            var regex = DmnServices.GetComparisonNumber(cellValue);
            var regex2 = DmnServices.GetRangeNumber(cellValue);

            if (int.TryParse(cellValue, out var intType)) return intType.ToString();
            if (long.TryParse(cellValue, out var longType)) return longType.ToString();
            if (double.TryParse(cellValue, out var doubleType)) return doubleType.ToString();
            if (bool.TryParse(cellValue, out var booleanType)) return booleanType.ToString().ToLower();
            var values = cellValue.Split(";");
            string newCellValue = cellValue;
            if (values != null && values.Any() && regex==null && regex2==null && !string.IsNullOrEmpty(newCellValue))
            {
                newCellValue = string.Empty;
                if (cellValue.StartsWith("\"")&& cellValue.EndsWith("\""))
                {
                    newCellValue = cellValue;
                }
                else
                {
                    for (int i = 0; i < values.Count(); i++)
                    {
                        newCellValue = i == 0 ? string.Concat("\"", values[i], "\"") : string.Concat(newCellValue, ",", "\"", values[i], "\"");
                    }
                }
            }

            return newCellValue;
        }


        private tInputClause[] CreateDmnInputs(Dictionary<string, Dictionary<string, string>> inputsDictionary, Dictionary<int, string> inputsTypes)
        {
            var inputs = new List<tInputClause>();
            int i = 0;
            foreach (var inputValue in inputsDictionary)
            {
                var input = inputValue.Value.FirstOrDefault();
                var inputId = variableId(input, out var inputLable);
                var haveType = inputsTypes.TryGetValue(i, out string type);
                var InputClause = new tInputClause()
                {
                    id = inputId,
                    label = inputLable,

                    inputExpression = new tLiteralExpression()
                    {
                        id = string.Concat("exp_", inputId),
                        label = string.Concat("label_", inputId),
                        Item = inputId,
                        typeRef = new XmlQualifiedName(type)
                    }
                };
                inputs.Add(InputClause);
                i++;
            }
            return inputs.ToArray();
        }

        private tOutputClause[] CreateDmnOutpus(Dictionary<string, Dictionary<string, string>> outputsDictionary, Dictionary<int, string> outputsType)
        {
            var outputs = new List<tOutputClause>();
            int i = 0;
            foreach (var entry in outputsDictionary)
            {
                var output = entry.Value.FirstOrDefault();

                var outputId = variableId(output, out var outputLabel);
                var haveType = outputsType.TryGetValue(i, out string type);

                var dmnOutputClause = new tOutputClause()
                {
                    id = string.Concat(outputId, "_Id"),
                    label = outputLabel,
                    name = outputId,
                    typeRef = new XmlQualifiedName(type)
                };
                outputs.Add(dmnOutputClause);
                i++;
            }
            return outputs.ToArray();
        }


        private static string variableId(KeyValuePair<string, string> variableValue, out string variableLable)
        {
            var inputId = variableValue.Value;
            variableLable = variableValue.Key;

            if (string.IsNullOrEmpty(inputId))
            {
                inputId = Regex.Replace(variableLable, @"\s+", "");
                inputId = inputId.Length <= 10 ? inputId : inputId.Substring(0, 10);
            }
            return inputId;
        }
    }
}
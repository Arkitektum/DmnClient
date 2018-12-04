using System;
using System.Collections.Generic;
using System.Drawing;
using System.Dynamic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DecisionModelNotation;
using DecisionModelNotation.Models;
using DecisionModelNotation.Shema;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;

namespace Excel
{
    public class ExcelServices
    {
        public Dictionary<string, Dictionary<string, string>> GetExcelHeaderName(ExcelWorksheet ws, string[] columnsIndexs, bool vaiableId = false)
        {

            var table = ws.Tables.First();
            var tableStartRow = table.Address.Start.Row;
            Dictionary<string, Dictionary<string, string>> dictionary = new Dictionary<string, Dictionary<string, string>>();

            var start = GetRange(columnsIndexs, out var end);
            if (!CheckColumnRange(table, start, end)) return dictionary;

            for (int col = start; col <= end; col++)
            {
                Dictionary<string, string> headerDictionary = new Dictionary<string, string>();
                var cellName = string.Concat(GetColumnName(col), tableStartRow);
                var cellValue = GetCellValue(ws, cellName);
                if (vaiableId)
                {
                    var idCellName = string.Concat(GetColumnName(col), tableStartRow + 1);
                    var cellIdValue = GetCellValue(ws, idCellName);
                    headerDictionary.Add(cellValue.ToString(), cellIdValue.ToString());
                }
                else
                {
                    headerDictionary.Add(cellValue.ToString(), string.Empty);
                }
                dictionary.Add(cellName, headerDictionary);
            }

            return dictionary;
        }

        public Dictionary<int, Dictionary<string, object>> GetRulesFromExcel(ExcelWorksheet ws, string[] columnsIndexs, bool variableId = false)
        {

            var table = ws.Tables.First();
            var tableStartRow = table.Address.Start.Row;
            var dictionary = new Dictionary<int, Dictionary<string, object>>();
            int startRow = variableId ? tableStartRow + 2 : tableStartRow + 1;

            var start = GetRange(columnsIndexs, out var end);
            if (!CheckColumnRange(table, start, end)) return dictionary;
            for (int i = startRow; i <= table.Address.End.Row; i++)
            {
                var valuesDictionary = new Dictionary<string, object>();
                for (int j = start; j <= end; j++)
                {
                    var cellName = string.Concat(GetColumnName(j), i);
                    var cellValue = ws.Cells[cellName].Value;
                    valuesDictionary.Add(cellName, cellValue);
                }
                dictionary.Add(i, valuesDictionary);
            }

            return dictionary;
        }
        public static void AddTableTitle(string tableName, ExcelWorksheet wsSheet1, tDecisionTable decisionTable, string tableId)
        {
            wsSheet1.Cells["B1"].Value = "DMN Navn:";
            wsSheet1.Cells["C1"].Value = tableName;

            wsSheet1.Cells["B2"].Value = "DMN id:";
            wsSheet1.Cells["C2"].Value = tableId;

            wsSheet1.Cells["B1"].Style.Font.Size = 12;
            wsSheet1.Cells["B1"].Style.Font.Bold = true;
            wsSheet1.Cells["B1"].Style.Font.Italic = true;
            wsSheet1.Cells["B2"].Style.Font.Size = 12;
            wsSheet1.Cells["B2"].Style.Font.Bold = true;
            wsSheet1.Cells["B2"].Style.Font.Italic = true;
        }
        public static void AddTableInputOutputTitle(ExcelWorksheet wsSheet, tDecisionTable decisionTable)
        {
            var totalInput = decisionTable.input.Count();
            var totalOutput = decisionTable.output.Count();
            const int stratRow = 4;
            const int stratColumn = 2;
            var endInputColumn = stratColumn + totalInput;
            var endColum = endInputColumn + totalOutput;

            //input
            using (ExcelRange rng = wsSheet.Cells[stratRow, stratColumn, stratRow, endInputColumn - 1])
            {
                InputOutputTitleFormat(rng, "Input");
            }

            //Output
            using (ExcelRange rng = wsSheet.Cells[stratRow, endInputColumn, stratRow, endColum])
            {
                InputOutputTitleFormat(rng, "Output");
            }

        }
        private static void InputOutputTitleFormat(ExcelRange excelRange, string text)
        {
            excelRange.Value = text;
            excelRange.Style.Font.Size = 12;
            excelRange.Style.Font.Bold = true;
            excelRange.Style.Font.Italic = true;
            excelRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            excelRange.Merge = true;
            excelRange.Style.Border.BorderAround(ExcelBorderStyle.Thin);

        }

        public static void CreateExcelTableFromDecisionTable(tDecisionTable decisionTable, ExcelWorksheet wsSheet, string tableName)
        {
            // palse Table in Excel
            const int stratRow = 5;
            const int stratColumn = 2;

            // Calculate size of the table
            var totalInput = decisionTable.input.Count();
            var totalOutput = decisionTable.output.Count();
            var totalRules = decisionTable.rule.Count();
            var endRow = stratRow + totalRules + 1;
            var endColum = stratColumn + totalInput + totalOutput;

            // Create Excel table Header
            using (ExcelRange rng = wsSheet.Cells[stratRow, stratColumn, endRow, endColum])
            {
                //Indirectly access ExcelTableCollection class
                ExcelTable table = wsSheet.Tables.Add(rng, tableName);
                var color = Color.FromArgb(250, 199, 111);
                //Set Columns position & name
                var i = 0;
                foreach (var inputClause in decisionTable.input)
                {
                    table.Columns[i].Name = inputClause.label;

                    //add input variable name
                    AddExcelCellByRowAndColumn(stratColumn + i, stratRow + 1,
                        inputClause.inputExpression.Item.ToString(), wsSheet, color);

                    i++;
                }

                foreach (var outputClause in decisionTable.output)
                {


                    table.Columns[i].Name = outputClause.label;

                    // Add Output variableId name
                    var variableId = outputClause.name ?? "";
                    AddExcelCellByRowAndColumn(stratColumn + i, stratRow + 1, variableId, wsSheet, color);

                    i++;
                }

                // Add empty cell for annotation
                table.Columns[i].Name = "Annotation";
                AddExcelCellByRowAndColumn(stratColumn + i, stratRow + 1, " ", wsSheet, color);
                //table.ShowHeader = false;
                //table.ShowFilter = true;
                //table.ShowTotal = true;
            }

            // Set Excel table content
            var inputColumn = stratColumn;
            var outputColumn = stratColumn + totalInput;
            var row = stratRow + 1;

            foreach (var rule in decisionTable.rule)
            {
                // input content
                row++;
                foreach (var tUnaryTestse in rule.inputEntry)
                {
                    AddExcelCellByRowAndColumn(inputColumn, row, tUnaryTestse.text, wsSheet);
                    inputColumn++;
                }
                inputColumn = stratColumn;

                // set Output result content
                foreach (var literalExpression in rule.outputEntry)
                {
                    AddExcelCellByRowAndColumn(outputColumn, row, literalExpression.Item.ToString(), wsSheet);
                    outputColumn++;
                }

                var annotationCellName = string.Concat(GetColumnName(endColum), row);
                using (ExcelRange rng = wsSheet.Cells[annotationCellName])
                {
                    rng.Value = rule.description;
                }

                outputColumn = stratColumn + totalInput;
            }

            //wsSheet1.Protection.IsProtected = false;
            //wsSheet1.Protection.AllowSelectLockedCells = false;
            wsSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
        }
        private static void AddExcelCellByRowAndColumn(int column, int row, string value, ExcelWorksheet wsSheet, Color? color = null)
        {
            var cellName = string.Concat(GetColumnName(column), row);
            using (ExcelRange rng1 = wsSheet.Cells[cellName])
            {
                if (color.HasValue)
                {
                    rng1.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng1.Style.Fill.BackgroundColor.SetColor(color.Value);
                }

                rng1.Value = value;
            }
        }

        //Excel to DMN
        private int GetRange(string[] columnIndexs, out int end)
        {
            var columnIndexsOrderBy = columnIndexs.OrderBy(d => d).ToArray();
            var start = GetColumnIndex(columnIndexsOrderBy[0]);
            end = GetColumnIndex(columnIndexsOrderBy.Last());
            return start;
        }

        public bool CheckColumnRange(ExcelTable table, int start, int end)
        {
            var tableStartColumn = table.Address.Start.Column;
            var tableEndColumn = table.Address.End.Column;
            return Enumerable.Range(tableStartColumn, tableEndColumn).Contains(start) &&
                   Enumerable.Range(tableStartColumn, tableEndColumn).Contains(end);
        }

        public int GetColumnIndex(string columnName)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            var index = letters.ToLower().IndexOf(columnName.ToLower());

            return index + 1;
        }

        public static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length - 1];

            return value;
        }

        private object GetCellValue(ExcelWorksheet ws, string cellName)
        {
            return ws.Cells[cellName].Value ?? string.Empty;
        }

        public Dictionary<int, string> GetRulesTypes(ExcelWorksheet ws, string[] columnsIndexs, bool variableId = false)
        {
            var table = ws.Tables.First();
            var tableStartRow = table.Address.Start.Row;
            var rulesTypes = new Dictionary<int, string>();
            int startRow = variableId ? tableStartRow + 2 : tableStartRow + 1;

            var startColumn = GetRange(columnsIndexs, out var endColumn);
            if (!CheckColumnRange(table, startColumn, endColumn)) return rulesTypes;
            int i = 0;
            for (int j = startColumn; j <= endColumn; j++)
            {
                var cellName = string.Concat(GetColumnName(j), startRow);
                var cellValue = ws.Cells[cellName].Value;
                var cellValueType = GetValueType(j, startRow, ws);
                rulesTypes.Add(i, cellValueType);
                i++;
            }

            return rulesTypes;
        }

        private string GetValueType(int columnIndex, int startRow, ExcelWorksheet ws)
        {
            var table = ws.Tables.First();
            string type = null;
            for (int i = startRow; i <= table.Address.End.Row; i++)
            {
                var cellName = string.Concat(GetColumnName(columnIndex), i);
                var cellValue = ws.Cells[cellName].Value;
                var typeTemp = ParseCellValue(cellValue);
                type = string.IsNullOrEmpty(type) ? typeTemp : type;
                type = GetType(type, typeTemp);
            }
            return type;
        }

        private static string GetType(string type, string typeTemp)
        {
            if (type != typeTemp && !string.IsNullOrEmpty(typeTemp))
            {
                if (type == "integer")
                {
                    switch (typeTemp)
                    {
                        case "double":
                            type = "double";
                            break;
                        case "long":
                            type = "long";
                            break;
                        default:
                            type = "string";
                            break;
                    }
                }

                if (type == "double" || type == "long")
                {
                    switch (typeTemp)
                    {
                        case "double":
                        case "long":
                        case "integer":
                            type = "double";
                            break;
                        case "string":
                        case "boolean":
                            type = "string";
                            break;
                    }
                }

                if (type == "boolean")
                    type = "string";
            }

            return type;
        }



        private string ParseCellValue(object cellValue)
        {
            //cellValue = cellValue.ToString().Trim();
            if (cellValue == null || string.IsNullOrEmpty(cellValue.ToString())) return string.Empty;
            //if (string.IsNullOrEmpty(cellValue.ToString()))return String.Empty;

            string cellValueString = DmnServices.GetComparisonNumber(cellValue.ToString()) ?? cellValue.ToString();

            var cellRangeNumber = DmnServices.GetRangeNumber(cellValue.ToString());
            if (cellRangeNumber != null)
            {
                var type1 = ParseCellValue(cellRangeNumber[0]);
                var type2 = ParseCellValue(cellRangeNumber[1]);

                if (type1 != type2)
                    return GetType(type1, type2);
                cellValueString = cellRangeNumber[0];
            }


            if (int.TryParse(cellValueString, out var intType)) return "integer";
            if (long.TryParse(cellValueString, out var longType)) return "long";
            if (double.TryParse(cellValueString, out var doubleType)) return "double";
            if (bool.TryParse(cellValueString, out var booleanType)) return "boolean";

            return "string";

        }

        public Dictionary<string, string> GetDmnInfo(ExcelWorksheet ws)
        {
            var dmnName = ws.Cells["C1"].Value ?? "DMN Table Name";
            var dmnId = ws.Cells["C2"].Value ?? "dmnId";

            return new Dictionary<string, string>() { { dmnId.ToString(), dmnName.ToString() } };
        }

        public static Dictionary<string, string[]> GetColumnRagngeInLeters(ExcelTable table, int inputsColumnsCount, int outputsColumnsCount)
        {
            var dictionary = new Dictionary<string, string[]>();
            var start = table.Address.Start.Column;
            var end = table.Address.End.Column;

            if ((inputsColumnsCount + outputsColumnsCount) > (end -(start-1)))
                return null;
            var outputsStart = start + inputsColumnsCount;
            var inputsColumnIndexes = new List<string>();
            for (int i = start; i < outputsStart; i++)
            {
                inputsColumnIndexes.Add(GetColumnName(i));
            }
            var outputsColumnIndexes = new List<string>();
            for (int i = outputsStart; i < outputsStart + outputsColumnsCount; i++)
            {
                outputsColumnIndexes.Add(GetColumnName(i));
            }
            dictionary.Add("inputsIndex", inputsColumnIndexes.ToArray());
            dictionary.Add("outputsIndex", outputsColumnIndexes.ToArray());
            return dictionary;
        }

        public Dictionary<int, string> GetAnnotationsRulesFromExcel(ExcelWorksheet ws, string[] inputsIndex, string[] outputsIndex, bool variableId)
        {
            var dictionary = new Dictionary<int, string>();

            var table = ws.Tables.First();
            var tableStartRow = table.Address.Start.Row;


            var startRow = variableId ? tableStartRow + 2 : tableStartRow + 1;

            var tableSize = table.Columns.Count;
            if (tableSize == outputsIndex.Count() + inputsIndex.Count() + 1)
            {
                var start = GetRange(inputsIndex, out var end);
                if (!CheckColumnRange(table, start, end)) return dictionary;
                for (var i = startRow; i <= table.Address.End.Row; i++)
                {
                    var cellName = string.Concat(GetColumnName(table.Address.End.Column), i);
                    var cellValue = ws.Cells[cellName].Value;

                    dictionary.Add(i, cellValue?.ToString());
                }
            }
            return dictionary;
        }

        public static void CreateSummaryExcelTableDataDictionary(List<DataDictionaryModel> dataDictionaryList, ExcelWorksheet wsSheet, string tableName, string[] objectPropertyNames)
        {
            // place Table in Excel
            SetExcelTablePosition(dataDictionaryList.Count, objectPropertyNames.Length - 1, out var stratColumn, out var endRow, out var endColum, out var stratRow);

            //Create Excel table  and set Header
            CreateExcelTableHeader(wsSheet, tableName, stratRow, stratColumn, endRow, endColum, objectPropertyNames);

            // Set Excel table content
            var dataRow = stratRow + 1;
            var dataColumn = stratColumn;

            foreach (var dataDictionary in dataDictionaryList)
            {
                AddExcelTableRowData(dataDictionary, dataColumn, dataRow, wsSheet, objectPropertyNames);
                dataRow++;
            }

            //wsSheet1.Protection.IsProtected = false;
            //wsSheet1.Protection.AllowSelectLockedCells = false;
            wsSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
        }

        public static void CreateDmnExcelTableDataDictionary(IEnumerable<DmnDataDictionaryModel> dmns, ExcelWorksheet wsSheet, string tableName, string[] objectPropertyNames)
        {
            // place Table in Excel
            SetExcelTablePosition(dmns.Count(), objectPropertyNames.Length - 1, out var stratColumn, out var endRow, out var endColum, out var stratRow);

            //Create Excel table  and set Header
            CreateExcelTableHeader(wsSheet, tableName, stratRow, stratColumn, endRow, endColum, objectPropertyNames);

            // Set Excel table content
            var dataRow = stratRow + 1;
            var dataColumn = stratColumn;

            foreach (var dmnData in dmns)
            {
                AddExcelTableRowData(dmnData, dataColumn, dataRow, wsSheet, objectPropertyNames);
                dataRow++;
            }
            wsSheet.Cells[wsSheet.Dimension.Address].AutoFitColumns();
        }

        private static void CreateExcelTableHeader(ExcelWorksheet wsSheet, string tableName, int stratRow, int stratColumn, int endRow,
            int endColum, string[] objectPropertyNames)
        {
            ExcelTable table;
            using (ExcelRange rng = wsSheet.Cells[stratRow, stratColumn, endRow, endColum])
            {
                //Add data dictionary Headers to Table
                table = wsSheet.Tables.Add(rng, tableName);
                AddExcelTableHeaderColumnsValues(table, objectPropertyNames);
            }
        }

        private static void SetExcelTablePosition(int rowNumber, int propertiesNamesCount, out int stratColumn, out int endRow, out int endColum, out int stratRow)
        {
            stratRow = 5;
            stratColumn = 2;

            var totalDataColumns = propertiesNamesCount;
            var totalDataRows = rowNumber;
            endRow = stratRow + totalDataRows;
            endColum = stratColumn + totalDataColumns;
        }

        private static void AddExcelTableRowData(object modelData, int dataColumn, int dataRow, ExcelWorksheet wsSheet, string[] dmnFields)
        {
            for (int i = 0; i < dmnFields.Length; i++)
            {
                var value = GetPropertyStringValue(modelData, dmnFields[i]);
                    AddExcelCellByRowAndColumn(dataColumn, dataRow, value, wsSheet);
                    dataColumn++;
            }
        }

        private static string GetPropertyStringValue(object objectData, string propertyName)
        {
            try
            {
                foreach (String part in propertyName.Split('.'))
                {
                    if (objectData == null) { return null; }
                    Type type = objectData.GetType();
                    PropertyInfo info = type.GetProperty(part);
                    if (info == null) { return null; }
                    objectData = info.GetValue(objectData, null);
                }
                var stringValue = objectData.ToString();
                return stringValue;
            }
            catch
            {
                return null;
            }
        }

        private static void AddExcelTableHeaderColumnsValues(ExcelTable table, string[] name)
        {
            for (int i = 0; i < name.Length; i++)
            {
                table.Columns[i].Name = name[i];
            }
        }
    }
}

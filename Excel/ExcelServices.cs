using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
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
                type = type ?? typeTemp;
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
            }
            return type;
        }

        private string ParseCellValue(object cellValue)
        {
            if (string.IsNullOrEmpty(cellValue.ToString())) return String.Empty;
            var regex = Regex.Match(cellValue?.ToString(), @"^[<,>,=]?[=]?\s?(?<number>\d+(\.\d+)?)$");
            string cellValueString = regex.Success ? regex.Groups["number"].ToString() : cellValue.ToString();

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
    }
}

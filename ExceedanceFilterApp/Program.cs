using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;

namespace ExceedanceFilterApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            string configCsvPath = @"FilterConfig.csv";
            string dataCsvPath = @"VT-AQL.csv";
            string outputExcelPath = @"FilteredOutput.xlsx";

            if (!File.Exists(configCsvPath))
            {
                Console.WriteLine("Missing required config CSV file.");
                return;
            }
            if (!File.Exists(dataCsvPath))
            {
                Console.WriteLine("Missing required data file.");
                return;
            }
            if (File.Exists(outputExcelPath))
            {
                File.Delete(outputExcelPath);
                Console.WriteLine("🗑️ Existing output file deleted.");
            }

            ExceedanceFilterToExcel(configCsvPath, dataCsvPath, outputExcelPath);
        }

        /// <summary>
        /// Load CSV into list of dictionaries
        /// </summary>
        public static List<Dictionary<string, string>> LoadCsvToList(string csvPath)
        {
            var result = new List<Dictionary<string, string>>();
            using (TextFieldParser parser = new TextFieldParser(csvPath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                // Read header
                if (parser.EndOfData) return result;
                string[] headers = parser.ReadFields();
                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    if (fields == null || fields.Length == 0) continue;
                    var row = new Dictionary<string, string>();
                    for (int i = 0; i < headers.Length && i < fields.Length; i++)
                    {
                        row[headers[i].Trim(' ', '\uFEFF', '"')] = fields[i].Trim(' ', '"');
                    }
                    result.Add(row);
                }
            }
            return result;
        }
        /// <summary>
        /// Determines the limit (High > Medium > Low > NA)
        /// </summary>
        public static string DetermineLimit(Dictionary<string, string> cfg, out string note)
        {
            note = null;
            if (!string.IsNullOrEmpty(cfg.GetValueOrDefault("High Limit")))
                return cfg["High Limit"];
            if (!string.IsNullOrEmpty(cfg.GetValueOrDefault("Medium Limit")))
                return cfg["Medium Limit"];
            if (!string.IsNullOrEmpty(cfg.GetValueOrDefault("Low Limit")))
                return cfg["Low Limit"];
            note = "Note: No Limits Given";
            return "NA";
        }

        /// <summary>
        /// Expands (ColumnN, ConditionN) into a list. Supports comma-splitting conditions.
        /// </summary>
        public static List<(string Column, string Condition)> GetConditionPairs(Dictionary<string, string> cfg)
        {
            var result = new List<(string, string)>();
            int n = 1;
            while (true)
            {
                var colKey = $"Column{n}";
                var condKey = $"Condition{n}";
                if (!cfg.ContainsKey(colKey) || !cfg.ContainsKey(condKey)) break;
                var column = cfg[colKey];
                var conds = cfg[condKey];
                if (!string.IsNullOrWhiteSpace(column) && !string.IsNullOrWhiteSpace(conds))
                {
                    foreach (var cond in conds.Split(','))
                    {
                        var cleanCond = cond.Trim();
                        if (!string.IsNullOrWhiteSpace(cleanCond))
                            result.Add((column, cleanCond));
                    }
                }
                n++;
            }
            return result;
        }

        /// <summary>
        /// Returns the full expanded AND-ed condition string (one per condition)
        /// </summary>
        public static string GetConditionString(List<(string Column, string Condition)> pairs)
        {
            if (pairs.Count == 0)
                return "";
            var expanded = new List<string>();
            foreach (var pair in pairs)
            {
                if (pair.Condition.Contains(","))
                {
                    var subs = pair.Condition.Split(',');
                    foreach (var sub in subs)
                    {
                        if (!string.IsNullOrWhiteSpace(sub))
                            expanded.Add($"{pair.Column}{sub.Trim()}");
                    }
                }
                else
                {
                    expanded.Add($"{pair.Column}{pair.Condition}");
                }
            }
            return string.Join(" && ", expanded);
        }

        /// <summary>
        /// All conditions must pass (multiple per column is AND logic)
        /// </summary>
        public static bool AllConditionsPass(Dictionary<string, string> row, List<(string Column, string Condition)> condPairs)
        {
            var condsByCol = condPairs
                .GroupBy(x => x.Column)
                .ToDictionary(g => g.Key, g => g.Select(x => x.Condition).ToList());

            foreach (var col in condsByCol.Keys)
            {
                if (!row.ContainsKey(col))
                    return false;
                string cellVal = row[col];
                foreach (var cond in condsByCol[col])
                {
                    if (!ConditionMatch(cellVal, cond))
                        return false;
                }
            }
            return true;
        }

        /// <summary>
        /// Compares a single value with a condition string (>, <, >=, <=, =)
        /// </summary>
        public static bool ConditionMatch(string cellVal, string cond)
        {
            // Numeric check
            if (double.TryParse(cellVal, out double cellNum))
            {
                if (cond.StartsWith(">="))
                    return cellNum >= double.Parse(cond.Substring(2));
                else if (cond.StartsWith("<="))
                    return cellNum <= double.Parse(cond.Substring(2));
                else if (cond.StartsWith(">"))
                    return cellNum > double.Parse(cond.Substring(1));
                else if (cond.StartsWith("<"))
                    return cellNum < double.Parse(cond.Substring(1));
                else if (cond.StartsWith("="))
                    return cellNum == double.Parse(cond.Substring(1));
                else
                    return false; // unsupported
            }
            // String equality only (e.g., =ON)
            else
            {
                if (cond.StartsWith("="))
                    return cellVal == cond.Substring(1);
                else
                    return false; // unsupported string op
            }
        }

        /// <summary>
        /// Main logic: applies all filter config, writes reference sheet, and logs to Excel and console
        /// </summary>
        public static void ExceedanceFilterToExcel(string configCsvPath, string dataCsvPath, string outputExcelPath)
        {
            var configRows = LoadCsvToList(configCsvPath);
            var dataRows = LoadCsvToList(dataCsvPath);
            var dataHeaders = dataRows.Count > 0 ? dataRows[0].Keys.ToList() : new List<string>();

            if (configRows.Count == 0 || dataRows.Count == 0)
            {
                Console.WriteLine("No data in config or data CSV.");
                return;
            }

            var sheetNameCounts = new Dictionary<string, int>();
            using var workbook = new XLWorkbook();

            // 1. Add RawData sheet for reference
            //var wsRaw = workbook.Worksheets.Add("RawData");
            //for (int c = 0; c < dataHeaders.Count; c++)
            //{
            //    wsRaw.Cell(1, c + 1).Value = dataHeaders[c];
            //    wsRaw.Cell(1, c + 1).Style.Font.Bold = true;
            //    wsRaw.Cell(1, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            //}
            //for (int r = 0; r < dataRows.Count; r++)
            //{
            //    for (int c = 0; c < dataHeaders.Count; c++)
            //    {
            //        wsRaw.Cell(r + 2, c + 1).Value = dataRows[r][dataHeaders[c]];
            //    }
            //}
            //wsRaw.Columns().AdjustToContents();
            //Console.WriteLine("RawData sheet written (all flight data included for reference).");

            int configRowIndex = 0;
            foreach (var cfg in configRows)
            {
                configRowIndex++;

                // SHEET NAME LOGIC (sanitize, deduplicate, truncate)
                string baseSheetNameRaw = cfg.GetValueOrDefault("SheetName") ?? "Sheet";
                char[] invalidChars = new char[] { '/', '\\', '?', '*', '[', ']', ':' };
                string baseSheetName = baseSheetNameRaw;
                foreach (var ch in invalidChars)
                    baseSheetName = baseSheetName.Replace(ch, '_');
                string sheetNameForCount = baseSheetName.Length > 31 ? baseSheetName.Substring(0, 31) : baseSheetName;
                string validSheetName = sheetNameForCount;
                if (sheetNameCounts.ContainsKey(sheetNameForCount))
                {
                    int nextIdx = ++sheetNameCounts[sheetNameForCount];
                    string suffix = $"_{nextIdx}";
                    int maxLen = 31 - suffix.Length;
                    validSheetName = (sheetNameForCount.Length > maxLen ? sheetNameForCount.Substring(0, maxLen) : sheetNameForCount) + suffix;
                }
                else
                {
                    sheetNameCounts[sheetNameForCount] = 0;
                }

                Console.WriteLine($"[Row {configRowIndex}] Config SheetName: \"{baseSheetNameRaw}\" → Excel SheetName: \"{validSheetName}\"");

                // Limit logic
                string noteLimit;
                string limit = DetermineLimit(cfg, out noteLimit);
                if (!string.IsNullOrEmpty(noteLimit))
                    Console.WriteLine($"[Row {configRowIndex}] {noteLimit}");

                // Condition pairs and string
                var conditionPairs = GetConditionPairs(cfg);

                // Debug print: See all parsed conditions for this config row
                foreach (var pair in conditionPairs)
                    Console.WriteLine($"  ConditionPair: {pair.Column} {pair.Condition}");

                string condString = GetConditionString(conditionPairs);


                // Filtering data rows: all conditions (possibly multiple for one column) must pass
                List<Dictionary<string, string>> filtered = new List<Dictionary<string, string>>();
                string noteCondition = null;
                if (conditionPairs.Count > 0)
                {
                    Console.WriteLine($"[Row {configRowIndex}] Filter condition: {condString}");
                    filtered = dataRows.Where(row => AllConditionsPass(row, conditionPairs)).ToList();
                }
                else
                {
                    noteCondition = "Note: No Condition Given";
                    Console.WriteLine($"[Row {configRowIndex}] {noteCondition}");
                }
                if (filtered.Count == 0)
                    Console.WriteLine($"[Row {configRowIndex}] Note: No Exceedance Detected");

                // --- Write to worksheet ---
                var ws = workbook.Worksheets.Add(validSheetName);

                int rowPtr = 1;

                // Title (left-aligned, as requested)
                ws.Cell(rowPtr, 1).Value = $"{baseSheetNameRaw} > {limit}";
                ws.Row(rowPtr).Style.Font.Bold = true;
                ws.Row(rowPtr).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                rowPtr++;

                // Note: Limits
                if (!string.IsNullOrEmpty(noteLimit))
                {
                    ws.Cell(rowPtr, 1).Value = noteLimit;
                    ws.Range(rowPtr, 1, rowPtr, dataHeaders.Count).Merge();
                    rowPtr++;
                }

                // Note: Conditions
                if (!string.IsNullOrEmpty(noteCondition))
                {
                    ws.Cell(rowPtr, 1).Value = noteCondition;
                    ws.Range(rowPtr, 1, rowPtr, dataHeaders.Count).Merge();
                    rowPtr++;
                }
                else
                {
                    ws.Cell(rowPtr, 1).Value = "Condition Applied: " + condString;
                    ws.Range(rowPtr, 1, rowPtr, dataHeaders.Count).Merge();
                    rowPtr++;
                }

                // Header row (always output header)
                for (int c = 0; c < dataHeaders.Count; c++)
                {
                    ws.Cell(rowPtr, c + 1).Value = dataHeaders[c];
                    ws.Cell(rowPtr, c + 1).Style.Font.Bold = true;
                    ws.Cell(rowPtr, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                rowPtr++;

                // Write filtered data, or note if empty
                if (filtered.Count > 0)
                {
                    for (int r = 0; r < filtered.Count; r++)
                    {
                        for (int c = 0; c < dataHeaders.Count; c++)
                        {
                            ws.Cell(rowPtr + r, c + 1).Value = filtered[r][dataHeaders[c]];
                        }
                    }
                }
                else
                {
                    ws.Cell(rowPtr, 1).Value = "Note: No Exceedance Detected";
                    ws.Range(rowPtr, 1, rowPtr, dataHeaders.Count).Merge();
                }

                ws.Columns().AdjustToContents();
                Console.WriteLine(new string('-', 40));
            }

            workbook.SaveAs(outputExcelPath);
            Console.WriteLine($"✅ Excel file written: {outputExcelPath}");
        }
    }
}

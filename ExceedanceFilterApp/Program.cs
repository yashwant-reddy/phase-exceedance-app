using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using ClosedXML.Excel;
using Microsoft.VisualBasic.FileIO;
using System.Globalization;

namespace ExceedanceFilterApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // --- Main entry: file paths are hardcoded for demonstration purposes.
            string configCsvPath = @"FilterConfig.csv";
            string dataCsvPath = @"VT-AQL.csv";
            string outputExcelPath = @"FilteredOutput.xlsx";

            // --- Check for presence of required input files.
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

            // --- Run the exceedance filtering/export logic.
            ExceedanceFilterToExcel(configCsvPath, dataCsvPath, outputExcelPath);
        }

        /// <summary>
        /// Loads a CSV into a list of dictionaries (header:value).
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
        /// Determines which limit is used (High > Medium > Low), 
        /// returns limit value, sets out param limitType ("High Limit", etc.), and returns null if none found.
        /// </summary>
        public static string DetermineLimit(Dictionary<string, string> cfg, out string limitType, out string note)
        {
            note = null;
            limitType = null;
            if (!string.IsNullOrEmpty(cfg.GetValueOrDefault("High Limit")))
            {
                limitType = "High Limit";
                return cfg["High Limit"];
            }
            if (!string.IsNullOrEmpty(cfg.GetValueOrDefault("Medium Limit")))
            {
                limitType = "Medium Limit";
                return cfg["Medium Limit"];
            }
            if (!string.IsNullOrEmpty(cfg.GetValueOrDefault("Low Limit")))
            {
                limitType = "Low Limit";
                return cfg["Low Limit"];
            }
            note = "Note: No Limits Given";
            return null;
        }

        /// <summary>
        /// Returns filter time (in seconds) based on which limit is active.
        /// If filter time is blank or missing for present limit, returns 1.
        /// </summary>
        public static int GetFilterTime(Dictionary<string, string> cfg, string limitType)
        {
            // Select filter time string based on limit type
            string filterTimeStr = null;
            if (limitType == "High Limit")
                filterTimeStr = cfg.GetValueOrDefault("High Filter Time");
            else if (limitType == "Medium Limit")
                filterTimeStr = cfg.GetValueOrDefault("Medium Filter Time");
            else if (limitType == "Low Limit")
                filterTimeStr = cfg.GetValueOrDefault("Low Filter Time");

            int filterTimeSeconds = 1; // Default value if missing or not valid
            if (!string.IsNullOrWhiteSpace(filterTimeStr))
            {
                int.TryParse(filterTimeStr.Trim(), out filterTimeSeconds);
                if (filterTimeSeconds < 1) filterTimeSeconds = 1;
            }
            return filterTimeSeconds;
        }

        /// <summary>
        /// Expands (ColumnN, ConditionN) into a list of pairs. Supports comma-splitting for multiple conditions.
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
        /// Returns a human-readable string summarizing all AND-ed conditions.
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
        /// All conditions must pass for a row. Multiple per column is AND logic.
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
        /// Compares a single value with a condition string (>, <, >=, <=, =), numeric or string.
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
        /// Parses "HH:MM:SS" string to TimeSpan, returns null if invalid.
        /// </summary>
        public static TimeSpan? ParseTime(string timeStr)
        {
            if (TimeSpan.TryParseExact(timeStr, "c", CultureInfo.InvariantCulture, out var t)) return t;
            if (TimeSpan.TryParseExact(timeStr, @"hh\:mm\:ss", CultureInfo.InvariantCulture, out t)) return t;
            if (TimeSpan.TryParse(timeStr, out t)) return t;
            return null;
        }

        /// <summary>
        /// Applies filter time grouping logic: returns only groups where the
        /// threshold is exceeded for >= minFilterTime seconds (consecutive).
        /// </summary>
        /// <summary>
        /// Filters using a sliding window of N seconds over the time column.
        /// Only includes rows that are part of a window where all times are present and all rows meet the condition.
        /// </summary>
        public static List<Dictionary<string, string>> ApplyFilterTimeGrouping(
            List<Dictionary<string, string>> filteredRows, // Already condition-checked
            string timeColumn,
            int windowSeconds)
        {
            if (filteredRows.Count == 0 || windowSeconds <= 1)
                return filteredRows;

            // Sort by time ascending
            var rowsSorted = filteredRows
                .Where(r => ParseTime(r[timeColumn]).HasValue)
                .OrderBy(r => ParseTime(r[timeColumn]).Value)
                .ToList();

            var timeToRow = rowsSorted
                .ToDictionary(r => ParseTime(r[timeColumn]).Value, r => r);

            var outputRows = new HashSet<Dictionary<string, string>>();

            for (int i = 0; i < rowsSorted.Count; i++)
            {
                var tStart = ParseTime(rowsSorted[i][timeColumn]).Value;
                bool windowOk = true;
                // Check for consecutive N-1 seconds present and valid
                for (int j = 0; j < windowSeconds; j++)
                {
                    var tNext = tStart.Add(TimeSpan.FromSeconds(j));
                    if (!timeToRow.ContainsKey(tNext))
                    {
                        windowOk = false;
                        break;
                    }
                }
                if (windowOk)
                {
                    // All rows in the window are valid, add all of them
                    for (int j = 0; j < windowSeconds; j++)
                    {
                        var tNext = tStart.Add(TimeSpan.FromSeconds(j));
                        outputRows.Add(timeToRow[tNext]);
                    }
                }
            }
            // Return rows in file order
            return filteredRows.Where(r => outputRows.Contains(r)).ToList();
        }

        /// <summary>
        /// Main logic: applies all filter config, writes output Excel, logs to console and sheet.
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

            int configRowIndex = 0;
            foreach (var cfg in configRows)
            {
                configRowIndex++;

                // --- Sanitize and deduplicate sheet name (Excel rules)
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

                // --- Determine which limit is being used (and its label, and note if missing)
                string noteLimit, usedLimitType;
                string usedLimit = DetermineLimit(cfg, out usedLimitType, out noteLimit);

                // --- If no limits at all, skip this config row
                if (usedLimit == null)
                {
                    Console.WriteLine($"[Row {configRowIndex}] {noteLimit}");
                    continue;
                }

                // --- Get filter time for the chosen limit, modular and reusable
                int filterTimeSeconds = GetFilterTime(cfg, usedLimitType);

                // --- Prepare condition pairs and string
                var conditionPairs = GetConditionPairs(cfg);
                string condString = GetConditionString(conditionPairs);

                // --- Filtering rows: all conditions must pass
                List<Dictionary<string, string>> filteredRows = new List<Dictionary<string, string>>();
                string noteCondition = null;
                if (conditionPairs.Count > 0)
                {
                    Console.WriteLine($"[Row {configRowIndex}] Filter condition: {condString}");
                    filteredRows = dataRows.Where(row => AllConditionsPass(row, conditionPairs)).ToList();
                }
                else
                {
                    noteCondition = "Note: No Condition Given";
                    Console.WriteLine($"[Row {configRowIndex}] {noteCondition}");
                }

                // --- Apply filter time logic ONLY if time column exists and limit is present
                string timeColumn = "HH:MM:SS";
                if (!string.IsNullOrWhiteSpace(usedLimit) && filterTimeSeconds > 1 && filteredRows.Count > 0 && dataHeaders.Contains(timeColumn))
                {
                    filteredRows = ApplyFilterTimeGrouping(filteredRows, timeColumn, filterTimeSeconds);
                }

                if (filteredRows.Count == 0)
                    Console.WriteLine($"[Row {configRowIndex}] Note: No Exceedance Detected");

                // --- Write to worksheet ---
                var ws = workbook.Worksheets.Add(validSheetName);
                int rowPtr = 1;

                // --- Title: include limit name, value (no parenthesis), and filter time in required format
                ws.Cell(rowPtr, 1).Value = $"{baseSheetNameRaw} > {usedLimitType} {usedLimit} (Filter Time: {filterTimeSeconds} sec)";
                ws.Row(rowPtr).Style.Font.Bold = true;
                ws.Row(rowPtr).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                rowPtr++;

                // --- Note: Limits
                if (!string.IsNullOrEmpty(noteLimit))
                {
                    ws.Cell(rowPtr, 1).Value = noteLimit;
                    ws.Range(rowPtr, 1, rowPtr, dataHeaders.Count).Merge();
                    rowPtr++;
                }

                // --- Note: Conditions
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

                // --- Header row (always output header)
                for (int c = 0; c < dataHeaders.Count; c++)
                {
                    ws.Cell(rowPtr, c + 1).Value = dataHeaders[c];
                    ws.Cell(rowPtr, c + 1).Style.Font.Bold = true;
                    ws.Cell(rowPtr, c + 1).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                }
                rowPtr++;

                // --- Write filtered data, or note if empty
                if (filteredRows.Count > 0)
                {
                    for (int r = 0; r < filteredRows.Count; r++)
                    {
                        for (int c = 0; c < dataHeaders.Count; c++)
                        {
                            ws.Cell(rowPtr + r, c + 1).Value = filteredRows[r][dataHeaders[c]];
                        }
                    }
                    // Add Row Count in the first empty row after data
                    int rowCountCellRow = rowPtr + filteredRows.Count + 1;
                    ws.Cell(rowCountCellRow, dataHeaders.Count + 1).Value = $"Row Count: {filteredRows.Count}";
                    ws.Cell(rowCountCellRow, dataHeaders.Count + 1).Style.Font.Bold = true;
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

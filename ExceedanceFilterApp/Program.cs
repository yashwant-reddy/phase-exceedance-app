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
            string outputExcelPath = @"ExceedanceOutput.xlsx";

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

        //wip
        public static List<Dictionary<string, string>> LoadCsvToList(string csvPath)
        {
            // Holds final CSV rows
            var result = new List<Dictionary<string, string>>();

            using (TextFieldParser parser = new TextFieldParser(csvPath))
            {
                // Configure parser for comma-delimited CSV
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;

                // If file is empty, return empty list
                if (parser.EndOfData) return result;

                // Read header row
                string[] headers = parser.ReadFields();

                // Clean header values (handle BOM, spaces, quotes)
                for (int i = 0; i < headers.Length; i++)
                {
                    headers[i] = headers[i].Trim(' ', '\uFEFF', '"');
                }

                // Read data rows
                while (!parser.EndOfData)
                {
                    string[] fields;

                    try
                    {
                        // Attempt to read a CSV row
                        fields = parser.ReadFields();
                    }
                    catch (MalformedLineException)
                    {
                        // Skip malformed rows instead of crashing
                        continue;
                    }

                    // Skip empty or invalid rows
                    if (fields == null || fields.Length == 0)
                        continue;

                    var row = new Dictionary<string, string>();

                    // Map fields to headers safely
                    for (int i = 0; i < headers.Length; i++)
                    {
                        // If field count is less than header count, fill empty
                        string value = i < fields.Length ? fields[i] : string.Empty;

                        row[headers[i]] = value.Trim(' ', '"');
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
        // ===============================================
        // Extract Column + Condition + Logic (AND/OR)
        // FIX: Supports multiple conditions (>7,<9)
        // ===============================================
        public static List<(string Column, string Condition, string Logic)> GetConditionPairs(Dictionary<string, string> cfg)
        {
            var result = new List<(string, string, string)>();
            int n = 1;

            // --- Loop through Column1, Column2... dynamically
            while (true)
            {
                var colKey = $"Column{n}";
                var condKey = $"Condition{n}";
                var logicKey = $"Logic{n}";

                // --- Stop if no more columns
                if (!cfg.ContainsKey(colKey) || !cfg.ContainsKey(condKey))
                    break;

                var column = cfg[colKey];
                var conds = cfg[condKey];
                var logic = cfg.GetValueOrDefault(logicKey)?.Trim().ToUpper();

                // --- Validate values
                if (!string.IsNullOrWhiteSpace(column) && !string.IsNullOrWhiteSpace(conds))
                {
                    // --- Split multiple conditions (>7,<9)
                    var splitConditions = conds.Split(',');

                    foreach (var cond in splitConditions)
                    {
                        var cleanCond = cond.Trim();

                        if (!string.IsNullOrWhiteSpace(cleanCond))
                        {
                            result.Add((column.Trim(), cleanCond, logic));
                        }
                    }
                }

                n++;
            }

            return result;
        }

        /// <summary>
        /// Returns a human-readable string summarizing all AND-ed conditions.
        /// </summary>
        // ===============================================
        // Build readable condition string with AND / OR
        // ===============================================
        public static string GetConditionString(List<(string Column, string Condition, string Logic)> pairs)
        {
            if (pairs.Count == 0)
                return "";

            string result = $"{pairs[0].Column}{pairs[0].Condition}";

            for (int i = 1; i < pairs.Count; i++)
            {
                string logic = pairs[i - 1].Logic ?? "AND";
                string symbol = logic == "OR" ? " || " : " && ";

                result += symbol + $"{pairs[i].Column}{pairs[i].Condition}";
            }

            return result;
        }


        /// <summary>
        /// All conditions must pass for a row. Multiple per column is AND logic.
        /// </summary>
        // ===============================================
        // Evaluate conditions using Logic1, Logic2...
        // OPTIMIZED: Uses columnMap (fast lookup)
        // ===============================================
        public static bool AllConditionsPass(
            Dictionary<string, string> row,
            List<(string Column, string Condition, string Logic)> condPairs,
            Dictionary<string, string> columnMap) // <-- performance fix
        {
            if (condPairs.Count == 0)
                return false;

            bool result = false;

            for (int i = 0; i < condPairs.Count; i++)
            {
                var pair = condPairs[i];

                // --- Fast column lookup (no expensive search)
                if (!columnMap.TryGetValue(pair.Column, out string actualColumn))
                    return false;

                string cellVal = row[actualColumn];

                bool currentMatch = ConditionMatch(cellVal, pair.Condition);

                if (i == 0)
                {
                    result = currentMatch;
                }
                else
                {
                    string logic = condPairs[i - 1].Logic ?? "AND";

                    // --- Short circuit (performance boost)
                    if (logic == "OR")
                    {
                        result = result || currentMatch;
                        if (result) return true;
                    }
                    else
                    {
                        result = result && currentMatch;
                        if (!result) return false;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Compares a single value with a condition string (>, <, >=, <=, =), numeric or string.
        /// </summary>
        // ===============================================
        // Compare value with condition
        // Supports: >, <, >=, <=, =, ==, !=
        // Handles numeric + string safely
        // ===============================================
        public static bool ConditionMatch(string cellVal, string cond)
        {
            // --- Clean inputs
            cellVal = cellVal?.Trim();
            cond = cond?.Trim();

            if (string.IsNullOrWhiteSpace(cellVal) || string.IsNullOrWhiteSpace(cond))
                return false;

            // --- Try parse number once
            bool isNumber = double.TryParse(cellVal, out double cellNum);

            // ===============================================
            // NUMERIC CONDITIONS
            // ===============================================
            if (isNumber)
            {
                double value;

                if (cond.StartsWith(">=") && double.TryParse(cond.Substring(2).Trim(), out value))
                    return cellNum >= value;

                if (cond.StartsWith("<=") && double.TryParse(cond.Substring(2).Trim(), out value))
                    return cellNum <= value;

                if (cond.StartsWith("!=") && double.TryParse(cond.Substring(2).Trim(), out value))
                    return cellNum != value;

                if (cond.StartsWith("==") && double.TryParse(cond.Substring(2).Trim(), out value))
                    return cellNum == value;

                if (cond.StartsWith(">") && double.TryParse(cond.Substring(1).Trim(), out value))
                    return cellNum > value;

                if (cond.StartsWith("<") && double.TryParse(cond.Substring(1).Trim(), out value))
                    return cellNum < value;

                if (cond.StartsWith("=") && double.TryParse(cond.Substring(1).Trim(), out value))
                    return cellNum == value;
            }

            // ===============================================
            // STRING CONDITIONS
            // ===============================================
            if (cond.StartsWith("!="))
                return !string.Equals(cellVal, cond.Substring(2).Trim(), StringComparison.OrdinalIgnoreCase);

            if (cond.StartsWith("==") || cond.StartsWith("="))
                return string.Equals(cellVal, cond.TrimStart('=').Trim(), StringComparison.OrdinalIgnoreCase);

            return false;
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
        /// Applies filter time grouping logic (UPDATED)
        /// FIX:
        /// - Ignores invalid "-" rows
        /// - Works with real flight data gaps
        /// - Keeps valid exceedance rows
        /// </summary>
        public static List<Dictionary<string, string>> ApplyFilterTimeGrouping(
            List<Dictionary<string, string>> filteredRows,
            string timeColumn,
            int windowSeconds)
        {
            // ===============================================
            // EARLY EXIT
            // ===============================================
            if (filteredRows.Count == 0 || windowSeconds <= 1)
                return filteredRows;

            // ===============================================
            // REMOVE INVALID ROWS ("-")
            // ===============================================
            var validRows = filteredRows
                .Where(row =>
                {
                    foreach (var val in row.Values)
                    {
                        if (val == "-" || string.IsNullOrWhiteSpace(val))
                            return false;
                    }
                    return true;
                })
                .ToList();

            // --- fallback safety
            if (validRows.Count == 0)
                return filteredRows;

            // ===============================================
            // PARSE + SORT BY TIME
            // ===============================================
            var rowsWithTime = validRows
                .Select(r => new
                {
                    Row = r,
                    Time = ParseTime(r[timeColumn])
                })
                .Where(x => x.Time.HasValue)
                .OrderBy(x => x.Time.Value)
                .ToList();

            var outputRows = new HashSet<Dictionary<string, string>>();

            // ===============================================
            // GROUP BASED ON TIME DIFFERENCE
            // ===============================================
            int startIndex = 0;

            while (startIndex < rowsWithTime.Count)
            {
                var startTime = rowsWithTime[startIndex].Time.Value;
                var group = new List<Dictionary<string, string>>
        {
            rowsWithTime[startIndex].Row
        };

                int nextIndex = startIndex + 1;

                while (nextIndex < rowsWithTime.Count)
                {
                    var currentTime = rowsWithTime[nextIndex].Time.Value;

                    if ((currentTime - startTime).TotalSeconds < windowSeconds)
                    {
                        group.Add(rowsWithTime[nextIndex].Row);
                        nextIndex++;
                    }
                    else
                    {
                        break;
                    }
                }

                // --- keep group (even single row)
                foreach (var row in group)
                {
                    outputRows.Add(row);
                }

                startIndex = nextIndex;
            }

            // ===============================================
            // RETURN ORIGINAL ORDER
            // ===============================================
            return filteredRows.Where(r => outputRows.Contains(r)).ToList();
        }

        /// <summary>
        /// Main processing method:
        /// - Loads config + data
        /// - Applies exceedance logic
        /// - ALWAYS creates sheet (even if no exceedance)
        /// - Writes debug logs
        /// </summary>
        public static void ExceedanceFilterToExcel(string configCsvPath, string dataCsvPath, string outputExcelPath)
        {
            // ===============================================
            // LOGGING SETUP (Buffered logging for performance)
            // ===============================================
            var logBuilder = new System.Text.StringBuilder();
            string logFilePath = Path.Combine(Path.GetDirectoryName(outputExcelPath) ?? "", "Exceedance_Debug_Log.txt");
            logBuilder.AppendLine($"===== START EXCEEDANCE PROCESS: {DateTime.Now} =====");

            // ===============================================
            // LOAD INPUT FILES (Config + Data CSV)
            // ===============================================
            var configRows = LoadCsvToList(configCsvPath);
            var dataRows = LoadCsvToList(dataCsvPath);
            var dataHeaders = dataRows.Count > 0 ? dataRows[0].Keys.ToList() : new List<string>();

            // --- Validate input
            if (configRows.Count == 0 || dataRows.Count == 0)
            {
                logBuilder.AppendLine("❌ Error: Config or Data CSV is empty.");
                File.WriteAllText(logFilePath, logBuilder.ToString());
                return;
            }

            // ===============================================
            // CREATE COLUMN MAP (Case-insensitive lookup)
            // ===============================================
            var columnMap = dataHeaders.ToDictionary(h => h.Trim(), h => h, StringComparer.OrdinalIgnoreCase);

            // --- Track duplicate sheet names
            var sheetNameCounts = new Dictionary<string, int>();

            // --- Flag to check if any sheet created
            bool isAnyWorksheetCreated = false;

            using var workbook = new XLWorkbook();
            int configRowIndex = 0;

            // ===============================================
            // LOOP THROUGH EACH CONFIG RULE
            // ===============================================
            foreach (var cfg in configRows)
            {
                configRowIndex++;

                // --- Get base sheet name
                string baseSheetNameRaw = cfg.GetValueOrDefault("SheetName") ?? $"Rule_{configRowIndex}";

                // ===============================================
                // STEP 1: DETERMINE LIMIT + FILTER TIME
                // ===============================================
                string usedLimitType, noteLimit;
                string usedLimit = DetermineLimit(cfg, out usedLimitType, out noteLimit);

                if (string.IsNullOrEmpty(usedLimit))
                {
                    logBuilder.AppendLine($"Skipping Config Row {configRowIndex}: No limit defined.");
                    continue;
                }

                int filterTimeSeconds = GetFilterTime(cfg, usedLimitType);

                // ===============================================
                // STEP 2: BUILD CONDITIONS
                // ===============================================
                var conditionPairs = GetConditionPairs(cfg);

                // ===============================================
                // STEP 3: APPLY INITIAL CONDITION FILTER
                // ===============================================
                var filteredRows = new List<Dictionary<string, string>>();

                // ===============================================
                // STEP 3: APPLY CLEAN + CONDITION FILTER
                // FIX: Skip rows with "-" in condition columns
                // ===============================================
                foreach (var row in dataRows)
                {
                    bool isValid = true;

                    // --- Check ONLY required columns (IMPORTANT)
                    foreach (var pair in conditionPairs)
                    {
                        if (!columnMap.TryGetValue(pair.Column, out string actualColumn))
                        {
                            isValid = false;
                            break;
                        }

                        var value = row[actualColumn];

                        if (string.IsNullOrWhiteSpace(value) || value.Trim() == "-")
                        {
                            isValid = false;
                            break;
                        }
                    }

                    if (!isValid)
                        continue;

                    // --- Apply condition
                    if (AllConditionsPass(row, conditionPairs, columnMap))
                    {
                        filteredRows.Add(row);
                    }
                }

                logBuilder.AppendLine($"Config {configRowIndex} ({baseSheetNameRaw}): {filteredRows.Count} rows passed initial conditions.");

                // ===============================================
                // STEP 4: APPLY TIME FILTER (PERSISTENCE)
                // ===============================================
                string timeColumn = "HH:MM:SS";

                if (filterTimeSeconds > 1 && filteredRows.Count > 0 && columnMap.ContainsKey(timeColumn))
                {
                    filteredRows = ApplyFilterTimeGrouping(filteredRows, columnMap[timeColumn], filterTimeSeconds);
                    logBuilder.AppendLine($"   -> After {filterTimeSeconds}s time filter: {filteredRows.Count} rows remain.");
                }

                // ===============================================
                // STEP 5: ALWAYS CREATE SHEET (FIXED)
                // ===============================================
                string validSheetName = GenerateValidSheetName(baseSheetNameRaw, sheetNameCounts);
                var ws = workbook.Worksheets.Add(validSheetName);
                isAnyWorksheetCreated = true;

                int rowPtr = 1;

                // --- Write Header (Rule + Limit)
                ws.Cell(rowPtr, 1).Value = $"{baseSheetNameRaw} | {usedLimitType}: {usedLimit}";
                ws.Cell(rowPtr, 1).Style.Font.Bold = true;
                rowPtr++;

                // --- Write Condition string
                ws.Cell(rowPtr++, 1).Value = "Condition: " + GetConditionString(conditionPairs);

                // ===============================================
                // STEP 6: HANDLE NO EXCEEDANCE CASE
                // ===============================================
                if (filteredRows.Count == 0)
                {
                    // --- Write message instead of skipping
                    ws.Cell(rowPtr, 1).Value = "No exceedance detected";
                    ws.Cell(rowPtr, 1).Style.Font.FontColor = XLColor.Red;
                    ws.Cell(rowPtr, 1).Style.Font.Bold = true;

                    logBuilder.AppendLine($"   -> Result: No exceedance detected. Sheet '{validSheetName}' created (empty).");

                    continue; // move to next config
                }

                // ===============================================
                // STEP 7: WRITE TABLE HEADER
                // ===============================================
                for (int c = 0; c < dataHeaders.Count; c++)
                {
                    var cell = ws.Cell(rowPtr, c + 1);
                    cell.Value = dataHeaders[c];
                    cell.Style.Fill.BackgroundColor = XLColor.LightGray;
                    cell.Style.Font.Bold = true;
                }

                rowPtr++;

                // ===============================================
                // STEP 8: WRITE FILTERED DATA ROWS
                // ===============================================
                foreach (var r in filteredRows)
                {
                    for (int c = 0; c < dataHeaders.Count; c++)
                    {
                        ws.Cell(rowPtr, c + 1).Value = r[dataHeaders[c]];
                    }
                    rowPtr++;
                }

                // --- Auto fit columns
                ws.Columns().AdjustToContents();

                logBuilder.AppendLine($"   -> Result: Success. Generated sheet '{validSheetName}' with {filteredRows.Count} rows.");
            }

            // ===============================================
            // FINAL STEP: HANDLE NO SHEETS CASE
            // ===============================================
            if (!isAnyWorksheetCreated)
            {
                var ws = workbook.Worksheets.Add("Summary");
                ws.Cell(1, 1).Value = "No exceedance data detected for the provided parameters.";
            }

            // ===============================================
            // SAVE FILE + WRITE LOG
            // ===============================================
            workbook.SaveAs(outputExcelPath);

            logBuilder.AppendLine($"===== PROCESS COMPLETE: {DateTime.Now} =====");

            File.WriteAllText(logFilePath, logBuilder.ToString());

            Console.WriteLine($"Done. Output: {outputExcelPath}. Log: {logFilePath}");
        }

        // Helper to keep the main method clean
        private static string GenerateValidSheetName(string name, Dictionary<string, int> counts)
        {
            char[] invalidChars = { '/', '\\', '?', '*', '[', ']', ':' };
            foreach (var ch in invalidChars) name = name.Replace(ch, '_');

            if (name.Length > 25) name = name.Substring(0, 25); // Leave room for suffix

            if (counts.ContainsKey(name))
            {
                counts[name]++;
                return $"{name}_{counts[name]}";
            }

            counts[name] = 0;
            return name;
        }

    }
}

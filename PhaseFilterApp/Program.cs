using System;
using System.IO;
using System.Linq;
using System.Globalization;
using System.Collections.Generic;
using ClosedXML.Excel;
using ExceedanceFilterApp; // Project reference required

namespace PhaseFilterApp
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // --- Console arg check and usage
            if (args.Length != 3)
            {
                Console.WriteLine("Usage: PhaseFilterApp [config.csv] [data.csv] [output.xlsx]");
                return;
            }
            string configPath = args[0];
            string dataPath = args[1];
            string outputPath = args[2];

            try
            {
                // --- Delete output if exists (like Exceedance app)
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                    Console.WriteLine("🗑️ Existing output file deleted.");
                }

                PhaseFilterToExcel(configPath, dataPath, outputPath);
                Console.WriteLine($"✅ Done! Excel file created: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("❌ Error: " + ex.Message);
            }
        }
        /// <summary>
        /// Advanced: Checks if the cell value matches *all* comma-separated conditions.
        /// Handles numeric and string (case-insensitive, trimmed) comparisons, and all operators.
        /// </summary>
        public static bool ConditionMatchAdvanced(string cellVal, string cond)
        {
            if (cond == null) return true;
            if (cellVal == null) cellVal = "";

            foreach (var c in cond.Split(','))
            {
                var condPart = c.Trim();
                if (string.IsNullOrEmpty(condPart)) continue;

                // Try numeric comparisons
                if (double.TryParse(cellVal.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out double cellNum))
                {
                    if (condPart.StartsWith(">="))
                    {
                        if (!(cellNum >= double.Parse(condPart.Substring(2), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (condPart.StartsWith("<="))
                    {
                        if (!(cellNum <= double.Parse(condPart.Substring(2), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (condPart.StartsWith(">"))
                    {
                        if (!(cellNum > double.Parse(condPart.Substring(1), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (condPart.StartsWith("<"))
                    {
                        if (!(cellNum < double.Parse(condPart.Substring(1), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (condPart.StartsWith("="))
                    {
                        if (!(cellNum == double.Parse(condPart.Substring(1), CultureInfo.InvariantCulture))) return false;
                    }
                    else
                    {
                        continue; // skip unsupported
                    }
                }
                else // String comparison
                {
                    string val = cellVal.Trim().Trim('"');
                    string expected = condPart.Replace("==", "").Trim().Trim('"');

                    if (condPart.StartsWith("=="))
                    {
                        if (!val.Equals(expected, StringComparison.OrdinalIgnoreCase)) return false;
                    }
                    else if (condPart.StartsWith("!="))
                    {
                        if (val.Equals(expected, StringComparison.OrdinalIgnoreCase)) return false;
                    }
                    else
                    {
                        if (!val.Equals(expected, StringComparison.OrdinalIgnoreCase)) return false;
                    }
                }
            }
            return true;
        }

        public static bool AllConditionsPassStrict(Dictionary<string, string> row, List<(string Column, string Condition)> condPairs)
        {
            foreach (var pair in condPairs)
            {
                if (!row.ContainsKey(pair.Column))
                    return false;
                string cellVal = row[pair.Column]?.Trim();
                string[] condParts = pair.Condition.Split(',');

                foreach (var cond in condParts)
                {
                    string part = cond.Trim();
                    // Numeric
                    if (part.StartsWith(">="))
                    {
                        if (!double.TryParse(cellVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                            !(val >= double.Parse(part.Substring(2), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (part.StartsWith("<="))
                    {
                        if (!double.TryParse(cellVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                            !(val <= double.Parse(part.Substring(2), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (part.StartsWith(">"))
                    {
                        if (!double.TryParse(cellVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                            !(val > double.Parse(part.Substring(1), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (part.StartsWith("<"))
                    {
                        if (!double.TryParse(cellVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                            !(val < double.Parse(part.Substring(1), CultureInfo.InvariantCulture))) return false;
                    }
                    else if (part.StartsWith("="))
                    {
                        if (!double.TryParse(cellVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                            !(val == double.Parse(part.Substring(1), CultureInfo.InvariantCulture))) return false;
                    }
                    else // String
                    {
                        if (!cellVal.Equals(part.Trim('"'), StringComparison.OrdinalIgnoreCase))
                            return false;
                    }
                }
            }
            return true;
        }

        /// <summary>
        /// Parses the config row to triplets: (Column, Condition, Logic) for AND/OR logic.
        /// </summary>
        public static List<(string Column, string Condition, string Logic)> GetConditionTriplets(Dictionary<string, string> cfg)
        {
            var result = new List<(string, string, string)>();
            int idx = 1;
            while (cfg.ContainsKey($"Column{idx}") && !string.IsNullOrWhiteSpace(cfg[$"Column{idx}"]))
            {
                string column = cfg[$"Column{idx}"].Trim();
                string condition = cfg.GetValueOrDefault($"Condition{idx}", "").Trim();
                string logic = cfg.GetValueOrDefault($"Logic{idx}", "").Trim().ToUpper(); // AND/OR/blank

                result.Add((column, condition, logic));
                idx++;
            }
            return result;
        }

        /// <summary>
        /// Evaluates all conditions for a row using full AND/OR logic, 
        /// and supports multiple conditions per column (e.g., >10 and <40).
        /// </summary>
        public static bool RowMatchesWithLogic(Dictionary<string, string> row, List<(string Column, string Condition, string Logic)> condTriplets)
        {
            if (condTriplets.Count == 0) return true;

            // Group by column for ANDs (multiple conditions on same column)
            var groups = new List<(List<(string Column, string Condition)>, string Logic)>();
            int idx = 0;
            while (idx < condTriplets.Count)
            {
                string currentCol = condTriplets[idx].Column;
                string logicToNext = condTriplets[idx].Logic;
                var block = new List<(string Column, string Condition)>();
                block.Add((condTriplets[idx].Column, condTriplets[idx].Condition));
                // Collect all subsequent triplets for the same column with AND
                int j = idx + 1;
                while (j < condTriplets.Count && condTriplets[j].Column == currentCol && condTriplets[j - 1].Logic == "AND")
                {
                    block.Add((condTriplets[j].Column, condTriplets[j].Condition));
                    logicToNext = condTriplets[j].Logic;
                    j++;
                }
                groups.Add((block, logicToNext));
                idx = j;
            }

            // Evaluate each group (AND within group, chain by Logic)
            bool result = true;
            for (int g = 0; g < groups.Count; g++)
            {
                var (block, logic) = groups[g];
                bool groupResult = true;
                foreach (var (col, cond) in block)
                {
                    if (!EvaluateCondition(row, col, cond))
                    {
                        groupResult = false;
                        break;
                    }
                }
                if (g == 0)
                {
                    result = groupResult;
                }
                else
                {
                    string prevLogic = groups[g - 1].Item2;
                    if (prevLogic == "OR")
                        result = result || groupResult;
                    else // AND or blank
                        result = result && groupResult;
                }
            }
            return result;
        }

        /// <summary>
        /// Evaluates a single condition (numeric or string match).
        /// Handles numeric ops (>, <, >=, <=, =) and string equality ("WOW").
        /// </summary>
        public static bool EvaluateCondition(Dictionary<string, string> row, string column, string condition)
        {
            if (!row.ContainsKey(column))
                return false;
            string cellVal = row[column]?.Trim();

            foreach (var part in condition.Split(',')) // support comma-separated multiple conditions, all must pass
            {
                var cond = part.Trim();
                if (string.IsNullOrEmpty(cond)) continue;

                // Numeric operators with trim
                if (cond.StartsWith(">="))
                {
                    string testVal = cellVal.Trim();
                    string compareVal = cond.Substring(2).Trim();
                    if (!double.TryParse(testVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                        !(val >= double.Parse(compareVal, CultureInfo.InvariantCulture))) return false;
                }
                else if (cond.StartsWith("<="))
                {
                    string testVal = cellVal.Trim();
                    string compareVal = cond.Substring(2).Trim();
                    if (!double.TryParse(testVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                        !(val <= double.Parse(compareVal, CultureInfo.InvariantCulture))) return false;
                }
                else if (cond.StartsWith(">"))
                {
                    string testVal = cellVal.Trim();
                    string compareVal = cond.Substring(1).Trim();
                    if (!double.TryParse(testVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                        !(val > double.Parse(compareVal, CultureInfo.InvariantCulture))) return false;
                }
                else if (cond.StartsWith("<"))
                {
                    string testVal = cellVal.Trim();
                    string compareVal = cond.Substring(1).Trim();
                    if (!double.TryParse(testVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                        !(val < double.Parse(compareVal, CultureInfo.InvariantCulture))) return false;
                }
                else if (cond.StartsWith("="))
                {
                    string testVal = cellVal.Trim();
                    string compareVal = cond.Substring(1).Trim();
                    if (!double.TryParse(testVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double val) ||
                        !(val == double.Parse(compareVal, CultureInfo.InvariantCulture))) return false;
                }
                else // String equality (e.g., WOW, etc.)
                {
                    if (!cellVal.Equals(cond.Trim('"'), StringComparison.OrdinalIgnoreCase))
                        return false;
                }
            }
            return true;
        }


        /// <summary>
        /// Builds a readable AND/OR logic condition string from phase config row for the Condition worksheet.
        /// </summary>
        public static string BuildHumanReadableConditionString(Dictionary<string, string> cfg)
        {
            List<string> conds = new List<string>();
            int idx = 1;
            while (cfg.ContainsKey($"Column{idx}") && !string.IsNullOrWhiteSpace(cfg[$"Column{idx}"]))
            {
                string col = cfg[$"Column{idx}"];
                string cond = cfg.GetValueOrDefault($"Condition{idx}", "");
                string logic = cfg.GetValueOrDefault($"Logic{idx}", "");
                string rendered = "";

                // Special case: 'WOW' as string match
                if (!string.IsNullOrEmpty(cond) && cond.Trim().ToUpper() == "WOW")
                {
                    rendered = $"{col} == \"WOW\"";
                }
                else if (!string.IsNullOrEmpty(cond) && (cond.StartsWith(">") || cond.StartsWith("<") || cond.StartsWith("==") || cond.StartsWith("!=")))
                {
                    rendered = $"{col} {cond}";
                }
                else if (!string.IsNullOrEmpty(cond))
                {
                    rendered = $"{col} == {cond}";
                }
                else
                {
                    rendered = $"{col}";
                }

                conds.Add(rendered);
                // Append logic if present and not last
                if (!string.IsNullOrWhiteSpace(logic))
                {
                    if (logic.Trim().ToUpper() == "AND")
                        conds.Add("&&");
                    else if (logic.Trim().ToUpper() == "OR")
                        conds.Add("||");
                }
                idx++;
            }
            // Remove trailing logic if any
            while (conds.Count > 0 && (conds.Last() == "&&" || conds.Last() == "||"))
                conds.RemoveAt(conds.Count - 1);
            return string.Join(" ", conds);
        }

        /// <summary>
        /// For each row, checks all phase "Start" conditions (with FilterTime) in configLines,
        /// and assigns the Phase column with all matching phases joined with " & " (e.g., "Engine On & Taxi Out").
        /// This supports parallel/overlapping phase labeling for FOQA/FDR data,
        /// and returns Condition sheet info as before (for backward compatibility).
        /// </summary>
        public static List<(string Phase, string Duration, string FilterTime, string ConditionString, int MatchCount)>
            AssignPhasesAndBuildConditionRows(
                List<Dictionary<string, string>> configLines,
                List<Dictionary<string, string>> dataRows
            )
        {
            // --- List for Condition worksheet (start/end conditions as in config)
            var conditionRows = new List<(string Phase, string Duration, string FilterTime, string ConditionString, int MatchCount)>();

            // --- Parallel phase labeling (overlap-aware): mark all Start conditions satisfied at each row
            for (int rowIdx = 0; rowIdx < dataRows.Count; rowIdx++)
            {
                List<string> matchedPhases = new List<string>();

                // Check all config Start rows for this row
                foreach (var cfg in configLines)
                {
                    string phase = cfg.GetValueOrDefault("FlightPhase") ?? "";
                    string duration = cfg.GetValueOrDefault("Duration") ?? "";
                    if (!duration.Equals("Start", StringComparison.OrdinalIgnoreCase))
                        continue;

                    string filterTimeStr = cfg.GetValueOrDefault("FilterTime") ?? "1";
                    int filterTime = 1;
                    int.TryParse(filterTimeStr, out filterTime);
                    if (filterTime < 1) filterTime = 1;

                    var condTriplets = GetConditionTriplets(cfg);

                    if (rowIdx <= dataRows.Count - filterTime)
                    {
                        bool allMatch = true;
                        for (int k = 0; k < filterTime; k++)
                        {
                            if (!RowMatchesWithLogic(dataRows[rowIdx + k], condTriplets))
                            {
                                allMatch = false;
                                break;
                            }
                        }
                        if (allMatch)
                            matchedPhases.Add(phase);
                    }
                }

                // Write all matching phases joined by " & " (or empty if none)
                dataRows[rowIdx]["Phase"] = matchedPhases.Count > 0 ? string.Join(" & ", matchedPhases) : "";
            }

            // --- Build Condition sheet rows (Start and End for all phases, always included even if never detected)
            foreach (var cfg in configLines)
            {
                string phase = cfg.GetValueOrDefault("FlightPhase") ?? "";
                string duration = cfg.GetValueOrDefault("Duration") ?? "";
                string filterTimeStr = cfg.GetValueOrDefault("FilterTime") ?? "1";
                string condStr = BuildHumanReadableConditionString(cfg);

                // Count block matches for this config row (for QA/summary)
                int matchCount = 0;
                var condTriplets = GetConditionTriplets(cfg);
                int filterTime = 1;
                int.TryParse(filterTimeStr, out filterTime);
                if (filterTime < 1) filterTime = 1;
                for (int i = 0; i <= dataRows.Count - filterTime;)
                {
                    bool allMatch = true;
                    for (int k = 0; k < filterTime; k++)
                    {
                        if (!RowMatchesWithLogic(dataRows[i + k], condTriplets))
                        {
                            allMatch = false;
                            break;
                        }
                    }
                    if (allMatch)
                    {
                        matchCount++;
                        i += filterTime;
                    }
                    else
                    {
                        i++;
                    }
                }

                conditionRows.Add((phase, duration, filterTimeStr, condStr, matchCount));
            }

            return conditionRows;
        }



        /// <summary>
        /// Loads, processes, and tags phases using robust single-pass helper logic.
        /// Ensures numbers are written as numbers and strings as strings in Excel.
        /// Adds a Condition worksheet and Sector column (before Phase).
        /// </summary>
        public static void PhaseFilterToExcel(string configCsvPath, string dataCsvPath, string outputExcelPath)
        {
            // --- Load all data rows (dictionary per row)
            var dataRows = ExceedanceFilterApp.Program.LoadCsvToList(dataCsvPath);

            // --- Always ensure output columns present and clean (remove others)
            foreach (var row in dataRows)
            {
                row["Altitude Rate"] = "";
                row["Pitch Rate"] = "";
                row["Sector"] = "";
                row["Phase"] = "";
            }

            // --- STEP 1: Calculate Altitude Rate and Pitch Rate
            for (int i = 1; i < dataRows.Count; i++)
            {
                double altCurr = 0, altPrev = 0, pitchCurr = 0, pitchPrev = 0;
                bool hasAltCurr = double.TryParse(dataRows[i].GetValueOrDefault("Pressure Altitude 1", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out altCurr);
                bool hasAltPrev = double.TryParse(dataRows[i - 1].GetValueOrDefault("Pressure Altitude 1", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out altPrev);
                bool hasPitchCurr = double.TryParse(dataRows[i].GetValueOrDefault("Pitch Attitude (pitch angle) INERTIAL CHA", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out pitchCurr);
                bool hasPitchPrev = double.TryParse(dataRows[i - 1].GetValueOrDefault("Pitch Attitude (pitch angle) INERTIAL CHA", ""), NumberStyles.Any, CultureInfo.InvariantCulture, out pitchPrev);

                dataRows[i]["Altitude Rate"] = (hasAltCurr && hasAltPrev) ? (altCurr - altPrev).ToString("F2", CultureInfo.InvariantCulture) : "";
                dataRows[i]["Pitch Rate"] = (hasPitchCurr && hasPitchPrev) ? (pitchCurr - pitchPrev).ToString("F2", CultureInfo.InvariantCulture) : "";
            }
            if (dataRows.Count > 0)
            {
                dataRows[0]["Altitude Rate"] = "0.00";
                dataRows[0]["Pitch Rate"] = "0.00";
            }

            // --- STEP 2: Fill sector column based on WOW transitions
            FillSectorByWOWTransitions(dataRows);

            // --- STEP 3: Assign phases and build Condition worksheet info in a single pass
            var configLines = ExceedanceFilterApp.Program.LoadCsvToList(configCsvPath);
            var conditionRows = AssignPhasesAndBuildConditionRows(configLines, dataRows);

            // --- Build final header list: all input columns in original order, then the new columns (if not present already)
            var extraCols = new[] { "Altitude Rate", "Pitch Rate", "Sector", "Phase" };
            var allHeaders = dataRows.Count > 0
                ? dataRows[0].Keys.ToList()
                : extraCols.ToList();
            foreach (var col in extraCols)
                if (!allHeaders.Contains(col)) allHeaders.Add(col);

            // --- Write to Excel: each value as number or string as per content
            using var workbook = new ClosedXML.Excel.XLWorkbook();
            var ws = workbook.Worksheets.Add("FlightData");

            // --- Write header row
            for (int col = 0; col < allHeaders.Count; col++)
            {
                ws.Cell(1, col + 1).Value = allHeaders[col];
                ws.Cell(1, col + 1).Style.Font.Bold = true;
            }

            // --- Write data rows: write numbers as numbers, others as string
            for (int r = 0; r < dataRows.Count; r++)
            {
                for (int c = 0; c < allHeaders.Count; c++)
                {
                    string value = dataRows[r].GetValueOrDefault(allHeaders[c], "");
                    // Try to parse as number using InvariantCulture
                    if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double numVal))
                    {
                        ws.Cell(r + 2, c + 1).Value = numVal;
                    }
                    else
                    {
                        ws.Cell(r + 2, c + 1).Value = value;
                    }
                }
            }

            // --- Create Condition worksheet: one row per config entry (Start and End), always present
            var wsCond = workbook.Worksheets.Add("Condition");
            wsCond.Cell(1, 1).Value = "FlightPhase";
            wsCond.Cell(1, 2).Value = "Duration";
            wsCond.Cell(1, 3).Value = "FilterTime (s)";
            wsCond.Cell(1, 4).Value = "Condition";
            wsCond.Cell(1, 5).Value = "Matched Row Count";
            wsCond.Row(1).Style.Font.Bold = true;

            int condRow = 2;
            foreach (var cfg in configLines)
            {
                string phase = cfg.GetValueOrDefault("FlightPhase") ?? "";
                string duration = cfg.GetValueOrDefault("Duration") ?? "";
                string filterTimeStr = cfg.GetValueOrDefault("FilterTime") ?? "1";
                string condStr = BuildHumanReadableConditionString(cfg);

                // --- Count block matches in data (optional, for QA)
                int matchCount = 0;
                var condTriplets = GetConditionTriplets(cfg);
                int filterTime = 1;
                int.TryParse(filterTimeStr, out filterTime);
                if (filterTime < 1) filterTime = 1;
                for (int i = 0; i <= dataRows.Count - filterTime;)
                {
                    bool allMatch = true;
                    for (int k = 0; k < filterTime; k++)
                    {
                        if (!RowMatchesWithLogic(dataRows[i + k], condTriplets))
                        {
                            allMatch = false;
                            break;
                        }
                    }
                    if (allMatch)
                    {
                        matchCount++;
                        i += filterTime;
                    }
                    else
                    {
                        i++;
                    }
                }

                wsCond.Cell(condRow, 1).Value = phase;
                wsCond.Cell(condRow, 2).Value = duration;
                wsCond.Cell(condRow, 3).Value = filterTimeStr;
                wsCond.Cell(condRow, 4).Value = condStr;
                wsCond.Cell(condRow, 5).Value = matchCount;
                condRow++;
            }


            workbook.SaveAs(outputExcelPath);
        }

        /// <summary>
        /// Build robust condition rows for Condition worksheet, including block match count.
        /// </summary>
        public static List<(string Phase, string Duration, string FilterTime, string ConditionString, int MatchCount)>
            BuildConditionRows(List<Dictionary<string, string>> configLines, List<Dictionary<string, string>> dataRows)
        {
            var conditionRows = new List<(string Phase, string Duration, string FilterTime, string ConditionString, int MatchCount)>();

            foreach (var cfg in configLines)
            {
                string phase = cfg.GetValueOrDefault("FlightPhase") ?? "";
                string duration = cfg.GetValueOrDefault("Duration") ?? "";
                string filterTimeStr = cfg.GetValueOrDefault("FilterTime") ?? "1";
                int filterTime = 1;
                int.TryParse(filterTimeStr, out filterTime);
                if (filterTime < 1) filterTime = 1;

                // --- Use new GetConditionPairsSimple for AllConditionsPass
                var condPairs = GetConditionTriplets(cfg);
                string condStr = BuildHumanReadableConditionString(cfg);

                // --- Count block matches in data (non-overlapping)
                int matchCount = 0;
                if (condPairs.Count > 0)
                {
                    for (int i = 0; i <= dataRows.Count - filterTime;)
                    {
                        bool allMatch = true;
                        for (int k = 0; k < filterTime; k++)
                        {
                            if (!RowMatchesWithLogic(dataRows[i + k], condPairs))
                            {
                                allMatch = false;
                                break;
                            }
                        }
                        if (allMatch)
                        {
                            matchCount++;
                            i += filterTime;
                        }
                        else
                        {
                            i++;
                        }
                    }
                }
                conditionRows.Add((phase, duration, filterTimeStr, condStr, matchCount));
            }
            return conditionRows;
        }

        /// <summary>
        /// Assigns sector labels ("NPS"/"NFS") only if a No WOW segment contains at least 60 rows of actual "No WOW".
        /// Blanks in the WOW column are ignored for threshold (do NOT count toward the 60).
        /// Segments are broken by WOW or 30 fully empty rows. No timestamp/duration required.
        /// </summary>
        public static void FillSectorByWOWTransitions(List<Dictionary<string, string>> dataRows)
        {
            // --- Ensure "Sector" column exists and is cleared for all rows
            foreach (var row in dataRows)
                row["Sector"] = "";

            const int MinNoWOWRowsForSector = 45;   // Minimum No WOW count for sector (FS/PS)
            const int MinTGWowRows = 5, MaxTGWowRows = 45; // Min/Max WOW rows for TG (now 45 inclusive)
            int sectorCount = 0, partialCount = 0, tgCount = 0;
            int rowIndex = 0;
            bool firstSectorIsPartial = false;

            while (rowIndex < dataRows.Count)
            {
                // --- 1. Find start of next No WOW segment
                while (rowIndex < dataRows.Count)
                {
                    string val = dataRows[rowIndex].GetValueOrDefault("WOW Discrete 1 LH MLG 1 on ground", "").Trim();
                    if (val.Equals("No WOW", StringComparison.OrdinalIgnoreCase))
                        break;
                    rowIndex++;
                }
                if (rowIndex >= dataRows.Count)
                    break; // No more No WOW found

                int startIdx = rowIndex;

                // --- 2. Find the end of this No WOW segment (broken by WOW or 30 fully empty rows)
                int endIdx = startIdx;
                int emptyStreak = 0;
                for (int i = startIdx + 1; i < dataRows.Count; i++)
                {
                    // Check for 30 fully empty rows
                    bool isFullyEmpty = true;
                    foreach (var val in dataRows[i].Values)
                    {
                        if (!string.IsNullOrWhiteSpace(val))
                        {
                            isFullyEmpty = false;
                            break;
                        }
                    }
                    if (isFullyEmpty)
                    {
                        emptyStreak++;
                        if (emptyStreak >= 30)
                        {
                            endIdx = i - emptyStreak;
                            break;
                        }
                    }
                    else
                    {
                        emptyStreak = 0; // Reset streak
                        string wowVal = dataRows[i].GetValueOrDefault("WOW Discrete 1 LH MLG 1 on ground", "").Trim();
                        if (wowVal.Equals("WOW", StringComparison.OrdinalIgnoreCase))
                        {
                            endIdx = i - 1;
                            break;
                        }
                        if (i == dataRows.Count - 1)
                            endIdx = i; // Last row
                    }
                }

                // --- 3. Count No WOW rows in this segment
                int noWOWRowCount = 0;
                for (int j = startIdx; j <= endIdx; j++)
                {
                    string v = dataRows[j].GetValueOrDefault("WOW Discrete 1 LH MLG 1 on ground", "").Trim();
                    if (v.Equals("No WOW", StringComparison.OrdinalIgnoreCase))
                        noWOWRowCount++;
                }

                // --- 4. Only assign sector if at least 45 actual "No WOW" rows in segment
                if (noWOWRowCount >= MinNoWOWRowsForSector)
                {
                    // --- Find last meaningful value before the segment for PS/FS logic
                    int prevIdx = startIdx - 1;
                    string prevVal = "";
                    while (prevIdx >= 0)
                    {
                        prevVal = dataRows[prevIdx].GetValueOrDefault("WOW Discrete 1 LH MLG 1 on ground", "").Trim();
                        if (!string.IsNullOrEmpty(prevVal))
                            break;
                        prevIdx--;
                    }

                    bool isFirstPartial = (prevIdx < 0);
                    bool isFullSector = (!isFirstPartial && prevVal.Equals("WOW", StringComparison.OrdinalIgnoreCase));
                    bool isLastPartial = (endIdx == dataRows.Count - 1);

                    string label;
                    if (isFirstPartial)
                    {
                        partialCount++;
                        label = $"{partialCount}PS";
                        firstSectorIsPartial = true;
                    }
                    else if (isLastPartial)
                    {
                        partialCount++;
                        int psLabelNum = firstSectorIsPartial ? sectorCount + partialCount : partialCount;
                        label = $"{psLabelNum}PS";
                    }
                    else if (isFullSector)
                    {
                        sectorCount++;
                        int fsLabelNum = firstSectorIsPartial ? sectorCount + partialCount : sectorCount;
                        label = $"{fsLabelNum}FS";
                    }
                    else
                    {
                        partialCount++;
                        label = $"{partialCount}PS";
                    }

                    // --- 5. Assign label to all No WOW and blank rows in the segment (skip fully empty rows)
                    for (int j = startIdx; j <= endIdx; j++)
                    {
                        bool isFullyEmpty = true;
                        foreach (var val in dataRows[j].Values)
                        {
                            if (!string.IsNullOrWhiteSpace(val))
                            {
                                isFullyEmpty = false;
                                break;
                            }
                        }
                        if (isFullyEmpty)
                            continue;

                        string v = dataRows[j].GetValueOrDefault("WOW Discrete 1 LH MLG 1 on ground", "").Trim();
                        if (v.Equals("No WOW", StringComparison.OrdinalIgnoreCase) || string.IsNullOrEmpty(v))
                            dataRows[j]["Sector"] = label;
                    }

                    // --- 6. Check for TG (Touch-and-Go) in following rows before the next long No WOW sector
                    int tgIdx = endIdx + 1;
                    while (tgIdx < dataRows.Count)
                    {
                        // Look for a WOW run
                        int wowStart = -1, wowLen = 0;
                        while (tgIdx < dataRows.Count)
                        {
                            string wowVal = dataRows[tgIdx].GetValueOrDefault("WOW Discrete 1 LH MLG 1 on ground", "").Trim();
                            if (wowVal.Equals("WOW", StringComparison.OrdinalIgnoreCase))
                            {
                                if (wowStart == -1) wowStart = tgIdx;
                                wowLen++;
                            }
                            else if (wowStart != -1)
                                break;
                            tgIdx++;
                        }
                        // TG: 5–45 WOW rows (inclusive), otherwise ignore
                        if (wowLen >= MinTGWowRows && wowLen <= MaxTGWowRows)
                        {
                            tgCount++;
                            string tgLabel = $"{tgCount}TG";
                            for (int t = wowStart; t < wowStart + wowLen; t++)
                                dataRows[t]["Sector"] = tgLabel;
                            // No WOW after TG, if it's long enough, continue as same sector
                            int afterTG = wowStart + wowLen;
                            int noWOWLen = 0, afterNoWOW = -1;
                            for (int k = afterTG; k < dataRows.Count; k++)
                            {
                                string nwVal = dataRows[k].GetValueOrDefault("WOW Discrete 1 LH MLG 1 on ground", "").Trim();
                                if (nwVal.Equals("No WOW", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (afterNoWOW == -1) afterNoWOW = k;
                                    noWOWLen++;
                                }
                                else if (afterNoWOW != -1)
                                    break;
                            }
                            // Continue same sector after TG if enough No WOW
                            if (noWOWLen >= MinNoWOWRowsForSector)
                            {
                                for (int t = afterNoWOW; t < afterNoWOW + noWOWLen; t++)
                                    dataRows[t]["Sector"] = label;
                                tgIdx = afterNoWOW + noWOWLen;
                            }
                            else
                            {
                                tgIdx = (afterNoWOW == -1 ? tgIdx : afterNoWOW + noWOWLen);
                            }
                        }
                        // If next is a long WOW (≥45), break and move to next sector
                        else if (wowLen >= MinNoWOWRowsForSector)
                        {
                            rowIndex = wowStart + wowLen;
                            break;
                        }
                        else // For short WOW blips, just skip (sector not broken)
                        {
                            tgIdx += 1;
                        }
                    }
                    // --- Done with this sector and TG (if any)
                    rowIndex = tgIdx;
                }
                else
                {
                    // Advance to next possible segment if segment too short for sector
                    rowIndex = endIdx + 1;
                }
            }
        }

        /// <summary>
        /// Helper to safely parse doubles from string dictionary (handles missing/blank).
        /// </summary>
        private static double TryParseDouble(string value)
        {
            if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double d))
                return d;
            return double.NaN;
        }

    }
}

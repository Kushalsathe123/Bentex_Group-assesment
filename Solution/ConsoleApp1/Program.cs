using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AchReconParser
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            string inputFilePath = "C:\\Users\\kusha\\Desktop\\New folder (5)\\ConsoleApp1\\ConsoleApp1\\BAI2 Sample ACH Recon Services ITEM LEVEL (1).txt";
            string outputFilePath = "C:\\Users\\kusha\\Desktop\\New folder (5)\\ConsoleApp1\\ConsoleApp1\\Output.xlsx";

            try
            {
                string fileContent = File.ReadAllText(inputFilePath);

                fileContent = fileContent.Replace("\r\n", "\n").Replace("\r", "\n");
                string[] allLines = fileContent.Split(new[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);

                var headerData = ParseHeaderData(allLines);
                var detailData = ParseDetailData(allLines);

                CreateExcelFile(outputFilePath, headerData, detailData);

                Console.WriteLine($"Excel file created successfully at: {outputFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        static Dictionary<string, string> ParseHeaderData(string[] lines)
        {
            var headerData = new Dictionary<string, string>();

            var line01 = lines.FirstOrDefault(l => l.StartsWith("01,"));
            if (line01 != null)
            {
                var parts = line01.Split(',');
                headerData.Add("0101", parts[0]); 
                headerData.Add("0102", parts[1]); 
                headerData.Add("0103", parts[2]); 
                headerData.Add("0104", parts[3]); 
            }

            var line02 = lines.FirstOrDefault(l => l.StartsWith("02,"));
            if (line02 != null)
            {
                var parts = line02.Split(',');
                headerData.Add("0201", parts[0]); 
                headerData.Add("0202", parts[1]); 
                headerData.Add("0203", parts[2]); 
                headerData.Add("0204", parts[3]); 
                headerData.Add("0205", parts[4]); 
            }

            var line03 = lines.FirstOrDefault(l => l.StartsWith("03,"));
            if (line03 != null)
            {
                var parts = line03.Split(',');
                headerData.Add("0301", parts[0]); 
                headerData.Add("0302", parts[1]);
                headerData.Add("0303", parts[2]); 
            }

            return headerData;
        }

        static List<Dictionary<string, string>> ParseDetailData(string[] lines)
        {
            var detailRecords = new List<Dictionary<string, string>>();
            Dictionary<string, string> currentRecord = null;
            StringBuilder continuationBuilder = new StringBuilder();

            for (int i = 0; i < lines.Length; i++)
            {
                string currentLine = lines[i].Trim();

                if (currentLine.StartsWith("16,"))
                {
                    if (currentRecord != null)
                    {
                        if (continuationBuilder.Length > 0)
                        {
                            ProcessContinuationData(currentRecord, continuationBuilder.ToString());
                            continuationBuilder.Clear();
                        }

                        detailRecords.Add(currentRecord);
                    }

                    currentRecord = new Dictionary<string, string>();
                    var parts = currentLine.Split(',');

                    currentRecord["Code"] = parts[0];
                    currentRecord["Transaction Code_02"] = parts.Length > 1 ? parts[1] : "";
                    currentRecord["Amount_03"] = parts.Length > 2 ? parts[2] : "";

                    if (parts.Length > 3 && parts[3].Length > 0 && char.IsLetter(parts[3][0]))
                    {
                        currentRecord["16_04"] = parts[3];
                        currentRecord["16_05"] = parts.Length > 4 ? parts[4] : "";

                        if (parts.Length > 7)
                        {
                            currentRecord["Batch"] = parts[7];
                        }
                    }
                    else
                    {
                        currentRecord["16_04"] = ""; 

                        currentRecord["16_05"] = parts.Length > 3 ? parts[3] : "";

                        if (parts.Length > 4 && parts[4].StartsWith("BATCH"))
                        {
                            currentRecord["Batch"] = parts[4];
                        }
                        else
                        {
                            for (int j = 3; j < parts.Length; j++)
                            {
                                if (parts[j].StartsWith("BATCH"))
                                {
                                    currentRecord["Batch"] = parts[j];
                                    break;
                                }
                            }
                        }
                    }
                }
                else if (currentLine.StartsWith("88,") && currentRecord != null)
                {
                    continuationBuilder.Append(currentLine.Substring(3)); 
                    continuationBuilder.Append(" ");
                }
            }

            if (currentRecord != null)
            {
                if (continuationBuilder.Length > 0)
                {
                    ProcessContinuationData(currentRecord, continuationBuilder.ToString());
                }
                detailRecords.Add(currentRecord);
            }

            EnsureAllFieldsPresent(detailRecords);

            return detailRecords;
        }

        static void ProcessContinuationData(Dictionary<string, string> record, string continuationData)
        {
            continuationData = continuationData.Replace("88,", " ").Trim();

            var indNamePattern = @"IND NAME=([^,]+)";
            var indNameMatch = Regex.Match(continuationData, indNamePattern);
            if (indNameMatch.Success)
            {
                string indName = indNameMatch.Groups[1].Value.Trim();
                indName = indName.Replace("88", "");
                record["IND NAME"] = indName;
            }

            ExtractFieldWithinRecord(continuationData, @"DFI BANK=([^,]+)", "DFI BANK", record);
            ExtractFieldWithinRecord(continuationData, @"DFI ACCT=([^,]+)", "DFI ACCT", record);
            ExtractFieldWithinRecord(continuationData, @"IND ID NO=([^,]+)", "IND ID NO", record);
            ExtractFieldWithinRecord(continuationData, @"TRACE NO=([^,]+)", "TRACE NO", record);

            var batchPattern = @"BATCH NUMBER=BATCH([0-9]+)(?=\s|,|$|SETT)";
            var batchMatch = Regex.Match(continuationData, batchPattern);

            if (batchMatch.Success)
            {
                string batchNumber = "BATCH" + batchMatch.Groups[1].Value.Trim();
                record["BATCH NUMBER"] = batchNumber;
            }
            else
            {
                var altBatchPattern = @"BATCH([0-9]+)(?=\s|,|$|SETT)";
                var altMatch = Regex.Match(continuationData, altBatchPattern);

                if (altMatch.Success)
                {
                    string batchNumber = "BATCH" + altMatch.Groups[1].Value.Replace(" ", "");
                    record["BATCH NUMBER"] = batchNumber;
                }
            }

            if (record.TryGetValue("BATCH NUMBER", out string batchVal) &&
                batchVal.StartsWith("BATCH") &&
                Regex.IsMatch(continuationData, @"\b" + Regex.Escape(batchVal) + @"\s+(\d+)(?=\s|,|$|SETT)"))
            {
                var numMatch = Regex.Match(continuationData, @"\b" + Regex.Escape(batchVal) + @"\s+(\d+)(?=\s|,|$|SETT)");
                if (numMatch.Success)
                {
                    record["BATCH NUMBER"] = batchVal + numMatch.Groups[1].Value;
                }
            }
        

        ExtractFieldWithinRecord(continuationData, @"SETT BANKREF=([^,]+)", "SETT BANKREF", record);
            string cleanContinuationData = continuationData.Replace("88,", " ").Replace("\n", " ").Replace("\r", " ");

            var custRefMatch = Regex.Match(cleanContinuationData, @"SETT CUSTREF=([^,]+?)(SETT AMOUNT|$)");
            if (custRefMatch.Success)
            {
                string custRef = custRefMatch.Groups[1].Value.Trim();
                record["SETT CUSTREF"] = custRef;
            }

            var settAmountMatch = Regex.Match(cleanContinuationData, @"SETT AMOUNT=([0-9.]+)");
            if (settAmountMatch.Success)
            {
                record["SETT AMOUNT"] = settAmountMatch.Groups[1].Value.Trim();
            }
        }

        static void ExtractFieldWithinRecord(string data, string pattern, string fieldName, Dictionary<string, string> record)
        {
            var match = Regex.Match(data, pattern);
            if (match.Success)
            {
                record[fieldName] = match.Groups[1].Value.Trim();
            }
        }

        static void EnsureAllFieldsPresent(List<Dictionary<string, string>> records)
        {
            HashSet<string> allFields = new HashSet<string>();
            foreach (var record in records)
            {
                foreach (var key in record.Keys)
                {
                    allFields.Add(key);
                }
            }

            foreach (var record in records)
            {
                foreach (var field in allFields)
                {
                    if (!record.ContainsKey(field))
                    {
                        record[field] = "";
                    }
                }
            }
        }

        static void CreateExcelFile(string filePath, Dictionary<string, string> headerData, List<Dictionary<string, string>> detailData)
        {
            using (var package = new ExcelPackage())
            {
                var headerSheet = package.Workbook.Worksheets.Add("Header");

                int headerCol = 1;
                foreach (var key in headerData.Keys)
                {
                    headerSheet.Cells[1, headerCol].Value = key;
                    headerCol++;
                }

                headerCol = 1;
                foreach (var key in headerData.Keys)
                {
                    headerSheet.Cells[2, headerCol].Value = headerData[key];
                    headerCol++;
                }

                headerSheet.Cells[1, 1, 1, headerData.Count].Style.Font.Bold = true;
                headerSheet.Cells[1, 1, 2, headerData.Count].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                headerSheet.Cells[1, 1, 2, headerData.Count].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                headerSheet.Cells[1, 1, 2, headerData.Count].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                headerSheet.Cells[1, 1, 2, headerData.Count].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                headerSheet.Cells.AutoFitColumns();

                var detailSheet = package.Workbook.Worksheets.Add("Detail");

                string[] orderedColumns = {
                    "Code", "Transaction Code_02", "Amount_03", "16_04", "16_05", "Batch",
                    "DFI BANK", "DFI ACCT", "IND ID NO", "IND NAME", "TRACE NO",
                    "BATCH NUMBER", "SETT BANKREF", "SETT CUSTREF", "SETT AMOUNT"
                };

                int col = 1;
                foreach (var column in orderedColumns)
                {
                    detailSheet.Cells[1, col].Value = column;
                    col++;
                }

                for (int i = 0; i < detailData.Count; i++)
                {
                    var record = detailData[i];
                    col = 1;
                    foreach (var column in orderedColumns)
                    {
                        detailSheet.Cells[i + 2, col].Value = record.ContainsKey(column) ? record[column] : "";
                        col++;
                    }
                }

                for (int i = 0; i < detailData.Count; i++)
                {
                    var record = detailData[i];
                    col = 1;
                    foreach (var column in orderedColumns)
                    {
                        string cellValue = record.ContainsKey(column) ? record[column] : "";

                        if (column != "IND NAME" && !string.IsNullOrEmpty(cellValue))
                        {
                            cellValue = cellValue.Replace(" ", "");
                        }

                        detailSheet.Cells[i + 2, col].Value = cellValue;
                        col++;
                    }
                }

                detailSheet.Cells[1, 1, 1, orderedColumns.Length].Style.Font.Bold = true;
                detailSheet.Cells[1, 1, detailData.Count + 1, orderedColumns.Length].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                detailSheet.Cells[1, 1, detailData.Count + 1, orderedColumns.Length].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                detailSheet.Cells[1, 1, detailData.Count + 1, orderedColumns.Length].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                detailSheet.Cells[1, 1, detailData.Count + 1, orderedColumns.Length].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                detailSheet.Cells.AutoFitColumns();

                package.SaveAs(new FileInfo(filePath));
            }
        }
    }
}
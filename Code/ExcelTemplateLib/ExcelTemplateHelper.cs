using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace ExcelTemplateLib
{
    public static class ExcelTemplateHelper
    {
        public static Stream HandleExcel(string fullFilePath, object obj)
        {
            var result = new MemoryStream();
            XSSFWorkbook wk = null;
            using (FileStream fs = File.Open(fullFilePath, FileMode.Open, FileAccess.Read))
            {
                int fileSize = (int)fs.Length;
                byte[] data = new byte[fileSize];
                fs.Position = 0;
                fs.Read(data, 0, fileSize);
                MemoryStream ms = new MemoryStream(data);

                wk = new XSSFWorkbook(ms);

                ISheet sheet = wk.GetSheetAt(0);
                var rowCount = sheet.LastRowNum + 1;

                for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
                {
                    var cells = GetCellsFromRow(sheet, rowIndex);
                    // Recognize does is loop row
                    // the loop row must bee first cell of row has {!1-n!} as prefix
                    // the value surround "{!" and "!}"
                    // "1-n", the "1" mean only 1 row will do loop, the "n" is mean we will replace "n" as index base 0
                    // "2-m", the "2" mean the 2 row will do loop, the "m" is mean we will replace "n" as index base 0
                    // for example, the first cell of row has "2-m", the other cell has path "{{test[m].name}}"
                    // if the test has 3 item, the first two row will be replace as test[0].name, the second will be test[1].name
                    if (cells.Count > 0)
                    {
                        var firstCellValue = cells[0].StringCellValue;
                        var matched = GetLoopMark(firstCellValue);
                        if (matched == null)
                        {
                            foreach (var cellObj in cells)
                            {
                                HandleCellData(cellObj, obj);
                            }
                        }
                        else
                        {
                            // It is loop start row index
                            var startLoopRowIndex = rowIndex;
                            // parse the loop marked
                            var prefixArr = matched.Split("-".ToCharArray());
                            // row count will be loop
                            var lineCount = int.Parse(prefixArr[0]);
                            // this is array flag that need to loop
                            var flag = prefixArr[1];
                            // if the lines have multiple rows.
                            // get all cell obj from rows

                            for (var i = 1; i < lineCount; i++)
                            {
                                rowIndex++;
                                cells.AddRange(GetCellsFromRow(sheet, rowIndex));
                            }

                            // Get all prop paths
                            var propPaths = GetPropPathsFromCells(cells);
                            // Get how much items need to loop.
                            // Only find first item include [n]
                            var loopObjPath = GetFirstLoopListObjectPath(propPaths, flag);
                            var loopObj = ObjectPathParser.GetDeepPropertyValue(obj, loopObjPath);
                            var loopCount = 0;
                            if (loopObj != null)
                            {
                                if (loopObj is System.Collections.IList)
                                {
                                    loopCount = ((System.Collections.IList)loopObj).Count;
                                }
                                else if (loopObj is Array)
                                {
                                    loopCount = ((Array)loopObj).Length;
                                }
                            }

                            // remove row from excel if loop index equal zore
                            if (loopCount == 0)
                            {
                                rowIndex = startLoopRowIndex - 1;
                                rowCount -= lineCount;
                                for (int i = startLoopRowIndex; i < startLoopRowIndex + lineCount; i++)
                                {
                                    sheet.RemoveRow(sheet.GetRow(i));
                                }
                            }
                            else
                            {
                                for (var innerIndex = 0; innerIndex < lineCount; innerIndex++)
                                {
                                    var sourceIndex = startLoopRowIndex + innerIndex;

                                    if (sheet.GetRow(sourceIndex) == null)
                                    {
                                        sheet.CreateRow(sourceIndex);
                                    }
                                }

                                // copy row first
                                for (var i = 1; i < loopCount; i++)
                                {
                                    for (var innerIndex = 0; innerIndex < lineCount; innerIndex++)
                                    {
                                        var sourceIndex = startLoopRowIndex + innerIndex;
                                        var targetIndex = startLoopRowIndex + (i * lineCount) + innerIndex;

                                        InsertAndCopyRow(sheet, sourceIndex, targetIndex);
                                        //sheet.CopyRow(sourceIndex, targetIndex);

                                    }
                                }

                                rowIndex += (loopCount - 1) * lineCount;
                                rowCount += (loopCount - 1) * lineCount;

                                // fill data 
                                for (var i = 0; i < loopCount; i++)
                                {
                                    for (var innerIndex = 0; innerIndex < lineCount; innerIndex++)
                                    {
                                        var currRowIndex = startLoopRowIndex + (i * lineCount) + innerIndex;
                                        var loopCells = GetCellsFromRow(sheet, currRowIndex);
                                        foreach (var cellObj in loopCells)
                                        {
                                            HandleCellData(cellObj, obj, flag, i);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                wk.Write(result);
                wk.Close();
            }

            return result;
        }

        private static void InsertAndCopyRow(ISheet sheet, int sourceRowIndex, int destRowIndex)
        {
            sheet.ShiftRows(destRowIndex, sheet.LastRowNum, 1, true, true);

            IRow sourceRow = null;
            IRow targetRow = null;
            ICell sourceCell = null;
            ICell targetCell = null;

            short m;

            sourceRow = sheet.GetRow(sourceRowIndex);
            targetRow = sheet.CreateRow(destRowIndex);
            targetRow.HeightInPoints = sourceRow.HeightInPoints;

            for (m = (short)sourceRow.FirstCellNum; m < sourceRow.LastCellNum; m++)
            {

                sourceCell = sourceRow.GetCell(m);
                targetCell = targetRow.CreateCell(m);
                if (sourceCell != null)
                {
                    targetCell.CellStyle = sourceCell.CellStyle;
                    targetCell.SetCellType(sourceCell.CellType);
                    targetCell.SetCellValue(sourceCell.StringCellValue);
                }
            }
        }

        private static string GetFirstLoopListObjectPath(List<string> propPaths, string flag)
        {
            foreach (var propPath in propPaths)
            {
                var matched = GetFirstLoopFlagMatch(propPath, flag);
                if (matched != null)
                {
                    var index = propPath.IndexOf(matched);
                    return propPath.Substring(0, index);
                }
            }

            return null;
        }


        /// <summary>
        /// Get cells object from row index
        /// </summary>
        /// <param name="sheet">excel sheel</param>
        /// <param name="rowIndex">row index</param>
        /// <returns>return cells of special row</returns>
        private static List<ICell> GetCellsFromRow(ISheet sheet, int rowIndex)
        {
            List<ICell> cells = new List<ICell>();
            var row = sheet.GetRow(rowIndex);
            if (row != null)
            {
                var cellCount = row.LastCellNum + 1;

                for (var cellIndex = 0; cellIndex < cellCount; cellIndex++)
                {
                    var cellObj = row.GetCell(cellIndex);
                    if (cellObj != null)
                    {
                        cells.Add(cellObj);
                    }
                }
            }

            return cells;
        }

        private static List<string> GetPropPathsFromCells(List<ICell> cellObjs)
        {
            var results = new List<string>();
            foreach (var cellObj in cellObjs)
            {
                var cellData = cellObj.StringCellValue;
                var propPaths = GetPropPathFromInput(cellData);
                results.AddRange(propPaths);
            }

            return results;
        }

        /// <summary>
        /// 处理单元格的数据
        /// </summary>
        /// <param name="cellObj">单元格对象</param>
        /// <param name="obj">绑定的对象</param>
        private static void HandleCellData(ICell cellObj, object obj)
        {
            var cellData = cellObj.StringCellValue;
            if (!string.IsNullOrWhiteSpace(cellData))
            {
                var propPaths = GetPropPathFromInput(cellData);
                var isSingleProperty = IsSingleProperty(cellData);
                foreach (var propPath in propPaths)
                {
                    var parsedContent = ObjectPathParser.GetDeepPropertyValue(obj, propPath);
                    if (parsedContent != null)
                    {
                        if (isSingleProperty)
                        {
                            if (parsedContent is DateTime)
                            {
                                cellObj.SetCellValue((DateTime)parsedContent);
                            }
                            else if (parsedContent is bool)
                            {
                                cellObj.SetCellValue((bool)parsedContent);
                            }
                            else
                            {
                                cellObj.SetCellValue(parsedContent.ToString());
                            }
                        }
                        else
                        {
                            cellData = cellData.Replace("{{" + propPath + "}}", parsedContent.ToString());
                        }
                    }
                    else
                    {
                        cellData = cellData.Replace("{{" + propPath + "}}", "");
                    }
                }

                if (!isSingleProperty)
                    cellObj.SetCellValue(cellData);
            }
        }


        /// <summary>
        /// 处理Loop的单元格的数据
        /// </summary>
        /// <param name="cellObj">单元格对象</param>
        /// <param name="obj">绑定的对象</param>
        private static void HandleCellData(ICell cellObj, object obj, string loopFlag, int index)
        {
            var cellData = cellObj.StringCellValue;

            if (!string.IsNullOrWhiteSpace(cellData))
            {
                cellData = RemoveLoopPrefixFlag(cellData);
                cellData = ReplaceLoopFlag(cellData, loopFlag, index);

                var propPaths = GetPropPathFromInput(cellData);
                foreach (var propPath in propPaths)
                {
                    var parsedContent = ObjectPathParser.GetDeepPropertyValue(obj, propPath);
                    if (parsedContent != null)
                        cellData = cellData.Replace("{{" + propPath + "}}", parsedContent.ToString());
                    else
                        cellData = cellData.Replace("{{" + propPath + "}}", "");
                }

                cellObj.SetCellValue(cellData);
            }
        }

        /// <summary>
        /// Get surround value frorm {!data!} from prefix of string
        /// </summary>
        /// <param name="input">input string that need match</param>
        /// <returns>the data</returns>
        private static string GetLoopMark(string input)
        {
            if (input == null) input = "";
            input = input.Trim();

            string matchPattern = @"^\{!(?<prefix>[0-9a-zA-Z-]+)!\}";
            Regex matchReg = new Regex(matchPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Match match = matchReg.Match(input);
            while (match.Success)
            {
                return match.Groups["prefix"].Value;
            }

            return null;
        }

        /// <summary>
        /// Get object path from input
        /// the path must surround "{{" and "}}"
        /// </summary>
        /// <param name="input">the input string</param>
        /// <returns>return matched object paths</returns>
        private static List<string> GetPropPathFromInput(string input)
        {
            var results = new List<string>();
            string matchPattern = @"\{\{(?<name>[0-9a-zA-Z.\[\] ]+)\}\}";
            Regex matchReg = new Regex(matchPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Match match = matchReg.Match(input);
            while (match.Success)
            {
                results.Add(match.Groups["name"].Value);
                match = match.NextMatch();
            }

            return results.Distinct().ToList();
        }

        /// <summary>
        /// Does the cell data only bind to property
        /// </summary>
        /// <param name="input">cell data</param>
        /// <returns>true -- only one bind, no other data. others is false</returns>
        private static bool IsSingleProperty(string input)
        {
            string matchPattern = @"\{\{(?<name>[0-9a-zA-Z.\[\] ]+)\}\}";
            Regex matchReg = new Regex(matchPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Match match = matchReg.Match(input);
            while (match.Success)
            {
                return input.Trim() == match.Value.Trim();
            }

            return false;
        }

        private static string GetFirstLoopFlagMatch(string input, string flag)
        {
            string matchPattern = @"\[\s*" + flag + @"\s*\]";
            Regex matchReg = new Regex(matchPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Match match = matchReg.Match(input);
            while (match.Success)
            {
                return match.Value;
            }

            return null;
        }

        private static string ReplaceLoopFlag(string input, string flag, int index)
        {
            string result = input;
            string matchPattern = @"\[\s*" + flag + @"\s*\]";
            Regex matchReg = new Regex(matchPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Match match = matchReg.Match(input);
            while (match.Success)
            {
                var matched = match.Value;
                result = result.Replace(matched, "[" + index.ToString() + "]");
                match = match.NextMatch();
            }

            return result;
        }

        private static string RemoveLoopPrefixFlag(string input)
        {
            if (input == null) input = "";
            input = input.Trim();

            string matchPattern = @"^\{!(?<prefix>[0-9a-zA-Z-]+)!\}";
            Regex matchReg = new Regex(matchPattern, RegexOptions.Multiline | RegexOptions.IgnoreCase);
            Match match = matchReg.Match(input);
            while (match.Success)
            {
                var prefix = match.Value;
                return input.Replace(prefix, "");
            }

            return input;
        }
    }
}

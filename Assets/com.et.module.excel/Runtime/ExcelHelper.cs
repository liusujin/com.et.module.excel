using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ETModel
{
    public class ExcelHelper
    {
        public const string ExcelPath = "../Excel";
        public const string ClientConfigPath = "./Assets/Addressables/Config/";
        public const string ServerConfigPath = "../Config/";

        public static bool IsClient = true;

        private static ExcelMD5Info md5Info;

        private static Dictionary<string, Func<string, string>> converters = new Dictionary<string, Func<string, string>>();

        public static string[] ConveterTypes => converters.Keys.ToArray();

        static ExcelHelper()
        {
            converters["int[]"] = input => string.IsNullOrEmpty(input) ? "[]" : $"[{input}]";
            converters["int32[]"] = input => string.IsNullOrEmpty(input) ? "[]" : $"[{input}]";
            converters["long[]"] = input => string.IsNullOrEmpty(input) ? "[]" : $"[{input}]";

            converters["string[]"] = input => string.IsNullOrEmpty(input) ? "[]" : $"[{string.Join(",", input.Split(new[] { "\n" }, StringSplitOptions.None).Select(p => $"\"{p}\""))}]";

            converters["int"] = input => input;
            converters["int32"] = input => input;
            converters["int64"] = input => input;
            converters["long"] = input => input;
            converters["float"] = input => input;
            converters["double"] = input => input;

            converters["bool"] = input => input.ToLower();

            converters["string"] = input => $"\"{input}\"";
        }

        public static void RegisterConverter(string type, Func<string, string> converter)
        {
            converters[type.ToLower()] = converter;
        }

        public static void ExportAllClass(string exportDirectory, string header)
        {
            if (!Directory.Exists(exportDirectory))
            {
                Directory.CreateDirectory(exportDirectory);
            }

            foreach (string filePath in Directory.GetFiles(ExcelPath))
            {
                if (filePath.Length < 12 || Path.GetFileName(filePath).StartsWith("~") || !filePath.EndsWith("Config.xlsx"))
                {
                    continue;
                }

                ExportClass(filePath, exportDirectory, header);
            }
        }

        public static void ExportClass(string fileName, string exportDirectory, string header)
        {
            XSSFWorkbook xssfWorkbook;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xssfWorkbook = new XSSFWorkbook(file);
            }

            string protoName = Path.GetFileNameWithoutExtension(fileName);
            string exportPath = Path.Combine(exportDirectory, $"{protoName}.cs");
            using (FileStream txt = new FileStream(exportPath, FileMode.Create))
            using (StreamWriter sw = new StreamWriter(txt))
            {
                StringBuilder sb = new StringBuilder();
                ISheet sheet = xssfWorkbook.GetSheetAt(0);
                string appType = sheet.GetRow(0)?.GetCell(0)?.ToString() ?? string.Empty;
                sb.Append(header);

                sb.Append($"\t[Config((int)({appType}))]\n");
                sb.Append($"\tpublic partial class {protoName}Category : AddressableCategory<{protoName}>\n");
                sb.Append("\t{\n");
                sb.Append("\t}\n\n");

                sb.Append($"\tpublic class {protoName}: IConfig\n");
                sb.Append("\t{\n");
                sb.Append("\t\tpublic long Id { get; set; }\n");

                int cellCount = sheet.GetRow(3).LastCellNum;

                for (int i = 2; i < cellCount; i++)
                {
                    string fieldDescription = sheet.GetRow(2)?.GetCell(i)?.ToString() ?? string.Empty;

                    if (fieldDescription.StartsWith("#"))
                    {
                        continue;
                    }
                    //s开头表示这个字段是服务端专用
                    if (fieldDescription.StartsWith("s") && IsClient)
                    {
                        continue;
                    }

                    string fieldName = sheet.GetRow(3)?.GetCell(i)?.ToString() ?? string.Empty;
                    if (fieldName == "Id" || fieldName == "_id")
                    {
                        continue;
                    }

                    string fieldType = sheet.GetRow(4)?.GetCell(i)?.ToString() ?? string.Empty;
                    if (fieldType == "" || fieldName == "")
                    {
                        continue;
                    }

                    sb.Append($"\t\tpublic {fieldType} {fieldName};\n");
                }

                sb.Append("\t}\n");
                sb.Append("}\n");

                sw.Write(sb.ToString());
            }
        }

        public static void ExportAll(string exportDirectory, bool forceRegenerate = true)
        {
            if (!Directory.Exists(exportDirectory))
            {
                Directory.CreateDirectory(exportDirectory);
            }

            string md5File = Path.Combine(ExcelPath, "md5.txt");
            if (!File.Exists(md5File))
            {
                md5Info = new ExcelMD5Info();
            }
            else
            {
                md5Info = JsonHelper.FromJson<ExcelMD5Info>(File.ReadAllText(md5File));
            }

            string[] files = Directory.GetFiles(ExcelPath);
            foreach (string filePath in files)
            {
                if (filePath.Length < 12 || Path.GetFileName(filePath).StartsWith("~") || !filePath.EndsWith("Config.xlsx"))
                {
                    continue;
                }
                string fileName = Path.GetFileName(filePath);
                string oldMD5 = md5Info.Get(fileName);
                string md5 = MD5Helper.FileMD5(filePath);
                md5Info.Add(fileName, md5);
                if (md5 == oldMD5 && !forceRegenerate)
                {
                    continue;
                }

                Export(filePath, exportDirectory);
            }

            File.WriteAllText(md5File, JsonHelper.ToJson(md5Info));
        }

        public static string Export(string fileName)
        {
            XSSFWorkbook xssfWorkbook;
            IFormulaEvaluator formulaEvaluator;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xssfWorkbook = new XSSFWorkbook(file);
                formulaEvaluator = xssfWorkbook.GetCreationHelper().CreateFormulaEvaluator();
                Dictionary<string, IFormulaEvaluator> evaluators = new Dictionary<string, IFormulaEvaluator>
                {
                    {Path.GetFileName(fileName),formulaEvaluator }
                };
                foreach (string linkedFileName in xssfWorkbook.ExternalLinksTable.Select(p => $"{ExcelPath}/{p.LinkedFileName}"))
                {
                    XSSFWorkbook linkedWorkBook;
                    using (FileStream linkedFile = new FileStream(linkedFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        linkedWorkBook = new XSSFWorkbook(linkedFile);
                    }
                    evaluators.Add(Path.GetFileName(linkedFileName), linkedWorkBook.GetCreationHelper().CreateFormulaEvaluator());
                }
                formulaEvaluator.SetupReferencedWorkbooks(evaluators);
            }
            ISheet sheet = xssfWorkbook.GetSheetAt(0);
            int cellCount = sheet.GetRow(3).LastCellNum;
            ExcelCellInfo[] cellInfos = new ExcelCellInfo[cellCount];
            for (int i = 2; i < cellCount; ++i)
            {
                string fieldDesc = GetCellString(sheet, 2, i, formulaEvaluator);
                string fieldName = GetCellString(sheet, 3, i, formulaEvaluator);
                string fieldType = GetCellString(sheet, 4, i, formulaEvaluator);
                cellInfos[i] = new ExcelCellInfo() { Name = fieldName, Type = fieldType, Description = fieldDesc };
            }

            StringBuilder sb = new StringBuilder();
            for (int rowIndex = 5; rowIndex <= sheet.LastRowNum; ++rowIndex)
            {
                string id = GetCellString(sheet, rowIndex, 2, formulaEvaluator);
                if (string.IsNullOrEmpty(id))
                {
                    continue;
                }
                sb.Append("{");
                IRow row = sheet.GetRow(rowIndex);
                for (int columnIndex = 2; columnIndex < cellCount; columnIndex++)
                {
                    string description = cellInfos[columnIndex].Description.ToLower();
                    if (description.StartsWith("#"))
                    {
                        continue;
                    }
                    //s开头表示这个字段是服务端专用
                    if (description.StartsWith("s") && IsClient)
                    {
                        continue;
                    }
                    //c开头表示这个字段是客户端专用
                    if (description.StartsWith("c") && !IsClient)
                    {
                        continue;
                    }
                    string fieldValue = GetCellString(row, columnIndex, formulaEvaluator);
                    if (columnIndex > 2)
                    {
                        sb.Append(",");
                    }
                    string fieldName = cellInfos[columnIndex].Name;
                    if (fieldName == "Id" || fieldName == "_id")
                    {
                        if (IsClient)
                        {
                            fieldName = "Id";
                        }
                        else
                        {
                            fieldName = "_id";
                        }
                    }
                    string fieldType = cellInfos[columnIndex].Type;
                    sb.Append($"\"{fieldName}\":{Convert(fieldType, fieldValue)}");
                }
                sb.Append("}");
                if (rowIndex < sheet.LastRowNum)
                {
                    sb.Append("\n");
                }
            }
            return sb.ToString();
        }

        public static void Export(string fileName, string exportDirectory)
        {
            XSSFWorkbook xssfWorkbook;
            IFormulaEvaluator formulaEvaluator;
            using (FileStream file = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                xssfWorkbook = new XSSFWorkbook(file);
                formulaEvaluator = xssfWorkbook.GetCreationHelper().CreateFormulaEvaluator();
                Dictionary<string, IFormulaEvaluator> evaluators = new Dictionary<string, IFormulaEvaluator>
                {
                    {Path.GetFileName(fileName),formulaEvaluator }
                };
                foreach (string linkedFileName in xssfWorkbook.ExternalLinksTable.Select(p => $"{ExcelPath}/{p.LinkedFileName}"))
                {
                    XSSFWorkbook linkedWorkBook;
                    using (FileStream linkedFile = new FileStream(linkedFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                    {
                        linkedWorkBook = new XSSFWorkbook(linkedFile);
                    }
                    evaluators.Add(Path.GetFileName(linkedFileName), linkedWorkBook.GetCreationHelper().CreateFormulaEvaluator());
                }
                formulaEvaluator.SetupReferencedWorkbooks(evaluators);
            }
            string protoName = Path.GetFileNameWithoutExtension(fileName);
            string exportPath = Path.Combine(exportDirectory, $"{protoName}.txt");
            using (FileStream txt = new FileStream(exportPath, FileMode.Create))
            using (StreamWriter sw = new StreamWriter(txt))
            {
                for (int i = 0; i < xssfWorkbook.NumberOfSheets; ++i)
                {
                    ISheet sheet = xssfWorkbook.GetSheetAt(i);
                    if (sheet.SheetName.StartsWith("~"))
                    {
                        continue;
                    }
                    ExportSheet(sheet, sw, formulaEvaluator);
                }
            }
        }

        public static void ExportSheet(ISheet sheet, StreamWriter sw, IFormulaEvaluator formulaEvaluator)
        {
            int cellCount = sheet.GetRow(3).LastCellNum;
            ExcelCellInfo[] cellInfos = new ExcelCellInfo[cellCount];
            for (int i = 2; i < cellCount; ++i)
            {
                string fieldDesc = GetCellString(sheet, 2, i, formulaEvaluator);
                string fieldName = GetCellString(sheet, 3, i, formulaEvaluator);
                string fieldType = GetCellString(sheet, 4, i, formulaEvaluator);
                cellInfos[i] = new ExcelCellInfo() { Name = fieldName, Type = fieldType, Description = fieldDesc };
            }

            for (int rowIndex = 5; rowIndex <= sheet.LastRowNum; ++rowIndex)
            {
                string id = GetCellString(sheet, rowIndex, 2, formulaEvaluator);
                if (string.IsNullOrEmpty(id))
                {
                    continue;
                }
                StringBuilder sb = new StringBuilder();
                sb.Append("{");
                IRow row = sheet.GetRow(rowIndex);
                for (int columnIndex = 2; columnIndex < cellCount; columnIndex++)
                {
                    string description = cellInfos[columnIndex].Description.ToLower();
                    if (description.StartsWith("#"))
                    {
                        continue;
                    }
                    //s开头表示这个字段是服务端专用
                    if (description.StartsWith("s") && IsClient)
                    {
                        continue;
                    }
                    //c开头表示这个字段是客户端专用
                    if (description.StartsWith("c") && !IsClient)
                    {
                        continue;
                    }
                    string fieldValue = GetCellString(row, columnIndex, formulaEvaluator);
                    if (columnIndex > 2)
                    {
                        sb.Append(",");
                    }
                    string fieldName = cellInfos[columnIndex].Name;
                    if (fieldName == "Id" || fieldName == "_id")
                    {
                        if (IsClient)
                        {
                            fieldName = "Id";
                        }
                        else
                        {
                            fieldName = "_id";
                        }
                    }
                    string fieldType = cellInfos[columnIndex].Type;
                    sb.Append($"\"{fieldName}\":{Convert(fieldType, fieldValue)}");
                }
                sb.Append("}");
                sw.WriteLine(sb.ToString());
            }
        }

        private static string Convert(string type, string input)
        {
            string key = type.ToLower();
            return converters.ContainsKey(key) ? converters[key](input) : throw new Exception($"不支持此类型: {type}");
        }

        private static string GetCellString(ISheet sheet, int rowIndex, int columnIndex, IFormulaEvaluator formulaEvaluator)
        {
            ICell cell = sheet.GetRow(rowIndex)?.GetCell(columnIndex);
            if (cell?.CellType == CellType.Formula)
            {
                return $"{formulaEvaluator.EvaluateInCell(cell)}";
            }
            else
            {
                return cell?.ToString() ?? string.Empty;
            }
        }

        private static string GetCellString(IRow row, int columnIndex, IFormulaEvaluator formulaEvaluator)
        {
            ICell cell = row?.GetCell(columnIndex);
            if (cell?.CellType == CellType.Formula)
            {
                formulaEvaluator.ClearAllCachedResultValues();
                return $"{formulaEvaluator.EvaluateInCell(cell)}";
            }
            else
            {
                return cell?.ToString() ?? "";
            }
        }

        public override string ToString()
        {
            return $"ClientConfigPath:{new DirectoryInfo(ClientConfigPath).FullName}\n" +
                $"ServerConfigPath:{new DirectoryInfo(ServerConfigPath).FullName}\n";
        }
    }
}
using System;
using System.Threading.Tasks;
using Microsoft.Playwright;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Text.Json;

namespace AutoFinan
{
    public class AppConfig
    {
        public string ExcelFilePath { get; set; } = "报销信息.xlsx";
        public string MappingFilePath { get; set; } = "标题-ID.xlsx";
        public string SheetName { get; set; } = "ChaiLv_sheet";
        public string MappingSheetName { get; set; } = "Sheet1";
        public Dictionary<string, ScreenPosition> ScreenPositions { get; set; } = new Dictionary<string, ScreenPosition>();
    }

    public class ScreenPosition
    {
        public int X { get; set; }
        public int Y { get; set; }
        public string Description { get; set; }
    }

    class Program
    {
        static async Task Main(string[] args)
        {
            // 设置EPPlus许可证上下文
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            Console.WriteLine("=== 财务报销自动化系统 ===");

            try
            {
                var automation = new ReimbursementAutomation();
                await automation.RunAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"程序运行出错: {ex.Message}");
                Console.WriteLine($"详细错误: {ex}");
            }

            Console.WriteLine("程序结束，按任意键退出...");
            Console.ReadKey();
        }
    }

    public class ReimbursementAutomation
    {
        private AppConfig config;
        private const string SubsequenceStartColumn = "子序列开始";
        private const string SubsequenceEndColumn = "子序列结束";
        private const string SubsequenceMarker = "是";
        private const string SubsequenceMarker2 = "1"; // 第二种子序列的标记

        private Dictionary<string, string> titleIdMapping;
        private Dictionary<string, Dictionary<string, string>> dropdownMappings;
        private IPlaywright playwright;
        private IBrowser browser;
        private IPage page;
        private string currentSubjectId; // 存储当前科目ID，用于金额填写
        private bool isInSecondSubsequence = false; // 标记是否在第二种子序列中
        private int subsequenceRowIndex = 0; // 第二种子序列中的行序号（从0开始）
        private PythonScriptExecutor pythonExecutor; // Python脚本执行器
        private ExcelWorksheet currentWorksheet; // 当前工作表引用
        private List<string> currentHeaders; // 当前表头
        private string lastSavedPdfPath; // 最后保存的PDF路径
        private ExcelPackage currentPackage; // 当前Excel包引用

        public async Task RunAsync()
        {
            Console.WriteLine("开始读取配置文件...");

            // 加载配置文件
            await LoadConfiguration();

            // 初始化Python脚本执行器
            try
            {
                pythonExecutor = new PythonScriptExecutor();
                Console.WriteLine("Python脚本执行器初始化成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Python脚本执行器初始化失败: {ex.Message}");
                Console.WriteLine("将跳过Python脚本相关功能");
                pythonExecutor = null;
            }

            Console.WriteLine("开始读取Excel文件...");

            // 提醒用户关闭Excel文件
            Console.WriteLine("⚠️  重要提醒：请确保Excel文件没有被其他程序（如Excel、WPS）打开");
            Console.WriteLine("   如果文件被占用，程序将无法写入数据到Excel中");

            // 获取当前执行目录
            string currentDirectory = Directory.GetCurrentDirectory();
            Console.WriteLine($"当前工作目录: {currentDirectory}");

            // 尝试多个可能的文件路径
            string[] possiblePaths = {
                config.ExcelFilePath,
                Path.Combine(currentDirectory, config.ExcelFilePath),
                Path.Combine(currentDirectory, "..", config.ExcelFilePath),
                Path.Combine(currentDirectory, "..", "..", config.ExcelFilePath),
                Path.Combine(currentDirectory, "..", "..", "..", config.ExcelFilePath),
                Path.Combine(currentDirectory, "..", "..", "..", "..", config.ExcelFilePath)
            };

            string actualExcelPath = null;
            foreach (string path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    actualExcelPath = path;
                    Console.WriteLine($"找到Excel文件: {path}");
                    break;
                }
            }

            if (actualExcelPath == null)
            {
                Console.WriteLine($"错误：找不到文件 {config.ExcelFilePath}");
                Console.WriteLine("尝试过的路径:");
                foreach (string path in possiblePaths)
                {
                    Console.WriteLine($"  {path}");
                }
                return;
            }

            // 加载标题-ID映射表
            await LoadTitleIdMapping(actualExcelPath);

            // 初始化下拉框映射
            InitializeDropdownMappings();

            if (!File.Exists(actualExcelPath))
            {
                Console.WriteLine($"错误：找不到文件 {actualExcelPath}");
                return;
            }

            // 启动浏览器
            await InitializeBrowser();

            // 导航到目标网页
            await NavigateToTargetPage();

            using (var package = new ExcelPackage(new FileInfo(actualExcelPath)))
            {
                var worksheet = package.Workbook.Worksheets[config.SheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"错误：找不到工作表 {config.SheetName}");
                    return;
                }

                Console.WriteLine($"成功加载工作表: {config.SheetName}");

                // 设置当前工作表和表头引用
                currentWorksheet = worksheet;
                currentHeaders = GetHeaders(worksheet);
                currentPackage = package; // 保存Excel包引用

                // 获取数据范围
                int rowCount = worksheet.Dimension?.Rows ?? 0;
                int colCount = worksheet.Dimension?.Columns ?? 0;

                Console.WriteLine($"数据范围: {rowCount} 行 x {colCount} 列");

                if (rowCount == 0 || colCount == 0)
                {
                    Console.WriteLine("错误：Excel文件中没有数据");
                    return;
                }

                // 获取列标题
                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    var header = worksheet.Cells[1, col].Value?.ToString() ?? "";
                    headers.Add(header);
                }

                Console.WriteLine("列标题:");
                for (int i = 0; i < headers.Count; i++)
                {
                    Console.WriteLine($"  {GetColumnName(i + 1)}: {headers[i]}");
                }

                // 找到子序列相关列的索引
                var subsequenceStartColumns = new List<int>();
                var subsequenceEndColumns = new List<int>();

                Console.WriteLine("查找子序列相关列...");
                Console.WriteLine($"要查找的子序列开始列名: '{SubsequenceStartColumn}'");
                Console.WriteLine($"要查找的子序列结束列名: '{SubsequenceEndColumn}'");

                for (int i = 0; i < headers.Count; i++)
                {
                    Console.WriteLine($"  列 {GetColumnName(i + 1)}: '{headers[i]}'");
                    if (headers[i] == SubsequenceStartColumn)
                    {
                        subsequenceStartColumns.Add(i + 1);
                        Console.WriteLine($"    找到子序列开始列: {GetColumnName(i + 1)} (索引: {i + 1})");
                    }
                    else if (headers[i] == SubsequenceEndColumn)
                    {
                        subsequenceEndColumns.Add(i + 1);
                        Console.WriteLine($"    找到子序列结束列: {GetColumnName(i + 1)} (索引: {i + 1})");
                    }
                }

                Console.WriteLine($"找到 {subsequenceStartColumns.Count} 个子序列开始列: {string.Join(", ", subsequenceStartColumns.Select(i => GetColumnName(i)))}");
                Console.WriteLine($"找到 {subsequenceEndColumns.Count} 个子序列结束列: {string.Join(", ", subsequenceEndColumns.Select(i => GetColumnName(i)))}");

                // 第一层循环：从上至下，读取每一行数据（广义上的行）
                Console.WriteLine("\n=== 开始第一层循环：处理每一行数据 ===");

                for (int row = 2; row <= rowCount; row++) // 从第2行开始（跳过标题行）
                {
                    Console.WriteLine($"\n--- 处理第 {row} 行数据 ---");

                    // 第二层循环：从左至右，处理当前行的每个单元格
                    Console.WriteLine($"开始第二层循环：处理第 {row} 行的单元格");

                    for (int col = 1; col <= colCount; col++)
                    {
                        var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                        var columnName = GetColumnName(col);
                        var headerName = headers[col - 1];

                        Console.WriteLine($"  读取单元格 {columnName}{row}，列标题: {headerName}，值: '{cellValue}'");

                        // 检查是否遇到子序列开始标记
                        Console.WriteLine($"      当前列 {col}，子序列开始列列表: [{string.Join(", ", subsequenceStartColumns)}]，是否包含当前列: {subsequenceStartColumns.Contains(col)}");
                        if (subsequenceStartColumns.Contains(col))
                        {
                            Console.WriteLine($"      检查子序列开始标记：当前列 {col}，单元格值 '{cellValue}'，第一种子序列标记 '{SubsequenceMarker}'，第二种子序列标记 '{SubsequenceMarker2}'");

                            if (cellValue == SubsequenceMarker)
                            {
                                Console.WriteLine($"检测到第一种子序列开始标记，进入子序列处理逻辑");
                                // 找到对应的子序列结束列
                                int subsequenceEndColIndex = FindCorrespondingEndColumn(col, subsequenceStartColumns, subsequenceEndColumns);
                                int nextRow = await ProcessSubsequence(worksheet, headers, row, rowCount, col, subsequenceEndColIndex);

                                // 如果子序列处理返回了下一行的行号，继续处理那一行
                                if (nextRow > 0)
                                {
                                    Console.WriteLine($"子序列处理完成，继续处理第 {nextRow} 行");
                                    await ProcessRowFromColumn(worksheet, headers, nextRow, subsequenceEndColIndex + 1, colCount);

                                    // 更新外层循环的行号，跳过已经处理过的行
                                    row = nextRow;
                                }
                                break; // 跳出当前行的列循环，继续处理下一行
                            }
                            else if (cellValue == SubsequenceMarker2)
                            {
                                Console.WriteLine($"检测到第二种子序列开始标记，进入第二种子序列处理逻辑");
                                // 找到对应的子序列结束列
                                int subsequenceEndColIndex = FindCorrespondingEndColumn(col, subsequenceStartColumns, subsequenceEndColumns);
                                int nextRow = await ProcessSecondSubsequence(worksheet, headers, row, rowCount, col, subsequenceEndColIndex);

                                // 如果子序列处理返回了下一行的行号，继续处理那一行
                                if (nextRow > 0)
                                {
                                    Console.WriteLine($"第二种子序列处理完成，继续处理第 {nextRow} 行");
                                    await ProcessRowFromColumn(worksheet, headers, nextRow, subsequenceEndColIndex + 1, colCount);

                                    // 更新外层循环的行号，跳过已经处理过的行
                                    row = nextRow;
                                }
                                break; // 跳出当前行的列循环，继续处理下一行
                            }
                            else
                            {
                                Console.WriteLine($"      子序列开始列的值 '{cellValue}' 不匹配任何子序列标记");
                            }
                        }

                        // 如果不是子序列开始列，则正常处理单元格
                        if (!subsequenceStartColumns.Contains(col) && (!string.IsNullOrEmpty(cellValue) || headers[col - 1].StartsWith("?")))
                        {
                            await ProcessCell(columnName, row, headerName, cellValue);
                        }
                    }

                    Console.WriteLine($"第 {row} 行数据处理完成");
                }

                Console.WriteLine("\n=== 所有数据处理完成 ===");

                // 保存Excel文件
                SaveExcelFile();
            }

            // 等待用户手动关闭浏览器
            Console.WriteLine(new string('=', 50));
            Console.WriteLine("所有操作已完成！");
            Console.WriteLine("浏览器将保持打开状态，您可以手动关闭。");
            Console.WriteLine(new string('=', 50));

            try
            {
                Console.WriteLine("按回车键关闭浏览器...");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"等待用户输入时出错: {ex.Message}");
            }
            finally
            {
                // 清理资源
                await Cleanup();
            }
        }

        private async Task LoadTitleIdMapping(string excelFilePath)
        {
            Console.WriteLine("开始加载标题-ID映射表...");

            // 获取Excel文件所在目录
            string excelDirectory = Path.GetDirectoryName(excelFilePath);
            string mappingFilePath = Path.Combine(excelDirectory, config.MappingFilePath);

            if (!File.Exists(mappingFilePath))
            {
                Console.WriteLine($"错误：找不到标题-ID映射文件 {mappingFilePath}");
                return;
            }

            titleIdMapping = new Dictionary<string, string>();

            using (var package = new ExcelPackage(new FileInfo(mappingFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[config.MappingSheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"错误：找不到工作表 {config.MappingSheetName}");
                    return;
                }

                int rowCount = worksheet.Dimension?.Rows ?? 0;
                Console.WriteLine($"标题-ID映射表行数: {rowCount}");

                for (int row = 2; row <= rowCount; row++) // 从第2行开始（跳过标题行）
                {
                    var title = worksheet.Cells[row, 1].Value?.ToString() ?? "";
                    var id = worksheet.Cells[row, 2].Value?.ToString() ?? "";

                    if (!string.IsNullOrEmpty(title) && !string.IsNullOrEmpty(id))
                    {
                        titleIdMapping[title] = id;
                        Console.WriteLine($"  映射: {title} -> {id}");
                    }
                }
            }

            Console.WriteLine($"成功加载 {titleIdMapping.Count} 个标题-ID映射");
        }

        private async Task<int> ProcessSubsequence(ExcelWorksheet worksheet, List<string> headers, int startRow, int totalRows, int subsequenceStartColIndex, int subsequenceEndColIndex)
        {
            Console.WriteLine($"\n=== 进入第三层循环：第一种子序列处理逻辑 ===");
            Console.WriteLine($"子序列处理从第 {startRow} 行开始");

            // 第三层循环：从上至下，处理子序列中的每一行
            for (int row = startRow; row <= totalRows; row++)
            {
                Console.WriteLine($"\n--- 处理子序列第 {row} 行 ---");

                // 只处理子序列范围内的列（从子序列开始列的下一列到子序列结束列的前一列）
                int startCol = subsequenceStartColIndex + 1; // 从子序列开始列的下一列开始
                int endCol = subsequenceEndColIndex > 0 ? subsequenceEndColIndex - 1 : headers.Count; // 到子序列结束列的前一列结束

                Console.WriteLine($"    子序列处理范围：从列 {GetColumnName(startCol)} 到列 {GetColumnName(endCol)}");

                // 处理子序列范围内的列（从左至右）
                for (int col = startCol; col <= endCol; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    var columnName = GetColumnName(col);
                    var headerName = headers[col - 1];

                    Console.WriteLine($"    子序列处理：读取单元格 {columnName}{row}，列标题: {headerName}，值: '{cellValue}'");

                    if (!string.IsNullOrEmpty(cellValue) || headerName.StartsWith("?"))
                    {
                        await ProcessCell(columnName, row, headerName, cellValue);
                    }
                }

                Console.WriteLine($"子序列第 {row} 行处理完成");

                // 检查当前行的子序列结束列是否标记为"是"
                if (subsequenceEndColIndex > 0)
                {
                    var currentRowSubsequenceEndValue = worksheet.Cells[row, subsequenceEndColIndex].Value?.ToString() ?? "";
                    if (currentRowSubsequenceEndValue == SubsequenceMarker)
                    {
                        Console.WriteLine($"检测到当前行({row})的子序列结束标记，结束子序列处理");
                        Console.WriteLine($"程序将从第 {row} 行继续正常处理逻辑");
                        return row; // 返回当前行的行号
                    }
                    else
                    {
                        Console.WriteLine($"当前行({row})的子序列结束列未标记为'是'，继续处理下一行");
                    }
                }
            }

            Console.WriteLine("=== 子序列处理逻辑结束 ===");
            return totalRows + 1; // 如果没有找到子序列结束标记，返回下一行
        }

        private async Task<int> ProcessSecondSubsequence(ExcelWorksheet worksheet, List<string> headers, int startRow, int totalRows, int subsequenceStartColIndex, int subsequenceEndColIndex)
        {
            Console.WriteLine($"\n=== 进入第三层循环：第二种子序列处理逻辑 ===");
            Console.WriteLine($"第二种子序列处理从第 {startRow} 行开始");

            // 设置第二种子序列标记
            isInSecondSubsequence = true;
            subsequenceRowIndex = 0; // 从0开始计数

            // 第三层循环：从上至下，处理子序列中的每一行
            for (int row = startRow; row <= totalRows; row++)
            {
                Console.WriteLine($"\n--- 处理第二种子序列第 {row} 行（子序列内序号：{subsequenceRowIndex}） ---");

                // 只处理子序列范围内的列（从子序列开始列的下一列到子序列结束列的前一列）
                int startCol = subsequenceStartColIndex + 1; // 从子序列开始列的下一列开始
                int endCol = subsequenceEndColIndex > 0 ? subsequenceEndColIndex - 1 : headers.Count; // 到子序列结束列的前一列结束

                Console.WriteLine($"    第二种子序列处理范围：从列 {GetColumnName(startCol)} 到列 {GetColumnName(endCol)}");

                // 处理子序列范围内的列（从左至右）
                for (int col = startCol; col <= endCol; col++)
                {
                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    var columnName = GetColumnName(col);
                    var headerName = headers[col - 1];

                    Console.WriteLine($"    第二种子序列处理：读取单元格 {columnName}{row}，列标题: {headerName}，值: '{cellValue}'");

                    if (!string.IsNullOrEmpty(cellValue) || headerName.StartsWith("?"))
                    {
                        await ProcessCell(columnName, row, headerName, cellValue);
                    }
                }

                Console.WriteLine($"第二种子序列第 {row} 行处理完成");

                // 检查当前行的子序列结束列是否标记为"1"
                if (subsequenceEndColIndex > 0)
                {
                    var currentRowSubsequenceEndValue = worksheet.Cells[row, subsequenceEndColIndex].Value?.ToString() ?? "";
                    if (currentRowSubsequenceEndValue == SubsequenceMarker2)
                    {
                        Console.WriteLine($"检测到当前行({row})的第二种子序列结束标记，结束子序列处理");
                        Console.WriteLine($"程序将从第 {row} 行继续正常处理逻辑");

                        // 重置第二种子序列标记
                        isInSecondSubsequence = false;
                        subsequenceRowIndex = 0;

                        return row; // 返回当前行的行号
                    }
                    else
                    {
                        Console.WriteLine($"当前行({row})的子序列结束列未标记为'1'，继续处理下一行");
                    }
                }

                // 增加子序列内行序号
                subsequenceRowIndex++;
            }

            Console.WriteLine("=== 第二种子序列处理逻辑结束 ===");

            // 重置第二种子序列标记
            isInSecondSubsequence = false;
            subsequenceRowIndex = 0;

            return totalRows + 1; // 如果没有找到子序列结束标记，返回下一行
        }

        /// <summary>
        /// 获取表头
        /// </summary>
        private List<string> GetHeaders(ExcelWorksheet worksheet)
        {
            var headers = new List<string>();
            int colCount = worksheet.Dimension?.Columns ?? 0;

            for (int col = 1; col <= colCount; col++)
            {
                var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? "";
                headers.Add(headerValue);
            }

            return headers;
        }

        /// <summary>
        /// 处理以?开头的列标题（程序自动填写）
        /// </summary>
        private async Task HandleQuestionMarkColumn(string columnName, int row, string headerName, string cellValue)
        {
            try
            {
                Console.WriteLine($"      检测到?列标题: {headerName}");
                Console.WriteLine($"      参数: columnName={columnName}, row={row}, headerName={headerName}, cellValue={cellValue}");

                // 移除?前缀，获取实际的列名
                string actualColumnName = headerName.Substring(1);
                Console.WriteLine($"      实际列名: {actualColumnName}");

                string valueToWrite = "";

                // 根据不同的列名处理不同的逻辑
                switch (actualColumnName.ToLower())
                {
                    case "pdf路径":
                        valueToWrite = lastSavedPdfPath ?? "未生成PDF文件";
                        Console.WriteLine($"      填写PDF路径: {valueToWrite}");
                        break;

                    case "预约号":
                        // 从lastSavedPdfPath中提取预约号，避免页面跳转后无法获取
                        if (!string.IsNullOrEmpty(lastSavedPdfPath))
                        {
                            var fileName = Path.GetFileNameWithoutExtension(lastSavedPdfPath);
                            var parts = fileName.Split('-');
                            if (parts.Length >= 1)
                            {
                                valueToWrite = parts[0];
                            }
                            else
                            {
                                valueToWrite = "未获取到预约号";
                            }
                        }
                        else
                        {
                            valueToWrite = "未获取到预约号";
                        }
                        Console.WriteLine($"      填写预约号: {valueToWrite}");
                        break;

                    case "涉及总金额":
                    case "涉及金额":
                        // 从lastSavedPdfPath中提取金额，避免页面跳转后无法获取
                        if (!string.IsNullOrEmpty(lastSavedPdfPath))
                        {
                            var fileName = Path.GetFileNameWithoutExtension(lastSavedPdfPath);
                            var parts = fileName.Split('-');
                            if (parts.Length >= 2)
                            {
                                valueToWrite = parts[1];
                            }
                            else
                            {
                                valueToWrite = "未获取到金额";
                            }
                        }
                        else
                        {
                            valueToWrite = "未获取到金额";
                        }
                        Console.WriteLine($"      填写涉及金额: {valueToWrite}");
                        break;

                    case "当前时间":
                        valueToWrite = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
                        Console.WriteLine($"      填写当前时间: {valueToWrite}");
                        break;

                    case "文件名":
                        if (!string.IsNullOrEmpty(lastSavedPdfPath))
                        {
                            valueToWrite = Path.GetFileName(lastSavedPdfPath);
                        }
                        else
                        {
                            valueToWrite = "未生成文件";
                        }
                        Console.WriteLine($"      填写文件名: {valueToWrite}");
                        break;

                    default:
                        Console.WriteLine($"      未知的?列名: {actualColumnName}，跳过处理");
                        return;
                }

                // 写入Excel单元格
                Console.WriteLine($"      准备写入Excel: 行={row}, 标题={headerName}, 值={valueToWrite}");
                WriteToExcelCell(row, headerName, valueToWrite);
                Console.WriteLine($"      ✓ 已填写{actualColumnName}: {valueToWrite}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      处理?列时出错: {ex.Message}");
                Console.WriteLine($"      详细错误: {ex}");
            }
        }

        /// <summary>
        /// 保存Excel文件
        /// </summary>
        private void SaveExcelFile()
        {
            try
            {
                if (currentPackage != null)
                {
                    Console.WriteLine("      正在保存Excel文件...");

                    // 获取文件路径信息
                    var fileInfo = currentPackage.File;
                    if (fileInfo != null)
                    {
                        Console.WriteLine($"      文件路径: {fileInfo.FullName}");
                        Console.WriteLine($"      文件是否存在: {fileInfo.Exists}");
                        Console.WriteLine($"      文件大小: {fileInfo.Length} 字节");

                        // 检查文件权限
                        try
                        {
                            using (var stream = fileInfo.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                            {
                                // 如果能打开文件，说明有读写权限
                                Console.WriteLine("      ✓ 文件权限检查通过");
                            }
                        }
                        catch (UnauthorizedAccessException)
                        {
                            Console.WriteLine("      ✗ 文件权限不足，无法写入");
                            return;
                        }
                        catch (IOException)
                        {
                            Console.WriteLine("      ✗ 文件被其他程序占用");
                            return;
                        }
                    }

                    // 保存文件
                    try
                    {
                        currentPackage.Save();
                        Console.WriteLine("      ✓ Excel文件保存成功");
                    }
                    catch (Exception saveEx)
                    {
                        Console.WriteLine($"      直接保存失败，尝试备用保存方法: {saveEx.Message}");

                        // 备用保存方法：尝试保存到临时文件
                        try
                        {
                            string tempPath = Path.Combine(Path.GetTempPath(), $"temp_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
                            currentPackage.SaveAs(new FileInfo(tempPath));
                            Console.WriteLine($"      ✓ 备用保存成功，文件保存到: {tempPath}");
                            Console.WriteLine($"      请手动将文件从 {tempPath} 复制到原位置");
                        }
                        catch (Exception backupEx)
                        {
                            Console.WriteLine($"      ✗ 备用保存也失败: {backupEx.Message}");
                            throw; // 重新抛出原始异常
                        }
                    }
                }
                else
                {
                    Console.WriteLine("      ✗ currentPackage为null，无法保存文件");
                }
            }
            catch (System.IO.IOException ex)
            {
                Console.WriteLine($"      ✗ 保存Excel文件失败: 文件可能被其他程序占用");
                Console.WriteLine($"      请关闭Excel、WPS等程序，然后重新运行");
                Console.WriteLine($"      详细错误: {ex.Message}");
                Console.WriteLine($"      错误类型: {ex.GetType().Name}");
            }
            catch (UnauthorizedAccessException ex)
            {
                Console.WriteLine($"      ✗ 保存Excel文件失败: 权限不足");
                Console.WriteLine($"      请以管理员身份运行程序，或检查文件权限");
                Console.WriteLine($"      详细错误: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      ✗ 保存Excel文件时发生未知错误: {ex.Message}");
                Console.WriteLine($"      错误类型: {ex.GetType().Name}");
                Console.WriteLine($"      详细错误: {ex}");
            }
        }

        /// <summary>
        /// 写入Excel单元格
        /// </summary>
        private void WriteToExcelCell(int row, string columnName, string value)
        {
            try
            {
                Console.WriteLine($"      开始写入Excel: 行={row}, 列名={columnName}, 值={value}");

                if (currentWorksheet == null)
                {
                    Console.WriteLine($"      警告: currentWorksheet为null，无法写入Excel");
                    return;
                }

                // 将列名转换为列索引
                int columnIndex = GetColumnIndex(columnName);
                Console.WriteLine($"      列名'{columnName}'对应的索引: {columnIndex}");

                if (columnIndex <= 0)
                {
                    Console.WriteLine($"      警告: 无法找到列 {columnName}");
                    Console.WriteLine($"      当前表头: {string.Join(", ", currentHeaders ?? new List<string>())}");
                    return;
                }

                // 尝试写入Excel单元格
                try
                {
                    currentWorksheet.Cells[row, columnIndex].Value = value;
                    Console.WriteLine($"      成功写入Excel: {columnName}{row} = {value}");

                    // 不立即保存，在主循环结束时统一保存
                }
                catch (System.IO.IOException ex)
                {
                    Console.WriteLine($"      ✗ 写入Excel失败: 文件可能被其他程序占用");
                    Console.WriteLine($"      请关闭Excel、WPS等程序，然后重新运行");
                    Console.WriteLine($"      详细错误: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      ✗ 写入Excel时发生未知错误: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      写入Excel单元格时出错: {ex.Message}");
                Console.WriteLine($"      详细错误: {ex}");
            }
        }

        /// <summary>
        /// 获取列索引
        /// </summary>
        private int GetColumnIndex(string columnName)
        {
            if (currentHeaders == null) return -1;

            for (int i = 0; i < currentHeaders.Count; i++)
            {
                if (currentHeaders[i] == columnName)
                {
                    return i + 1; // 返回1-based索引
                }
            }
            return -1;
        }

        private async Task ProcessCell(string columnName, int row, string headerName, string cellValue)
        {
            // 这里实现具体的单元格处理逻辑
            // 根据不同的列标题和值，执行不同的操作

            Console.WriteLine($"    执行操作：{columnName}{row} - {headerName} = '{cellValue}'");

            // 1. 处理以?开头的列标题（程序自动填写）
            if (headerName.StartsWith("?"))
            {
                Console.WriteLine($"      检测到?列标题，开始自动填写");
                Console.WriteLine($"      列名: {columnName}, 标题: {headerName}, 值: {cellValue}");
                await HandleQuestionMarkColumn(columnName, row, headerName, cellValue);
                return; // ?列不需要对网页进行操作
            }
            // 2. 等待操作（列标题为"等待"）
            else if (headerName == "等待")
            {
                Console.WriteLine($"      检测到等待操作: {cellValue}");
                await WaitOperation(cellValue);
            }
            // 3. 回车键操作（列标题为"回车"，值为"$点击"）
            else if (headerName == "回车" && cellValue == "$点击")
            {
                Console.WriteLine($"      检测到回车键操作");
                await PressEnterKey();
            }
            // 4. 按钮点击操作（以$开头）
            else if (cellValue == "$点击" || cellValue == "$预约")
            {
                Console.WriteLine($"      检测到按钮点击操作: {headerName}");
                await ClickButton(headerName);
            }
            // 5. Radio按钮点击操作（以$$开头）
            else if (cellValue.StartsWith("$$"))
            {
                string radioValue = cellValue.Substring(2); // 去掉$$前缀
                Console.WriteLine($"      检测到Radio按钮操作: {radioValue}");
                await ClickRadioButton(radioValue);
            }
            // 6. 银行卡选择操作（以*开头）
            else if (cellValue.StartsWith("*"))
            {
                Console.WriteLine($"      检测到银行卡选择操作: {cellValue}");
                await SelectCardByNumber(cellValue);
            }
            // 7. 科目输入框操作（以#开头）
            else if (cellValue.StartsWith("#"))
            {
                Console.WriteLine($"      检测到科目输入框操作: {cellValue}");
                await FillSubjectInput(headerName, cellValue);
            }
            // 8. 下拉框选择操作
            else if (IsDropdownField(headerName))
            {
                Console.WriteLine($"      检测到下拉框选择操作: {headerName} = {cellValue}");
                await SelectDropdown(headerName, cellValue);
            }
            // 9. 日期选择操作（日期字段或格式：yyyy-mm-dd）
            else if (IsDateField(headerName) || IsDate(cellValue))
            {
                Console.WriteLine($"      检测到日期选择操作: {cellValue}");
                await SelectDate(headerName, cellValue);
            }
            // 10. 金额输入框操作（需要与科目配对）
            else if (headerName == "金额" && !string.IsNullOrEmpty(currentSubjectId))
            {
                Console.WriteLine($"      检测到金额输入框操作: {cellValue}");
                await FillAmountInput(currentSubjectId, cellValue);
                currentSubjectId = null; // 清空当前科目ID
            }
            // 11. 一般输入框操作
            else
            {
                Console.WriteLine($"      检测到输入框操作: {cellValue}");
                await FillInput(headerName, cellValue);
            }

            // 模拟异步操作
            await Task.Delay(100);
        }

        private async Task FillInput(string headerName, string value)
        {
            try
            {
                string elementId = GetElementId(headerName);
                if (string.IsNullOrEmpty(elementId))
                {
                    Console.WriteLine($"      警告：未找到标题 '{headerName}' 对应的元素ID");
                    return;
                }

                Console.WriteLine($"      填写输入框: {headerName} -> {elementId} = {value}");

                // 实现实际的输入框填写逻辑
                bool filled = false;

                // 方法1: 优先在iframe中查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var inputElement = frame.Locator($"#{elementId}").First;
                        if (await inputElement.CountAsync() > 0)
                        {
                            await inputElement.FillAsync(value);
                            Console.WriteLine($"      在iframe中成功填写输入框 {elementId}: {value}");

                            // 如果是银行卡相关字段，触发事件来弹出选择窗口
                            if (IsBankCardField(headerName))
                            {
                                await TriggerBankCardSelection(inputElement);
                            }

                            // 如果是奖助学金项目号，自动按回车键并选择表格第一行
                            if (headerName == "奖助学金项目号")
                            {
                                Console.WriteLine("      检测到奖助学金项目号，自动按回车键");
                                await inputElement.PressAsync("Enter");
                                Console.WriteLine("      成功在奖助学金项目号输入框中按回车键");
                                await Task.Delay(1000); // 等待页面响应

                                // 自动选择表格第一行
                                await SelectFirstTableRow();
                            }
                            // 如果是奖助学金工号，自动按回车键
                            else if (headerName == "奖助学金工号")
                            {
                                Console.WriteLine("      检测到奖助学金工号，自动按回车键");
                                await inputElement.PressAsync("Enter");
                                Console.WriteLine("      成功在奖助学金工号输入框中按回车键");
                                await Task.Delay(1000); // 等待页面响应
                            }

                            filled = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找输入框失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 如果iframe中找不到，尝试在主页面查找
                if (!filled)
                {
                    try
                    {
                        await page.WaitForSelectorAsync($"#{elementId}", new PageWaitForSelectorOptions { Timeout = 3000 });
                        await page.FillAsync($"#{elementId}", value);
                        Console.WriteLine($"      在主页面成功填写输入框 {elementId}: {value}");

                        // 如果是银行卡相关字段，触发事件来弹出选择窗口
                        if (IsBankCardField(headerName))
                        {
                            var inputElement = page.Locator($"#{elementId}").First;
                            await TriggerBankCardSelection(inputElement);
                        }

                        // 如果是奖助学金项目号，自动按回车键并选择表格第一行
                        if (headerName == "奖助学金项目号")
                        {
                            Console.WriteLine("      检测到奖助学金项目号，自动按回车键");
                            var inputElement = page.Locator($"#{elementId}").First;
                            await inputElement.PressAsync("Enter");
                            Console.WriteLine("      成功在奖助学金项目号输入框中按回车键");
                            await Task.Delay(1000); // 等待页面响应

                            // 自动选择表格第一行
                            await SelectFirstTableRow();
                        }
                        // 如果是奖助学金工号，自动按回车键
                        else if (headerName == "奖助学金工号")
                        {
                            Console.WriteLine("      检测到奖助学金工号，自动按回车键");
                            var inputElement = page.Locator($"#{elementId}").First;
                            await inputElement.PressAsync("Enter");
                            Console.WriteLine("      成功在奖助学金工号输入框中按回车键");
                            await Task.Delay(1000); // 等待页面响应
                        }

                        filled = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面查找输入框失败: {ex.Message}");
                    }
                }

                // 方法3: 如果还是找不到，尝试通过name属性查找
                if (!filled)
                {
                    foreach (var frame in frames)
                    {
                        try
                        {
                            var inputElement = frame.Locator($"input[name='{elementId}']").First;
                            if (await inputElement.CountAsync() > 0)
                            {
                                await inputElement.FillAsync(value);
                                Console.WriteLine($"      在iframe中通过name属性成功填写输入框 {elementId}: {value}");

                                // 如果是银行卡相关字段，触发事件来弹出选择窗口
                                if (IsBankCardField(headerName))
                                {
                                    await TriggerBankCardSelection(inputElement);
                                }

                                // 如果是奖助学金项目号，自动按回车键并选择表格第一行
                                if (headerName == "奖助学金项目号")
                                {
                                    Console.WriteLine("      检测到奖助学金项目号，自动按回车键");
                                    await inputElement.PressAsync("Enter");
                                    Console.WriteLine("      成功在奖助学金项目号输入框中按回车键");
                                    await Task.Delay(1000); // 等待页面响应

                                    // 自动选择表格第一行
                                    await SelectFirstTableRow();
                                }
                                // 如果是奖助学金工号，自动按回车键
                                else if (headerName == "奖助学金工号")
                                {
                                    Console.WriteLine("      检测到奖助学金工号，自动按回车键");
                                    await inputElement.PressAsync("Enter");
                                    Console.WriteLine("      成功在奖助学金工号输入框中按回车键");
                                    await Task.Delay(1000); // 等待页面响应
                                }

                                filled = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      在iframe中通过name属性查找失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                if (!filled)
                {
                    Console.WriteLine($"      最终失败：无法找到输入框 {elementId}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      填写输入框失败: {ex.Message}");
            }
        }

        private async Task ClickButton(string headerName)
        {
            try
            {
                // 特殊处理：打印确认单按钮 - 使用固定选择器直接点击
                if (headerName == "打印确认单" || headerName == "打印确认单按钮")
                {
                    Console.WriteLine("      检测到打印确认单按钮，使用特殊处理方式");
                    await ClickPrintConfirmButton();
                    return;
                }

                // 特殊处理：预约按钮（表格中的按钮）- 需要优先处理，不需要查找ID
                if (headerName == "预约按钮")
                {
                    Console.WriteLine("      检测到预约按钮，在表格中查找第一行的预约按钮");
                    await ClickAppointmentButton();
                    return;
                }

                // 特殊处理：确定按钮（layui弹窗）- 需要优先处理
                if (headerName == "确定按钮")
                {
                    Console.WriteLine("      检测到确定按钮，在layui弹窗中查找确定按钮");
                    await ClickLayuiConfirmButton();
                    return;
                }

                string elementId = GetElementId(headerName);
                if (string.IsNullOrEmpty(elementId))
                {
                    Console.WriteLine($"      警告：未找到标题 '{headerName}' 对应的按钮ID");
                    return;
                }

                Console.WriteLine($"      点击按钮: {headerName} -> {elementId}");

                // 特殊处理：JavaScript函数调用
                if (elementId.StartsWith("navToPrj(") && elementId.EndsWith(")"))
                {
                    Console.WriteLine($"      检测到JavaScript函数调用: {elementId}");
                    try
                    {
                        await page.EvaluateAsync(elementId);
                        Console.WriteLine($"      成功执行JavaScript函数: {elementId}");
                        await Task.Delay(2000); // 等待页面跳转

                        // 打印页面信息以确认是否跳转到新页面
                        try
                        {
                            var currentUrl = page.Url;
                            var currentTitle = await page.TitleAsync();
                            Console.WriteLine($"      当前页面URL: {currentUrl}");
                            Console.WriteLine($"      当前页面标题: {currentTitle}");

                            // 检查是否有新标签页打开
                            var contexts = browser.Contexts;
                            Console.WriteLine($"      当前浏览器共有 {contexts.Count} 个上下文");

                            foreach (var context in contexts)
                            {
                                var pages = context.Pages;
                                Console.WriteLine($"      上下文中有 {pages.Count} 个页面");

                                for (int i = 0; i < pages.Count; i++)
                                {
                                    var pageUrl = pages[i].Url;
                                    Console.WriteLine($"        页面 {i + 1}: {pageUrl}");
                                }
                            }

                            // 检查当前页面URL是否变化
                            if (currentUrl.Contains("WF_GF6_NEW") || currentUrl.Contains("WF_YB6"))
                            {
                                Console.WriteLine("      确认：当前页面已跳转到新页面");
                            }
                            else
                            {
                                Console.WriteLine("      注意：当前页面URL未变化，尝试切换到新标签页");

                                // 尝试切换到新标签页
                                try
                                {
                                    foreach (var context in contexts)
                                    {
                                        var pages = context.Pages;
                                        for (int i = 0; i < pages.Count; i++)
                                        {
                                            var pageUrl = pages[i].Url;
                                            if (pageUrl.Contains("WF_GF6_NEW") || pageUrl.Contains("WF_YB6"))
                                            {
                                                // 切换到新标签页
                                                page = pages[i];
                                                Console.WriteLine($"      成功切换到新标签页: {pageUrl}");
                                                return;
                                            }
                                        }
                                    }
                                    Console.WriteLine("      未找到目标新标签页");
                                }
                                catch (Exception switchEx)
                                {
                                    Console.WriteLine($"      切换标签页失败: {switchEx.Message}");
                                }
                            }
                        }
                        catch (Exception infoEx)
                        {
                            Console.WriteLine($"      获取页面信息失败: {infoEx.Message}");
                        }

                        return;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      JavaScript函数执行失败: {ex.Message}");
                        return;
                    }
                }

                // 特殊处理：登录按钮需要等待验证码输入
                if (headerName == "登录按钮")
                {
                    Console.WriteLine("      检测到登录按钮，等待用户输入验证码...");
                    Console.WriteLine(new string('=', 50));
                    Console.WriteLine("请在下方输入验证码:");
                    Console.WriteLine(new string('=', 50));

                    string captcha = Console.ReadLine();
                    Console.WriteLine($"      用户输入验证码: {captcha}");

                    // 填写验证码
                    await FillCaptcha(captcha);
                }

                // 特殊处理：网上预约报账按钮（导航按钮）
                if (headerName == "网上预约报账按钮")
                {
                    Console.WriteLine("      检测到网上预约报账按钮，使用导航功能");
                    await ClickNavigationButton();
                    return;
                }

                // 特殊处理：打印确认单按钮
                if (headerName == "打印确认单")
                {
                    Console.WriteLine("      检测到打印确认单按钮，先执行正常点击，然后调用Python脚本处理后续操作");
                    // 不直接返回，继续执行下面的按钮点击逻辑
                }



                // 实现实际的按钮点击逻辑
                bool clicked = false;

                // 等待页面完全加载
                await Task.Delay(500);

                // 方法1: 优先在iframe中通过btnname属性查找
                var frames = page.Frames;
                Console.WriteLine($"      开始查找按钮，共有 {frames.Count} 个iframe");
                foreach (var frame in frames)
                {
                    try
                    {
                        Console.WriteLine($"      在iframe中通过btnname查找按钮: button[btnname='{elementId}']");
                        var buttonElement = frame.Locator($"button[btnname='{elementId}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      在iframe中通过btnname成功点击按钮: {elementId}");
                            clicked = true;
                            break;
                        }
                        else
                        {
                            Console.WriteLine($"      在iframe中未找到btnname为 '{elementId}' 的按钮");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中通过btnname查找按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 在iframe中通过ID查找
                if (!clicked)
                {
                    foreach (var frame in frames)
                    {
                        try
                        {
                            Console.WriteLine($"      在iframe中通过ID查找按钮: #{elementId}");
                            var buttonElement = frame.Locator($"#{elementId}").First;
                            if (await buttonElement.CountAsync() > 0)
                            {
                                await buttonElement.ClickAsync();
                                Console.WriteLine($"      在iframe中通过ID成功点击按钮: {elementId}");
                                clicked = true;
                                break;
                            }
                            else
                            {
                                Console.WriteLine($"      在iframe中未找到ID为 '{elementId}' 的按钮");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      在iframe中通过ID查找按钮失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                // 方法3: 在主页面通过btnname属性查找
                if (!clicked)
                {
                    try
                    {
                        var buttonElement = page.Locator($"button[btnname='{elementId}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      在主页面通过btnname成功点击按钮: {elementId}");
                            clicked = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面通过btnname查找按钮失败: {ex.Message}");
                    }
                }

                // 方法4: 在主页面通过ID查找
                if (!clicked)
                {
                    try
                    {
                        await page.WaitForSelectorAsync($"#{elementId}", new PageWaitForSelectorOptions { Timeout = 3000 });
                        await page.ClickAsync($"#{elementId}");
                        Console.WriteLine($"      在主页面通过ID成功点击按钮: {elementId}");
                        clicked = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面通过ID查找按钮失败: {ex.Message}");
                    }
                }

                // 方法5: 尝试其他选择器
                if (!clicked)
                {
                    string[] alternativeSelectors = {
                        $"button[guid*='{elementId}']",
                        $"button:has-text('{elementId}')",
                        $"input[btnname='{elementId}']",
                        $"[btnname='{elementId}']"
                    };

                    foreach (string selector in alternativeSelectors)
                    {
                        try
                        {
                            // 在iframe中查找
                            foreach (var frame in frames)
                            {
                                try
                                {
                                    var buttonElement = frame.Locator(selector).First;
                                    if (await buttonElement.CountAsync() > 0)
                                    {
                                        await buttonElement.ClickAsync();
                                        Console.WriteLine($"      在iframe中使用备用选择器成功点击按钮: {selector}");
                                        clicked = true;
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }
                            }

                            if (clicked) break;

                            // 在主页面查找
                            try
                            {
                                var buttonElement = page.Locator(selector).First;
                                if (await buttonElement.CountAsync() > 0)
                                {
                                    await buttonElement.ClickAsync();
                                    Console.WriteLine($"      在主页面使用备用选择器成功点击按钮: {selector}");
                                    clicked = true;
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      备用选择器 {selector} 失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                if (!clicked)
                {
                    Console.WriteLine($"      最终失败：无法找到按钮 {elementId}");
                }
                else
                {
                    Console.WriteLine($"      按钮点击状态: clicked = {clicked}");
                }

                // 等待按钮点击后的页面加载
                await Task.Delay(1000);

                // 特殊处理：如果是打印确认单按钮，调用Python脚本处理后续操作
                if (headerName == "打印确认单" && clicked)
                {
                    Console.WriteLine("      按钮点击成功，开始调用Python脚本处理后续操作...");
                    await HandlePrintConfirmButton();
                }
                else if (headerName == "打印确认单" && !clicked)
                {
                    Console.WriteLine("      警告：打印确认单按钮点击失败，但尝试调用Python脚本...");
                    // 即使点击失败，也尝试调用Python脚本，因为可能按钮已经被点击了
                    await HandlePrintConfirmButton();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击按钮失败: {ex.Message}");
            }
        }



        private async Task FillCaptcha(string captcha)
        {
            try
            {
                Console.WriteLine($"      开始填写验证码: {captcha}");

                // 尝试常见的验证码输入框选择器
                string[] captchaSelectors = {
                    "input[name='captcha']",
                    "input[id*='captcha']",
                    "input[placeholder*='验证码']",
                    "input[placeholder*='captcha']",
                    "#captcha",
                    ".captcha-input"
                };

                bool captchaFilled = false;
                foreach (string selector in captchaSelectors)
                {
                    try
                    {
                        await page.WaitForSelectorAsync(selector, new PageWaitForSelectorOptions { Timeout = 1000 });
                        await page.FillAsync(selector, captcha);
                        Console.WriteLine($"      成功填写验证码: {captcha}");
                        captchaFilled = true;
                        break;
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (!captchaFilled)
                {
                    Console.WriteLine("      警告：未找到验证码输入框，请手动输入验证码");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      填写验证码失败: {ex.Message}");
            }
        }

        private async Task WaitOperation(string waitValue)
        {
            try
            {
                // 解析等待时间
                if (double.TryParse(waitValue, out double waitSeconds))
                {
                    Console.WriteLine($"      开始等待 {waitSeconds} 秒...");

                    // 等待指定的秒数
                    await Task.Delay((int)(waitSeconds * 1000));

                    Console.WriteLine($"      等待 {waitSeconds} 秒完成");
                }
                else
                {
                    Console.WriteLine($"      警告：无法解析等待时间 '{waitValue}'，期望数字格式");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      等待操作失败: {ex.Message}");
            }
        }

        private async Task PressEnterKey()
        {
            try
            {
                Console.WriteLine("      开始模拟回车键操作...");

                // 方法1: 在当前页面按回车键
                try
                {
                    await page.Keyboard.PressAsync("Enter");
                    Console.WriteLine("      成功在当前页面按回车键");
                    await Task.Delay(500); // 等待页面响应
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      在当前页面按回车键失败: {ex.Message}");
                }

                // 方法2: 在iframe中按回车键
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        await frame.EvaluateAsync("document.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true}));");
                        Console.WriteLine("      成功在iframe中按回车键");
                        await Task.Delay(500); // 等待页面响应
                        return;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中按回车键失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法3: 在焦点元素上按回车键
                try
                {
                    await page.EvaluateAsync("document.activeElement && document.activeElement.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', code: 'Enter', keyCode: 13, which: 13, bubbles: true}));");
                    Console.WriteLine("      成功在焦点元素上按回车键");
                    await Task.Delay(500); // 等待页面响应
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      在焦点元素上按回车键失败: {ex.Message}");
                }

                Console.WriteLine("      警告：无法执行回车键操作");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      回车键操作失败: {ex.Message}");
            }
        }

        private async Task SelectFirstTableRow()
        {
            try
            {
                Console.WriteLine("      开始选择表格第一行...");
                await Task.Delay(1000); // 等待表格加载

                // 方法1: 在主页面查找表格第一行
                try
                {
                    var firstRow = page.Locator("#gridWF_GF6_418 tr[id*='418_']").First;
                    if (await firstRow.CountAsync() > 0)
                    {
                        await firstRow.ClickAsync();
                        Console.WriteLine("      成功在主页面点击表格第一行");
                        await Task.Delay(500); // 等待页面响应
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      在主页面查找表格第一行失败: {ex.Message}");
                }

                // 方法2: 在iframe中查找表格第一行
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var firstRow = frame.Locator("#gridWF_GF6_418 tr[id*='418_']").First;
                        if (await firstRow.CountAsync() > 0)
                        {
                            await firstRow.ClickAsync();
                            Console.WriteLine("      成功在iframe中点击表格第一行");
                            await Task.Delay(500); // 等待页面响应
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找表格第一行失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法3: 使用更通用的选择器
                try
                {
                    var firstRow = page.Locator("table.ui-jqgrid-btable tr[id*='418_']").First;
                    if (await firstRow.CountAsync() > 0)
                    {
                        await firstRow.ClickAsync();
                        Console.WriteLine("      成功使用通用选择器点击表格第一行");
                        await Task.Delay(500); // 等待页面响应
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      使用通用选择器查找表格第一行失败: {ex.Message}");
                }

                // 方法4: 在iframe中使用通用选择器
                foreach (var frame in frames)
                {
                    try
                    {
                        var firstRow = frame.Locator("table.ui-jqgrid-btable tr[id*='418_']").First;
                        if (await firstRow.CountAsync() > 0)
                        {
                            await firstRow.ClickAsync();
                            Console.WriteLine("      成功在iframe中使用通用选择器点击表格第一行");
                            await Task.Delay(500); // 等待页面响应
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中使用通用选择器查找表格第一行失败: {ex.Message}");
                        continue;
                    }
                }

                Console.WriteLine("      警告：无法找到表格第一行");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      选择表格第一行失败: {ex.Message}");
            }
        }

        private async Task ClickNavigationButton()
        {
            try
            {
                Console.WriteLine("      开始处理网上预约报账导航按钮...");

                // 方法1: 通过onclick属性查找
                try
                {
                    var navigationElement = page.Locator("div[onclick*='navToPrj(\"WF_YB6\")']").First;
                    if (await navigationElement.CountAsync() > 0)
                    {
                        await navigationElement.ClickAsync();
                        Console.WriteLine("      成功点击网上预约报账导航按钮（通过onclick属性）");
                        await Task.Delay(2000); // 等待页面跳转
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      通过onclick属性查找失败: {ex.Message}");
                }

                // 方法2: 通过JavaScript直接调用
                try
                {
                    await page.EvaluateAsync("navToPrj('WF_YB6')");
                    Console.WriteLine("      成功调用navToPrj('WF_YB6')函数");
                    await Task.Delay(2000); // 等待页面跳转
                    return;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      JavaScript调用失败: {ex.Message}");
                }

                // 方法3: 通过class和onclick组合查找
                try
                {
                    var syslinkElement = page.Locator("div.syslink[onclick*='WF_YB6']").First;
                    if (await syslinkElement.CountAsync() > 0)
                    {
                        await syslinkElement.ClickAsync();
                        Console.WriteLine("      成功点击网上预约报账导航按钮（通过class+onclick）");
                        await Task.Delay(2000); // 等待页面跳转
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      通过class+onclick查找失败: {ex.Message}");
                }

                // 方法4: 通过第一个syslink元素查找（如果只有一个导航选项）
                try
                {
                    var firstSyslink = page.Locator("div.syslink").First;
                    if (await firstSyslink.CountAsync() > 0)
                    {
                        await firstSyslink.ClickAsync();
                        Console.WriteLine("      成功点击第一个导航按钮");
                        await Task.Delay(2000); // 等待页面跳转
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      点击第一个导航按钮失败: {ex.Message}");
                }

                Console.WriteLine("      警告：无法找到网上预约报账导航按钮");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击导航按钮失败: {ex.Message}");
            }
        }



        private async Task ClickAppointmentButton()
        {
            try
            {
                Console.WriteLine("      开始处理预约按钮...");

                // 等待页面加载
                await Task.Delay(2000);

                // 方法1: 在iframe中查找表格中的第一行预约按钮
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        // 查找表格中的第一行预约按钮 - 使用更精确的选择器
                        var firstAppointmentButton = frame.Locator("table.ui-jqgrid-btable tr:first-child button[btnname='预约']").First;
                        if (await firstAppointmentButton.CountAsync() > 0)
                        {
                            await firstAppointmentButton.ClickAsync();
                            Console.WriteLine("      在iframe中成功点击第一行预约按钮");
                            await Task.Delay(2000); // 等待页面响应
                            return;
                        }

                        // 备用方法：通过ID模式查找第一个预约按钮（ID格式：colbtn数字_0_0）
                        var firstAppointmentButtonById = frame.Locator("button[id^='colbtn'][id$='_0_0']").First;
                        if (await firstAppointmentButtonById.CountAsync() > 0)
                        {
                            await firstAppointmentButtonById.ClickAsync();
                            Console.WriteLine("      在iframe中通过ID模式成功点击第一行预约按钮");
                            await Task.Delay(2000); // 等待页面响应
                            return;
                        }

                        // 备用方法2：查找表格中第一行的按钮
                        var firstRowButton = frame.Locator("table.ui-jqgrid-btable tr:first-child td:last-child button").First;
                        if (await firstRowButton.CountAsync() > 0)
                        {
                            await firstRowButton.ClickAsync();
                            Console.WriteLine("      在iframe中成功点击表格第一行最后一个单元格的按钮");
                            await Task.Delay(2000); // 等待页面响应
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找预约按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 在主页面查找表格中的第一行预约按钮
                try
                {
                    var firstAppointmentButton = page.Locator("table.ui-jqgrid-btable tr:first-child button[btnname='预约']").First;
                    if (await firstAppointmentButton.CountAsync() > 0)
                    {
                        await firstAppointmentButton.ClickAsync();
                        Console.WriteLine("      在主页面成功点击第一行预约按钮");
                        await Task.Delay(2000); // 等待页面响应
                        return;
                    }

                    // 备用方法：通过ID模式查找第一个预约按钮（ID格式：colbtn数字_0_0）
                    var firstAppointmentButtonById = page.Locator("button[id^='colbtn'][id$='_0_0']").First;
                    if (await firstAppointmentButtonById.CountAsync() > 0)
                    {
                        await firstAppointmentButtonById.ClickAsync();
                        Console.WriteLine("      在主页面通过ID模式成功点击第一行预约按钮");
                        await Task.Delay(2000); // 等待页面响应
                        return;
                    }

                    // 备用方法2：查找表格中第一行的按钮
                    var firstRowButton = page.Locator("table.ui-jqgrid-btable tr:first-child td:last-child button").First;
                    if (await firstRowButton.CountAsync() > 0)
                    {
                        await firstRowButton.ClickAsync();
                        Console.WriteLine("      在主页面成功点击表格第一行最后一个单元格的按钮");
                        await Task.Delay(2000); // 等待页面响应
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      在主页面查找预约按钮失败: {ex.Message}");
                }

                // 方法3: 使用更通用的选择器
                try
                {
                    var firstAppointmentButton = page.Locator("button[btnname='预约']").First;
                    if (await firstAppointmentButton.CountAsync() > 0)
                    {
                        await firstAppointmentButton.ClickAsync();
                        Console.WriteLine("      成功点击第一个预约按钮");
                        await Task.Delay(2000); // 等待页面响应
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      使用通用选择器查找预约按钮失败: {ex.Message}");
                }

                // 方法4: 在iframe中使用通用选择器
                foreach (var frame in frames)
                {
                    try
                    {
                        var firstAppointmentButton = frame.Locator("button[btnname='预约']").First;
                        if (await firstAppointmentButton.CountAsync() > 0)
                        {
                            await firstAppointmentButton.ClickAsync();
                            Console.WriteLine("      在iframe中成功点击第一个预约按钮");
                            await Task.Delay(2000); // 等待页面响应
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中使用通用选择器查找预约按钮失败: {ex.Message}");
                        continue;
                    }
                }

                Console.WriteLine("      警告：无法找到预约按钮");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击预约按钮失败: {ex.Message}");
            }
        }

        private async Task ClickRadioButton(string radioValue)
        {
            try
            {
                string elementId = GetElementId(radioValue);
                if (string.IsNullOrEmpty(elementId))
                {
                    Console.WriteLine($"      警告：未找到Radio值 '{radioValue}' 对应的元素ID");
                    return;
                }

                Console.WriteLine($"      点击Radio按钮: {radioValue} -> {elementId}");

                // 实现实际的Radio按钮点击逻辑
                bool clicked = false;

                // 等待页面完全加载
                await Task.Delay(500);

                // 方法1: 优先在iframe中通过value属性查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var radioElement = frame.Locator($"input[type='radio'][name='{elementId}'][value='{elementId}']").First;
                        if (await radioElement.CountAsync() > 0)
                        {
                            await radioElement.ClickAsync();
                            Console.WriteLine($"      在iframe中通过value成功点击Radio按钮: {elementId}");
                            clicked = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中通过value查找Radio按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 在iframe中通过ID查找
                if (!clicked)
                {
                    foreach (var frame in frames)
                    {
                        try
                        {
                            var radioElement = frame.Locator($"#{elementId}").First;
                            if (await radioElement.CountAsync() > 0)
                            {
                                await radioElement.ClickAsync();
                                Console.WriteLine($"      在iframe中通过ID成功点击Radio按钮: {elementId}");
                                clicked = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      在iframe中通过ID查找Radio按钮失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                // 方法3: 在主页面通过value属性查找
                if (!clicked)
                {
                    try
                    {
                        var radioElement = page.Locator($"input[type='radio'][name='{elementId}'][value='{elementId}']").First;
                        if (await radioElement.CountAsync() > 0)
                        {
                            await radioElement.ClickAsync();
                            Console.WriteLine($"      在主页面通过value成功点击Radio按钮: {elementId}");
                            clicked = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面通过value查找Radio按钮失败: {ex.Message}");
                    }
                }

                // 方法4: 在主页面通过ID查找
                if (!clicked)
                {
                    try
                    {
                        await page.WaitForSelectorAsync($"#{elementId}", new PageWaitForSelectorOptions { Timeout = 3000 });
                        await page.ClickAsync($"#{elementId}");
                        Console.WriteLine($"      在主页面通过ID成功点击Radio按钮: {elementId}");
                        clicked = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面通过ID查找Radio按钮失败: {ex.Message}");
                    }
                }

                // 方法5: 尝试通过文本内容查找
                if (!clicked)
                {
                    try
                    {
                        // 查找包含指定文本的span元素，然后找到其父级的radio按钮
                        var spanElement = page.Locator($"span:has-text('{radioValue}')").First;
                        if (await spanElement.CountAsync() > 0)
                        {
                            // 找到span的父级li元素中的radio按钮
                            var radioElement = spanElement.Locator("xpath=../input[@type='radio']").First;
                            if (await radioElement.CountAsync() > 0)
                            {
                                await radioElement.ClickAsync();
                                Console.WriteLine($"      通过文本内容成功点击Radio按钮: {radioValue}");
                                clicked = true;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      通过文本内容查找Radio按钮失败: {ex.Message}");
                    }
                }

                // 方法6: 尝试其他选择器
                if (!clicked)
                {
                    string[] alternativeSelectors = {
                         $"input[type='radio'][value*='{elementId}']",
                         $"input[type='radio'][name*='{elementId}']",
                         $"input[type='radio']:has-text('{radioValue}')"
                     };

                    foreach (string selector in alternativeSelectors)
                    {
                        try
                        {
                            // 在iframe中查找
                            foreach (var frame in frames)
                            {
                                try
                                {
                                    var radioElement = frame.Locator(selector).First;
                                    if (await radioElement.CountAsync() > 0)
                                    {
                                        await radioElement.ClickAsync();
                                        Console.WriteLine($"      在iframe中使用备用选择器成功点击Radio按钮: {selector}");
                                        clicked = true;
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    continue;
                                }
                            }

                            if (clicked) break;

                            // 在主页面查找
                            try
                            {
                                var radioElement = page.Locator(selector).First;
                                if (await radioElement.CountAsync() > 0)
                                {
                                    await radioElement.ClickAsync();
                                    Console.WriteLine($"      在主页面使用备用选择器成功点击Radio按钮: {selector}");
                                    clicked = true;
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                continue;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      备用选择器 {selector} 失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                if (!clicked)
                {
                    Console.WriteLine($"      最终失败：无法找到Radio按钮 {elementId}");
                }

                // 等待Radio按钮点击后的页面加载
                await Task.Delay(1000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击Radio按钮失败: {ex.Message}");
            }
        }

        private async Task SelectDate(string headerName, string dateValue)
        {
            try
            {
                string elementId = GetElementId(headerName);
                if (string.IsNullOrEmpty(elementId))
                {
                    Console.WriteLine($"      警告：未找到标题 '{headerName}' 对应的日期控件ID");
                    return;
                }

                Console.WriteLine($"      选择日期: {headerName} -> {elementId} = {dateValue}");

                // 实现实际的日期选择逻辑
                bool selected = false;

                // 方法1: 优先在iframe中查找日期输入框
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var dateElement = frame.Locator($"#{elementId}").First;
                        if (await dateElement.CountAsync() > 0)
                        {
                            // 点击日期输入框以触发日历控件
                            await dateElement.ClickAsync();
                            Console.WriteLine($"      在iframe中点击日期输入框 {elementId} 触发日历控件");

                            // 等待日历控件出现
                            await Task.Delay(2000);

                            // 尝试通过JavaScript直接设置日期值
                            try
                            {
                                await frame.EvaluateAsync($"document.getElementById('{elementId}').value = '{dateValue}'");
                                Console.WriteLine($"      在iframe中通过JavaScript设置日期 {elementId}: {dateValue}");

                                // 触发change事件
                                await frame.EvaluateAsync($"document.getElementById('{elementId}').dispatchEvent(new Event('change', {{ bubbles: true }}))");
                                Console.WriteLine($"      在iframe中触发change事件 {elementId}");

                                selected = true;
                                break;
                            }
                            catch (Exception jsEx)
                            {
                                Console.WriteLine($"      通过JavaScript设置日期失败: {jsEx.Message}");

                                // 如果JavaScript方法失败，尝试通过日历控件选择日期
                                selected = await SelectDateFromCalendar(frame, dateValue);
                                if (selected) break;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找日期输入框失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 如果iframe中找不到，尝试在主页面查找
                if (!selected)
                {
                    try
                    {
                        await page.WaitForSelectorAsync($"#{elementId}", new PageWaitForSelectorOptions { Timeout = 3000 });

                        // 点击日期输入框以触发日历控件
                        await page.ClickAsync($"#{elementId}");
                        Console.WriteLine($"      在主页面点击日期输入框 {elementId} 触发日历控件");

                        // 等待日历控件出现
                        await Task.Delay(2000);

                        // 尝试通过JavaScript直接设置日期值
                        try
                        {
                            await page.EvaluateAsync($"document.getElementById('{elementId}').value = '{dateValue}'");
                            Console.WriteLine($"      在主页面通过JavaScript设置日期 {elementId}: {dateValue}");

                            // 触发change事件
                            await page.EvaluateAsync($"document.getElementById('{elementId}').dispatchEvent(new Event('change', {{ bubbles: true }}))");
                            Console.WriteLine($"      在主页面触发change事件 {elementId}");

                            selected = true;
                        }
                        catch (Exception jsEx)
                        {
                            Console.WriteLine($"      通过JavaScript设置日期失败: {jsEx.Message}");

                            // 如果JavaScript方法失败，尝试通过日历控件选择日期
                            selected = await SelectDateFromCalendar(page, dateValue);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面查找日期输入框失败: {ex.Message}");
                    }
                }

                // 方法3: 如果还是找不到，尝试通过name属性查找
                if (!selected)
                {
                    foreach (var frame in frames)
                    {
                        try
                        {
                            var dateElement = frame.Locator($"input[name='{elementId}']").First;
                            if (await dateElement.CountAsync() > 0)
                            {
                                // 点击日期输入框以触发日历控件
                                await dateElement.ClickAsync();
                                Console.WriteLine($"      在iframe中通过name属性点击日期输入框 {elementId} 触发日历控件");

                                // 等待日历控件出现
                                await Task.Delay(2000);

                                // 尝试通过JavaScript直接设置日期值
                                try
                                {
                                    await frame.EvaluateAsync($"document.querySelector('input[name=\"{elementId}\"]').value = '{dateValue}'");
                                    Console.WriteLine($"      在iframe中通过name属性JavaScript设置日期 {elementId}: {dateValue}");

                                    // 触发change事件
                                    await frame.EvaluateAsync($"document.querySelector('input[name=\"{elementId}\"]').dispatchEvent(new Event('change', {{ bubbles: true }}))");
                                    Console.WriteLine($"      在iframe中通过name属性触发change事件 {elementId}");

                                    selected = true;
                                    break;
                                }
                                catch (Exception jsEx)
                                {
                                    Console.WriteLine($"      通过name属性JavaScript设置日期失败: {jsEx.Message}");

                                    // 如果JavaScript方法失败，尝试通过日历控件选择日期
                                    selected = await SelectDateFromCalendar(frame, dateValue);
                                    if (selected) break;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      在iframe中通过name属性查找失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                if (!selected)
                {
                    Console.WriteLine($"      最终失败：无法找到日期输入框 {elementId}");
                }

                // 等待日期选择后的页面加载
                await Task.Delay(1000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      选择日期失败: {ex.Message}");
            }
        }

        private async Task<bool> SelectDateFromCalendar(IPage page, string dateValue)
        {
            try
            {
                Console.WriteLine($"      尝试通过日历控件选择日期: {dateValue}");

                // 解析日期
                if (DateTime.TryParse(dateValue, out DateTime targetDate))
                {
                    // 等待日历控件出现
                    await page.WaitForSelectorAsync(".ui-datepicker", new PageWaitForSelectorOptions { Timeout = 5000 });
                    Console.WriteLine($"      日历控件已出现");

                    // 选择年份
                    var yearSelect = page.Locator(".ui-datepicker-year").First;
                    if (await yearSelect.CountAsync() > 0)
                    {
                        await yearSelect.SelectOptionAsync(targetDate.Year.ToString());
                        Console.WriteLine($"      选择年份: {targetDate.Year}");
                    }

                    // 选择月份
                    var monthSelect = page.Locator(".ui-datepicker-month").First;
                    if (await monthSelect.CountAsync() > 0)
                    {
                        await monthSelect.SelectOptionAsync((targetDate.Month - 1).ToString());
                        Console.WriteLine($"      选择月份: {targetDate.Month}");
                    }

                    // 选择日期
                    var dayLink = page.Locator($".ui-datepicker-calendar td[data-year='{targetDate.Year}'][data-month='{targetDate.Month - 1}'] a:has-text('{targetDate.Day}')").First;
                    if (await dayLink.CountAsync() > 0)
                    {
                        await dayLink.ClickAsync();
                        Console.WriteLine($"      选择日期: {targetDate.Day}");
                        return true;
                    }
                    else
                    {
                        Console.WriteLine($"      未找到日期链接: {targetDate.Day}");
                    }
                }
                else
                {
                    Console.WriteLine($"      无法解析日期格式: {dateValue}");
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      通过日历控件选择日期失败: {ex.Message}");
                return false;
            }
        }

        private async Task<bool> SelectDateFromCalendar(IFrame frame, string dateValue)
        {
            try
            {
                Console.WriteLine($"      在iframe中尝试通过日历控件选择日期: {dateValue}");

                // 解析日期
                if (DateTime.TryParse(dateValue, out DateTime targetDate))
                {
                    // 等待日历控件出现
                    await frame.WaitForSelectorAsync(".ui-datepicker", new FrameWaitForSelectorOptions { Timeout = 5000 });
                    Console.WriteLine($"      iframe中日历控件已出现");

                    // 选择年份
                    var yearSelect = frame.Locator(".ui-datepicker-year").First;
                    if (await yearSelect.CountAsync() > 0)
                    {
                        await yearSelect.SelectOptionAsync(targetDate.Year.ToString());
                        Console.WriteLine($"      在iframe中选择年份: {targetDate.Year}");
                    }

                    // 选择月份
                    var monthSelect = frame.Locator(".ui-datepicker-month").First;
                    if (await monthSelect.CountAsync() > 0)
                    {
                        await monthSelect.SelectOptionAsync((targetDate.Month - 1).ToString());
                        Console.WriteLine($"      在iframe中选择月份: {targetDate.Month}");
                    }

                    // 选择日期
                    var dayLink = frame.Locator($".ui-datepicker-calendar td[data-year='{targetDate.Year}'][data-month='{targetDate.Month - 1}'] a:has-text('{targetDate.Day}')").First;
                    if (await dayLink.CountAsync() > 0)
                    {
                        await dayLink.ClickAsync();
                        Console.WriteLine($"      在iframe中选择日期: {targetDate.Day}");
                        return true;
                    }
                    else
                    {
                        Console.WriteLine($"      在iframe中未找到日期链接: {targetDate.Day}");
                    }
                }
                else
                {
                    Console.WriteLine($"      无法解析日期格式: {dateValue}");
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      在iframe中通过日历控件选择日期失败: {ex.Message}");
                return false;
            }
        }

        private string GetElementId(string title)
        {
            // 如果在第二种子序列中，使用 "表头-子序列内行序号" 的格式查找ID
            if (isInSecondSubsequence)
            {
                string lookupKey = $"{title}-{subsequenceRowIndex}";
                Console.WriteLine($"      第二种子序列ID查找：{title} -> {lookupKey}");

                if (titleIdMapping.ContainsKey(lookupKey))
                {
                    return titleIdMapping[lookupKey];
                }
                else
                {
                    Console.WriteLine($"      警告：在第二种子序列中未找到ID映射：{lookupKey}");
                    return null;
                }
            }
            else
            {
                // 普通ID查找方式
                if (titleIdMapping.ContainsKey(title))
                {
                    return titleIdMapping[title];
                }
                return null;
            }
        }

        private async Task ProcessRowFromColumn(ExcelWorksheet worksheet, List<string> headers, int row, int startCol, int totalCols)
        {
            Console.WriteLine($"从第 {row} 行的第 {GetColumnName(startCol)} 列开始处理");

            // 找到子序列相关列的索引
            var subsequenceStartColumns = new List<int>();
            var subsequenceEndColumns = new List<int>();

            for (int i = 0; i < headers.Count; i++)
            {
                if (headers[i] == SubsequenceStartColumn)
                {
                    subsequenceStartColumns.Add(i + 1);
                }
                else if (headers[i] == SubsequenceEndColumn)
                {
                    subsequenceEndColumns.Add(i + 1);
                }
            }

            for (int col = startCol; col <= totalCols; col++)
            {
                var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                var columnName = GetColumnName(col);
                var headerName = headers[col - 1];

                Console.WriteLine($"  读取单元格 {columnName}{row}，列标题: {headerName}，值: '{cellValue}'");

                // 检查是否遇到子序列开始标记
                Console.WriteLine($"      当前列 {col}，子序列开始列列表: [{string.Join(", ", subsequenceStartColumns)}]，是否包含当前列: {subsequenceStartColumns.Contains(col)}");
                if (subsequenceStartColumns.Contains(col))
                {
                    Console.WriteLine($"      检查子序列开始标记：当前列 {col}，单元格值 '{cellValue}'，第一种子序列标记 '{SubsequenceMarker}'，第二种子序列标记 '{SubsequenceMarker2}'");

                    if (cellValue == SubsequenceMarker)
                    {
                        Console.WriteLine($"检测到第一种子序列开始标记，进入子序列处理逻辑");
                        // 找到对应的子序列结束列
                        int subsequenceEndColIndex = FindCorrespondingEndColumn(col, subsequenceStartColumns, subsequenceEndColumns);
                        int nextRow = await ProcessSubsequence(worksheet, headers, row, worksheet.Dimension?.Rows ?? 0, col, subsequenceEndColIndex);

                        // 如果子序列处理返回了下一行的行号，继续处理那一行
                        if (nextRow > 0)
                        {
                            Console.WriteLine($"子序列处理完成，继续处理第 {nextRow} 行");
                            await ProcessRowFromColumn(worksheet, headers, nextRow, subsequenceEndColIndex + 1, totalCols);
                        }
                        return; // 退出当前方法
                    }
                    else if (cellValue == SubsequenceMarker2)
                    {
                        Console.WriteLine($"检测到第二种子序列开始标记，进入第二种子序列处理逻辑");
                        // 找到对应的子序列结束列
                        int subsequenceEndColIndex = FindCorrespondingEndColumn(col, subsequenceStartColumns, subsequenceEndColumns);
                        int nextRow = await ProcessSecondSubsequence(worksheet, headers, row, worksheet.Dimension?.Rows ?? 0, col, subsequenceEndColIndex);

                        // 如果子序列处理返回了下一行的行号，继续处理那一行
                        if (nextRow > 0)
                        {
                            Console.WriteLine($"第二种子序列处理完成，继续处理第 {nextRow} 行");
                            await ProcessRowFromColumn(worksheet, headers, nextRow, subsequenceEndColIndex + 1, totalCols);
                        }
                        return; // 退出当前方法
                    }
                    else
                    {
                        Console.WriteLine($"      子序列开始列的值 '{cellValue}' 不匹配任何子序列标记");
                    }
                }

                // 如果不是子序列开始列，则正常处理单元格
                if (!subsequenceStartColumns.Contains(col) && (!string.IsNullOrEmpty(cellValue) || headerName.StartsWith("?")))
                {
                    await ProcessCell(columnName, row, headerName, cellValue);
                }
            }

            Console.WriteLine($"第 {row} 行从第 {GetColumnName(startCol)} 列开始的处理完成");
        }

        private string GetColumnName(int columnIndex)
        {
            string columnName = "";
            while (columnIndex > 0)
            {
                columnIndex--;
                columnName = (char)('A' + columnIndex % 26) + columnName;
                columnIndex /= 26;
            }
            return columnName;
        }

        private int FindCorrespondingEndColumn(int startColumn, List<int> startColumns, List<int> endColumns)
        {
            // 找到当前开始列在所有开始列中的索引
            int startIndex = startColumns.IndexOf(startColumn);

            // 如果找到对应的结束列，返回它；否则返回最后一个结束列
            if (startIndex >= 0 && startIndex < endColumns.Count)
            {
                return endColumns[startIndex];
            }
            else if (endColumns.Count > 0)
            {
                return endColumns[endColumns.Count - 1];
            }

            return -1; // 没有找到对应的结束列
        }

        private bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }

        private bool IsDate(string value)
        {
            // 检查是否为yyyy-mm-dd格式
            return Regex.IsMatch(value, @"^\d{4}-\d{2}-\d{2}$");
        }

        private bool IsDateField(string headerName)
        {
            // 检查是否是日期字段
            string[] dateFields = { "起", "迄", "开始日期", "结束日期", "日期" };
            return dateFields.Any(field => headerName.Contains(field));
        }

        private void InitializeDropdownMappings()
        {
            dropdownMappings = new Dictionary<string, Dictionary<string, string>>
            {
                ["支付方式"] = new Dictionary<string, string>
                {
                    ["个人转卡"] = "10",
                    ["转账汇款"] = "2",
                    ["合同支付"] = "11",
                    ["混合支付"] = "14",
                    ["冲销其它项目借款"] = "9",
                    ["公务卡认证还款"] = "15"
                },
                ["人员类型"] = new Dictionary<string, string>
                {
                    ["院士"] = "院士",
                    ["国家级人才或同等层次人才"] = "国家级人才或同等层次人才",
                    ["2级教授"] = "2级教授",
                    ["高级职称人员"] = "高级职称人员",
                    ["其他人员"] = "其他人员"
                },
                ["人员类别"] = new Dictionary<string, string>
                {
                    ["学生"] = "T^行内转卡/4-学生",
                    ["退休人员"] = "T^行内转卡/2-退休人员",
                    ["在职人员"] = "T^行内转卡/2-在职人员",
                    ["长期聘用人员"] = "T^行内转卡/2-长期聘用人员",
                    ["全部人员"] = "TF^行内转卡/2-全部人员",
                    ["博士后"] = "T^行内转卡/2-博士后",
                    ["校外人员"] = "F^行内转卡/7-校外人员"
                },
                ["省份"] = new Dictionary<string, string>
                {
                    ["北京市"] = "北京市",
                    ["天津市"] = "天津市",
                    ["河北省（石家庄、廊坊、保定）"] = "河北省（石家庄、廊坊、保定）",
                    ["河北省（其他地区）"] = "河北省（其他地区）",
                    ["山西省（太原、大同、晋城）"] = "山西省（太原、大同、晋城）",
                    ["山西省（其他地区）"] = "山西省（其他地区）",
                    ["内蒙古（呼和浩特）"] = "内蒙古（呼和浩特）",
                    ["内蒙古（其他地区）"] = "内蒙古（其他地区）",
                    ["辽宁省（沈阳）"] = "辽宁省（沈阳）",
                    ["辽宁省（其他地区）"] = "辽宁省（其他地区）",
                    ["大连市"] = "大连市",
                    ["吉林省（长春）"] = "吉林省（长春）",
                    ["吉林省（其他地区）"] = "吉林省（其他地区）",
                    ["黑龙江省（哈尔滨）"] = "黑龙江省（哈尔滨）",
                    ["黑龙江省（其他地区）"] = "黑龙江省（其他地区）",
                    ["上海市"] = "上海市",
                    ["江苏省（南京、苏州、无锡、常州、镇江）"] = "江苏省（南京、苏州、无锡、常州、镇江）",
                    ["江苏省（其他地区）"] = "江苏省（其他地区）",
                    ["浙江省（杭州）"] = "浙江省（杭州）",
                    ["浙江省（其他地区）"] = "浙江省（其他地区）",
                    ["宁波市"] = "宁波市",
                    ["安徽省"] = "安徽省",
                    ["福建省（福州、泉州、平潭综合实验区）"] = "福建省（福州、泉州、平潭综合实验区）",
                    ["福建省（其他地区）"] = "福建省（其他地区）",
                    ["厦门市"] = "厦门市",
                    ["江西省"] = "江西省",
                    ["山东省（济南、淄博、枣庄、东营、潍坊、济宁、泰安）"] = "山东省（济南、淄博、枣庄、东营、潍坊、济宁、泰安）",
                    ["山东省（其他地区）"] = "山东省（其他地区）",
                    ["青岛市"] = "青岛市",
                    ["河南省（郑州）"] = "河南省（郑州）",
                    ["河南省（其他地区）"] = "河南省（其他地区）",
                    ["湖北省（武汉）"] = "湖北省（武汉）",
                    ["湖北省（其他地区）"] = "湖北省（其他地区）",
                    ["湖南省（长沙）"] = "湖南省（长沙）",
                    ["湖南省（其他地区）"] = "湖南省（其他地区）",
                    ["广东省（广州、珠海、佛山、东莞、中山、江门）"] = "广东省（广州、珠海、佛山、东莞、中山、江门）",
                    ["广东省（其他地区）"] = "广东省（其他地区）",
                    ["深圳市"] = "深圳市",
                    ["广西（南宁）"] = "广西（南宁）",
                    ["广西（其他地区）"] = "广西（其他地区）",
                    ["海南省(海口、文昌、澄迈县）"] = "海南省(海口、文昌、澄迈县）",
                    ["海南省（其他地区）"] = "海南省（其他地区）",
                    ["重庆市（9个中心城区、北部新区）"] = "重庆市（9个中心城区、北部新区）",
                    ["重庆市(其他地区)"] = "重庆市(其他地区)",
                    ["四川省（成都）"] = "四川省（成都）",
                    ["四川省（其他地区）"] = "四川省（其他地区）",
                    ["贵州省（贵阳）"] = "贵州省（贵阳）",
                    ["贵州省（其他地区）"] = "贵州省（其他地区）",
                    ["云南省（昆明、大理州、丽江、迪庆州、西双版纳州）"] = "云南省（昆明、大理州、丽江、迪庆州、西双版纳州）",
                    ["云南省（其他地区）"] = "云南省（其他地区）",
                    ["西藏（拉萨）"] = "西藏（拉萨）",
                    ["西藏（其他地区）"] = "西藏（其他地区）",
                    ["陕西省（西安）"] = "陕西省（西安）",
                    ["陕西省（其他地区）"] = "陕西省（其他地区）",
                    ["甘肃省（兰州）"] = "甘肃省（兰州）",
                    ["甘肃省(其他地区)"] = "甘肃省(其他地区)",
                    ["青海省（西宁）"] = "青海省（西宁）",
                    ["青海省（其他地区）"] = "青海省（其他地区）",
                    ["宁夏（银川）"] = "宁夏（银川）",
                    ["宁夏（其他地区）"] = "宁夏（其他地区）",
                    ["新疆（乌鲁木齐）"] = "新疆（乌鲁木齐）",
                    ["新疆（其他地区）"] = "新疆（其他地区）"
                },
                ["是否安排伙食"] = new Dictionary<string, string>
                {
                    ["未安排"] = "未安排",
                    ["安排"] = "安排"
                },
                ["是否安排交通"] = new Dictionary<string, string>
                {
                    ["未安排"] = "未安排",
                    ["安排"] = "安排"
                },
                ["酬金性质"] = new Dictionary<string, string>
                {
                    ["劳务费"] = "13",
                    ["省部级及以上奖励"] = "45"
                },
                ["安排状态"] = new Dictionary<string, string>
                {
                    ["未安排"] = "未安排",
                    ["安排"] = "安排"
                },
                ["交通费"] = new Dictionary<string, string>
                {
                    ["未安排"] = "未安排",
                    ["安排"] = "安排"
                },
                ["奖助学金性质"] = new Dictionary<string, string>
                {
                    ["导师助研助学金"] = "导师助研助学金",
                    ["科研经费博士助研费（基本）"] = "科研经费博士助研费（基本）",
                    ["科研经费博士助研费（奖励）"] = "科研经费博士助研费（奖励）"
                }
            };

            Console.WriteLine($"初始化下拉框映射，共 {dropdownMappings.Count} 个字段");
        }

        private bool IsDropdownField(string headerName)
        {
            return dropdownMappings.ContainsKey(headerName);
        }

        private async Task SelectDropdown(string headerName, string displayValue)
        {
            try
            {
                string elementId = GetElementId(headerName);
                if (string.IsNullOrEmpty(elementId))
                {
                    Console.WriteLine($"      警告：未找到标题 '{headerName}' 对应的下拉框ID");
                    return;
                }

                Console.WriteLine($"      选择下拉框: {headerName} -> {elementId} = {displayValue}");

                // 获取映射的值
                string mappedValue = GetDropdownMappedValue(headerName, displayValue);
                if (string.IsNullOrEmpty(mappedValue))
                {
                    Console.WriteLine($"      警告：未找到显示值 '{displayValue}' 对应的映射值");
                    return;
                }

                Console.WriteLine($"      下拉框映射: {displayValue} -> {mappedValue}");

                // 实现实际的下拉框选择逻辑
                bool selected = false;

                // 方法1: 优先在iframe中查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var selectElement = frame.Locator($"#{elementId}").First;
                        if (await selectElement.CountAsync() > 0)
                        {
                            await selectElement.SelectOptionAsync(mappedValue);
                            Console.WriteLine($"      在iframe中成功选择下拉框 {elementId}: {mappedValue}");
                            selected = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找下拉框失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 如果iframe中找不到，尝试在主页面查找
                if (!selected)
                {
                    try
                    {
                        await page.WaitForSelectorAsync($"#{elementId}", new PageWaitForSelectorOptions { Timeout = 3000 });
                        await page.SelectOptionAsync($"#{elementId}", mappedValue);
                        Console.WriteLine($"      在主页面成功选择下拉框 {elementId}: {mappedValue}");
                        selected = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面查找下拉框失败: {ex.Message}");
                    }
                }

                if (!selected)
                {
                    Console.WriteLine($"      最终失败：无法找到下拉框 {elementId}");
                }

                // 等待下拉框选择后的页面加载
                await Task.Delay(500);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      选择下拉框失败: {ex.Message}");
            }
        }

        private string GetDropdownMappedValue(string headerName, string displayValue)
        {
            if (dropdownMappings.ContainsKey(headerName) &&
                dropdownMappings[headerName].ContainsKey(displayValue))
            {
                return dropdownMappings[headerName][displayValue];
            }
            return null;
        }

        private async Task FillSubjectInput(string headerName, string cellValue)
        {
            try
            {
                // 提取科目名称（去掉#前缀）
                string subjectName = cellValue.Substring(1);

                Console.WriteLine($"      处理科目输入框: {headerName} = {cellValue}");
                Console.WriteLine($"      提取科目名称: {subjectName}");

                // 在标题-ID表中查找对应的输入框ID
                string elementId = GetElementId(subjectName);
                if (string.IsNullOrEmpty(elementId))
                {
                    Console.WriteLine($"      警告：未找到科目 '{subjectName}' 对应的ID映射");
                    return;
                }

                Console.WriteLine($"      找到科目输入框ID: {subjectName} -> {elementId}");

                // 存储当前科目ID，供后续金额填写使用
                currentSubjectId = elementId;

                // 特殊处理：科目和金额填写（需要等待页面加载）
                if (headerName == "科目" || headerName == "金额")
                {
                    Console.WriteLine($"      特殊处理{headerName}填写，等待页面加载完成...");
                    await Task.Delay(5000); // 等待页面加载
                    Console.WriteLine($"      页面加载等待完成，开始填写科目: {subjectName}");
                }

                // 实现实际的科目输入框填写逻辑
                bool filled = false;

                // 方法1: 优先在iframe中查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var inputElement = frame.Locator($"#{elementId}").First;
                        if (await inputElement.CountAsync() > 0)
                        {
                            // 对于科目输入框，我们需要等待下一个单元格（金额）的值
                            // 这里先记录科目ID，等待后续的金额值
                            Console.WriteLine($"      找到科目输入框 {elementId}，等待金额值...");
                            filled = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找科目输入框失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 如果iframe中找不到，尝试在主页面查找
                if (!filled)
                {
                    try
                    {
                        await page.WaitForSelectorAsync($"#{elementId}", new PageWaitForSelectorOptions { Timeout = 3000 });
                        Console.WriteLine($"      在主页面找到科目输入框 {elementId}，等待金额值...");
                        filled = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面查找科目输入框失败: {ex.Message}");
                    }
                }

                // 方法3: 如果还是找不到，尝试通过name属性查找
                if (!filled)
                {
                    foreach (var frame in frames)
                    {
                        try
                        {
                            var inputElement = frame.Locator($"input[name='{elementId}']").First;
                            if (await inputElement.CountAsync() > 0)
                            {
                                Console.WriteLine($"      在iframe中通过name属性找到科目输入框 {elementId}，等待金额值...");
                                filled = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      在iframe中通过name属性查找失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                if (!filled)
                {
                    Console.WriteLine($"      最终失败：无法找到科目输入框 {elementId}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      填写科目输入框失败: {ex.Message}");
            }
        }

        private async Task FillAmountInput(string subjectId, string amountValue)
        {
            try
            {
                Console.WriteLine($"      填写金额到科目输入框: {subjectId} = {amountValue}");

                // 实现实际的金额输入框填写逻辑
                bool filled = false;

                // 方法1: 优先在iframe中查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var inputElement = frame.Locator($"#{subjectId}").First;
                        if (await inputElement.CountAsync() > 0)
                        {
                            await inputElement.FillAsync(amountValue);
                            Console.WriteLine($"      在iframe中成功填写金额到科目输入框 {subjectId}: {amountValue}");
                            filled = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找科目输入框失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 如果iframe中找不到，尝试在主页面查找
                if (!filled)
                {
                    try
                    {
                        await page.WaitForSelectorAsync($"#{subjectId}", new PageWaitForSelectorOptions { Timeout = 3000 });
                        await page.FillAsync($"#{subjectId}", amountValue);
                        Console.WriteLine($"      在主页面成功填写金额到科目输入框 {subjectId}: {amountValue}");
                        filled = true;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面查找科目输入框失败: {ex.Message}");
                    }
                }

                // 方法3: 如果还是找不到，尝试通过name属性查找
                if (!filled)
                {
                    foreach (var frame in frames)
                    {
                        try
                        {
                            var inputElement = frame.Locator($"input[name='{subjectId}']").First;
                            if (await inputElement.CountAsync() > 0)
                            {
                                await inputElement.FillAsync(amountValue);
                                Console.WriteLine($"      在iframe中通过name属性成功填写金额到科目输入框 {subjectId}: {amountValue}");
                                filled = true;
                                break;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"      在iframe中通过name属性查找失败: {ex.Message}");
                            continue;
                        }
                    }
                }

                if (!filled)
                {
                    Console.WriteLine($"      最终失败：无法找到科目输入框 {subjectId} 来填写金额");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      填写金额到科目输入框失败: {ex.Message}");
            }
        }

        private async Task SelectCardByNumber(string cardValue)
        {
            try
            {
                // 提取卡号尾号（去掉*前缀）
                string cardTail = cardValue.Substring(1);

                Console.WriteLine($"      开始选择卡号尾号: {cardTail}");

                // 等待页面加载完成
                Console.WriteLine($"      等待页面加载...");
                await Task.Delay(3000);

                bool selected = false;

                // 等待jQuery UI对话框出现
                Console.WriteLine($"      等待银行卡选择对话框出现...");
                await Task.Delay(2000); // 等待对话框加载

                // 在对话框中查找银行卡表格
                selected = await SelectCardInDialog(cardTail);

                // 方法1: 优先在iframe中查找包含指定卡号尾号的td元素，然后找到同行的radio按钮
                var frames = page.Frames;
                Console.WriteLine($"      开始搜索 {frames.Count} 个iframe中的银行卡表格");

                // 添加调试信息：检查页面中是否有银行卡相关的元素
                try
                {
                    var allRadioButtons = page.Locator("input[type='radio'][name='rdoacnt']").AllAsync();
                    Console.WriteLine($"      主页面找到 {allRadioButtons.Result.Count} 个银行卡radio按钮");

                    // 检查页面中是否有包含"卡号"的文本
                    var cardNumberTexts = page.Locator("text=卡号").AllAsync();
                    Console.WriteLine($"      主页面找到 {cardNumberTexts.Result.Count} 个包含'卡号'的文本元素");

                    foreach (var frame in frames)
                    {
                        var frameRadioButtons = frame.Locator("input[type='radio'][name='rdoacnt']").AllAsync();
                        Console.WriteLine($"      iframe中找到 {frameRadioButtons.Result.Count} 个银行卡radio按钮");

                        var frameCardTexts = frame.Locator("text=卡号").AllAsync();
                        Console.WriteLine($"      iframe中找到 {frameCardTexts.Result.Count} 个包含'卡号'的文本元素");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      调试信息获取失败: {ex.Message}");
                }

                foreach (var frame in frames)
                {
                    try
                    {
                        // 先检查这个iframe中是否有银行卡表格
                        var tableElements = frame.Locator("table").AllAsync();
                        var tableCount = tableElements.Result.Count;
                        Console.WriteLine($"      当前iframe中找到 {tableCount} 个表格");

                        if (tableCount > 0)
                        {
                            // 直接使用Python代码的XPath选择器逻辑
                            string radioSelector = $"//tr[td[contains(text(), '{cardTail}')]]/td/input[@type='radio'][@name='rdoacnt']";
                            Console.WriteLine($"      尝试XPath选择器: {radioSelector}");

                            try
                            {
                                var radioElement = frame.Locator(radioSelector).First;
                                var radioCount = await radioElement.CountAsync();
                                Console.WriteLine($"      找到 {radioCount} 个匹配的radio按钮");

                                if (radioCount > 0)
                                {
                                    await radioElement.ClickAsync();
                                    Console.WriteLine($"      在iframe中成功选择卡号尾号 {cardTail} 对应的radio按钮");
                                    await Task.Delay(1000); // 等待元素响应
                                    selected = true;
                                    break;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"      XPath选择器执行失败: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找radio按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 如果iframe中找不到，尝试在主页面查找
                if (!selected)
                {
                    try
                    {
                        string radioSelector = $"//tr[td[contains(text(), '{cardTail}')]]/td/input[@type='radio'][@name='rdoacnt']";
                        var radioElement = page.Locator(radioSelector).First;
                        if (await radioElement.CountAsync() > 0)
                        {
                            await radioElement.ClickAsync();
                            Console.WriteLine($"      在主页面成功选择卡号尾号 {cardTail} 对应的radio按钮");
                            await Task.Delay(1000); // 等待元素响应
                            selected = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面查找radio按钮失败: {ex.Message}");
                    }
                }

                // 方法3: 尝试更通用的选择器 - 使用onclick属性中包含卡号尾号的方式
                if (!selected)
                {
                    try
                    {
                        string radioSelector = $"input[type='radio'][name='rdoacnt'][onclick*='{cardTail}']";
                        Console.WriteLine($"      尝试onclick选择器: {radioSelector}");

                        // 先在主页面尝试
                        var radioElement = page.Locator(radioSelector).First;
                        var mainPageCount = await radioElement.CountAsync();
                        Console.WriteLine($"      主页面找到 {mainPageCount} 个匹配的radio按钮");

                        if (mainPageCount > 0)
                        {
                            await radioElement.ClickAsync();
                            Console.WriteLine($"      通过onclick属性成功选择卡号尾号 {cardTail} 对应的radio按钮");
                            await Task.Delay(1000); // 等待元素响应
                            selected = true;
                        }

                        // 在iframe中尝试
                        if (!selected)
                        {
                            foreach (var frame in frames)
                            {
                                try
                                {
                                    var radioElement2 = frame.Locator(radioSelector).First;
                                    var frameCount = await radioElement2.CountAsync();
                                    Console.WriteLine($"      当前iframe中找到 {frameCount} 个匹配的radio按钮");

                                    if (frameCount > 0)
                                    {
                                        await radioElement2.ClickAsync();
                                        Console.WriteLine($"      在iframe中通过onclick属性成功选择卡号尾号 {cardTail} 对应的radio按钮");
                                        await Task.Delay(1000); // 等待元素响应
                                        selected = true;
                                        break;
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"      在iframe中通过onclick属性查找失败: {ex.Message}");
                                    continue;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      通过onclick属性查找失败: {ex.Message}");
                    }
                }

                if (!selected)
                {
                    Console.WriteLine($"      警告：未找到卡号尾号 {cardTail} 对应的radio按钮");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      选择银行卡失败: {ex.Message}");
            }
        }

        private bool IsBankCardField(string headerName)
        {
            // 检查是否是银行卡相关字段
            string[] bankCardFields = { "转卡信息工号", "卡号尾号", "银行卡", "银行账号", "差旅转卡工号" };
            return bankCardFields.Any(field => headerName.Contains(field));
        }

        private async Task TriggerBankCardSelection(ILocator inputElement)
        {
            try
            {
                Console.WriteLine($"      触发银行卡选择事件...");

                // 方法1: 按回车键
                await inputElement.PressAsync("Enter");
                Console.WriteLine($"      已按回车键触发银行卡选择");

                // 等待一下，让事件处理完成
                await Task.Delay(1000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      触发银行卡选择事件失败: {ex.Message}");
            }
        }

        private async Task<bool> SelectCardInDialog(string cardTail)
        {
            try
            {
                Console.WriteLine($"      在对话框中查找卡号尾号: {cardTail}");

                // 等待对话框加载完成
                await page.WaitForSelectorAsync(".ui-dialog", new PageWaitForSelectorOptions { Timeout = 5000 });
                Console.WriteLine($"      对话框加载完成");

                // 方法1: 使用XPath查找包含卡号尾号的td元素所在的行，然后找到该行中的radio按钮
                try
                {
                    string radioSelector = $"//tr[td[contains(text(), '{cardTail}')]]/td/input[@type='radio'][@name='rdoacnt']";
                    Console.WriteLine($"      尝试XPath选择器: {radioSelector}");

                    var radioElement = page.Locator(radioSelector).First;
                    var radioCount = await radioElement.CountAsync();
                    Console.WriteLine($"      找到 {radioCount} 个匹配的radio按钮");

                    if (radioCount > 0)
                    {
                        await radioElement.ClickAsync();
                        Console.WriteLine($"      在对话框中成功选择卡号尾号 {cardTail} 对应的radio按钮");
                        await Task.Delay(1000); // 等待元素响应
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      XPath选择器执行失败: {ex.Message}");
                }

                // 方法2: 使用onclick属性中包含卡号尾号的方式
                try
                {
                    string radioSelector = $"input[type='radio'][name='rdoacnt'][onclick*='{cardTail}']";
                    Console.WriteLine($"      尝试onclick选择器: {radioSelector}");

                    var radioElement = page.Locator(radioSelector).First;
                    var radioCount = await radioElement.CountAsync();
                    Console.WriteLine($"      找到 {radioCount} 个匹配的radio按钮");

                    if (radioCount > 0)
                    {
                        await radioElement.ClickAsync();
                        Console.WriteLine($"      在对话框中通过onclick属性成功选择卡号尾号 {cardTail} 对应的radio按钮");
                        await Task.Delay(1000); // 等待元素响应
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      onclick选择器执行失败: {ex.Message}");
                }

                Console.WriteLine($"      在对话框中未找到卡号尾号 {cardTail} 对应的radio按钮");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      在对话框中选择银行卡失败: {ex.Message}");
                return false;
            }
        }

        private async Task InitializeBrowser()
        {
            try
            {
                Console.WriteLine("正在启动浏览器...");

                playwright = await Playwright.CreateAsync();
                browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                {
                    Headless = false, // 显示浏览器窗口
                    SlowMo = 100 // 放慢操作速度，便于观察
                });

                page = await browser.NewPageAsync();

                // 设置页面超时时间
                page.SetDefaultTimeout(10000); // 10秒



                Console.WriteLine("浏览器启动成功");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"浏览器启动失败: {ex.Message}");
                throw;
            }
        }



        private async Task NavigateToTargetPage()
        {
            try
            {
                Console.WriteLine("正在导航到目标网页...");

                // 目标URL（电子科技大学财务系统）
                string targetUrl = "https://cwcx.uestc.edu.cn/WFManager/home.jsp";

                await page.GotoAsync(targetUrl, new PageGotoOptions { Timeout = 30000 });

                Console.WriteLine($"成功导航到页面: {targetUrl}");

                // 等待页面加载完成
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });

                Console.WriteLine("页面加载完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"页面导航失败: {ex.Message}");
                throw;
            }
        }

        private async Task ClickLayuiConfirmButton()
        {
            try
            {
                Console.WriteLine("      开始处理layui弹窗确定按钮...");

                // 等待页面加载
                await Task.Delay(2000);

                // 方法1: 在主页面查找layui弹窗的确定按钮
                try
                {
                    // 使用layui弹窗确定按钮的选择器
                    var confirmButton = page.Locator(".layui-layer-btn0").First;
                    if (await confirmButton.CountAsync() > 0)
                    {
                        await confirmButton.ClickAsync();
                        Console.WriteLine("      在主页面成功点击layui弹窗确定按钮");
                        await Task.Delay(2000); // 等待页面响应
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      在主页面查找layui弹窗确定按钮失败: {ex.Message}");
                }

                // 方法2: 在iframe中查找layui弹窗的确定按钮
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var confirmButton = frame.Locator(".layui-layer-btn0").First;
                        if (await confirmButton.CountAsync() > 0)
                        {
                            await confirmButton.ClickAsync();
                            Console.WriteLine("      在iframe中成功点击layui弹窗确定按钮");
                            await Task.Delay(2000); // 等待页面响应
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中查找layui弹窗确定按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法3: 使用更通用的选择器
                try
                {
                    // 查找包含"确定"文本的按钮
                    var confirmButton = page.Locator("a:has-text('确定')").First;
                    if (await confirmButton.CountAsync() > 0)
                    {
                        await confirmButton.ClickAsync();
                        Console.WriteLine("      使用文本选择器成功点击确定按钮");
                        await Task.Delay(2000); // 等待页面响应
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      使用文本选择器查找确定按钮失败: {ex.Message}");
                }

                // 方法4: 在iframe中使用文本选择器
                foreach (var frame in frames)
                {
                    try
                    {
                        var confirmButton = frame.Locator("a:has-text('确定')").First;
                        if (await confirmButton.CountAsync() > 0)
                        {
                            await confirmButton.ClickAsync();
                            Console.WriteLine("      在iframe中使用文本选择器成功点击确定按钮");
                            await Task.Delay(2000); // 等待页面响应
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中使用文本选择器查找确定按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法5: 使用layui弹窗的其他选择器
                try
                {
                    // 查找layui弹窗中的确定按钮
                    var confirmButton = page.Locator(".layui-layer-dialog .layui-layer-btn a:first-child").First;
                    if (await confirmButton.CountAsync() > 0)
                    {
                        await confirmButton.ClickAsync();
                        Console.WriteLine("      使用layui弹窗选择器成功点击确定按钮");
                        await Task.Delay(2000); // 等待页面响应
                        return;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      使用layui弹窗选择器查找确定按钮失败: {ex.Message}");
                }

                Console.WriteLine("      警告：无法找到layui弹窗确定按钮");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击layui弹窗确定按钮失败: {ex.Message}");
            }
        }

        private async Task LoadConfiguration()
        {
            try
            {
                Console.WriteLine("开始加载配置文件...");

                // 获取当前执行目录
                string currentDirectory = Directory.GetCurrentDirectory();
                Console.WriteLine($"当前工作目录: {currentDirectory}");

                // 尝试多个可能的配置文件路径
                string[] possibleConfigPaths = {
                    "config.json",
                    Path.Combine(currentDirectory, "config.json"),
                    Path.Combine(currentDirectory, "..", "config.json"),
                    Path.Combine(currentDirectory, "..", "..", "config.json"),
                    Path.Combine(currentDirectory, "..", "..", "..", "config.json"),
                    Path.Combine(currentDirectory, "..", "..", "..", "..", "config.json")
                };

                string configPath = null;
                foreach (string path in possibleConfigPaths)
                {
                    if (File.Exists(path))
                    {
                        configPath = path;
                        Console.WriteLine($"找到配置文件: {path}");
                        break;
                    }
                }

                if (configPath == null)
                {
                    Console.WriteLine("警告：找不到配置文件，使用默认配置");
                    config = new AppConfig();
                    Console.WriteLine("使用默认配置:");
                    Console.WriteLine($"  ExcelFilePath: {config.ExcelFilePath}");
                    Console.WriteLine($"  MappingFilePath: {config.MappingFilePath}");
                    Console.WriteLine($"  SheetName: {config.SheetName}");
                    Console.WriteLine($"  MappingSheetName: {config.MappingSheetName}");
                    return;
                }

                // 读取配置文件
                string jsonContent = await File.ReadAllTextAsync(configPath);
                config = JsonSerializer.Deserialize<AppConfig>(jsonContent);

                Console.WriteLine("配置文件加载成功:");
                Console.WriteLine($"  ExcelFilePath: {config.ExcelFilePath}");
                Console.WriteLine($"  MappingFilePath: {config.MappingFilePath}");
                Console.WriteLine($"  SheetName: {config.SheetName}");
                Console.WriteLine($"  MappingSheetName: {config.MappingSheetName}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"加载配置文件失败: {ex.Message}");
                Console.WriteLine("使用默认配置");
                config = new AppConfig();
            }
        }

        private async Task Cleanup()
        {
            try
            {
                if (page != null)
                {
                    await page.CloseAsync();
                }

                if (browser != null)
                {
                    await browser.CloseAsync();
                }

                if (playwright != null)
                {
                    playwright.Dispose();
                }

                Console.WriteLine("浏览器已关闭");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"清理资源时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 特殊处理打印确认单按钮点击
        /// </summary>
        private async Task ClickPrintConfirmButton()
        {
            try
            {
                Console.WriteLine("      ========================================");
                Console.WriteLine("      开始特殊处理打印确认单按钮...");
                Console.WriteLine("      ========================================");

                // 在点击按钮之前，先提取预约号和申请总金额信息
                Console.WriteLine("      在点击按钮之前，先提取预约号和申请总金额信息...");
                var (appointmentNumber, totalAmount) = await ExtractAppointmentInfoFromPage();

                Console.WriteLine($"      提取到的预约号: {appointmentNumber}");
                Console.WriteLine($"      提取到的申请总金额: {totalAmount}");

                // 直接使用iframe查找方法
                Console.WriteLine("      直接在iframe中查找打印确认单按钮...");

                bool buttonClicked = false;

                try
                {
                    Console.WriteLine("      在iframe中查找按钮 #BtnPrint");
                    var frames = page.Frames;
                    Console.WriteLine($"      找到 {frames.Count} 个iframe");

                    // 查找所有iframe
                    for (int i = 0; i < frames.Count; i++)
                    {
                        try
                        {
                            Console.WriteLine($"      尝试在iframe {i + 1} 中查找按钮...");
                            await frames[i].WaitForSelectorAsync("#BtnPrint", new FrameWaitForSelectorOptions { Timeout = 2000 });
                            await frames[i].ClickAsync("#BtnPrint");
                            Console.WriteLine($"      ✓ 成功：在iframe {i + 1} 中点击了打印确认单按钮");
                            buttonClicked = true;
                            break;
                        }
                        catch (Exception frameEx)
                        {
                            Console.WriteLine($"      在iframe {i + 1} 中查找失败: {frameEx.Message}");
                            continue;
                        }
                    }

                    if (!buttonClicked)
                    {
                        Console.WriteLine($"      所有 {frames.Count} 个iframe都查找失败，直接执行Python脚本");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      iframe查找失败: {ex.Message}");
                }

                // 无论是否点击成功，都直接执行Python脚本
                if (buttonClicked)
                {
                    Console.WriteLine("      ✓ 打印确认单按钮点击成功！");
                }
                else
                {
                    Console.WriteLine("      ✗ 无法点击打印确认单按钮，但继续执行Python脚本");
                }

                Console.WriteLine("      立即开始调用Python脚本处理后续操作...");
                await HandlePrintConfirmButton(appointmentNumber, totalAmount);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      特殊处理打印确认单按钮时出错: {ex.Message}");
                Console.WriteLine($"      详细错误信息: {ex}");

                // 即使出错，也尝试调用Python脚本
                try
                {
                    Console.WriteLine("      尝试调用Python脚本作为备用方案...");
                    await HandlePrintConfirmButton("null", "0.00");
                }
                catch (Exception pyEx)
                {
                    Console.WriteLine($"      Python脚本调用也失败: {pyEx.Message}");
                }
            }
        }

        /// <summary>
        /// 处理打印确认单按钮点击
        /// </summary>
        private async Task HandlePrintConfirmButton(string appointmentNumber = "null", string totalAmount = "0.00")
        {
            try
            {
                Console.WriteLine("      ========================================");
                Console.WriteLine("      开始处理打印确认单按钮后续操作...");
                Console.WriteLine("      ========================================");

                // 使用传入的预约号和申请总金额信息
                Console.WriteLine($"      使用传入的预约号: {appointmentNumber}");
                Console.WriteLine($"      使用传入的申请总金额: {totalAmount}");

                // 如果传入的是默认值，尝试重新提取
                if (appointmentNumber == "null" || totalAmount == "0.00")
                {
                    Console.WriteLine("      传入的值为默认值，尝试重新从网页提取...");
                    var (newAppointmentNumber, newTotalAmount) = await ExtractAppointmentInfoFromPage();

                    if (appointmentNumber == "null" && newAppointmentNumber != "null")
                    {
                        appointmentNumber = newAppointmentNumber;
                        Console.WriteLine($"      重新提取到预约号: {appointmentNumber}");
                    }

                    if (totalAmount == "0.00" && newTotalAmount != "0.00")
                    {
                        totalAmount = newTotalAmount;
                        Console.WriteLine($"      重新提取到申请总金额: {totalAmount}");
                    }
                }

                // 检查Python脚本执行器是否可用
                if (pythonExecutor == null)
                {
                    Console.WriteLine("      警告：Python脚本执行器未初始化，无法执行后续的鼠标键盘自动化");
                    Console.WriteLine("      打印确认单按钮已点击，但无法处理后续操作");
                    return;
                }

                Console.WriteLine("      调用Python脚本处理后续的鼠标键盘自动化操作...");

                // 直接调用Python脚本
                string scriptPath = "test_mouse_keyboard.py";
                string configPath = "config.json";

                // 尝试找到项目根目录的Python脚本
                string[] possibleScriptPaths = {
                    scriptPath, // 当前目录
                    Path.Combine("..", "..", "..", "..", scriptPath), // 项目根目录
                    Path.Combine("..", "..", "..", scriptPath), // 上级目录
                    Path.Combine("..", "..", scriptPath) // 上上级目录
                };

                string actualScriptPath = null;
                foreach (string path in possibleScriptPaths)
                {
                    if (File.Exists(path))
                    {
                        actualScriptPath = path;
                        Console.WriteLine($"      找到Python脚本: {path}");
                        break;
                    }
                }

                if (actualScriptPath == null)
                {
                    Console.WriteLine($"      ✗ 未找到Python脚本: {scriptPath}");
                    return;
                }

                Console.WriteLine($"      检查脚本文件: {scriptPath}");
                if (!File.Exists(scriptPath))
                {
                    Console.WriteLine($"      ✗ 脚本文件不存在: {scriptPath}");
                    Console.WriteLine($"      当前目录: {Directory.GetCurrentDirectory()}");
                    Console.WriteLine("      请确保test_mouse_keyboard.py文件在当前目录中");
                    return;
                }
                Console.WriteLine($"      ✓ 脚本文件存在: {scriptPath}");

                Console.WriteLine($"      检查配置文件: {configPath}");
                if (!File.Exists(configPath))
                {
                    Console.WriteLine($"      ✗ 配置文件不存在: {configPath}");
                    Console.WriteLine("      请确保config.json文件在当前目录中");
                    return;
                }
                Console.WriteLine($"      ✓ 配置文件存在: {configPath}");

                // 从config.json读取文件夹路径
                string folderPath = "C:\\Users\\FH\\Documents\\报销单"; // 默认路径
                try
                {
                    // 尝试多个可能的配置文件路径
                    string[] possibleConfigPaths = {
                        configPath, // 当前目录
                        Path.Combine("..", "..", "..", "..", configPath), // 项目根目录
                        Path.Combine("..", "..", "..", configPath), // 上级目录
                        Path.Combine("..", "..", configPath) // 上上级目录
                    };

                    string actualConfigPath = null;
                    foreach (string path in possibleConfigPaths)
                    {
                        if (File.Exists(path))
                        {
                            actualConfigPath = path;
                            Console.WriteLine($"      找到配置文件: {path}");
                            break;
                        }
                    }

                    if (actualConfigPath != null)
                    {
                        string configContent = File.ReadAllText(actualConfigPath);
                        var config = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(configContent);
                        if (config.ContainsKey("SaveFolderPath"))
                        {
                            folderPath = config["SaveFolderPath"].ToString();
                            Console.WriteLine($"      从配置文件读取保存路径: {folderPath}");
                        }
                        else
                        {
                            Console.WriteLine($"      配置文件中未找到SaveFolderPath，使用默认路径: {folderPath}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"      未找到配置文件，使用默认路径: {folderPath}");
                    }
                }
                catch (Exception configEx)
                {
                    Console.WriteLine($"      读取配置文件失败，使用默认路径: {configEx.Message}");
                }

                // 执行Python脚本处理后续操作（如打印对话框等）
                // 使用预约号-申请总金额-时间的格式命名文件
                string timeStamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileName = $"{appointmentNumber}-{totalAmount}-{timeStamp}.pdf";
                string arguments = $"--config config.json --folder \"{folderPath}\" --file \"{fileName}\"";

                Console.WriteLine($"      执行命令: python {scriptPath} {arguments}");
                Console.WriteLine($"      工作目录: {Directory.GetCurrentDirectory()}");

                Console.WriteLine($"      开始执行Python脚本...");
                Console.WriteLine($"      Python解释器: {pythonExecutor.GetType().GetField("_pythonExecutable", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)?.GetValue(pythonExecutor)}");

                var result = await pythonExecutor.ExecuteScriptAsync(
                    scriptPath: actualScriptPath,
                    arguments: arguments,
                    timeoutMilliseconds: 120000 // 2分钟超时
                );

                Console.WriteLine($"      执行结果 - 成功: {result.Success}, 退出码: {result.ExitCode}");
                Console.WriteLine($"      标准输出长度: {result.Output?.Length ?? 0}");
                Console.WriteLine($"      错误输出长度: {result.Error?.Length ?? 0}");

                if (!string.IsNullOrEmpty(result.Output))
                {
                    Console.WriteLine($"      标准输出内容: {result.Output}");
                }

                if (!string.IsNullOrEmpty(result.Error))
                {
                    Console.WriteLine($"      错误输出内容: {result.Error}");
                }

                if (result.Success)
                {
                    Console.WriteLine("      ✓ 打印确认单后续处理成功！");

                    // 更新最后保存的PDF路径
                    string fullPdfPath = Path.Combine(folderPath, fileName);
                    lastSavedPdfPath = fullPdfPath;
                    Console.WriteLine($"      ✓ 已更新最后保存的PDF路径: {fullPdfPath}");

                    if (!string.IsNullOrEmpty(result.Output))
                    {
                        Console.WriteLine($"      输出: {result.Output}");
                    }
                }
                else
                {
                    Console.WriteLine("      ✗ 打印确认单后续处理失败");
                    if (!string.IsNullOrEmpty(result.Error))
                    {
                        Console.WriteLine($"      错误: {result.Error}");
                    }
                    else
                    {
                        Console.WriteLine("      错误: 无错误信息");
                    }
                    if (!string.IsNullOrEmpty(result.Output))
                    {
                        Console.WriteLine($"      输出: {result.Output}");
                    }
                    Console.WriteLine("      请检查：");
                    Console.WriteLine("      1. Python环境是否正确安装");
                    Console.WriteLine("      2. test_mouse_keyboard.py脚本是否存在");
                    Console.WriteLine("      3. config.json中的坐标配置是否正确");
                    Console.WriteLine("      4. 运行 debug_python_execution.bat 进行详细诊断");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      处理打印确认单按钮时出错: {ex.Message}");
                Console.WriteLine($"      详细错误信息: {ex}");
            }
        }

        /// <summary>
        /// 从网页中提取预约号和申请总金额信息
        /// </summary>
        /// <returns>预约号和申请总金额的元组</returns>
        private async Task<(string appointmentNumber, string totalAmount)> ExtractAppointmentInfoFromPage()
        {
            string appointmentNumber = "null";
            string totalAmount = "0.00";

            try
            {
                Console.WriteLine("      开始从网页提取信息...");

                // 优先在iframe 8中查找（根据日志显示，预约号和金额在这里）
                Console.WriteLine("      优先在iframe 8中查找...");
                var frames = page.Frames;
                if (frames.Count >= 8)
                {
                    var frame8 = frames[7]; // 第8个iframe（索引为7）
                    var frameUrl = frame8.Url;
                    var frameName = frame8.Name;
                    Console.WriteLine($"      iframe 8 URL: {frameUrl}");
                    Console.WriteLine($"      iframe 8 Name: {frameName}");

                    // 检查是否是目标iframe（包含printYB/ybprint.jsp）
                    if (frameUrl.Contains("printYB/ybprint.jsp"))
                    {
                        Console.WriteLine("      ✓ 确认是目标iframe，开始提取信息...");
                        var result = await ExtractFromPrintContentInFrame(frame8);
                        appointmentNumber = result.appointmentNumber;
                        totalAmount = result.totalAmount;

                        if (appointmentNumber != "null" && totalAmount != "0.00")
                        {
                            Console.WriteLine($"      ✓ 在iframe 8中成功提取到所有信息");
                            // 清理数据
                            appointmentNumber = CleanAppointmentNumber(appointmentNumber);
                            totalAmount = CleanTotalAmount(totalAmount);
                            Console.WriteLine($"      最终提取结果 - 预约号: {appointmentNumber}, 申请总金额: {totalAmount}");
                            return (appointmentNumber, totalAmount);
                        }
                    }
                    else
                    {
                        Console.WriteLine("      ✗ iframe 8不是目标iframe，尝试其他方法...");
                    }
                }

                // 如果iframe 8没找到，尝试从主页面提取
                if (appointmentNumber == "null" || totalAmount == "0.00")
                {
                    Console.WriteLine("      iframe 8未找到信息，尝试从主页面提取...");
                    if (appointmentNumber == "null")
                    {
                        appointmentNumber = await ExtractAppointmentNumberFromPage();
                    }
                    if (totalAmount == "0.00")
                    {
                        totalAmount = await ExtractTotalAmountFromPage();
                    }
                }

                // 如果主页面没有找到，尝试从其他iframe中提取
                if (appointmentNumber == "null" || totalAmount == "0.00")
                {
                    Console.WriteLine("      主页面未找到信息，尝试从其他iframe中提取...");
                    Console.WriteLine($"      找到 {frames.Count} 个iframe");

                    for (int i = 0; i < frames.Count; i++)
                    {
                        // 跳过iframe 8，因为已经尝试过了
                        if (i == 7) continue;

                        try
                        {
                            Console.WriteLine($"      尝试从iframe {i + 1} 中提取信息...");

                            // 先尝试在该iframe的printContent中查找
                            var frameResult = await ExtractFromPrintContentInFrame(frames[i]);
                            if (appointmentNumber == "null" && frameResult.appointmentNumber != "null")
                            {
                                appointmentNumber = frameResult.appointmentNumber;
                            }
                            if (totalAmount == "0.00" && frameResult.totalAmount != "0.00")
                            {
                                totalAmount = frameResult.totalAmount;
                            }

                            // 如果printContent中没找到，再用原来的方法
                            if (appointmentNumber == "null")
                            {
                                appointmentNumber = await ExtractAppointmentNumberFromFrame(frames[i]);
                            }
                            if (totalAmount == "0.00")
                            {
                                totalAmount = await ExtractTotalAmountFromFrame(frames[i]);
                            }

                            // 如果都找到了，就跳出循环
                            if (appointmentNumber != "null" && totalAmount != "0.00")
                            {
                                Console.WriteLine($"      ✓ 在iframe {i + 1} 中找到所有信息");
                                break;
                            }
                        }
                        catch (Exception frameEx)
                        {
                            Console.WriteLine($"      从iframe {i + 1} 提取信息时出错: {frameEx.Message}");
                        }
                    }
                }

                // 清理数据
                appointmentNumber = CleanAppointmentNumber(appointmentNumber);
                totalAmount = CleanTotalAmount(totalAmount);

                Console.WriteLine($"      最终提取结果 - 预约号: {appointmentNumber}, 申请总金额: {totalAmount}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      提取网页信息时出错: {ex.Message}");
                Console.WriteLine($"      使用默认值 - 预约号: {appointmentNumber}, 申请总金额: {totalAmount}");
            }

            return (appointmentNumber, totalAmount);
        }

        /// <summary>
        /// 从printContent容器中提取信息（主页面）
        /// </summary>
        private async Task<(string appointmentNumber, string totalAmount)> ExtractFromPrintContent()
        {
            string appointmentNumber = "null";
            string totalAmount = "0.00";

            try
            {
                Console.WriteLine("      开始查找printContent容器...");

                // 等待页面稳定
                await Task.Delay(2000);

                // 查找printContent容器
                var printContentCount = await page.Locator("#printContent").CountAsync();
                Console.WriteLine($"      主页面找到 {printContentCount} 个 printContent 容器");

                if (printContentCount > 0)
                {
                    var printContent = page.Locator("#printContent").First;

                    // 等待元素可见
                    try
                    {
                        await printContent.WaitForAsync(new LocatorWaitForOptions { Timeout = 5000 });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      等待printContent可见时超时: {ex.Message}");
                    }

                    // 获取容器的HTML内容
                    var html = await printContent.InnerHTMLAsync();
                    Console.WriteLine($"      获取到printContent HTML内容，长度: {html?.Length ?? 0}");

                    if (!string.IsNullOrEmpty(html))
                    {
                        // 使用我们的解析函数
                        var dictionary = ParsePrintContentHtmlToDictionary(html);
                        Console.WriteLine($"      解析出 {dictionary.Count} 个键值对");

                        // 查找预约号
                        foreach (var key in new[] { "预约号", "预约号：", "预约号:" })
                        {
                            if (dictionary.ContainsKey(key))
                            {
                                appointmentNumber = dictionary[key];
                                Console.WriteLine($"      ✓ 在printContent中找到预约号: {appointmentNumber}");
                                break;
                            }
                        }

                        // 查找申请总金额
                        foreach (var key in new[] { "申请总金额", "申请总金额：", "申请总金额:", "总计" })
                        {
                            if (dictionary.ContainsKey(key))
                            {
                                var value = dictionary[key];
                                // 从包含金额的文本中提取数字
                                var match = System.Text.RegularExpressions.Regex.Match(value, @"([\d,]+\.?\d*)");
                                if (match.Success)
                                {
                                    totalAmount = match.Groups[1].Value;
                                    Console.WriteLine($"      ✓ 在printContent中找到申请总金额: {totalAmount}");
                                    break;
                                }
                            }
                        }

                        // 如果还没找到，直接在HTML中用正则查找
                        if (appointmentNumber == "null")
                        {
                            var appointmentMatch = System.Text.RegularExpressions.Regex.Match(html, @"预约号[：:]\s*</td>\s*<td[^>]*>([^<]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            if (appointmentMatch.Success)
                            {
                                appointmentNumber = appointmentMatch.Groups[1].Value.Trim();
                                Console.WriteLine($"      ✓ 通过正则在printContent中找到预约号: {appointmentNumber}");
                            }
                        }

                        if (totalAmount == "0.00")
                        {
                            var amountMatch = System.Text.RegularExpressions.Regex.Match(html, @"申请总金额[：:]\s*([0-9,]+\.?\d*)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            if (amountMatch.Success)
                            {
                                totalAmount = amountMatch.Groups[1].Value.Trim();
                                Console.WriteLine($"      ✓ 通过正则在printContent中找到申请总金额: {totalAmount}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("      ✗ printContent容器内容为空");
                    }
                }
                else
                {
                    Console.WriteLine("      ✗ 主页面未找到printContent容器");

                    // 尝试查找其他可能的容器
                    var possibleContainers = new[] { ".printdiv", "[id*='print']", "[class*='print']" };
                    foreach (var selector in possibleContainers)
                    {
                        var count = await page.Locator(selector).CountAsync();
                        if (count > 0)
                        {
                            Console.WriteLine($"      找到可能的容器: {selector} (数量: {count})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      从printContent提取信息时出错: {ex.Message}");
                Console.WriteLine($"      详细错误: {ex}");
            }

            return (appointmentNumber, totalAmount);
        }

        /// <summary>
        /// 从printContent容器中提取信息（iframe）
        /// </summary>
        private async Task<(string appointmentNumber, string totalAmount)> ExtractFromPrintContentInFrame(IFrame frame)
        {
            string appointmentNumber = "null";
            string totalAmount = "0.00";

            try
            {
                Console.WriteLine("      在iframe中查找printContent容器...");

                // 获取iframe信息
                var frameUrl = frame.Url;
                var frameName = frame.Name;
                Console.WriteLine($"      iframe URL: {frameUrl}");
                Console.WriteLine($"      iframe Name: {frameName}");

                // 查找printContent容器
                var printContentCount = await frame.Locator("#printContent").CountAsync();
                Console.WriteLine($"      在iframe中找到 {printContentCount} 个 printContent 容器");

                if (printContentCount > 0)
                {
                    var printContent = frame.Locator("#printContent").First;

                    // 等待元素可见
                    try
                    {
                        await printContent.WaitForAsync(new LocatorWaitForOptions { Timeout = 3000 });
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      等待iframe中printContent可见时超时: {ex.Message}");
                    }

                    // 获取容器的HTML内容
                    var html = await printContent.InnerHTMLAsync();
                    Console.WriteLine($"      在iframe中获取到printContent HTML内容，长度: {html?.Length ?? 0}");

                    if (!string.IsNullOrEmpty(html))
                    {
                        // 使用我们的解析函数
                        var dictionary = ParsePrintContentHtmlToDictionary(html);
                        Console.WriteLine($"      在iframe中解析出 {dictionary.Count} 个键值对");

                        // 查找预约号
                        foreach (var key in new[] { "预约号", "预约号：", "预约号:" })
                        {
                            if (dictionary.ContainsKey(key))
                            {
                                appointmentNumber = dictionary[key];
                                Console.WriteLine($"      ✓ 在iframe的printContent中找到预约号: {appointmentNumber}");
                                break;
                            }
                        }

                        // 查找申请总金额
                        foreach (var key in new[] { "申请总金额", "申请总金额：", "申请总金额:", "总计" })
                        {
                            if (dictionary.ContainsKey(key))
                            {
                                var value = dictionary[key];
                                // 从包含金额的文本中提取数字
                                var match = System.Text.RegularExpressions.Regex.Match(value, @"([\d,]+\.?\d*)");
                                if (match.Success)
                                {
                                    totalAmount = match.Groups[1].Value;
                                    Console.WriteLine($"      ✓ 在iframe的printContent中找到申请总金额: {totalAmount}");
                                    break;
                                }
                            }
                        }

                        // 如果还没找到，直接在HTML中用正则查找
                        if (appointmentNumber == "null")
                        {
                            var appointmentMatch = System.Text.RegularExpressions.Regex.Match(html, @"预约号[：:]\s*</td>\s*<td[^>]*>([^<]+)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            if (appointmentMatch.Success)
                            {
                                appointmentNumber = appointmentMatch.Groups[1].Value.Trim();
                                Console.WriteLine($"      ✓ 通过正则在iframe的printContent中找到预约号: {appointmentNumber}");
                            }
                        }

                        if (totalAmount == "0.00")
                        {
                            var amountMatch = System.Text.RegularExpressions.Regex.Match(html, @"申请总金额[：:]\s*([0-9,]+\.?\d*)", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                            if (amountMatch.Success)
                            {
                                totalAmount = amountMatch.Groups[1].Value.Trim();
                                Console.WriteLine($"      ✓ 通过正则在iframe的printContent中找到申请总金额: {totalAmount}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("      ✗ iframe中printContent容器内容为空");
                    }
                }
                else
                {
                    Console.WriteLine("      ✗ iframe中未找到printContent容器");

                    // 尝试查找其他可能的容器
                    var possibleContainers = new[] { ".printdiv", "[id*='print']", "[class*='print']", "table", "tbody" };
                    foreach (var selector in possibleContainers)
                    {
                        var count = await frame.Locator(selector).CountAsync();
                        if (count > 0)
                        {
                            Console.WriteLine($"      在iframe中找到可能的容器: {selector} (数量: {count})");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      从iframe的printContent提取信息时出错: {ex.Message}");
                Console.WriteLine($"      详细错误: {ex}");
            }

            return (appointmentNumber, totalAmount);
        }

        /// <summary>
        /// 从主页面提取预约号
        /// </summary>
        private async Task<string> ExtractAppointmentNumberFromPage()
        {
            try
            {
                // 根据实际HTML结构调整选择器
                var selectors = new[]
                {
                    "td:has-text('预约号：') + td",
                    "xpath=//td[contains(normalize-space(.), '预约号')]/following-sibling::td[1]",
                    "tr.text_fd td:nth-child(2)",
                    "tbody tr.text_fd td:nth-child(2)",
                    "tr[class*='text_fd'] td:nth-child(2)",
                    "td[width='13%']"
                };

                foreach (var selector in selectors)
                {
                    try
                    {
                        await page.Locator(selector).First.WaitForAsync(new LocatorWaitForOptions { Timeout = 5000 });
                        var element = page.Locator(selector).First;
                        var text = await element.TextContentAsync();
                        text = text?.Replace('\u00A0', ' ').Trim();
                        if (!string.IsNullOrEmpty(text))
                        {
                            Console.WriteLine($"      ✓ 在主页面找到预约号: {text.Trim()}");
                            return text.Trim();
                        }
                    }
                    catch
                    {
                        // 继续尝试下一个选择器
                    }
                }

                Console.WriteLine("      ✗ 在主页面未找到预约号");
                return "null";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      从主页面提取预约号时出错: {ex.Message}");
                return "null";
            }
        }

        /// <summary>
        /// 从主页面提取申请总金额
        /// </summary>
        private async Task<string> ExtractTotalAmountFromPage()
        {
            try
            {
                // 根据实际HTML结构调整选择器
                var selectors = new[]
                {
                    "xpath=//td[contains(normalize-space(.), '申请总金额')]",
                    "td.title_fd:has-text('申请总金额')",
                    "td[class*='title_fd']:has-text('申请总金额')",
                    "td:has-text('申请总金额')",
                    "td[width='60%']:has-text('申请总金额')",
                };

                foreach (var selector in selectors)
                {
                    try
                    {
                        await page.Locator(selector).First.WaitForAsync(new LocatorWaitForOptions { Timeout = 5000 });
                        var element = page.Locator(selector).First;
                        var text = await element.TextContentAsync();
                        text = text?.Replace('\u00A0', ' ').Trim();
                        if (!string.IsNullOrEmpty(text) && text.Contains("申请总金额"))
                        {
                            // 提取金额部分 - 匹配"申请总金额: 1500.00"格式
                            var amountMatch = System.Text.RegularExpressions.Regex.Match(text, @"申请总金额\s*[:：]\s*([\d,]+\.?\d*)");
                            if (amountMatch.Success)
                            {
                                var amount = amountMatch.Groups[1].Value;
                                Console.WriteLine($"      ✓ 在主页面找到申请总金额: {amount}");
                                return amount;
                            }
                        }
                    }
                    catch
                    {
                        // 继续尝试下一个选择器
                    }
                }

                Console.WriteLine("      ✗ 在主页面未找到申请总金额");
                return "0.00";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      从主页面提取申请总金额时出错: {ex.Message}");
                return "0.00";
            }
        }

        /// <summary>
        /// 从iframe中提取预约号
        /// </summary>
        private async Task<string> ExtractAppointmentNumberFromFrame(IFrame frame)
        {
            try
            {
                var selectors = new[]
                {
                    "td:has-text('预约号：') + td",
                    "xpath=//td[contains(normalize-space(.), '预约号')]/following-sibling::td[1]",
                    "tbody tr.text_fd td:nth-child(2)",
                    "tr.text_fd td:nth-child(2)",
                    "[class*='text_fd'] td:nth-child(2)"
                };

                foreach (var selector in selectors)
                {
                    try
                    {
                        // 先等待元素出现
                        await frame.Locator(selector).First.WaitForAsync(new LocatorWaitForOptions { Timeout = 1000 });

                        // 然后获取元素
                        var element = frame.Locator(selector).First;
                        var text = await element.TextContentAsync();
                        text = text?.Replace('\u00A0', ' ').Trim();
                        if (!string.IsNullOrEmpty(text))
                        {
                            Console.WriteLine($"      ✓ 在iframe中找到预约号: {text.Trim()}");
                            return text.Trim();
                        }
                    }
                    catch
                    {
                        // 继续尝试下一个选择器
                    }
                }

                return "null";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      从iframe提取预约号时出错: {ex.Message}");
                return "null";
            }
        }
        /// <summary>
        /// 从iframe中提取申请总金额
        /// </summary>
        private async Task<string> ExtractTotalAmountFromFrame(IFrame frame)
        {
            try
            {
                var selectors = new[]
                {
                    "xpath=//td[contains(normalize-space(.), '申请总金额')]",
                    "td:has-text('申请总金额')",
                    "[class*='title_fd']:has-text('申请总金额')"
                };

                foreach (var selector in selectors)
                {
                    try
                    {
                        await frame.Locator(selector).First.WaitForAsync(new LocatorWaitForOptions { Timeout = 3000 });
                        var element = frame.Locator(selector).First;
                        var text = await element.TextContentAsync();
                        text = text?.Replace('\u00A0', ' ').Trim();
                        if (!string.IsNullOrEmpty(text) && text.Contains("申请总金额"))
                        {
                            var amountMatch = System.Text.RegularExpressions.Regex.Match(text, @"申请总金额\s*[:：]\s*([\d,]+\.?\d*)");
                            if (amountMatch.Success)
                            {
                                var amount = amountMatch.Groups[1].Value;
                                Console.WriteLine($"      ✓ 在iframe中找到申请总金额: {amount}");
                                return amount;
                            }
                        }
                    }
                    catch
                    {
                        // 继续尝试下一个选择器
                    }
                }

                return "0.00";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      从iframe提取申请总金额时出错: {ex.Message}");
                return "0.00";
            }
        }

        /// <summary>
        /// 清理预约号数据
        /// </summary>
        private string CleanAppointmentNumber(string appointmentNumber)
        {
            if (string.IsNullOrEmpty(appointmentNumber) || appointmentNumber == "null")
                return "null";

            // 移除空格、换行符等
            return appointmentNumber.Trim().Replace("\n", "").Replace("\r", "");
        }

        /// <summary>
        /// 清理申请总金额数据
        /// </summary>
        private string CleanTotalAmount(string totalAmount)
        {
            if (string.IsNullOrEmpty(totalAmount) || totalAmount == "0.00")
                return "0.00";

            // 移除空格、换行符、逗号等
            var cleaned = totalAmount.Trim().Replace("\n", "").Replace("\r", "").Replace(",", "");

            // 确保是数字格式
            if (decimal.TryParse(cleaned, out decimal amount))
            {
                return amount.ToString("F2");
            }

            return "0.00";
        }

        // 将 printContent 的 HTML 字符串解析为键值对字典（通用）
        private Dictionary<string, string> ParsePrintContentHtmlToDictionary(string html)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrWhiteSpace(html)) return result;

            try
            {
                // 规范化
                string normalized = html.Replace("\r", "").Replace("\n", "");

                // 提取所有行
                var rowMatches = System.Text.RegularExpressions.Regex.Matches(
                    normalized,
                    "<tr[\\s\\S]*?>([\\s\\S]*?)</tr>",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // 逐行解析
                var rows = new List<List<string>>();
                foreach (System.Text.RegularExpressions.Match row in rowMatches)
                {
                    var cells = new List<string>();
                    var cellMatches = System.Text.RegularExpressions.Regex.Matches(
                        row.Groups[1].Value,
                        "<td[\\s\\S]*?>([\\s\\S]*?)</td>",
                        System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                    foreach (System.Text.RegularExpressions.Match cell in cellMatches)
                    {
                        string cellHtml = cell.Groups[1].Value;
                        string text = CleanCellText(cellHtml);
                        if (text.Length > 0)
                        {
                            cells.Add(text);
                        }
                        else
                        {
                            cells.Add("");
                        }
                    }

                    if (cells.Count > 0)
                    {
                        rows.Add(cells);
                    }
                }

                // 解析规则：
                // 1) 单元格内的“键：值”对（可能包含多个）
                // 2) 邻接单元格“键” -> “值”
                for (int r = 0; r < rows.Count; r++)
                {
                    var cells = rows[r];

                    // 规则1：同一单元格内的多对键值
                    foreach (var cell in cells)
                    {
                        foreach (var kv in ExtractInlinePairs(cell))
                        {
                            result[kv.Key] = kv.Value;
                        }
                    }

                    // 规则2：相邻单元格配对
                    for (int i = 0; i < cells.Count - 1; i++)
                    {
                        string left = cells[i];
                        string right = FindNextNonEmpty(cells, i + 1);
                        if (string.IsNullOrEmpty(left) || string.IsNullOrEmpty(right)) continue;

                        // 左侧以冒号结尾或包含冒号
                        if (left.EndsWith("：") || left.EndsWith(":") || left.Contains("：") || left.Contains(":"))
                        {
                            string key = TrimKey(left);
                            string value = right.Trim();
                            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                            {
                                result[key] = value;
                            }
                            continue;
                        }

                        // 启发式：左侧是较短的标签，右侧是值（且右侧不像标签）
                        if (IsLikelyLabel(left) && !LooksLikeLabel(right))
                        {
                            string key = TrimKey(left);
                            string value = right.Trim();
                            if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                            {
                                result[key] = value;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      解析HTML为字典时出错: {ex.Message}");
            }

            return result;
        }

        private static string CleanCellText(string cellHtml)
        {
            // 移除所有标签
            string text = System.Text.RegularExpressions.Regex.Replace(cellHtml, "<[^>]+>", " ", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
            // HTML 实体解码
            text = System.Net.WebUtility.HtmlDecode(text);
            // 替换不间断空格
            text = text.Replace('\u00A0', ' ');
            // 压缩空白
            text = System.Text.RegularExpressions.Regex.Replace(text, "\\s+", " ").Trim();
            return text;
        }

        private static IEnumerable<KeyValuePair<string, string>> ExtractInlinePairs(string text)
        {
            var list = new List<KeyValuePair<string, string>>();
            if (string.IsNullOrEmpty(text)) return list;

            // 匹配形如 “键：值” 或 “键: 值”，同一文本中可能有多对
            var matches = System.Text.RegularExpressions.Regex.Matches(
                text,
                "([\\u4e00-\\u9fa5A-Za-z0-9（）()\\s]+?)\\s*[：:]\\s*([^：:]+?)(?=(?:[\\s]+[\\u4e00-\\u9fa5A-Za-z0-9（）()]+\\s*[：:])|$)");

            foreach (System.Text.RegularExpressions.Match m in matches)
            {
                string key = TrimKey(m.Groups[1].Value);
                string value = m.Groups[2].Value.Trim();
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(value))
                {
                    list.Add(new KeyValuePair<string, string>(key, value));
                }
            }
            return list;
        }

        private static string FindNextNonEmpty(List<string> cells, int start)
        {
            for (int i = start; i < cells.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(cells[i])) return cells[i];
            }
            return string.Empty;
        }

        private static string TrimKey(string key)
        {
            if (string.IsNullOrEmpty(key)) return key;
            key = key.Trim();
            key = key.TrimEnd('：', ':');
            key = key.Replace(" ", ""); // 去掉内部空格以稳定匹配
            return key;
        }

        private static bool IsLikelyLabel(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            if (text.Contains("：") || text.Contains(":")) return true;
            // 短文本、主要是中文或少量字母，认为是标签
            var t = text.Trim();
            if (t.Length <= 8)
            {
                return System.Text.RegularExpressions.Regex.IsMatch(t, "^[\\u4e00-\\u9fa5A-Za-z（）()]+$");
            }
            return false;
        }

        private static bool LooksLikeLabel(string text)
        {
            if (string.IsNullOrEmpty(text)) return false;
            if (text.Contains("：") || text.Contains(":")) return true;
            // 如果是明显的金额/日期/数字，则不是标签
            if (System.Text.RegularExpressions.Regex.IsMatch(text, "^[0-9，,\\.]+$")) return false;
            if (System.Text.RegularExpressions.Regex.IsMatch(text, "^\\d{4}-\\d{2}-\\d{2}")) return false;
            return IsLikelyLabel(text);
        }
    }
}
using System;
using System.Threading.Tasks;
using Microsoft.Playwright;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text.RegularExpressions;

namespace AutoFinan
{
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
        private const string ExcelFilePath = "报销信息.xlsx";
        private const string MappingFilePath = "标题-ID.xlsx";
        private const string SheetName = "ChaiLv_sheet";
        private const string MappingSheetName = "Sheet1"; // 标题-ID映射表的工作表名
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

        public async Task RunAsync()
        {
            Console.WriteLine("开始读取Excel文件...");
            
            // 获取当前执行目录
            string currentDirectory = Directory.GetCurrentDirectory();
            Console.WriteLine($"当前工作目录: {currentDirectory}");
            
            // 尝试多个可能的文件路径
            string[] possiblePaths = {
                ExcelFilePath,
                Path.Combine(currentDirectory, ExcelFilePath),
                Path.Combine(currentDirectory, "..", ExcelFilePath),
                Path.Combine(currentDirectory, "..", "..", ExcelFilePath),
                Path.Combine(currentDirectory, "..", "..", "..", ExcelFilePath),
                Path.Combine(currentDirectory, "..", "..", "..", "..", ExcelFilePath)
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
                Console.WriteLine($"错误：找不到文件 {ExcelFilePath}");
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
                var worksheet = package.Workbook.Worksheets[SheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"错误：找不到工作表 {SheetName}");
                    return;
                }

                Console.WriteLine($"成功加载工作表: {SheetName}");
                
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
                        if (!subsequenceStartColumns.Contains(col) && !string.IsNullOrEmpty(cellValue))
                        {
                            await ProcessCell(columnName, row, headerName, cellValue);
                        }
                    }
                    
                    Console.WriteLine($"第 {row} 行数据处理完成");
                }
                
                Console.WriteLine("\n=== 所有数据处理完成 ===");
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
            string mappingFilePath = Path.Combine(excelDirectory, MappingFilePath);
            
            if (!File.Exists(mappingFilePath))
            {
                Console.WriteLine($"错误：找不到标题-ID映射文件 {mappingFilePath}");
                return;
            }

            titleIdMapping = new Dictionary<string, string>();

            using (var package = new ExcelPackage(new FileInfo(mappingFilePath)))
            {
                var worksheet = package.Workbook.Worksheets[MappingSheetName];
                if (worksheet == null)
                {
                    Console.WriteLine($"错误：找不到工作表 {MappingSheetName}");
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
                    
                    if (!string.IsNullOrEmpty(cellValue))
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
                    
                    if (!string.IsNullOrEmpty(cellValue))
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

        private async Task ProcessCell(string columnName, int row, string headerName, string cellValue)
        {
            // 这里实现具体的单元格处理逻辑
            // 根据不同的列标题和值，执行不同的操作
            
            Console.WriteLine($"    执行操作：{columnName}{row} - {headerName} = '{cellValue}'");
            
            // 1. 等待操作（列标题为"等待"）
            if (headerName == "等待")
            {
                Console.WriteLine($"      检测到等待操作: {cellValue}");
                await WaitOperation(cellValue);
            }
            // 2. 按钮点击操作（以$开头）
            else if (cellValue == "$点击")
            {
                Console.WriteLine($"      检测到按钮点击操作: {headerName}");
                await ClickButton(headerName);
            }
            // 3. Radio按钮点击操作（以$$开头）
            else if (cellValue.StartsWith("$$"))
            {
                string radioValue = cellValue.Substring(2); // 去掉$$前缀
                Console.WriteLine($"      检测到Radio按钮操作: {radioValue}");
                await ClickRadioButton(radioValue);
            }
            // 4. 银行卡选择操作（以*开头）
            else if (cellValue.StartsWith("*"))
            {
                Console.WriteLine($"      检测到银行卡选择操作: {cellValue}");
                await SelectCardByNumber(cellValue);
            }
            // 5. 科目输入框操作（以#开头）
            else if (cellValue.StartsWith("#"))
            {
                Console.WriteLine($"      检测到科目输入框操作: {cellValue}");
                await FillSubjectInput(headerName, cellValue);
            }
            // 6. 下拉框选择操作
            else if (IsDropdownField(headerName))
            {
                Console.WriteLine($"      检测到下拉框选择操作: {headerName} = {cellValue}");
                await SelectDropdown(headerName, cellValue);
            }
                         // 7. 日期选择操作（日期字段或格式：yyyy-mm-dd）
             else if (IsDateField(headerName) || IsDate(cellValue))
             {
                 Console.WriteLine($"      检测到日期选择操作: {cellValue}");
                 await SelectDate(headerName, cellValue);
             }
            // 8. 金额输入框操作（需要与科目配对）
            else if (headerName == "金额" && !string.IsNullOrEmpty(currentSubjectId))
            {
                Console.WriteLine($"      检测到金额输入框操作: {cellValue}");
                await FillAmountInput(currentSubjectId, cellValue);
                currentSubjectId = null; // 清空当前科目ID
            }
            // 9. 一般输入框操作
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
                string elementId = GetElementId(headerName);
                if (string.IsNullOrEmpty(elementId))
                {
                    Console.WriteLine($"      警告：未找到标题 '{headerName}' 对应的按钮ID");
                    return;
                }

                Console.WriteLine($"      点击按钮: {headerName} -> {elementId}");
                
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
                
                // 实现实际的按钮点击逻辑
                bool clicked = false;
                
                // 等待页面完全加载
                await Task.Delay(500);
                
                // 方法1: 优先在iframe中通过btnname属性查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var buttonElement = frame.Locator($"button[btnname='{elementId}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      在iframe中通过btnname成功点击按钮: {elementId}");
                            clicked = true;
                            break;
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
                            var buttonElement = frame.Locator($"#{elementId}").First;
                            if (await buttonElement.CountAsync() > 0)
                            {
                                await buttonElement.ClickAsync();
                                Console.WriteLine($"      在iframe中通过ID成功点击按钮: {elementId}");
                                clicked = true;
                                break;
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
                
                // 等待按钮点击后的页面加载
                await Task.Delay(1000);
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
                if (!subsequenceStartColumns.Contains(col) && !string.IsNullOrEmpty(cellValue))
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
                 ["安排状态"] = new Dictionary<string, string>
                 {
                     ["未安排"] = "未安排",
                     ["安排"] = "安排"
                 },
                 ["交通费"] = new Dictionary<string, string>
                 {
                     ["未安排"] = "未安排",
                     ["安排"] = "安排"
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
            string[] bankCardFields = { "转卡信息工号", "卡号尾号", "银行卡", "银行账号" };
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
                
                await page.GotoAsync(targetUrl);
                
                Console.WriteLine($"成功导航到页面: {targetUrl}");
                
                // 等待页面加载完成
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                
                Console.WriteLine("页面加载完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"页面导航失败: {ex.Message}");
                throw;
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
    }
}

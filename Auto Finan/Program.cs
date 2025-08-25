using System;
using System.Threading.Tasks;
using Microsoft.Playwright;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;
using System.IO;

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
        private const string SheetName = "BaoXiao_sheet";
        private const string SubsequenceStartColumn = "子序列开始";
        private const string SubsequenceEndColumn = "子序列结束";
        private const string SubsequenceMarker = "是";

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
            
            string actualFilePath = null;
            foreach (string path in possiblePaths)
            {
                if (File.Exists(path))
                {
                    actualFilePath = path;
                    Console.WriteLine($"找到Excel文件: {path}");
                    break;
                }
            }
            
            if (actualFilePath == null)
            {
                Console.WriteLine($"错误：找不到文件 {ExcelFilePath}");
                Console.WriteLine("尝试过的路径:");
                foreach (string path in possiblePaths)
                {
                    Console.WriteLine($"  {path}");
                }
                return;
            }

            using (var package = new ExcelPackage(new FileInfo(actualFilePath)))
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
                int subsequenceStartColIndex = -1;
                int subsequenceEndColIndex = -1;
                
                for (int i = 0; i < headers.Count; i++)
                {
                    if (headers[i] == SubsequenceStartColumn)
                        subsequenceStartColIndex = i + 1;
                    else if (headers[i] == SubsequenceEndColumn)
                        subsequenceEndColIndex = i + 1;
                }

                Console.WriteLine($"子序列开始列: {GetColumnName(subsequenceStartColIndex)} (索引: {subsequenceStartColIndex})");
                Console.WriteLine($"子序列结束列: {GetColumnName(subsequenceEndColIndex)} (索引: {subsequenceEndColIndex})");

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
                if (col == subsequenceStartColIndex && cellValue == SubsequenceMarker)
                {
                    Console.WriteLine($"检测到子序列开始标记，进入子序列处理逻辑");
                    int nextRow = await ProcessSubsequence(worksheet, headers, row, rowCount, subsequenceStartColIndex, subsequenceEndColIndex);
                    
                    // 如果子序列处理返回了下一行的行号，继续处理那一行
                    if (nextRow > 0)
                    {
                        Console.WriteLine($"子序列处理完成，继续处理第 {nextRow} 行");
                        await ProcessRowFromColumn(worksheet, headers, nextRow, subsequenceEndColIndex + 1, colCount);
                    }
                    break; // 跳出当前行的列循环，继续处理下一行
                }
                        
                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            await ProcessCell(columnName, row, headerName, cellValue);
                        }
                    }
                    
                    Console.WriteLine($"第 {row} 行数据处理完成");
                }
                
                Console.WriteLine("\n=== 所有数据处理完成 ===");
            }
        }

        private async Task<int> ProcessSubsequence(ExcelWorksheet worksheet, List<string> headers, int startRow, int totalRows, int subsequenceStartColIndex, int subsequenceEndColIndex)
        {
            Console.WriteLine($"\n=== 进入第三层循环：子序列处理逻辑 ===");
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
                
                // 检查下一行的子序列结束列是否标记为"是"
                if (subsequenceEndColIndex > 0 && row + 1 <= totalRows)
                {
                    var nextRowSubsequenceEndValue = worksheet.Cells[row + 1, subsequenceEndColIndex].Value?.ToString() ?? "";
                    if (nextRowSubsequenceEndValue == SubsequenceMarker)
                    {
                        Console.WriteLine($"检测到下一行({row + 1})的子序列结束标记，结束子序列处理");
                        Console.WriteLine($"程序将从第 {row + 1} 行继续正常处理逻辑");
                        return row + 1; // 返回下一行的行号
                    }
                    else
                    {
                        Console.WriteLine($"下一行({row + 1})的子序列结束列未标记为'是'，继续处理下一行");
                    }
                }
            }
            
            Console.WriteLine("=== 子序列处理逻辑结束 ===");
            return 0; // 如果没有找到子序列结束标记，返回0
        }

        private async Task ProcessCell(string columnName, int row, string headerName, string cellValue)
        {
            // 这里实现具体的单元格处理逻辑
            // 根据不同的列标题和值，执行不同的操作
            
            Console.WriteLine($"    执行操作：{columnName}{row} - {headerName} = '{cellValue}'");
            
            // 示例操作类型判断
            if (cellValue.StartsWith("$"))
            {
                Console.WriteLine($"      检测到按钮操作: {cellValue}");
                // 这里可以添加按钮点击逻辑
            }
            else if (cellValue.StartsWith("$$"))
            {
                Console.WriteLine($"      检测到单选按钮操作: {cellValue}");
                // 这里可以添加单选按钮选择逻辑
            }
            else if (cellValue.StartsWith("@"))
            {
                Console.WriteLine($"      检测到导航操作: {cellValue}");
                // 这里可以添加导航操作逻辑
            }
            else if (cellValue.StartsWith("*"))
            {
                Console.WriteLine($"      检测到银行卡选择操作: {cellValue}");
                // 这里可以添加银行卡选择逻辑
            }
            else if (cellValue.StartsWith("#"))
            {
                Console.WriteLine($"      检测到科目操作: {cellValue}");
                // 这里可以添加科目处理逻辑
            }
            else if (IsNumeric(cellValue))
            {
                Console.WriteLine($"      检测到数值输入: {cellValue}");
                // 这里可以添加数值输入逻辑
            }
            else
            {
                Console.WriteLine($"      检测到文本输入: {cellValue}");
                // 这里可以添加文本输入逻辑
            }
            
            // 模拟异步操作
            await Task.Delay(100);
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

        private bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }

        private async Task ProcessRowFromColumn(ExcelWorksheet worksheet, List<string> headers, int row, int startCol, int totalCols)
        {
            Console.WriteLine($"从第 {row} 行的第 {GetColumnName(startCol)} 列开始处理");
            
            for (int col = startCol; col <= totalCols; col++)
            {
                var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                var columnName = GetColumnName(col);
                var headerName = headers[col - 1];
                
                Console.WriteLine($"  读取单元格 {columnName}{row}，列标题: {headerName}，值: '{cellValue}'");
                
                if (!string.IsNullOrEmpty(cellValue))
                {
                    await ProcessCell(columnName, row, headerName, cellValue);
                }
            }
            
            Console.WriteLine($"第 {row} 行从第 {GetColumnName(startCol)} 列开始的处理完成");
        }
    }
}

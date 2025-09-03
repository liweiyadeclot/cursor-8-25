using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace AutoFinan
{
    /// <summary>
    /// Python脚本调用器
    /// </summary>
    public class PythonInterop
    {
        private readonly string _pythonPath;
        private readonly string _scriptPath;
        
        public PythonInterop(string pythonPath = "python", string scriptPath = "mouse_keyboard_automation.py")
        {
            _pythonPath = pythonPath;
            _scriptPath = scriptPath;
        }
        
        /// <summary>
        /// 执行Python脚本并返回结果
        /// </summary>
        /// <param name="arguments">脚本参数</param>
        /// <returns>执行结果</returns>
        public async Task<string> ExecutePythonScriptAsync(string arguments = "")
        {
            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = _pythonPath,
                    Arguments = $"{_scriptPath} {arguments}",
                    UseShellExecute = false,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    CreateNoWindow = true,
                    StandardOutputEncoding = Encoding.UTF8,
                    StandardErrorEncoding = Encoding.UTF8
                };
                
                using var process = new Process { StartInfo = startInfo };
                process.Start();
                
                var output = await process.StandardOutput.ReadToEndAsync();
                var error = await process.StandardError.ReadToEndAsync();
                
                await process.WaitForExitAsync();
                
                if (process.ExitCode != 0)
                {
                    throw new Exception($"Python脚本执行失败: {error}");
                }
                
                return output;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行Python脚本时出错: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 执行鼠标键盘自动化操作
        /// </summary>
        /// <param name="operation">操作类型</param>
        /// <param name="parameters">参数</param>
        /// <returns>执行结果</returns>
        public async Task<bool> ExecuteMouseKeyboardOperationAsync(string operation, string parameters = "")
        {
            try
            {
                var arguments = $"--operation {operation} {parameters}";
                var result = await ExecutePythonScriptAsync(arguments);
                
                Console.WriteLine($"Python执行结果: {result}");
                return result.Contains("成功") || result.Contains("✓");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"执行鼠标键盘操作失败: {ex.Message}");
                return false;
            }
        }
        
        /// <summary>
        /// 执行文件保存流程
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="fileName">文件名</param>
        /// <param name="coordinates">坐标配置</param>
        /// <returns>是否成功</returns>
        public async Task<bool> ExecuteFileSaveProcessAsync(string filePath, string fileName, string coordinates)
        {
            var parameters = $"--filepath \"{filePath}\" --filename \"{fileName}\" --coordinates \"{coordinates}\"";
            return await ExecuteMouseKeyboardOperationAsync("file_save", parameters);
        }
        
        /// <summary>
        /// 执行打印对话框处理
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <param name="fileName">文件名</param>
        /// <returns>是否成功</returns>
        public async Task<bool> ExecutePrintDialogProcessAsync(string filePath, string fileName)
        {
            var parameters = $"--filepath \"{filePath}\" --filename \"{fileName}\"";
            return await ExecuteMouseKeyboardOperationAsync("print_dialog", parameters);
        }
        
        /// <summary>
        /// 获取鼠标位置
        /// </summary>
        /// <returns>鼠标位置字符串</returns>
        public async Task<string> GetMousePositionAsync()
        {
            try
            {
                var result = await ExecutePythonScriptAsync("--operation get_position");
                return result.Trim();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取鼠标位置失败: {ex.Message}");
                return "0, 0";
            }
        }
        
        /// <summary>
        /// 检查Python环境
        /// </summary>
        /// <returns>是否可用</returns>
        public async Task<bool> CheckPythonEnvironmentAsync()
        {
            try
            {
                var result = await ExecutePythonScriptAsync("--check");
                return result.Contains("Python环境正常");
            }
            catch
            {
                return false;
            }
        }
    }
}

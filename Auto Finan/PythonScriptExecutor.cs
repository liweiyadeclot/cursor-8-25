using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace AutoFinan
{
    /// <summary>
    /// Python脚本执行器
    /// </summary>
    public class PythonScriptExecutor
    {
        private string _pythonExecutable;
        
        /// <summary>
        /// Python脚本执行器
        /// </summary>
        /// <param name="pythonPath">Python解释器路径（如果为null则自动查找）</param>
        public PythonScriptExecutor(string pythonPath = null)
        {
            _pythonExecutable = pythonPath ?? FindPythonExecutable();
            if (string.IsNullOrEmpty(_pythonExecutable))
            {
                throw new InvalidOperationException("未找到Python解释器，请安装Python或指定Python路径");
            }
        }
        
        /// <summary>
        /// 自动查找Python解释器
        /// </summary>
        private string FindPythonExecutable()
        {
            // 常见的Python安装路径
            var possiblePaths = new[]
            {
                "python",        // 如果Python在PATH环境变量中
                "python3",
                @"python\python.exe",  // 相对路径的嵌入式Python
                @"C:\Python39\python.exe",
                @"C:\Python310\python.exe",
                @"C:\Python311\python.exe",
                @"C:\Python312\python.exe",
                @"C:\Program Files\Python39\python.exe",
                @"C:\Program Files\Python310\python.exe",
                @"C:\Program Files\Python311\python.exe",
                @"C:\Program Files\Python312\python.exe",
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), 
                            @"Programs\Python\Python39\python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), 
                            @"Programs\Python\Python310\python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), 
                            @"Programs\Python\Python311\python.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), 
                            @"Programs\Python\Python312\python.exe")
            };
            
            foreach (var path in possiblePaths)
            {
                if (CheckPythonPath(path))
                {
                    Console.WriteLine($"找到Python解释器: {path}");
                    return path;
                }
            }
            
            return null;
        }
        
        /// <summary>
        /// 检查Python路径是否有效
        /// </summary>
        private bool CheckPythonPath(string pythonPath)
        {
            try
            {
                // 如果是相对路径或文件名，尝试直接运行
                if (!pythonPath.Contains(@"\") && !pythonPath.Contains("/"))
                {
                    ProcessStartInfo start = new ProcessStartInfo(pythonPath, "--version")
                    {
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    };
                    
                    using (Process process = Process.Start(start))
                    {
                        process.WaitForExit(2000);
                        return process.ExitCode == 0;
                    }
                }
                
                // 如果是绝对路径，检查文件是否存在
                return File.Exists(pythonPath);
            }
            catch
            {
                return false;
            }
        }
        
        /// <summary>
        /// 执行Python脚本
        /// </summary>
        /// <param name="scriptPath">Python脚本路径</param>
        /// <param name="arguments">脚本参数</param>
        /// <param name="workingDirectory">工作目录</param>
        /// <param name="timeoutMilliseconds">超时时间（毫秒）</param>
        public async Task<PythonExecutionResult> ExecuteScriptAsync(
            string scriptPath, 
            string arguments = "", 
            string workingDirectory = null,
            int timeoutMilliseconds = 30000)
        {
            if (!File.Exists(scriptPath))
            {
                return new PythonExecutionResult
                {
                    Success = false,
                    Output = $"Python脚本文件不存在: {scriptPath}",
                    ExitCode = -1
                };
            }
            
            workingDirectory ??= Path.GetDirectoryName(scriptPath);
            
            var startInfo = new ProcessStartInfo
            {
                FileName = _pythonExecutable,
                Arguments = $"{scriptPath} {arguments}",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                WorkingDirectory = workingDirectory
            };
            
            try
            {
                using (var process = new Process())
                {
                    process.StartInfo = startInfo;
                    
                    var outputBuilder = new StringBuilder();
                    var errorBuilder = new StringBuilder();
                    
                    process.OutputDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            outputBuilder.AppendLine(e.Data);
                            Console.WriteLine($"[Python输出] {e.Data}");
                        }
                    };
                    
                    process.ErrorDataReceived += (sender, e) =>
                    {
                        if (!string.IsNullOrEmpty(e.Data))
                        {
                            errorBuilder.AppendLine(e.Data);
                            Console.WriteLine($"[Python错误] {e.Data}");
                        }
                    };
                    
                    process.Start();
                    
                    // 开始异步读取输出
                    process.BeginOutputReadLine();
                    process.BeginErrorReadLine();
                    
                    // 等待进程完成或超时
                    var completed = await Task.Run(() => process.WaitForExit(timeoutMilliseconds));
                    
                    if (!completed)
                    {
                        process.Kill();
                        return new PythonExecutionResult
                        {
                            Success = false,
                            Output = "执行超时",
                            ExitCode = -1
                        };
                    }
                    
                    // 确保所有输出都被读取
                    await Task.Delay(100);
                    
                    string output = outputBuilder.ToString();
                    string error = errorBuilder.ToString();
                    
                    return new PythonExecutionResult
                    {
                        Success = process.ExitCode == 0,
                        Output = output,
                        Error = error,
                        ExitCode = process.ExitCode
                    };
                }
            }
            catch (Exception ex)
            {
                return new PythonExecutionResult
                {
                    Success = false,
                    Output = $"执行失败: {ex.Message}",
                    ExitCode = -1
                };
            }
        }
        
        /// <summary>
        /// 同步执行Python脚本
        /// </summary>
        public PythonExecutionResult ExecuteScript(
            string scriptPath, 
            string arguments = "", 
            string workingDirectory = null,
            int timeoutMilliseconds = 30000)
        {
            return ExecuteScriptAsync(scriptPath, arguments, workingDirectory, timeoutMilliseconds).Result;
        }
        
        /// <summary>
        /// 执行自动化点击流程
        /// </summary>
        public async Task<bool> ExecuteAutoClickerAsync(
            string folderPath, 
            string fileName, 
            string configPath = "config.yaml",
            int timeoutMilliseconds = 60000)
        {
            string scriptPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "auto_clicker.py");
            
            string arguments = $"--folder \"{folderPath}\" --file \"{fileName}\" --config \"{configPath}\"";
            
            var result = await ExecuteScriptAsync(scriptPath, arguments, timeoutMilliseconds: timeoutMilliseconds);
            
            if (!result.Success)
            {
                Console.WriteLine($"执行失败: {result.Error}");
                return false;
            }
            
            Console.WriteLine("执行成功！");
            Console.WriteLine(result.Output);
            return true;
        }
        
        /// <summary>
        /// 执行鼠标键盘自动化脚本
        /// </summary>
        public async Task<bool> ExecuteMouseKeyboardAutomationAsync(
            string scriptPath = null,
            string arguments = "",
            int timeoutMilliseconds = 120000)
        {
            scriptPath ??= Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "mouse_keyboard_automation.py");
            
            var result = await ExecuteScriptAsync(scriptPath, arguments, timeoutMilliseconds: timeoutMilliseconds);
            
            if (!result.Success)
            {
                Console.WriteLine($"鼠标键盘自动化执行失败: {result.Error}");
                return false;
            }
            
            Console.WriteLine("鼠标键盘自动化执行成功！");
            Console.WriteLine(result.Output);
            return true;
        }
        
        /// <summary>
        /// 执行登录自动化脚本
        /// </summary>
        public async Task<bool> ExecuteLoginAutomationAsync(
            string scriptPath = null,
            string arguments = "",
            int timeoutMilliseconds = 60000)
        {
            scriptPath ??= Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "login_automation.py");
            
            var result = await ExecuteScriptAsync(scriptPath, arguments, timeoutMilliseconds: timeoutMilliseconds);
            
            if (!result.Success)
            {
                Console.WriteLine($"登录自动化执行失败: {result.Error}");
                return false;
            }
            
            Console.WriteLine("登录自动化执行成功！");
            Console.WriteLine(result.Output);
            return true;
        }
    }

    /// <summary>
    /// Python执行结果
    /// </summary>
    public class PythonExecutionResult
    {
        public bool Success { get; set; }
        public string Output { get; set; }
        public string Error { get; set; }
        public int ExitCode { get; set; }
    }
}

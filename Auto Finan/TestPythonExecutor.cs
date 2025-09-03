using System;
using System.Threading.Tasks;

namespace AutoFinan
{
    /// <summary>
    /// PythonScriptExecutor测试类
    /// </summary>
    public class TestPythonExecutor
    {
        /// <summary>
        /// 测试Python脚本执行器
        /// </summary>
        public static async Task TestPythonExecution()
        {
            try
            {
                Console.WriteLine("开始测试Python脚本执行器...");
                
                // 创建Python脚本执行器实例
                var executor = new PythonScriptExecutor();
                
                // 测试1: 执行一个简单的Python脚本
                Console.WriteLine("\n=== 测试1: 执行简单Python脚本 ===");
                var result1 = await executor.ExecuteScriptAsync("test_mouse_keyboard.py");
                Console.WriteLine($"执行结果: {(result1.Success ? "成功" : "失败")}");
                Console.WriteLine($"输出: {result1.Output}");
                if (!string.IsNullOrEmpty(result1.Error))
                {
                    Console.WriteLine($"错误: {result1.Error}");
                }
                
                // 测试2: 执行鼠标键盘自动化脚本
                Console.WriteLine("\n=== 测试2: 执行鼠标键盘自动化脚本 ===");
                var result2 = await executor.ExecuteMouseKeyboardAutomationAsync();
                Console.WriteLine($"鼠标键盘自动化执行结果: {(result2 ? "成功" : "失败")}");
                
                // 测试3: 执行登录自动化脚本
                Console.WriteLine("\n=== 测试3: 执行登录自动化脚本 ===");
                var result3 = await executor.ExecuteLoginAutomationAsync();
                Console.WriteLine($"登录自动化执行结果: {(result3 ? "成功" : "失败")}");
                
                Console.WriteLine("\n测试完成！");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"测试过程中发生错误: {ex.Message}");
                Console.WriteLine($"详细错误信息: {ex}");
            }
        }
        
        /// <summary>
        /// 测试Python版本检查
        /// </summary>
        public static void TestPythonVersion()
        {
            try
            {
                Console.WriteLine("检查Python版本...");
                var executor = new PythonScriptExecutor();
                
                // 执行Python版本检查
                var result = executor.ExecuteScript("", "--version");
                Console.WriteLine($"Python版本检查结果: {result.Output}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Python版本检查失败: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 主测试方法（已移除，避免与Program.cs的Main方法冲突）
        /// </summary>
        public static async Task RunTests()
        {
            Console.WriteLine("Python脚本执行器测试程序");
            Console.WriteLine("==========================");
            
            // 测试Python版本
            TestPythonVersion();
            
            // 测试Python脚本执行
            await TestPythonExecution();
            
            Console.WriteLine("\n测试完成！");
        }
    }
}

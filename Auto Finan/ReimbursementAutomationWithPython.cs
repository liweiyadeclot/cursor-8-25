using System;
using System.Threading.Tasks;

namespace AutoFinan
{
    /// <summary>
    /// 集成了Python脚本调用的报销自动化类
    /// </summary>
    public class ReimbursementAutomationWithPython
    {
        private readonly PythonInterop _pythonInterop;
        private readonly ReimbursementAutomation _reimbursementAutomation;
        
        public ReimbursementAutomationWithPython()
        {
            _pythonInterop = new PythonInterop();
            _reimbursementAutomation = new ReimbursementAutomation();
        }
        
        /// <summary>
        /// 运行完整的报销自动化流程（包含Python脚本处理打印对话框）
        /// </summary>
        public async Task RunAsync()
        {
            try
            {
                // 1. 检查Python环境
                Console.WriteLine("=== 检查Python环境 ===");
                bool envOk = await _pythonInterop.CheckPythonEnvironmentAsync();
                if (!envOk)
                {
                    Console.WriteLine("警告：Python环境异常，将跳过打印对话框自动处理");
                    Console.WriteLine("请确保已安装Python和pyautogui库");
                }
                else
                {
                    Console.WriteLine("✓ Python环境正常");
                }
                
                // 2. 运行原有的报销自动化流程
                Console.WriteLine("\n=== 开始报销自动化流程 ===");
                await _reimbursementAutomation.RunAsync();
                
                // 3. 如果Python环境正常，提示用户点击打印按钮
                if (envOk)
                {
                    Console.WriteLine("\n=== 准备处理打印对话框 ===");
                    Console.WriteLine("报销流程已完成，请点击报销确认单按钮...");
                    Console.WriteLine("程序将在2秒后自动处理打印对话框...");
                    
                    // 等待2秒
                    await Task.Delay(2000);
                    
                    // 自动执行打印对话框处理
                    await AutoHandlePrintDialog();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"报销自动化流程执行失败: {ex.Message}");
                throw;
            }
        }
        
        /// <summary>
        /// 自动处理打印对话框
        /// </summary>
        private async Task AutoHandlePrintDialog()
        {
            try
            {
                Console.WriteLine("开始自动处理打印对话框...");
                
                // 使用配置文件中的路径和当前时间戳生成文件名
                string filePath = @"C:\Users\FH\PycharmProjects\CursorCode8-5\pdf_output";
                string fileName = $"报销单_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                
                bool success = await _pythonInterop.ExecutePrintDialogProcessAsync(filePath, fileName);
                
                if (success)
                {
                    Console.WriteLine($"✓ 打印对话框处理成功！");
                    Console.WriteLine($"文件已保存到: {System.IO.Path.Combine(filePath, fileName)}");
                }
                else
                {
                    Console.WriteLine("✗ 打印对话框处理失败");
                    Console.WriteLine("请检查：");
                    Console.WriteLine("1. 打印对话框是否正确显示");
                    Console.WriteLine("2. 坐标配置是否正确");
                    Console.WriteLine("3. 文件路径是否存在");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自动处理打印对话框时出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 手动触发打印对话框处理（用于测试）
        /// </summary>
        public async Task ManualTriggerPrintDialog()
        {
            Console.WriteLine("=== 手动触发打印对话框处理 ===");
            await AutoHandlePrintDialog();
        }
        
        /// <summary>
        /// 获取鼠标位置（用于调试坐标）
        /// </summary>
        public async Task GetMousePosition()
        {
            try
            {
                string mousePos = await _pythonInterop.GetMousePositionAsync();
                Console.WriteLine($"当前鼠标位置: {mousePos}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"获取鼠标位置失败: {ex.Message}");
            }
        }
    }
}

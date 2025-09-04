using System;
using System.Threading.Tasks;
using Microsoft.Playwright;
using System.Collections.Generic;
using System.Linq;
using OfficeOpenXml;
using System.IO;
using System.Text.Json;

namespace AutoFinan
{
    public class ResearchFinanceAutomation
    {
        private IPlaywright playwright;
        private IBrowser browser;
        private IPage page;
        private bool isLoggedIn = false;

        /// <summary>
        /// 启动浏览器并导航到科研财务系统
        /// </summary>
        public async Task InitializeAsync()
        {
            try
            {
                Console.WriteLine("=== 科研财务系统自动化 ===");
                Console.WriteLine("正在启动浏览器...");

                playwright = await Playwright.CreateAsync();
                browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                {
                    Headless = false, // 显示浏览器窗口
                    SlowMo = 100 // 放慢操作速度，便于观察
                });

                page = await browser.NewPageAsync();
                page.SetDefaultTimeout(10000); // 10秒超时

                Console.WriteLine("浏览器启动成功");
                Console.WriteLine("正在导航到科研财务系统...");

                // 导航到目标网页
                await page.GotoAsync("https://www.kycw.uestc.edu.cn/WFManager/home.jsp");
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

                Console.WriteLine("成功导航到科研财务系统");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"初始化失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 执行登录流程
        /// </summary>
        public async Task<bool> LoginAsync()
        {
            try
            {
                Console.WriteLine("开始执行登录流程...");

                // 等待登录表单加载
                await page.WaitForSelectorAsync("#uid", new PageWaitForSelectorOptions { Timeout = 5000 });
                Console.WriteLine("登录表单加载完成");

                // 填写用户名
                await page.FillAsync("#uid", "5070016");
                Console.WriteLine("已填写用户名: 5070016");

                // 填写密码
                await page.FillAsync("#pwd", "Kp5070016");
                Console.WriteLine("已填写密码: Kp5070016");

                // 等待用户输入验证码
                Console.WriteLine("请查看验证码图片并在下方输入验证码:");
                Console.WriteLine("验证码输入框ID: chkcode1");
                Console.WriteLine("验证码图片ID: checkcodeImg");
                Console.WriteLine("登录按钮ID: zhLogin");
                Console.WriteLine("----------------------------------------");

                // 高亮验证码输入框
                await page.EvaluateAsync(@"
                    document.getElementById('chkcode1').style.border = '2px solid red';
                    document.getElementById('chkcode1').style.backgroundColor = '#fff3cd';
                    document.getElementById('chkcode1').focus();
                ");

                // 等待用户在控制台输入验证码
                Console.Write("请输入验证码: ");
                var captchaInput = Console.ReadLine()?.Trim();

                if (string.IsNullOrEmpty(captchaInput))
                {
                    Console.WriteLine("错误：验证码不能为空");
                    return false;
                }

                Console.WriteLine($"收到验证码: {captchaInput}");

                // 将验证码填充到网页输入框
                await page.FillAsync("#chkcode1", captchaInput);
                Console.WriteLine("验证码已填充到网页输入框");

                // 移除高亮样式
                await page.EvaluateAsync(@"
                    document.getElementById('chkcode1').style.border = '';
                    document.getElementById('chkcode1').style.backgroundColor = '';
                ");

                Console.WriteLine("准备点击登录按钮...");

                // 等待登录按钮可用
                Console.WriteLine("等待登录按钮可用...");
                try
                {
                    // 首先尝试等待按钮可见
                    await page.WaitForSelectorAsync("#zhLogin", new PageWaitForSelectorOptions { Timeout = 5000 });
                    Console.WriteLine("登录按钮已找到");

                    // 检查按钮是否可用
                    var buttonElement = page.Locator("#zhLogin");
                    var isDisabled = await buttonElement.GetAttributeAsync("disabled");

                    if (isDisabled != null)
                    {
                        Console.WriteLine("登录按钮当前被禁用，等待启用...");
                        // 等待按钮启用
                        await page.WaitForSelectorAsync("#zhLogin:not([disabled])", new PageWaitForSelectorOptions { Timeout = 10000 });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"等待登录按钮时出错: {ex.Message}");
                    Console.WriteLine("尝试直接点击登录按钮...");
                }

                // 点击登录按钮
                Console.WriteLine("正在点击登录按钮...");
                await page.ClickAsync("#zhLogin");
                Console.WriteLine("已点击登录按钮");

                // 添加调试信息
                Console.WriteLine("等待页面响应...");
                await Task.Delay(3000); // 等待3秒

                // 检查当前URL
                var currentUrlAfterClick = page.Url;
                Console.WriteLine($"点击登录按钮后的URL: {currentUrlAfterClick}");

                // 等待更长时间让页面完全加载
                await Task.Delay(5000);

                // 检查是否登录成功 - 使用更宽松的判断条件
                try
                {
                    // 检查URL变化 - 如果URL发生了变化，通常表示登录成功
                    var finalUrl = page.Url;
                    Console.WriteLine($"最终URL: {finalUrl}");

                    // 如果URL包含登录相关的关键词，可能还在登录页面
                    if (finalUrl.Contains("login") || finalUrl.Contains("auth"))
                    {
                        Console.WriteLine("仍在登录页面，登录可能失败");
                        return false;
                    }

                    // 检查是否有明显的错误信息
                    var pageContent = await page.ContentAsync();
                    if (pageContent.Contains("用户名或密码错误") ||
                        pageContent.Contains("验证码错误") ||
                        pageContent.Contains("登录失败"))
                    {
                        Console.WriteLine("页面包含明确的错误信息，登录失败");
                        return false;
                    }

                    // 检查是否有登录成功的标识
                    var successIndicators = new[] { "欢迎", "您好", "登录成功", "首页", "主页面" };
                    var hasSuccessIndicator = false;

                    foreach (var indicator in successIndicators)
                    {
                        try
                        {
                            var locator = page.Locator($"text={indicator}");
                            if (await locator.CountAsync() > 0)
                            {
                                Console.WriteLine($"找到成功标识: {indicator}");
                                hasSuccessIndicator = true;
                                break;
                            }
                        }
                        catch
                        {
                            // 忽略单个选择器的错误
                        }
                    }

                    if (hasSuccessIndicator)
                    {
                        Console.WriteLine("登录成功！");
                        isLoggedIn = true;
                        return true;
                    }

                    // 如果没有明确的成功标识，但URL已经变化且没有错误信息，也认为登录成功
                    if (finalUrl != "https://www.kycw.uestc.edu.cn/WFManager/home.jsp")
                    {
                        Console.WriteLine("URL已变化且无错误信息，认为登录成功");
                        isLoggedIn = true;
                        return true;
                    }

                    Console.WriteLine("登录状态不明确，但无明确错误，认为登录成功");
                    isLoggedIn = true;
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"检查登录状态时出错: {ex.Message}");
                    // 即使检查出错，如果URL已经变化，也认为可能登录成功
                    var currentUrl = page.Url;
                    if (currentUrl != "https://www.kycw.uestc.edu.cn/WFManager/home.jsp")
                    {
                        Console.WriteLine("检查出错但URL已变化，认为登录成功");
                        isLoggedIn = true;
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"登录过程出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 导航到指定页面
        /// </summary>
        public async Task NavigateToPageAsync(string url)
        {
            try
            {
                if (!isLoggedIn)
                {
                    Console.WriteLine("错误：尚未登录，请先执行登录");
                    return;
                }

                Console.WriteLine($"正在导航到: {url}");
                await page.GotoAsync(url);
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                Console.WriteLine("页面导航完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"页面导航失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 点击数据查询及公示导航按钮
        /// </summary>
        public async Task<bool> ClickDataQueryAndPublicityButtonAsync()
        {
            try
            {
                Console.WriteLine("正在精确查找'数据查询及公示'菜单项...");
                
                var startTime = DateTime.Now;
                var timeout = TimeSpan.FromMinutes(2); // 2分钟超时
                var retryInterval = TimeSpan.FromSeconds(5); // 5秒重试间隔
                var attemptCount = 0;

                while (DateTime.Now - startTime < timeout)
                {
                    attemptCount++;
                    Console.WriteLine($"\n=== 第 {attemptCount} 次尝试查找 ===");
                    Console.WriteLine($"已用时: {DateTime.Now - startTime:mm\\:ss}, 剩余时间: {timeout - (DateTime.Now - startTime):mm\\:ss}");

                    // 首先在主页面查找
                    Console.WriteLine("1. 在主页面查找...");
                    var buttonLocator = await FindButtonInPage(page, "主页面");
                    if (buttonLocator != null)
                    {
                        var result = await ClickButton(buttonLocator, "主页面");
                        if (result)
                        {
                            await ReadExcelAfterNavigationAsync();
                        }
                        return result;
                    }

                    // 如果主页面没找到，在所有iframe中查找
                    Console.WriteLine("2. 在所有iframe中查找...");
                    var frames = page.Frames;
                    Console.WriteLine($"找到 {frames.Count} 个iframe");

                    for (int i = 0; i < frames.Count; i++)
                    {
                        var frame = frames[i];
                        Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                        
                        try
                        {
                            var frameButtonLocator = await FindButtonInFrame(frame, $"iframe {i + 1}");
                            if (frameButtonLocator != null)
                            {
                                var result = await ClickButton(frameButtonLocator, $"iframe {i + 1}");
                                if (result)
                                {
                                    await ReadExcelAfterNavigationAsync();
                                }
                                return result;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                        }
                    }

                    // 如果还没超时，等待5秒后重试
                    if (DateTime.Now - startTime < timeout)
                    {
                        Console.WriteLine($"未找到按钮，等待 {retryInterval.TotalSeconds} 秒后重试...");
                        await Task.Delay(retryInterval);
                    }
                }

                // 超时退出
                var totalTime = DateTime.Now - startTime;
                Console.WriteLine($"\n❌ 查找超时！在 {totalTime:mm\\:ss} 内未找到'数据查询及公示'按钮");
                Console.WriteLine("可能原因：");
                Console.WriteLine("1. 网页加载缓慢");
                Console.WriteLine("2. 网络连接问题");
                Console.WriteLine("3. 页面结构发生变化");
                Console.WriteLine("4. 服务器响应慢");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找按钮时发生错误: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 在指定页面中查找按钮
        /// </summary>
        private async Task<ILocator> FindButtonInPage(IPage pageOrFrame, string pageName)
        {
            try
            {
                Console.WriteLine($"  在{pageName}中使用精确选择器查找...");

                // 使用非常精确的选择器，基于HTML结构
                var exactSelector = "ul#ulLevel1List > li.level1menu:has(a[onclick*='项目执行情况查询'])";
                var buttonLocator = pageOrFrame.Locator(exactSelector);
                var count = await buttonLocator.CountAsync();

                if (count > 0)
                {
                    Console.WriteLine($"  在{pageName}中找到按钮，使用选择器: {exactSelector}");
                    return buttonLocator;
                }

                Console.WriteLine($"  在{pageName}中使用精确选择器未找到，尝试备用选择器...");

                // 备用选择器
                var backupSelectors = new[]
                {
                    "li.level1menu:has(a:has(span:text('数据查询及公示')))",
                    "a:has(span:text('数据查询及公示'))",
                    "text=数据查询及公示",
                    "a[onclick*='项目执行情况查询']",
                    "a[onclick*='数据查询及公示']"
                };

                foreach (var selector in backupSelectors)
                {
                    try
                    {
                        buttonLocator = pageOrFrame.Locator(selector);
                        count = await buttonLocator.CountAsync();
                        if (count > 0)
                        {
                            Console.WriteLine($"  在{pageName}中找到按钮，使用选择器: {selector}");
                            return buttonLocator;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  在{pageName}中使用选择器 {selector} 失败: {ex.Message}");
                    }
                }

                // 尝试查找包含文本的所有元素
                Console.WriteLine($"  在{pageName}中尝试查找包含'数据查询及公示'文本的所有元素...");
                var allElements = await pageOrFrame.Locator("*:has-text('数据查询及公示')").AllAsync();
                Console.WriteLine($"  在{pageName}中找到 {allElements.Count} 个包含'数据查询及公示'文本的元素");

                foreach (var element in allElements)
                {
                    try
                    {
                        var tagName = await element.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
                        var isClickable = await element.EvaluateAsync<bool>("el => !!(el.onclick || el.getAttribute('href') || el.tagName.toLowerCase() === 'a')");
                        
                        Console.WriteLine($"    元素: 标签={tagName}, 可点击={isClickable}");
                        
                        if (isClickable)
                        {
                            // 创建一个定位器来包装这个元素
                            var elementIndex = -1;
                            for (int j = 0; j < allElements.Count; j++)
                            {
                                if (allElements[j] == element)
                                {
                                    elementIndex = j;
                                    break;
                                }
                            }
                            
                            if (elementIndex >= 0)
                            {
                                var elementLocator = pageOrFrame.Locator($"*:has-text('数据查询及公示')").Nth(elementIndex);
                                Console.WriteLine($"  在{pageName}中找到可点击的按钮元素");
                                return elementLocator;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"    检查元素时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"  在{pageName}中未找到按钮");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  在{pageName}中查找按钮时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 在指定iframe中查找按钮
        /// </summary>
        private async Task<ILocator> FindButtonInFrame(IFrame frame, string frameName)
        {
            try
            {
                Console.WriteLine($"  在{frameName}中使用精确选择器查找...");

                // 使用非常精确的选择器，基于HTML结构
                var exactSelector = "ul#ulLevel1List > li.level1menu:has(a[onclick*='项目执行情况查询'])";
                var buttonLocator = frame.Locator(exactSelector);
                var count = await buttonLocator.CountAsync();

                if (count > 0)
                {
                    Console.WriteLine($"  在{frameName}中找到按钮，使用选择器: {exactSelector}");
                    return buttonLocator;
                }

                Console.WriteLine($"  在{frameName}中使用精确选择器未找到，尝试备用选择器...");

                // 备用选择器
                var backupSelectors = new[]
                {
                    "li.level1menu:has(a:has(span:text('数据查询及公示')))",
                    "a:has(span:text('数据查询及公示'))",
                    "text=数据查询及公示",
                    "a[onclick*='项目执行情况查询']",
                    "a[onclick*='数据查询及公示']"
                };

                foreach (var selector in backupSelectors)
                {
                    try
                    {
                        buttonLocator = frame.Locator(selector);
                        count = await buttonLocator.CountAsync();
                        if (count > 0)
                        {
                            Console.WriteLine($"  在{frameName}中找到按钮，使用选择器: {selector}");
                            return buttonLocator;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  在{frameName}中使用选择器 {selector} 失败: {ex.Message}");
                    }
                }

                // 尝试查找包含文本的所有元素
                Console.WriteLine($"  在{frameName}中尝试查找包含'数据查询及公示'文本的所有元素...");
                var allElements = await frame.Locator("*:has-text('数据查询及公示')").AllAsync();
                Console.WriteLine($"  在{frameName}中找到 {allElements.Count} 个包含'数据查询及公示'文本的元素");

                foreach (var element in allElements)
                {
                    try
                    {
                        var tagName = await element.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
                        var isClickable = await element.EvaluateAsync<bool>("el => !!(el.onclick || el.getAttribute('href') || el.tagName.toLowerCase() === 'a')");
                        
                        Console.WriteLine($"    元素: 标签={tagName}, 可点击={isClickable}");
                        
                        if (isClickable)
                        {
                            // 创建一个定位器来包装这个元素
                            var elementIndex = -1;
                            for (int j = 0; j < allElements.Count; j++)
                            {
                                if (allElements[j] == element)
                                {
                                    elementIndex = j;
                                    break;
                                }
                            }
                            
                            if (elementIndex >= 0)
                            {
                                var elementLocator = frame.Locator($"*:has-text('数据查询及公示')").Nth(elementIndex);
                                Console.WriteLine($"  在{frameName}中找到可点击的按钮元素");
                                return elementLocator;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"    检查元素时出错: {ex.Message}");
                    }
                }

                Console.WriteLine($"  在{frameName}中未找到按钮");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"  在{frameName}中查找按钮时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 点击按钮
        /// </summary>
        private async Task<bool> ClickButton(ILocator buttonLocator, string pageName)
        {
            try
            {
                Console.WriteLine($"在{pageName}中找到按钮，准备点击...");

                // 获取按钮文本用于确认
                var buttonText = await buttonLocator.TextContentAsync();
                Console.WriteLine($"按钮文本: {buttonText?.Trim()}");

                // 等待2秒让按钮完全加载
                Console.WriteLine("等待2秒让按钮完全加载...");
                await Task.Delay(2000);

                // 点击按钮
                await buttonLocator.ClickAsync(new LocatorClickOptions
                {
                    Timeout = 5000,
                    Force = true
                });

                Console.WriteLine($"在{pageName}中点击成功！等待页面加载...");
                await Task.Delay(3000);

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"在{pageName}中点击失败: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 导航成功后读取Excel文件
        /// </summary>
        private async Task ReadExcelAfterNavigationAsync()
        {
            try
            {
                Console.WriteLine("导航成功，开始读取Excel文件...");
                var projectNumbers = await ReadProjectNumbersFromExcelAsync();
                
                if (projectNumbers.Count > 0)
                {
                    Console.WriteLine($"成功读取到 {projectNumbers.Count} 个项目编号:");
                    for (int i = 0; i < projectNumbers.Count; i++)
                    {
                        Console.WriteLine($"  {i + 1}. {projectNumbers[i]}");
                    }
                    
                    // 开始自动搜索项目编号
                    await AutoSearchProjectNumbers(projectNumbers);
                }
                else
                {
                    Console.WriteLine("未读取到任何项目编号");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取Excel文件时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 读取Excel文件中的项目编号
        /// </summary>
        public async Task<List<string>> ReadProjectNumbersFromExcelAsync()
        {
            try
            {
                Console.WriteLine("开始读取Excel文件中的项目编号...");
                
                // 读取配置文件 - 使用逐级向上查找的方法
                var configPath = FindConfigFile();
                if (string.IsNullOrEmpty(configPath))
                {
                    Console.WriteLine("错误：找不到配置文件 config.json");
                    return new List<string>();
                }

                var configContent = await File.ReadAllTextAsync(configPath);
                
                // 使用更灵活的JSON解析，处理可能的非字符串值
                var jsonOptions = new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true,
                    AllowTrailingCommas = true
                };
                
                var config = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(configContent, jsonOptions);
                
                if (!config.ContainsKey("ExcelFilePath"))
                {
                    Console.WriteLine("错误：配置文件中未找到ExcelFilePath");
                    return new List<string>();
                }

                var excelFilePath = config["ExcelFilePath"]?.ToString();
                if (string.IsNullOrEmpty(excelFilePath))
                {
                    Console.WriteLine("错误：ExcelFilePath配置项为空");
                    return new List<string>();
                }
                
                // 查找Excel文件 - 使用逐级向上查找的方法
                var fullExcelPath = FindExcelFile(excelFilePath);
                if (string.IsNullOrEmpty(fullExcelPath))
                {
                    Console.WriteLine($"错误：找不到Excel文件: {excelFilePath}");
                    return new List<string>();
                }

                Console.WriteLine($"正在读取Excel文件: {fullExcelPath}");
                
                // 设置EPPlus许可证上下文
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                
                using (var package = new ExcelPackage(new FileInfo(fullExcelPath)))
                {
                    var worksheet = package.Workbook.Worksheets["0-Proj信息"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("错误：未找到名为'0-Proj信息'的工作表");
                        return new List<string>();
                    }

                    Console.WriteLine("成功找到'0-Proj信息'工作表");
                    
                    // 查找"项目编号"列
                    var projectNumberColumn = -1;
                    var maxColumns = worksheet.Dimension?.Columns ?? 0;
                    
                    for (int col = 1; col <= maxColumns; col++)
                    {
                        var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                        if (headerValue == "项目编号")
                        {
                            projectNumberColumn = col;
                            Console.WriteLine($"找到'项目编号'列，位于第{col}列");
                            break;
                        }
                    }

                    if (projectNumberColumn == -1)
                    {
                        Console.WriteLine("错误：未找到'项目编号'列");
                        return new List<string>();
                    }

                                         // 查找"预算执行是否需要更新"列
                     var updateColumn = -1;
                     for (int col = 1; col <= maxColumns; col++)
                     {
                         var headerValue = worksheet.Cells[1, col].Value?.ToString()?.Trim();
                         if (headerValue == "预算执行是否需要更新")
                         {
                             updateColumn = col;
                             Console.WriteLine($"找到'预算执行是否需要更新'列，位于第{col}列");
                             break;
                         }
                     }

                     if (updateColumn == -1)
                     {
                         Console.WriteLine("错误：未找到'预算执行是否需要更新'列");
                         return new List<string>();
                     }

                     // 读取项目编号数据，只包含需要更新的项目
                     var projectNumbers = new List<string>();
                     var maxRows = worksheet.Dimension?.Rows ?? 0;
                     
                     for (int row = 2; row <= maxRows; row++) // 从第2行开始（跳过标题行）
                     {
                         var cellValue = worksheet.Cells[row, projectNumberColumn].Value?.ToString()?.Trim();
                         var updateValue = worksheet.Cells[row, updateColumn].Value?.ToString()?.Trim();
                         
                         if (!string.IsNullOrEmpty(cellValue))
                         {
                             // 检查是否需要更新
                             if (updateValue == "是")
                             {
                                 projectNumbers.Add(cellValue);
                                 Console.WriteLine($"读取到项目编号: {cellValue} (需要更新)");
                             }
                             else
                             {
                                 Console.WriteLine($"跳过项目编号: {cellValue} (不需要更新，值为: {updateValue})");
                             }
                         }
                     }

                    Console.WriteLine($"成功读取到 {projectNumbers.Count} 个项目编号");
                    return projectNumbers;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"读取Excel文件时出错: {ex.Message}");
                Console.WriteLine($"异常详情: {ex}");
                return new List<string>();
            }
        }

        /// <summary>
        /// 查找页面中的特定信息
        /// </summary>
        public async Task<Dictionary<string, string>> ExtractPageInfoAsync()
        {
            try
            {
                if (!isLoggedIn)
                {
                    Console.WriteLine("错误：尚未登录，请先执行登录");
                    return new Dictionary<string, string>();
                }

                Console.WriteLine("开始提取页面信息...");
                var pageInfo = new Dictionary<string, string>();

                // 获取页面标题
                var title = await page.TitleAsync();
                pageInfo["页面标题"] = title;
                Console.WriteLine($"页面标题: {title}");

                // 获取当前URL
                var currentUrl = page.Url;
                pageInfo["当前URL"] = currentUrl;
                Console.WriteLine($"当前URL: {currentUrl}");

                Console.WriteLine($"成功提取 {pageInfo.Count} 条页面信息");
                return pageInfo;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取页面信息失败: {ex.Message}");
                return new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// 等待用户操作完成
        /// </summary>
        public async Task WaitForUserOperationAsync()
        {
            Console.WriteLine("等待用户操作完成...");
            Console.WriteLine("按回车键继续程序执行...");

            try
            {
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"等待用户输入时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 关闭浏览器
        /// </summary>
        public async Task CloseAsync()
        {
            try
            {
                if (browser != null)
                {
                    await browser.CloseAsync();
                    Console.WriteLine("浏览器已关闭");
                }

                if (playwright != null)
                {
                    playwright.Dispose();
                    Console.WriteLine("Playwright资源已释放");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"关闭资源时出错: {ex.Message}");
            }
        }

        /// <summary>
        /// 检查登录状态
        /// </summary>
        public bool IsLoggedIn => isLoggedIn;

        /// <summary>
        /// 获取当前页面对象（供外部使用）
        /// </summary>
        public IPage CurrentPage => page;

        /// <summary>
        /// 查找配置文件 - 使用逐级向上查找的方法
        /// </summary>
        private string FindConfigFile()
        {
            try
            {
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
                    Console.WriteLine("尝试过的配置文件路径:");
                    foreach (string path in possibleConfigPaths)
                    {
                        Console.WriteLine($"  {path}");
                    }
                }

                return configPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找配置文件时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 查找Excel文件 - 使用逐级向上查找的方法
        /// </summary>
        private string FindExcelFile(string excelFileName)
        {
            try
            {
                // 获取当前执行目录
                string currentDirectory = Directory.GetCurrentDirectory();
                Console.WriteLine($"查找Excel文件: {excelFileName}");

                // 尝试多个可能的文件路径
                string[] possiblePaths = {
                    excelFileName,
                    Path.Combine(currentDirectory, excelFileName),
                    Path.Combine(currentDirectory, "..", excelFileName),
                    Path.Combine(currentDirectory, "..", "..", excelFileName),
                    Path.Combine(currentDirectory, "..", "..", "..", excelFileName),
                    Path.Combine(currentDirectory, "..", "..", "..", "..", excelFileName)
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
                    Console.WriteLine("尝试过的Excel文件路径:");
                    foreach (string path in possiblePaths)
                    {
                        Console.WriteLine($"  {path}");
                    }
                }

                return actualExcelPath;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找Excel文件时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 查找搜索输入框 - 先在主页面查找，然后在iframe中查找
        /// </summary>
        private async Task<ILocator> FindSearchInputBox()
        {
            try
            {
                Console.WriteLine("正在查找搜索输入框 #gridWF_KYCW_7203_qspi...");
                
                // 首先在主页面查找
                Console.WriteLine("1. 在主页面查找搜索输入框...");
                var mainPageLocator = page.Locator("#gridWF_KYCW_7203_qspi");
                var mainPageCount = await mainPageLocator.CountAsync();
                
                if (mainPageCount > 0)
                {
                    Console.WriteLine("在主页面中找到搜索输入框");
                    return mainPageLocator;
                }
                
                // 如果主页面没找到，在所有iframe中查找
                Console.WriteLine("2. 在所有iframe中查找搜索输入框...");
                var frames = page.Frames;
                Console.WriteLine($"找到 {frames.Count} 个iframe");
                
                for (int i = 0; i < frames.Count; i++)
                {
                    var frame = frames[i];
                    Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                    
                    try
                    {
                        var frameLocator = frame.Locator("#gridWF_KYCW_7203_qspi");
                        var frameCount = await frameLocator.CountAsync();
                        
                        if (frameCount > 0)
                        {
                            Console.WriteLine($"在iframe {i + 1} 中找到搜索输入框");
                            return frameLocator;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                    }
                }
                
                Console.WriteLine("在所有iframe中都未找到搜索输入框");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找搜索输入框时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 等待表格数据加载并点击第一行第二列
        /// </summary>
        /// <returns>true表示成功找到数据并点击，false表示未找到数据</returns>
        private async Task<bool> WaitForTableAndClickFirstRow()
        {
            try
            {
                Console.WriteLine("等待表格数据加载...");
                
                // 等待表格加载 - 先在主页面查找，然后在iframe中查找
                var tableLocator = await FindTableInPage();
                if (tableLocator == null)
                {
                    Console.WriteLine("警告：未找到表格，跳过点击操作");
                    return false;
                }
                
                // 等待表格有数据行（不是只有表头）
                Console.WriteLine("等待表格数据行加载...");
                var maxWaitTime = TimeSpan.FromSeconds(10);
                var startTime = DateTime.Now;
                var hasDataRows = false;
                
                while (DateTime.Now - startTime < maxWaitTime)
                {
                    try
                    {
                        // 查找数据行（排除第一行表头）
                        var dataRows = tableLocator.Locator("tbody tr:not(.jqgfirstrow)");
                        var rowCount = await dataRows.CountAsync();
                        
                        if (rowCount > 0)
                        {
                            Console.WriteLine($"表格已加载，找到 {rowCount} 行数据");
                            hasDataRows = true;
                            break;
                        }
                        
                        await Task.Delay(500);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"检查表格行数时出错: {ex.Message}");
                        await Task.Delay(500);
                    }
                }
                
                if (!hasDataRows)
                {
                    Console.WriteLine("警告：表格数据行加载超时，跳过点击操作");
                    return false;
                }
                
                // 点击第一行第二列
                Console.WriteLine("准备点击第一行第二列...");
                try
                {
                    // 尝试多种选择器来找到第一行第二列
                    var firstRowSecondCell = null as ILocator;
                    var cellCount = 0;
                    
                    // 方法1：使用更直接的选择器
                    var selector1 = "tbody tr[role='row']:not(.jqgfirstrow) td:nth-child(2)";
                    firstRowSecondCell = tableLocator.Locator(selector1);
                    cellCount = await firstRowSecondCell.CountAsync();
                    
                    if (cellCount == 0)
                    {
                        // 方法2：使用更宽松的选择器
                        Console.WriteLine("方法1失败，尝试方法2...");
                        var selector2 = "tbody tr:not(.jqgfirstrow) td:nth-child(2)";
                        firstRowSecondCell = tableLocator.Locator(selector2);
                        cellCount = await firstRowSecondCell.CountAsync();
                    }
                    
                    if (cellCount == 0)
                    {
                        // 方法3：先找到第一行，再找第二列
                        Console.WriteLine("方法2失败，尝试方法3...");
                        var firstRow = tableLocator.Locator("tbody tr:not(.jqgfirstrow)").First;
                        if (await firstRow.CountAsync() > 0)
                        {
                            firstRowSecondCell = firstRow.Locator("td:nth-child(2)");
                            cellCount = await firstRowSecondCell.CountAsync();
                        }
                    }
                    
                    if (cellCount == 0)
                    {
                        // 方法4：使用最宽松的选择器
                        Console.WriteLine("方法3失败，尝试方法4...");
                        var selector4 = "tbody tr td:nth-child(2)";
                        firstRowSecondCell = tableLocator.Locator(selector4);
                        cellCount = await firstRowSecondCell.CountAsync();
                    }
                    
                    if (cellCount > 0)
                    {
                        // 获取单元格内容用于确认
                        var cellText = await firstRowSecondCell.TextContentAsync();
                        Console.WriteLine($"找到第一行第二列，内容: {cellText?.Trim()}");
                        
                        // 点击单元格
                        await firstRowSecondCell.ClickAsync();
                        Console.WriteLine("成功点击第一行第二列，触发跳转");
                        
                        // 等待跳转完成
                        await Task.Delay(2000);
                        return true;
                    }
                    else
                    {
                        Console.WriteLine("警告：所有方法都失败，未找到第一行第二列，跳过点击操作");
                        
                        // 调试信息：显示所有找到的行和列
                        try
                        {
                            var allRows = tableLocator.Locator("tbody tr");
                            var rowCount = await allRows.CountAsync();
                            Console.WriteLine($"调试信息：表格总共有 {rowCount} 行");
                            
                            for (int i = 0; i < Math.Min(rowCount, 3); i++) // 只显示前3行
                            {
                                var row = allRows.Nth(i);
                                var cells = row.Locator("td");
                                var cellCountInRow = await cells.CountAsync();
                                Console.WriteLine($"第 {i + 1} 行有 {cellCountInRow} 列");
                                
                                if (cellCountInRow > 1)
                                {
                                    var secondCell = cells.Nth(1);
                                    var secondCellText = await secondCell.TextContentAsync();
                                    Console.WriteLine($"第 {i + 1} 行第2列内容: {secondCellText?.Trim()}");
                                }
                            }
                        }
                        catch (Exception debugEx)
                        {
                            Console.WriteLine($"调试信息获取失败: {debugEx.Message}");
                        }
                        return false;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"点击第一行第二列时出错: {ex.Message}");
                    return false;
                }
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"等待表格并点击第一行时出错: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 查找表格 - 先在主页面查找，然后在iframe中查找
        /// </summary>
        private async Task<ILocator> FindTableInPage()
        {
            try
            {
                Console.WriteLine("正在查找表格 #gridWF_KYCW_7203...");
                
                // 首先在主页面查找
                Console.WriteLine("1. 在主页面查找表格...");
                var mainPageLocator = page.Locator("#gridWF_KYCW_7203");
                var mainPageCount = await mainPageLocator.CountAsync();
                
                if (mainPageCount > 0)
                {
                    Console.WriteLine("在主页面中找到表格");
                    return mainPageLocator;
                }
                
                // 如果主页面没找到，在所有iframe中查找
                Console.WriteLine("2. 在所有iframe中查找表格...");
                var frames = page.Frames;
                Console.WriteLine($"找到 {frames.Count} 个iframe");
                
                for (int i = 0; i < frames.Count; i++)
                {
                    var frame = frames[i];
                    Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                    
                    try
                    {
                        var frameLocator = frame.Locator("#gridWF_KYCW_7203");
                        var frameCount = await frameLocator.CountAsync();
                        
                        if (frameCount > 0)
                        {
                            Console.WriteLine($"在iframe {i + 1} 中找到表格");
                            return frameLocator;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                    }
                }
                
                Console.WriteLine("在所有iframe中都未找到表格");
                return null;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找表格时出错: {ex.Message}");
                return null;
            }
        }

                 /// <summary>
         /// 在当前页面中查找并点击"预算情况"按钮
         /// </summary>
         private async Task ClickBudgetButtonInCurrentPage()
         {
             try
             {
                 Console.WriteLine("在当前页面中查找'预算情况'按钮...");
                 
                 // 等待页面完全加载
                 await Task.Delay(2000);
                 
                 // 首先在主页面查找
                 Console.WriteLine("1. 在主页面查找'预算情况'按钮...");
                 var mainPageLocator = page.Locator("a:has(span:text('预算情况'))");
                 var mainPageCount = await mainPageLocator.CountAsync();
                 
                 if (mainPageCount > 0)
                 {
                     Console.WriteLine("在主页面中找到'预算情况'按钮");
                     await mainPageLocator.ClickAsync();
                     Console.WriteLine("成功点击'预算情况'按钮");
                     return;
                 }
                 
                 // 如果主页面没找到，在所有iframe中查找
                 Console.WriteLine("2. 在所有iframe中查找'预算情况'按钮...");
                 var frames = page.Frames;
                 Console.WriteLine($"找到 {frames.Count} 个iframe");
                 
                 for (int i = 0; i < frames.Count; i++)
                 {
                     var frame = frames[i];
                     Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                     
                     try
                     {
                         // 使用多种选择器查找按钮
                         var selectors = new[]
                         {
                             "a:has(span:text('预算情况'))",
                             "a[href*='wgPager_WF_KY_7942d100001_2']",
                             "a[innerwinno='16171d100001']",
                             "text=预算情况",
                             "ul#wg_ul_WF_KY_7942d100001 li:has(a:has(span:text('预算情况')))",
                             "li:has(a:has(span:text('预算情况')))"
                         };
                         
                         foreach (var selector in selectors)
                         {
                             try
                             {
                                 var frameLocator = frame.Locator(selector);
                                 var frameCount = await frameLocator.CountAsync();
                                 
                                 if (frameCount > 0)
                                 {
                                     Console.WriteLine($"在iframe {i + 1} 中找到'预算情况'按钮，使用选择器: {selector}");
                                     
                                     // 获取按钮信息用于确认
                                     var buttonText = await frameLocator.TextContentAsync();
                                     var buttonHref = await frameLocator.GetAttributeAsync("href");
                                     Console.WriteLine($"按钮文本: {buttonText?.Trim()}, 链接: {buttonHref}");
                                     
                                     // 点击按钮
                                     await frameLocator.ClickAsync();
                                     Console.WriteLine($"在iframe {i + 1} 中成功点击'预算情况'按钮");
                                     
                                     // 等待点击响应
                                     await Task.Delay(2000);
                                     return;
                                 }
                             }
                             catch (Exception ex)
                             {
                                 Console.WriteLine($"在iframe {i + 1} 中使用选择器 {selector} 失败: {ex.Message}");
                             }
                         }
                     }
                     catch (Exception ex)
                     {
                         Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                     }
                 }
                 
                 Console.WriteLine("警告：在所有iframe中都未找到'预算情况'按钮");
             }
             catch (Exception ex)
             {
                 Console.WriteLine($"查找并点击'预算情况'按钮时出错: {ex.Message}");
             }
         }
         
                 /// <summary>
        /// 在当前页面中查找并点击"当前执行预算"按钮
        /// </summary>
        private async Task ClickCurrentExecutionBudgetButton(string currentProjectNumber = null)
         {
             try
             {
                 Console.WriteLine("在当前页面中查找'当前执行预算'按钮...");
                 
                 // 等待页面完全加载
                 await Task.Delay(2000);
                 
                 // 首先在主页面查找
                 Console.WriteLine("1. 在主页面查找'当前执行预算'按钮...");
                 var mainPageLocator = page.Locator("a:has(span:text('当前执行预算'))");
                 var mainPageCount = await mainPageLocator.CountAsync();
                 
                 if (mainPageCount > 0)
                 {
                     Console.WriteLine("在主页面中找到'当前执行预算'按钮");
                     await mainPageLocator.ClickAsync();
                     Console.WriteLine("成功点击'当前执行预算'按钮");
                     
                     // 等待点击响应
                     await Task.Delay(2000);
                     
                     // 等待10秒后开始提取表格数据
                     Console.WriteLine("等待10秒后开始提取表格数据...");
                     await Task.Delay(10000);
                     
                     // 提取表格数据并写入Excel
                     await ExtractTableDataAndWriteToExcel(currentProjectNumber);
                     
                     // 点击关闭按钮
                     await ClickCloseButton();
                     
                     return;
                 }
                 
                 // 如果主页面没找到，在所有iframe中查找
                 Console.WriteLine("2. 在所有iframe中查找'当前执行预算'按钮...");
                 var frames = page.Frames;
                 Console.WriteLine($"找到 {frames.Count} 个iframe");
                 
                 for (int i = 0; i < frames.Count; i++)
                 {
                     var frame = frames[i];
                     Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                     
                     try
                     {
                         // 使用多种选择器查找按钮
                         var selectors = new[]
                         {
                             "a:has(span:text('当前执行预算'))",
                             "a[href*='wgPager_WF_KY_11107d100001_1']",
                             "a[innerwinno='12799d100001']",
                             "text=当前执行预算",
                             "ul#wg_ul_WF_KY_11107d100001 li:has(a:has(span:text('当前执行预算')))",
                             "li:has(a:has(span:text('当前执行预算')))"
                         };
                         
                         foreach (var selector in selectors)
                         {
                             try
                             {
                                 var frameLocator = frame.Locator(selector);
                                 var frameCount = await frameLocator.CountAsync();
                                 
                                 if (frameCount > 0)
                                 {
                                     Console.WriteLine($"在iframe {i + 1} 中找到'当前执行预算'按钮，使用选择器: {selector}");
                                     
                                     // 获取按钮信息用于确认
                                     var buttonText = await frameLocator.TextContentAsync();
                                     var buttonHref = await frameLocator.GetAttributeAsync("href");
                                     Console.WriteLine($"按钮文本: {buttonText?.Trim()}, 链接: {buttonHref}");
                                     
                                     // 点击按钮
                                     await frameLocator.ClickAsync();
                                     Console.WriteLine($"在iframe {i + 1} 中成功点击'当前执行预算'按钮");
                                     
                                     // 等待点击响应
                                     await Task.Delay(2000);
                                     
                                     // 等待10秒后开始提取表格数据
                                     Console.WriteLine("等待10秒后开始提取表格数据...");
                                     await Task.Delay(10000);
                                     
                                     // 提取表格数据并写入Excel
                                     await ExtractTableDataAndWriteToExcel(currentProjectNumber);
                                     
                                     // 点击关闭按钮
                                     await ClickCloseButton();
                                     
                                     return;
                                 }
                             }
                             catch (Exception ex)
                             {
                                 Console.WriteLine($"在iframe {i + 1} 中使用选择器 {selector} 失败: {ex.Message}");
                             }
                         }
                     }
                     catch (Exception ex)
                     {
                         Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                     }
                 }
                 
                 Console.WriteLine("警告：在所有iframe中都未找到'当前执行预算'按钮");
             }
             catch (Exception ex)
             {
                 Console.WriteLine($"查找并点击'当前执行预算'按钮时出错: {ex.Message}");
             }
         }
         
         /// <summary>
         /// 点击关闭按钮
         /// </summary>
         private async Task ClickCloseButton()
         {
             try
             {
                 Console.WriteLine("正在查找'关闭'按钮...");
                 
                 // 等待页面完全加载
                 await Task.Delay(1000);
                 
                 // 首先在主页面查找
                 Console.WriteLine("1. 在主页面查找'关闭'按钮...");
                 var mainPageLocator = page.Locator("button:has(span:text('关闭'))");
                 var mainPageCount = await mainPageLocator.CountAsync();
                 
                 if (mainPageCount > 0)
                 {
                     Console.WriteLine("在主页面中找到'关闭'按钮");
                     await mainPageLocator.ClickAsync();
                     Console.WriteLine("成功点击'关闭'按钮");
                     return;
                 }
                 
                 // 如果主页面没找到，在所有iframe中查找
                 Console.WriteLine("2. 在所有iframe中查找'关闭'按钮...");
                 var frames = page.Frames;
                 Console.WriteLine($"找到 {frames.Count} 个iframe");
                 
                 for (int i = 0; i < frames.Count; i++)
                 {
                     var frame = frames[i];
                     Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                     
                     try
                     {
                         // 使用多种选择器查找按钮
                         var selectors = new[]
                         {
                             "button:has(span:text('关闭'))",
                             "button.ui-button:has(span.ui-button-text:text('关闭'))",
                             "button[class*='ui-button']:has(span:text('关闭'))",
                             "text=关闭"
                         };
                         
                         foreach (var selector in selectors)
                         {
                             try
                             {
                                 var frameLocator = frame.Locator(selector);
                                 var frameCount = await frameLocator.CountAsync();
                                 
                                 if (frameCount > 0)
                                 {
                                     Console.WriteLine($"在iframe {i + 1} 中找到'关闭'按钮，使用选择器: {selector}");
                                     
                                     // 获取按钮信息用于确认
                                     var buttonText = await frameLocator.TextContentAsync();
                                     Console.WriteLine($"按钮文本: {buttonText?.Trim()}");
                                     
                                     // 点击按钮
                                     await frameLocator.ClickAsync();
                                     Console.WriteLine($"在iframe {i + 1} 中成功点击'关闭'按钮");
                                     
                                     // 等待点击响应
                                     await Task.Delay(1000);
                                     return;
                                 }
                             }
                             catch (Exception ex)
                             {
                                 Console.WriteLine($"在iframe {i + 1} 中使用选择器 {selector} 失败: {ex.Message}");
                             }
                         }
                     }
                     catch (Exception ex)
                     {
                         Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                     }
                 }
                 
                 Console.WriteLine("警告：在所有iframe中都未找到'关闭'按钮");
             }
             catch (Exception ex)
             {
                 Console.WriteLine($"查找并点击'关闭'按钮时出错: {ex.Message}");
             }
         }
         
         /// <summary>
         /// 在指定页面中查找并点击"预算情况"按钮（已废弃，保留供参考）
         /// </summary>
         private async Task<bool> ClickBudgetButtonInPage(IPage targetPage)
         {
             try
             {
                 Console.WriteLine("正在查找'预算情况'按钮...");
                 
                 // 首先在主页面查找
                 Console.WriteLine("1. 在主页面查找'预算情况'按钮...");
                 var mainPageLocator = targetPage.Locator("a:has(span:text('预算情况'))");
                 var mainPageCount = await mainPageLocator.CountAsync();
                 
                 if (mainPageCount > 0)
                 {
                     Console.WriteLine("在主页面中找到'预算情况'按钮");
                     await mainPageLocator.ClickAsync();
                     Console.WriteLine("成功点击'预算情况'按钮");
                     return true;
                 }
                 
                 // 如果主页面没找到，在所有iframe中查找
                 Console.WriteLine("2. 在所有iframe中查找'预算情况'按钮...");
                 var frames = targetPage.Frames;
                 Console.WriteLine($"找到 {frames.Count} 个iframe");
                 
                 for (int i = 0; i < frames.Count; i++)
                 {
                     var frame = frames[i];
                     Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                     
                     try
                     {
                         // 使用多种选择器查找按钮
                         var selectors = new[]
                         {
                             "a:has(span:text('预算情况'))",
                             "a[href*='wgPager_WF_KY_7942d100001_2']",
                             "a[innerwinno='16171d100001']",
                             "text=预算情况"
                         };
                         
                         foreach (var selector in selectors)
                         {
                             try
                             {
                                 var frameLocator = frame.Locator(selector);
                                 var frameCount = await frameLocator.CountAsync();
                                 
                                 if (frameCount > 0)
                                 {
                                     Console.WriteLine($"在iframe {i + 1} 中找到'预算情况'按钮，使用选择器: {selector}");
                                     
                                     // 获取按钮信息用于确认
                                     var buttonText = await frameLocator.TextContentAsync();
                                     var buttonHref = await frameLocator.GetAttributeAsync("href");
                                     Console.WriteLine($"按钮文本: {buttonText?.Trim()}, 链接: {buttonHref}");
                                     
                                     // 点击按钮
                                     await frameLocator.ClickAsync();
                                     Console.WriteLine($"在iframe {i + 1} 中成功点击'预算情况'按钮");
                                     
                                     // 等待点击响应
                                     await Task.Delay(2000);
                                     return true;
                                 }
                             }
                             catch (Exception ex)
                             {
                                 Console.WriteLine($"在iframe {i + 1} 中使用选择器 {selector} 失败: {ex.Message}");
                             }
                         }
                     }
                     catch (Exception ex)
                     {
                         Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                     }
                 }
                 
                 Console.WriteLine("在所有iframe中都未找到'预算情况'按钮");
                 return false;
             }
             catch (Exception ex)
             {
                 Console.WriteLine($"查找并点击'预算情况'按钮时出错: {ex.Message}");
                 return false;
             }
         }
         
         /// <summary>
         /// 自动搜索项目编号
         /// </summary>
         private async Task AutoSearchProjectNumbers(List<string> projectNumbers)
        {
            try
            {
                Console.WriteLine("\n=== 开始自动搜索项目编号 ===");
                
                // 等待页面完全加载
                await Task.Delay(3000);
                
                // 查找搜索输入框 - 先在主页面查找，然后在iframe中查找
                var searchInputLocator = await FindSearchInputBox();
                if (searchInputLocator == null)
                {
                    Console.WriteLine("错误：未找到搜索输入框 #gridWF_KYCW_7203_qspi");
                    return;
                }
                
                Console.WriteLine("找到搜索输入框，开始自动搜索...");
                
                int processedCount = 0;
                int skippedCount = 0;
                
                foreach (var projectNumber in projectNumbers)
                {
                    try
                    {
                        // 检查是否以"无编号"开头
                        if (projectNumber.StartsWith("无编号", StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"跳过项目编号: {projectNumber} (以'无编号'开头)");
                            skippedCount++;
                            continue;
                        }
                        
                        Console.WriteLine($"\n--- 处理项目编号: {projectNumber} ---");
                        
                        // 清空输入框
                        await searchInputLocator.ClearAsync();
                        Console.WriteLine("已清空搜索输入框");
                        
                        // 输入项目编号
                        await searchInputLocator.FillAsync(projectNumber);
                        Console.WriteLine($"已输入项目编号: {projectNumber}");
                        
                        // 等待一下确保输入完成
                        await Task.Delay(500);
                        
                        // 模拟按回车键
                        await searchInputLocator.PressAsync("Enter");
                        Console.WriteLine("已按回车键，开始搜索...");
                        
                                                 // 等待搜索结果加载
                         await Task.Delay(2000);
                         
                         // 等待表格数据加载并点击第一行第二列
                         var tableFound = await WaitForTableAndClickFirstRow();
                         
                         //if (!tableFound)
                         //{
                         //    // 如果未找到表格数据，直接填写"未查询到"
                         //    Console.WriteLine($"项目编号 {projectNumber} 未查询到数据，填写'未查询到'");
                         //    await WriteNotFoundToExcel(projectNumber);
                         //    processedCount++;
                         //    continue;
                         //}
                         
                         // 等待页面加载完成后，在当前页面中查找并点击"预算情况"按钮
                         await Task.Delay(3000); // 等待页面完全加载
                         await ClickBudgetButtonInCurrentPage();
                         
                         // 等待页面加载完成后，在当前页面中查找并点击"当前执行预算"按钮
                         await Task.Delay(3000); // 等待页面完全加载
                         await ClickCurrentExecutionBudgetButton(projectNumber);
                         
                         processedCount++;
                         Console.WriteLine($"项目编号 {projectNumber} 搜索完成");
                         
                         // 在搜索下一个项目前稍作等待
                         await Task.Delay(1000);
                        
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"处理项目编号 {projectNumber} 时出错: {ex.Message}");
                    }
                }
                
                Console.WriteLine($"\n=== 自动搜索完成 ===");
                Console.WriteLine($"成功处理: {processedCount} 个项目编号");
                Console.WriteLine($"跳过处理: {skippedCount} 个项目编号");
                Console.WriteLine($"总计项目: {projectNumbers.Count} 个");
                
            }
            catch (Exception ex)
            {
                Console.WriteLine($"自动搜索项目编号时出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 提取表格数据并写入Excel
        /// </summary>
        private async Task ExtractTableDataAndWriteToExcel(string currentProjectNumber = null)
        {
            try
            {
                Console.WriteLine("开始提取表格数据...");
                
                // 查找表格 - 先在主页面查找，然后在iframe中查找
                var (tableData, projectBalance) = await FindAndExtractTableData();
                if (tableData == null || !tableData.Any())
                {
                    Console.WriteLine("警告：未找到表格数据或表格为空");
                    return;
                }
                
                // 将提取的数据写入Excel
                await WriteTableDataToExcel(tableData, projectBalance, currentProjectNumber);
                
                Console.WriteLine("表格数据提取和写入Excel完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取表格数据时出错: {ex.Message}");
            }
        }
        
        /// <summary>
        /// 查找并提取表格数据
        /// </summary>
        private async Task<(List<(string budgetItem, string balance)> details, string projectBalance)> FindAndExtractTableData()
        {
            try
            {
                Console.WriteLine("查找表格 #gridWF_KY_7944d100001...");
                
                // 首先在主页面查找
                var mainPageTable = page.Locator("#gridWF_KY_7944d100001");
                var mainPageCount = await mainPageTable.CountAsync();
                
                if (mainPageCount > 0)
                {
                    Console.WriteLine("在主页面中找到表格");
                    return await ExtractDataFromTable(mainPageTable);
                }
                
                // 如果主页面没找到，在所有iframe中查找
                Console.WriteLine("在主页面中未找到表格，在所有iframe中查找...");
                var frames = page.Frames;
                Console.WriteLine($"找到 {frames.Count} 个iframe");
                
                for (int i = 0; i < frames.Count; i++)
                {
                    var frame = frames[i];
                    Console.WriteLine($"检查iframe {i + 1}: {frame.Name ?? "未命名"} - {frame.Url}");
                    
                    try
                    {
                        // 使用多种选择器查找表格
                        var selectors = new[]
                        {
                            "#gridWF_KY_7944d100001",
                            "#gview_gridWF_KY_7944d100001",
                            "table[id*='gridWF_KY_7944d100001']"
                        };
                        
                        foreach (var selector in selectors)
                        {
                            try
                            {
                                var frameTable = frame.Locator(selector);
                                var frameCount = await frameTable.CountAsync();
                                
                                if (frameCount > 0)
                                {
                                    Console.WriteLine($"在iframe {i + 1} 中找到表格，使用选择器: {selector}");
                                    return await ExtractDataFromTable(frameTable);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"在iframe {i + 1} 中使用选择器 {selector} 失败: {ex.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                    }
                }
                
                Console.WriteLine("警告：在所有iframe中都未找到表格");
                return (null, "");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查找表格时出错: {ex.Message}");
                return (null, "");
            }
        }
        
        /// <summary>
        /// 从表格中提取数据
        /// </summary>
        private async Task<(List<(string budgetItem, string balance)> details, string projectBalance)> ExtractDataFromTable(ILocator tableLocator)
        {
            try
            {
                Console.WriteLine("开始从表格中提取数据...");
                
                // 等待表格数据加载
                await Task.Delay(2000);
                
                // 查找所有数据行（排除表头）
                var dataRows = tableLocator.Locator("tbody tr:not(.jqgfirstrow)");
                var rowCount = await dataRows.CountAsync();
                
                Console.WriteLine($"找到 {rowCount} 行数据");
                
                if (rowCount == 0)
                {
                    Console.WriteLine("表格中没有数据行");
                    return (new List<(string budgetItem, string balance)>(), "");
                }
                
                var extractedData = new List<(string budgetItem, string balance)>();
                string projectBalance = "";
                
                // 遍历每一行数据
                for (int i = 0; i < rowCount; i++)
                {
                    try
                    {
                        var row = dataRows.Nth(i);
                        
                        // 查找预算项列和余额列
                        // 根据用户提供的HTML，预算项是第二列，余额是第十一列
                        var budgetItemCell = row.Locator("td:nth-child(2)");
                        var balanceCell = row.Locator("td:nth-child(11)");
                        
                        var budgetItem = await budgetItemCell.TextContentAsync();
                        var balance = await balanceCell.TextContentAsync();
                        
                        if (!string.IsNullOrWhiteSpace(budgetItem) && !string.IsNullOrWhiteSpace(balance))
                        {
                            budgetItem = budgetItem.Trim();
                            balance = balance.Trim();
                            
                            // 对于第一行数据，直接作为项目余额
                            if (i == 0)
                            {
                                projectBalance = balance;
                                Console.WriteLine($"第一行余额作为项目余额: {projectBalance}");
                            }
                            
                            // 只提取余额不为0的项目
                            if (balance != "0" && balance != "0.00" && balance != "0.0")
                            {
                                extractedData.Add((budgetItem, balance));
                                Console.WriteLine($"提取数据: 预算项={budgetItem}, 余额={balance}");
                            }
                            else
                            {
                                Console.WriteLine($"跳过余额为0的项目: 预算项={budgetItem}, 余额={balance}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"提取第 {i + 1} 行数据时出错: {ex.Message}");
                    }
                }
                
                Console.WriteLine($"成功提取 {extractedData.Count} 条明细数据");
                if (!string.IsNullOrEmpty(projectBalance))
                {
                    Console.WriteLine($"项目余额: {projectBalance}");
                }
                
                return (extractedData, projectBalance);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"从表格中提取数据时出错: {ex.Message}");
                return (new List<(string budgetItem, string balance)>(), "");
            }
        }
        
        /// <summary>
        /// 将表格数据写入Excel
        /// </summary>
        private async Task WriteTableDataToExcel(List<(string budgetItem, string balance)> tableData, string projectBalance, string currentProjectNumber = null)
        {
            try
            {
                Console.WriteLine("开始将表格数据写入Excel...");
                
                // 读取配置文件
                var configPath = FindConfigFile();
                if (string.IsNullOrEmpty(configPath))
                {
                    Console.WriteLine("错误：未找到配置文件");
                    return;
                }
                
                var configContent = await File.ReadAllTextAsync(configPath);
                var jsonOptions = new System.Text.Json.JsonSerializerOptions
                {
                    PropertyNameCaseInsensitive = true
                };
                var config = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(configContent, jsonOptions);
                
                if (config == null || !config.ContainsKey("ExcelFilePath"))
                {
                    Console.WriteLine("错误：配置文件中未找到ExcelFilePath");
                    return;
                }
                
                var excelFilePath = config["ExcelFilePath"]?.ToString();
                if (string.IsNullOrEmpty(excelFilePath))
                {
                    Console.WriteLine("错误：ExcelFilePath为空");
                    return;
                }
                
                // 查找Excel文件
                var actualExcelPath = FindExcelFile(excelFilePath);
                if (string.IsNullOrEmpty(actualExcelPath))
                {
                    Console.WriteLine($"错误：Excel文件不存在: {excelFilePath}");
                    return;
                }
                
                Console.WriteLine($"找到Excel文件: {actualExcelPath}");
                
                // 打开Excel文件
                using (var package = new ExcelPackage(new FileInfo(actualExcelPath)))
                {
                    var worksheet = package.Workbook.Worksheets["0-Proj信息"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("错误：未找到工作表 '0-Proj信息'");
                        return;
                    }
                    
                                         // 查找"项目预算执行情况"列
                     var budgetColumnIndex = -1;
                     var balanceColumnIndex = -1;
                     var headers = worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column];
                     
                     for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                     {
                         var headerValue = worksheet.Cells[1, col].Value?.ToString();
                         if (headerValue == "项目预算执行情况")
                         {
                             budgetColumnIndex = col;
                         }
                         else if (headerValue == "项目余额")
                         {
                             balanceColumnIndex = col;
                         }
                     }
                     
                     if (budgetColumnIndex == -1)
                     {
                         Console.WriteLine("错误：未找到'项目预算执行情况'列");
                         return;
                     }
                     
                     if (balanceColumnIndex == -1)
                     {
                         Console.WriteLine("警告：未找到'项目余额'列，将不会填写余额信息");
                     }
                     
                     Console.WriteLine($"找到'项目预算执行情况'列，列索引: {budgetColumnIndex}");
                     if (balanceColumnIndex != -1)
                     {
                         Console.WriteLine($"找到'项目余额'列，列索引: {balanceColumnIndex}");
                     }
                    
                    // 查找当前项目编号对应的行
                    var targetRow = -1;
                    
                    if (!string.IsNullOrEmpty(currentProjectNumber))
                    {
                        // 查找"项目编号"列
                        var projectNumberColumnIndex = -1;
                        for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var headerValue = worksheet.Cells[1, col].Value?.ToString();
                            if (headerValue == "项目编号")
                            {
                                projectNumberColumnIndex = col;
                                break;
                            }
                        }
                        
                        if (projectNumberColumnIndex != -1)
                        {
                            // 在项目编号列中查找当前项目编号
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var cellValue = worksheet.Cells[row, projectNumberColumnIndex].Value?.ToString();
                                if (cellValue == currentProjectNumber)
                                {
                                    targetRow = row;
                                    Console.WriteLine($"找到项目编号 {currentProjectNumber} 对应的行: {row}");
                                    break;
                                }
                            }
                        }
                    }
                    
                                         // 如果没找到对应行，报告错误并跳过
                     if (targetRow == -1)
                     {
                         Console.WriteLine($"错误：未找到项目编号 {currentProjectNumber} 对应的行，跳过数据写入");
                         return;
                     }
                    
                                                              // 格式化数据为"预算项: 余额"的形式，每个项目后面加换行符
                     var formattedData = string.Join("; \n", tableData.Select(item => $"{item.budgetItem}: {item.balance}"));
                     
                     // 写入预算执行情况数据
                     worksheet.Cells[targetRow, budgetColumnIndex].Value = formattedData;
                     Console.WriteLine($"成功将预算执行情况数据写入Excel第 {targetRow} 行，列 {budgetColumnIndex}");
                     Console.WriteLine($"写入的数据: {formattedData}");
                     
                                           // 如果有项目余额列，将余额信息也写入
                      if (balanceColumnIndex != -1)
                      {
                          // 优先使用从表格中提取的项目余额
                          if (!string.IsNullOrEmpty(projectBalance))
                          {
                              // 直接使用表格中的项目余额
                              if (decimal.TryParse(projectBalance, out decimal extractedBalance))
                              {
                                  worksheet.Cells[targetRow, balanceColumnIndex].Value = extractedBalance;
                                  Console.WriteLine($"成功将表格中的项目余额写入Excel第 {targetRow} 行，列 {balanceColumnIndex}");
                                  Console.WriteLine($"写入的余额: {extractedBalance}");
                              }
                              else
                              {
                                  Console.WriteLine($"警告：无法解析表格中的项目余额: {projectBalance}");
                                  // 如果解析失败，回退到计算方式
                                  await WriteCalculatedBalance(worksheet, targetRow, balanceColumnIndex, tableData);
                              }
                          }
                          else
                          {
                              // 如果没有找到项目余额，使用计算方式
                              Console.WriteLine("未找到表格中的项目余额，使用计算方式");
                              await WriteCalculatedBalance(worksheet, targetRow, balanceColumnIndex, tableData);
                          }
                      }
                     
                                           // 保存Excel文件
                      package.Save();
                  }
              }
              catch (Exception ex)
              {
                  Console.WriteLine($"将表格数据写入Excel时出错: {ex.Message}");
              }
          }
          
                     /// <summary>
           /// 写入计算得出的项目余额
           /// </summary>
           private async Task WriteCalculatedBalance(ExcelWorksheet worksheet, int targetRow, int balanceColumnIndex, List<(string budgetItem, string balance)> tableData)
           {
               try
               {
                   // 计算总余额（所有余额的总和）
                   decimal totalBalance = 0;
                   foreach (var item in tableData)
                   {
                       if (decimal.TryParse(item.balance, out decimal balance))
                       {
                           totalBalance += balance;
                       }
                   }
                   
                   // 写入计算得出的总余额到项目余额列
                   worksheet.Cells[targetRow, balanceColumnIndex].Value = totalBalance;
                   Console.WriteLine($"成功将计算得出的项目余额写入Excel第 {targetRow} 行，列 {balanceColumnIndex}");
                   Console.WriteLine($"写入的余额: {totalBalance}");
               }
               catch (Exception ex)
               {
                   Console.WriteLine($"写入计算得出的项目余额时出错: {ex.Message}");
               }
           }
           
           /// <summary>
           /// 对于未查询到的项目，在Excel中填写"未查询到"
           /// </summary>
           private async Task WriteNotFoundToExcel(string projectNumber)
           {
               try
               {
                   Console.WriteLine($"开始为项目编号 {projectNumber} 填写'未查询到'...");
                   
                   // 读取配置文件
                   var configPath = FindConfigFile();
                   if (string.IsNullOrEmpty(configPath))
                   {
                       Console.WriteLine("错误：未找到配置文件");
                       return;
                   }
                   
                   var configContent = await File.ReadAllTextAsync(configPath);
                   var jsonOptions = new System.Text.Json.JsonSerializerOptions
                   {
                       PropertyNameCaseInsensitive = true
                   };
                   var config = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, object>>(configContent, jsonOptions);
                   
                   if (config == null || !config.ContainsKey("ExcelFilePath"))
                   {
                       Console.WriteLine("错误：配置文件中未找到ExcelFilePath");
                       return;
                   }
                   
                   var excelFilePath = config["ExcelFilePath"]?.ToString();
                   if (string.IsNullOrEmpty(excelFilePath))
                   {
                       Console.WriteLine("错误：ExcelFilePath为空");
                       return;
                   }
                   
                   // 查找Excel文件
                   var actualExcelPath = FindExcelFile(excelFilePath);
                   if (string.IsNullOrEmpty(actualExcelPath))
                   {
                       Console.WriteLine($"错误：Excel文件不存在: {excelFilePath}");
                       return;
                   }
                   
                   Console.WriteLine($"找到Excel文件: {actualExcelPath}");
                   
                   // 打开Excel文件
                   using (var package = new ExcelPackage(new FileInfo(actualExcelPath)))
                   {
                       var worksheet = package.Workbook.Worksheets["0-Proj信息"];
                       if (worksheet == null)
                       {
                           Console.WriteLine("错误：未找到工作表 '0-Proj信息'");
                           return;
                       }
                       
                       // 查找"项目预算执行情况"列
                       var budgetColumnIndex = -1;
                       for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                       {
                           var headerValue = worksheet.Cells[1, col].Value?.ToString();
                           if (headerValue == "项目预算执行情况")
                           {
                               budgetColumnIndex = col;
                               break;
                           }
                       }
                       
                       if (budgetColumnIndex == -1)
                       {
                           Console.WriteLine("错误：未找到'项目预算执行情况'列");
                           return;
                       }
                       
                       // 查找当前项目编号对应的行
                       var targetRow = -1;
                       
                       // 查找"项目编号"列
                       var projectNumberColumnIndex = -1;
                       for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                       {
                           var headerValue = worksheet.Cells[1, col].Value?.ToString();
                           if (headerValue == "项目编号")
                           {
                               projectNumberColumnIndex = col;
                               break;
                           }
                       }
                       
                       if (projectNumberColumnIndex != -1)
                       {
                           // 在项目编号列中查找当前项目编号
                           for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                           {
                               var cellValue = worksheet.Cells[row, projectNumberColumnIndex].Value?.ToString();
                               if (cellValue == projectNumber)
                               {
                                   targetRow = row;
                                   Console.WriteLine($"找到项目编号 {projectNumber} 对应的行: {row}");
                                   break;
                               }
                           }
                       }
                       
                       // 如果没找到对应行，报告错误并跳过
                       if (targetRow == -1)
                       {
                           Console.WriteLine($"错误：未找到项目编号 {projectNumber} 对应的行，跳过数据写入");
                           return;
                       }
                       
                       // 写入"未查询到"
                       worksheet.Cells[targetRow, budgetColumnIndex].Value = "未查询到";
                       Console.WriteLine($"成功将'未查询到'写入Excel第 {targetRow} 行，列 {budgetColumnIndex}");
                       
                       // 保存Excel文件
                       package.Save();
                       Console.WriteLine($"项目编号 {projectNumber} 的'未查询到'状态已保存到Excel");
                   }
               }
               catch (Exception ex)
               {
                   Console.WriteLine($"为项目编号 {projectNumber} 填写'未查询到'时出错: {ex.Message}");
               }
           }
     }
}
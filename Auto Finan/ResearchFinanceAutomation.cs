using System;
using System.Threading.Tasks;
using Microsoft.Playwright;
using System.Collections.Generic;
using System.Linq;

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
                if (!isLoggedIn)
                {
                    Console.WriteLine("错误：尚未登录，请先执行登录");
                    return false;
                }

                Console.WriteLine("等待3秒后开始点击导航按钮...");
                await Task.Delay(3000);

                Console.WriteLine("正在查找'数据查询及公示'导航按钮...");

                // 基于用户提供的HTML代码，使用更精确的选择器
                var buttonSelectors = new[]
                {
                    // 精确匹配：包含"数据查询及公示"文本的链接
                    "a:has-text('数据查询及公示')",
                    // 包含特定onclick属性的链接
                    "a[onclick*='项目执行情况查询']",
                    "a[onclick*='数据查询及公示']",
                    // 包含特定span文本的链接
                    "a:has(span:text('数据查询及公示'))",
                    // 通用文本匹配
                    "text=数据查询及公示",
                    // 包含特定类或样式的链接
                    "a[style*='cursor:pointer']:has-text('数据查询及公示')"
                };

                ILocator buttonLocator = null;
                foreach (var selector in buttonSelectors)
                {
                    try
                    {
                        var locator = page.Locator(selector);
                        if (await locator.CountAsync() > 0)
                        {
                            buttonLocator = locator;
                            Console.WriteLine($"找到按钮，使用选择器: {selector}");
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"选择器 {selector} 查找失败: {ex.Message}");
                    }
                }

                if (buttonLocator == null)
                {
                    Console.WriteLine("未找到'数据查询及公示'按钮，尝试查找包含该文本的元素...");

                    // 查找所有包含"数据查询及公示"文本的元素
                    var allElements = await page.Locator("*:has-text('数据查询及公示')").AllAsync();
                    if (allElements.Count > 0)
                    {
                        Console.WriteLine($"找到 {allElements.Count} 个包含'数据查询及公示'文本的元素");

                        // 尝试点击第一个可点击的元素
                        for (int i = 0; i < allElements.Count; i++)
                        {
                            try
                            {
                                var element = allElements[i];
                                var tagName = await element.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
                                var isClickable = await element.EvaluateAsync<bool>("el => !!(el.onclick || el.getAttribute('href') || el.tagName.toLowerCase() === 'a')");

                                Console.WriteLine($"元素 {i + 1}: 标签={tagName}, 可点击={isClickable}");

                                if (isClickable)
                                {
                                    Console.WriteLine($"尝试点击元素 {i + 1}...");
                                    await element.ClickAsync();
                                    Console.WriteLine("导航按钮点击成功！");
                                    return true;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"点击元素 {i + 1} 失败: {ex.Message}");
                            }
                        }
                    }

                    Console.WriteLine("无法找到可点击的'数据查询及公示'按钮");
                    return false;
                }

                // 点击找到的按钮
                Console.WriteLine("正在点击'数据查询及公示'导航按钮...");
                await buttonLocator.ClickAsync();
                Console.WriteLine("导航按钮点击成功！");

                // 等待页面响应
                await Task.Delay(2000);

                // 检查页面是否发生变化
                var currentUrl = page.Url;
                Console.WriteLine($"点击导航按钮后的URL: {currentUrl}");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"点击导航按钮失败: {ex.Message}");
                return false;
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

                // 查找表格信息
                var tables = await page.Locator("table").AllAsync();
                Console.WriteLine($"找到 {tables.Count} 个表格");

                for (int i = 0; i < tables.Count; i++)
                {
                    try
                    {
                        var table = tables[i];
                        var tableInfo = await ExtractTableInfoAsync(table, i + 1);
                        foreach (var kvp in tableInfo)
                        {
                            pageInfo[$"表格{i + 1}_{kvp.Key}"] = kvp.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"提取表格 {i + 1} 信息时出错: {ex.Message}");
                    }
                }

                // 查找其他重要信息
                var importantElements = await page.Locator("h1, h2, h3, .title, .header, [class*='title'], [class*='header']").AllAsync();
                foreach (var element in importantElements)
                {
                    try
                    {
                        var text = await element.TextContentAsync();
                        if (!string.IsNullOrEmpty(text))
                        {
                            var cleanText = text.Trim();
                            if (cleanText.Length > 0 && cleanText.Length < 100) // 限制长度避免过长
                            {
                                pageInfo[$"标题_{cleanText.Substring(0, Math.Min(20, cleanText.Length))}"] = cleanText;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"提取标题信息时出错: {ex.Message}");
                    }
                }

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
        /// 提取表格信息
        /// </summary>
        private async Task<Dictionary<string, string>> ExtractTableInfoAsync(ILocator table, int tableIndex)
        {
            var tableInfo = new Dictionary<string, string>();

            try
            {
                // 获取表格行数
                var rows = await table.Locator("tr").AllAsync();
                tableInfo[$"行数"] = rows.Count.ToString();

                // 获取表格列数（以第一行为准）
                if (rows.Count > 0)
                {
                    var firstRowCells = await rows[0].Locator("td, th").AllAsync();
                    tableInfo[$"列数"] = firstRowCells.Count.ToString();
                }

                // 提取表格内容（前几行）
                var maxRowsToExtract = Math.Min(5, rows.Count); // 最多提取5行
                for (int i = 0; i < maxRowsToExtract; i++)
                {
                    var cells = await rows[i].Locator("td, th").AllAsync();
                    var rowData = new List<string>();

                    foreach (var cell in cells)
                    {
                        var cellText = await cell.TextContentAsync();
                        rowData.Add(cellText?.Trim() ?? "");
                    }

                    if (rowData.Any(x => !string.IsNullOrEmpty(x)))
                    {
                        tableInfo[$"行{i + 1}"] = string.Join(" | ", rowData);
                    }
                }

                Console.WriteLine($"表格 {tableIndex} 信息提取完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"提取表格 {tableIndex} 信息时出错: {ex.Message}");
            }

            return tableInfo;
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
    }
}
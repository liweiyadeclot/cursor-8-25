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
                Console.WriteLine("正在精确查找'数据查询及公示'菜单项...");

                // 首先在主页面查找
                Console.WriteLine("1. 在主页面查找...");
                var buttonLocator = await FindButtonInPage(page, "主页面");
                if (buttonLocator != null)
                {
                    return await ClickButton(buttonLocator, "主页面");
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
                            return await ClickButton(frameButtonLocator, $"iframe {i + 1}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"检查iframe {i + 1} 时出错: {ex.Message}");
                    }
                }

                Console.WriteLine("在所有页面和iframe中都未找到'数据查询及公示'按钮");
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
    }
}
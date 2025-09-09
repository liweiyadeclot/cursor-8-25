using System;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Playwright;
using OfficeOpenXml;
using System.IO;
using System.Collections.Generic;
using System.Linq;

namespace AutoFinan
{
    public class SubjectBalanceQueryAutomation
    {
        private IPlaywright playwright;
        private IBrowser browser;
        private IPage page;

        public async Task RunAsync()
        {
            try
            {
                await InitializeBrowser();
                await NavigateToTargetPage();
                await LoginWithUserInputAsync();

                // 直接点击网上预约报账导航
                await ClickNavigationButton();

                // 点击申请报销单按钮
                await ClickApplyReimbursementButton();

                // 点击已阅读并同意按钮
                await ClickAgreeButton();

                // 处理科目余额查询
                await ProcessSubjectBalanceQuery();
            }
            catch (TimeoutException ex)
            {
                Console.WriteLine($"查询项目科目余额流程超时: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"查询项目科目余额流程出错: {ex.Message}");
            }
            finally
            {
                // 开发阶段保持浏览器打开，不关闭
                Console.WriteLine("开发阶段：浏览器保持打开状态");
                // try { if (browser != null) await browser.CloseAsync(); } catch { }
                // try { if (playwright != null) playwright.Dispose(); } catch { }
            }
        }

        private async Task InitializeBrowser()
        {
            Console.WriteLine("正在启动浏览器...");
            playwright = await Playwright.CreateAsync();
            browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
            {
                Headless = false,
                SlowMo = 100
            });
            page = await browser.NewPageAsync();
            page.SetDefaultTimeout(10000);
            Console.WriteLine("浏览器启动成功");
        }

        private async Task NavigateToTargetPage()
        {
            Console.WriteLine("正在导航到目标网页...");
            string targetUrl = "https://cwcx.uestc.edu.cn/WFManager/home.jsp";
            await page.GotoAsync(targetUrl, new PageGotoOptions { Timeout = 30000 });
            await page.WaitForLoadStateAsync(LoadState.NetworkIdle, new PageWaitForLoadStateOptions { Timeout = 30000 });
            Console.WriteLine($"成功导航到页面: {targetUrl}");
        }

        private async Task LoginWithUserInputAsync()
        {
            // 开发阶段使用固定账号密码
            string username = "5130008";
            string password = "Uestc418";
            
            Console.WriteLine("开发阶段使用固定账号密码：");
            Console.WriteLine($"用户名: {username}");
            Console.WriteLine($"密码: {new string('*', password.Length)}");
            Console.Write("验证码: ");
            string captcha = Console.ReadLine()?.Trim() ?? string.Empty;

            bool filled = false;
            var frames = page.Frames;

            var userSelectors = new[] {
                "#txtUserName", "input[name='username']", "#username", "#zh", "input[name='zh']",
                "#userid", "input[name='userid']", "#gh", "input[name='gh']",
                "input[placeholder*='工号']", "input[placeholder*='用户名']", "input[aria-label*='工号']", "input[aria-label*='用户名']"
            };
            var passSelectors = new[] {
                "#txtPassword", "input[name='password']", "#password", "#mm", "input[name='mm']",
                "#pwd", "input[name='pwd']",
                "input[placeholder*='密码']", "input[aria-label*='密码']"
            };
            var codeSelectors = new[] {
                "#txtValidateCode", "#validateCode", "input[name='validateCode']", "#yzm", "input[name='yzm']",
                "#code", "input[name='code']",
                "input[placeholder*='验证码']", "input[aria-label*='验证码']"
            };
            var loginBtnSelectors = new[] { "#zhLogin:not([disabled])", "#zhLogin", "button:has-text('登录')", "input[type='submit']" };

            async Task<bool> TryFillInPage(IPage anyPage)
            {
                try
                {
                    foreach (var s in userSelectors)
                    {
                        var loc = anyPage.Locator(s).First;
                        if (await loc.CountAsync() > 0) { await loc.ClickAsync(); await loc.FillAsync(""); await loc.FillAsync(username); Console.WriteLine($"已填写工号/用户名: 选择器 {s}"); break; }
                    }
                    foreach (var s in passSelectors)
                    {
                        var loc = anyPage.Locator(s).First;
                        if (await loc.CountAsync() > 0) { await loc.ClickAsync(); await loc.FillAsync(""); await loc.FillAsync(password); Console.WriteLine($"已填写密码: 选择器 {s}"); break; }
                    }
                    foreach (var s in codeSelectors)
                    {
                        var loc = anyPage.Locator(s).First;
                        if (await loc.CountAsync() > 0) { await loc.ClickAsync(); await loc.FillAsync(""); await loc.FillAsync(captcha); Console.WriteLine($"已填写验证码: 选择器 {s}"); break; }
                    }

                    foreach (var s in loginBtnSelectors)
                    {
                        var loc = anyPage.Locator(s).First;
                        if (await loc.CountAsync() > 0)
                        {
                            await loc.ClickAsync();
                            Console.WriteLine($"已点击登录按钮: 选择器 {s}");
                            try { await anyPage.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 10000 }); } catch { }
                            return true;
                        }
                    }

                    // 如果没有找到登录按钮，尝试在密码框回车提交
                    foreach (var s in passSelectors)
                    {
                        var loc = anyPage.Locator(s).First;
                        if (await loc.CountAsync() > 0)
                        {
                            await loc.PressAsync("Enter");
                            Console.WriteLine("未找到登录按钮，已在密码框按下回车尝试提交");
                            try { await anyPage.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 10000 }); } catch { }
                            return true;
                        }
                    }
                }
                catch { }
                return false;
            }

            async Task<bool> TryFillInFrame(IFrame anyFrame)
            {
                try
                {
                    foreach (var s in userSelectors)
                    {
                        var loc = anyFrame.Locator(s).First;
                        if (await loc.CountAsync() > 0) { await loc.ClickAsync(); await loc.FillAsync(""); await loc.FillAsync(username); Console.WriteLine($"已填写工号/用户名(iframe): 选择器 {s}"); break; }
                    }
                    foreach (var s in passSelectors)
                    {
                        var loc = anyFrame.Locator(s).First;
                        if (await loc.CountAsync() > 0) { await loc.ClickAsync(); await loc.FillAsync(""); await loc.FillAsync(password); Console.WriteLine($"已填写密码(iframe): 选择器 {s}"); break; }
                    }
                    foreach (var s in codeSelectors)
                    {
                        var loc = anyFrame.Locator(s).First;
                        if (await loc.CountAsync() > 0) { await loc.ClickAsync(); await loc.FillAsync(""); await loc.FillAsync(captcha); Console.WriteLine($"已填写验证码(iframe): 选择器 {s}"); break; }
                    }

                    foreach (var s in loginBtnSelectors)
                    {
                        var loc = anyFrame.Locator(s).First;
                        if (await loc.CountAsync() > 0)
                        {
                            await loc.ClickAsync();
                            Console.WriteLine($"已点击登录按钮(iframe): 选择器 {s}");
                            try { await anyFrame.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 10000 }); } catch { }
                            return true;
                        }
                    }

                    // 如果没有找到登录按钮，尝试在密码框回车提交
                    foreach (var s in passSelectors)
                    {
                        var loc = anyFrame.Locator(s).First;
                        if (await loc.CountAsync() > 0)
                        {
                            await loc.PressAsync("Enter");
                            Console.WriteLine("未找到登录按钮(iframe)，已在密码框按下回车尝试提交");
                            try { await anyFrame.WaitForLoadStateAsync(LoadState.NetworkIdle, new() { Timeout = 10000 }); } catch { }
                            return true;
                        }
                    }
                }
                catch { }
                return false;
            }

            filled = await TryFillInPage(page);
            if (!filled)
            {
                foreach (var f in frames)
                {
                    if (await TryFillInFrame(f)) { filled = true; break; }
                }
            }
            if (!filled)
            {
                throw new Exception("未找到登录输入框或登录按钮");
            }

            var loginOk = await WaitForLoginSuccessAsync(5, verbose: false);
            if (!loginOk)
            {
                Console.WriteLine("警告：登录成功确认未命中，但将继续尝试后续按钮点击");
            }
        }

        private async Task<bool> WaitForLoginSuccessAsync(int timeoutSeconds = 60, bool verbose = true)
        {
            var timeout = TimeSpan.FromSeconds(timeoutSeconds);
            var start = DateTime.Now;
            int attempt = 0;
            while (DateTime.Now - start < timeout)
            {
                try
                {
                    attempt++;
                    if (verbose) Console.WriteLine($"登录成功判定第 {attempt} 次检查...");
                    // 处理可能的layui弹窗
                    var layerOk = page.Locator(".layui-layer-btn0").First;
                    if (await layerOk.CountAsync() > 0)
                    {
                        try { await layerOk.ClickAsync(); if (verbose) Console.WriteLine("已点击layui弹窗确定"); } catch { }
                    }

                    if (await page.Locator("#spUsername").CountAsync() > 0) return true;
                    if (await page.Locator("text=网上预约报账").CountAsync() > 0) return true;
                    if (await page.Locator("[id*='yybz']").CountAsync() > 0) return true;

                    foreach (var f in page.Frames)
                    {
                        if (await f.Locator("#spUsername").CountAsync() > 0) return true;
                        if (await f.Locator("text=网上预约报账").CountAsync() > 0) return true;
                        if (await f.Locator("[id*='yybz']").CountAsync() > 0) return true;
                    }
                }
                catch { }
                await Task.Delay(1000);
            }
            return false;
        }

        private string ReadPassword()
        {
            var sb = new StringBuilder();
            while (true)
            {
                var key = Console.ReadKey(true);
                if (key.Key == ConsoleKey.Enter) break;
                if (key.Key == ConsoleKey.Backspace)
                {
                    if (sb.Length > 0) sb.Length--;
                    continue;
                }
                if (!char.IsControl(key.KeyChar)) sb.Append(key.KeyChar);
            }
            Console.WriteLine();
            return sb.ToString();
        }

        /// <summary>
        /// 直接调用 Program.cs 中的 ClickNavigationButton 逻辑
        /// </summary>
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
                        Console.WriteLine("      点击按钮: 网上预约报账按钮 -> navToPrj('WF_YB6')");
                        Console.WriteLine("      检测到JavaScript函数调用: navToPrj('WF_YB6')");
                        Console.WriteLine("      成功执行JavaScript函数: navToPrj('WF_YB6')");
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
                    Console.WriteLine("      点击按钮: 网上预约报账按钮 -> navToPrj('WF_YB6')");
                    Console.WriteLine("      检测到JavaScript函数调用: navToPrj('WF_YB6')");
                    Console.WriteLine("      成功执行JavaScript函数: navToPrj('WF_YB6')");
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
                        Console.WriteLine("      点击按钮: 网上预约报账按钮 -> navToPrj('WF_YB6')");
                        Console.WriteLine("      检测到JavaScript函数调用: navToPrj('WF_YB6')");
                        Console.WriteLine("      成功执行JavaScript函数: navToPrj('WF_YB6')");
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

        /// <summary>
        /// 点击申请报销单按钮
        /// </summary>
        private async Task ClickApplyReimbursementButton()
        {
            try
            {
                Console.WriteLine("      开始查找申请报销单按钮...");

                // 使用btnname属性查找按钮
                string btnname = "申请报销单";
                bool clicked = false;

                // 设置超时时间（1分钟）
                var startTime = DateTime.Now;
                var timeout = TimeSpan.FromMinutes(1);
                int attemptCount = 0;

                while (!clicked && DateTime.Now - startTime < timeout)
                {
                    attemptCount++;
                    Console.WriteLine($"      尝试第 {attemptCount} 次查找申请报销单按钮...");

                    // 等待页面完全加载
                    await Task.Delay(2000);

                // 方法1: 优先在iframe中通过btnname属性查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var buttonElement = frame.Locator($"button[btnname='{btnname}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      在iframe中通过btnname成功点击申请报销单按钮: {btnname}");
                            clicked = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中通过btnname查找申请报销单按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 在主页面通过btnname属性查找
                if (!clicked)
                {
                    try
                    {
                        var buttonElement = page.Locator($"button[btnname='{btnname}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      在主页面通过btnname成功点击申请报销单按钮: {btnname}");
                            clicked = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面通过btnname查找申请报销单按钮失败: {ex.Message}");
                    }
                }

                // 方法3: 尝试其他选择器
                if (!clicked)
                {
                    string[] alternativeSelectors = {
                        $"button:has-text('{btnname}')",
                        $"input[type='button'][value='{btnname}']",
                        $"input[type='submit'][value='{btnname}']",
                        $"a:has-text('{btnname}')"
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
                                        Console.WriteLine($"      在iframe中使用备用选择器成功点击申请报销单按钮: {selector}");
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
                                    Console.WriteLine($"      在主页面使用备用选择器成功点击申请报销单按钮: {selector}");
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

                    // 如果找到了按钮，跳出循环
                    if (clicked)
                    {
                        Console.WriteLine($"      申请报销单按钮点击成功，共尝试了 {attemptCount} 次");
                        break;
                    }

                    // 如果没找到，等待5秒后重试
                    Console.WriteLine($"      第 {attemptCount} 次未找到申请报销单按钮，5秒后重试...");
                    await Task.Delay(5000);
                }

                if (!clicked)
                {
                    Console.WriteLine($"      超时：在 {timeout.TotalMinutes} 分钟内无法找到申请报销单按钮");
                    throw new TimeoutException($"无法找到申请报销单按钮，共尝试了 {attemptCount} 次");
                }

                // 等待按钮点击后的页面加载
                await Task.Delay(2000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击申请报销单按钮失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 点击已阅读并同意按钮
        /// </summary>
        private async Task ClickAgreeButton()
        {
            try
            {
                Console.WriteLine("      开始查找已阅读并同意按钮...");

                // 使用btnname属性查找按钮
                string btnname = "已阅读并同意";
                bool clicked = false;

                // 设置超时时间（1分钟）
                var startTime = DateTime.Now;
                var timeout = TimeSpan.FromMinutes(1);
                int attemptCount = 0;

                while (!clicked && DateTime.Now - startTime < timeout)
                {
                    attemptCount++;
                    Console.WriteLine($"      尝试第 {attemptCount} 次查找已阅读并同意按钮...");

                    // 等待页面完全加载
                    await Task.Delay(2000);

                // 方法1: 优先在iframe中通过btnname属性查找
                var frames = page.Frames;
                foreach (var frame in frames)
                {
                    try
                    {
                        var buttonElement = frame.Locator($"button[btnname='{btnname}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      在iframe中通过btnname成功点击已阅读并同意按钮: {btnname}");
                            clicked = true;
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe中通过btnname查找已阅读并同意按钮失败: {ex.Message}");
                        continue;
                    }
                }

                // 方法2: 在主页面通过btnname属性查找
                if (!clicked)
                {
                    try
                    {
                        var buttonElement = page.Locator($"button[btnname='{btnname}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      在主页面通过btnname成功点击已阅读并同意按钮: {btnname}");
                            clicked = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在主页面通过btnname查找已阅读并同意按钮失败: {ex.Message}");
                    }
                }

                // 方法3: 尝试其他选择器
                if (!clicked)
                {
                    string[] alternativeSelectors = {
                        $"button:has-text('{btnname}')",
                        $"input[type='button'][value='{btnname}']",
                        $"input[type='submit'][value='{btnname}']",
                        $"a:has-text('{btnname}')"
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
                                        Console.WriteLine($"      在iframe中使用备用选择器成功点击已阅读并同意按钮: {selector}");
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
                                    Console.WriteLine($"      在主页面使用备用选择器成功点击已阅读并同意按钮: {selector}");
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

                    // 如果找到了按钮，跳出循环
                    if (clicked)
                    {
                        Console.WriteLine($"      已阅读并同意按钮点击成功，共尝试了 {attemptCount} 次");
                        break;
                    }

                    // 如果没找到，等待5秒后重试
                    Console.WriteLine($"      第 {attemptCount} 次未找到已阅读并同意按钮，5秒后重试...");
                    await Task.Delay(5000);
                }

                if (!clicked)
                {
                    Console.WriteLine($"      超时：在 {timeout.TotalMinutes} 分钟内无法找到已阅读并同意按钮");
                    throw new TimeoutException($"无法找到已阅读并同意按钮，共尝试了 {attemptCount} 次");
                }

                // 等待按钮点击后的页面加载
                await Task.Delay(2000);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击已阅读并同意按钮失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 处理科目余额查询
        /// </summary>
        private async Task ProcessSubjectBalanceQuery()
        {
            try
            {
                Console.WriteLine("      开始处理科目余额查询...");

                // 获取当前登录用户名
                string currentUsername = await GetCurrentUsername();
                if (string.IsNullOrEmpty(currentUsername))
                {
                    Console.WriteLine("      警告：无法获取当前登录用户名");
                    return;
                }

                Console.WriteLine($"      当前登录用户: {currentUsername}");

                // 读取Excel文件中的0-科目余额sheet
                var projectList = ReadSubjectBalanceSheet();
                if (projectList == null || projectList.Count == 0)
                {
                    Console.WriteLine("      警告：未找到有效的项目数据");
                    return;
                }

                Console.WriteLine($"      找到 {projectList.Count} 个项目记录");

                // 检查是否有任何项目有负责人信息
                bool hasResponsiblePerson = projectList.Any(p => !string.IsNullOrEmpty(p.ResponsiblePerson));
                
                if (!hasResponsiblePerson)
                {
                    Console.WriteLine("      警告：所有项目的负责人列都为空");
                    Console.WriteLine("      请选择处理方式：");
                    Console.WriteLine("      1. 处理所有项目（忽略负责人匹配）");
                    Console.WriteLine("      2. 跳过所有项目");
                    Console.Write("      请输入选择 (1 或 2): ");
                    
                    string choice = Console.ReadLine()?.Trim();
                    if (choice != "1")
                    {
                        Console.WriteLine("      跳过所有项目");
                        return;
                    }
                    
                    Console.WriteLine("      将处理所有项目，忽略负责人匹配");
                }

                // 遍历每个项目
                for (int i = 0; i < projectList.Count; i++)
                {
                    var project = projectList[i];
                    Console.WriteLine($"      处理项目 {i + 1}/{projectList.Count}: {project.ProjectNumber} (负责人: {project.ResponsiblePerson})");

                    // 检查负责人是否匹配当前用户（如果负责人列为空且用户选择处理所有项目，则跳过匹配检查）
                    if (hasResponsiblePerson && !IsUserMatched(project.ResponsiblePerson, currentUsername))
                    {
                        Console.WriteLine($"      跳过项目 {project.ProjectNumber}，负责人不匹配：{project.ResponsiblePerson} != {currentUsername}");
                        continue;
                    }

                    if (hasResponsiblePerson)
                    {
                        Console.WriteLine($"      项目 {project.ProjectNumber} 负责人匹配，开始处理...");
                    }
                    else
                    {
                        Console.WriteLine($"      项目 {project.ProjectNumber} 开始处理（忽略负责人匹配）...");
                    }

                    // 处理匹配的项目
                    await ProcessSingleProject(project, i + 2); // +2 因为Excel行从1开始，且第一行是表头
                }

                Console.WriteLine("      科目余额查询处理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      处理科目余额查询失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 获取当前登录用户名
        /// </summary>
        private async Task<string> GetCurrentUsername()
        {
            try
            {
                string rawUsername = "";
                
                // 尝试从#spUsername元素获取用户名
                var usernameElement = page.Locator("#spUsername").First;
                if (await usernameElement.CountAsync() > 0)
                {
                    rawUsername = await usernameElement.TextContentAsync();
                }
                else
                {
                    // 如果主页面找不到，尝试在iframe中查找
                    var frames = page.Frames;
                    foreach (var frame in frames)
                    {
                        try
                        {
                            var frameUsernameElement = frame.Locator("#spUsername").First;
                            if (await frameUsernameElement.CountAsync() > 0)
                            {
                                rawUsername = await frameUsernameElement.TextContentAsync();
                                break;
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }

                if (string.IsNullOrEmpty(rawUsername))
                {
                    Console.WriteLine("      警告：无法获取当前用户名");
                    return "";
                }

                string cleanUsername = rawUsername.Trim();
                Console.WriteLine($"      获取到原始用户名: '{cleanUsername}'");

                // 如果用户名包含"教师"后缀，去掉后缀
                if (cleanUsername.EndsWith("教师"))
                {
                    string usernameWithoutSuffix = cleanUsername.Replace("教师", "").Trim();
                    Console.WriteLine($"      去掉'教师'后缀后的用户名: '{usernameWithoutSuffix}'");
                    return usernameWithoutSuffix;
                }

                return cleanUsername;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      获取当前用户名失败: {ex.Message}");
                return "";
            }
        }

        /// <summary>
        /// 读取0-科目余额sheet
        /// </summary>
        private List<ProjectInfo> ReadSubjectBalanceSheet()
        {
            try
            {
                // 设置EPPlus许可证上下文
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                // 查找Excel文件
                string excelPath = FindExcelFile("420财务050823.xlsx");
                if (string.IsNullOrEmpty(excelPath))
                {
                    Console.WriteLine("      警告：未找到Excel文件");
                    return null;
                }

                Console.WriteLine($"      找到Excel文件: {excelPath}");

                using (var package = new ExcelPackage(new FileInfo(excelPath)))
                {
                    var worksheet = package.Workbook.Worksheets["0-科目余额"];
                    if (worksheet == null)
                    {
                        Console.WriteLine("      警告：未找到0-科目余额工作表");
                        return null;
                    }

                    var projectList = new List<ProjectInfo>();
                    int rowCount = worksheet.Dimension?.Rows ?? 0;
                    int colCount = worksheet.Dimension?.Columns ?? 0;

                    Console.WriteLine($"      调试：工作表有 {rowCount} 行，{colCount} 列");

                    // 显示表头信息
                    if (rowCount > 0)
                    {
                        Console.WriteLine("      调试：表头信息：");
                        for (int col = 1; col <= Math.Min(colCount, 5); col++)
                        {
                            string headerText = worksheet.Cells[1, col].Text?.Trim() ?? "";
                            Console.WriteLine($"        列 {col}: '{headerText}'");
                        }
                    }

                    // 假设第一行是表头，从第二行开始读取数据
                    for (int row = 2; row <= rowCount; row++)
                    {
                        string projectNumber = worksheet.Cells[row, 1].Text?.Trim();
                        string responsiblePerson = worksheet.Cells[row, 3].Text?.Trim(); // 第三列是负责人
                        
                        if (!string.IsNullOrEmpty(projectNumber))
                        {
                            var projectInfo = new ProjectInfo
                            {
                                ProjectNumber = projectNumber,
                                ResponsiblePerson = responsiblePerson ?? "",
                                ExcelRow = row
                            };
                            projectList.Add(projectInfo);
                            
                            // 调试输出：显示读取到的负责人信息
                            Console.WriteLine($"      调试：项目 {projectNumber}，负责人列内容: '{responsiblePerson}'");
                        }
                    }

                    Console.WriteLine($"      从Excel读取到 {projectList.Count} 个项目");
                    return projectList;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      读取Excel文件失败: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 检查用户是否匹配（模糊匹配）
        /// </summary>
        private bool IsUserMatched(string responsiblePerson, string currentUsername)
        {
            if (string.IsNullOrEmpty(responsiblePerson) || string.IsNullOrEmpty(currentUsername))
                return false;

            // 清理字符串：移除换行符、多余空格等
            string cleanResponsiblePerson = CleanUserName(responsiblePerson);
            string cleanCurrentUsername = CleanUserName(currentUsername);

            Console.WriteLine($"      匹配检查：负责人='{cleanResponsiblePerson}' vs 当前用户='{cleanCurrentUsername}'");

            // 多种匹配方式
            // 1. 完全匹配
            if (cleanResponsiblePerson.Equals(cleanCurrentUsername, StringComparison.OrdinalIgnoreCase))
                return true;

            // 2. 包含匹配
            if (cleanResponsiblePerson.Contains(cleanCurrentUsername) || cleanCurrentUsername.Contains(cleanResponsiblePerson))
                return true;

            // 注意：GetCurrentUsername已经处理了"教师"后缀的去除，所以这里不需要再次处理

            return false;
        }

        /// <summary>
        /// 清理用户名，移除换行符和多余空格
        /// </summary>
        private string CleanUserName(string userName)
        {
            if (string.IsNullOrEmpty(userName))
                return "";

            return userName
                .Replace("\r", "")
                .Replace("\n", "")
                .Replace("\t", " ")
                .Trim();
        }

        /// <summary>
        /// 处理单个项目
        /// </summary>
        private async Task ProcessSingleProject(ProjectInfo project, int excelRow)
        {
            try
            {
                Console.WriteLine($"      处理项目: {project.ProjectNumber}");

                // 1. 填写项目编号到输入框
                await FillProjectNumber(project.ProjectNumber);

                // 2. 填写金额1到addition输入框
                await FillAdditionAmount("1");

                // 3. 选择支付方式为"个人转卡"
                await SelectPaymentType("个人转卡");

                // 4. 点击下一步按钮
                await ClickNextStepButton();

                // 5. 提取表格数据
                try
                {
                    var tableData = await ExtractTableData();
                    Console.WriteLine($"      项目 {project.ProjectNumber} 表格数据提取完成，共 {tableData.Count} 行");
                    
                    // 输出表格数据摘要
                    foreach (var row in tableData)
                    {
                        Console.WriteLine($"        报销项: {row["报销项"]}, 余额: {row["余额"]}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      项目 {project.ProjectNumber} 表格数据提取失败: {ex.Message}");
                    Console.WriteLine($"      跳过表格数据提取，继续处理下一个项目");
                }

                Console.WriteLine($"      项目 {project.ProjectNumber} 处理完成");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      处理项目 {project.ProjectNumber} 失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 填写项目编号
        /// </summary>
        private async Task FillProjectNumber(string projectNumber)
        {
            try
            {
                string inputId = "formWF_YB6_230_yta-uni_prj_code";
                var frames = page.Frames;

                foreach (var frame in frames)
                {
                    try
                    {
                        var inputElement = frame.Locator($"#{inputId}").First;
                        if (await inputElement.CountAsync() > 0)
                        {
                            await inputElement.ClickAsync();
                            await inputElement.FillAsync("");
                            await inputElement.FillAsync(projectNumber);
                            Console.WriteLine($"      成功填写项目编号: {projectNumber}");
                            return;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                Console.WriteLine($"      警告：未找到项目编号输入框 {inputId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      填写项目编号失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 填写金额到addition输入框
        /// </summary>
        private async Task FillAdditionAmount(string amount)
        {
            try
            {
                string inputId = "formWF_YB6_230_yta-addition";
                var frames = page.Frames;

                foreach (var frame in frames)
                {
                    try
                    {
                        var inputElement = frame.Locator($"#{inputId}").First;
                        if (await inputElement.CountAsync() > 0)
                        {
                            await inputElement.ClickAsync();
                            await inputElement.FillAsync("");
                            await inputElement.FillAsync(amount);
                            Console.WriteLine($"      成功填写金额: {amount}");
                            return;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                Console.WriteLine($"      警告：未找到金额输入框 {inputId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      填写金额失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 选择支付方式
        /// </summary>
        private async Task SelectPaymentType(string paymentType)
        {
            try
            {
                string selectId = "formWF_YB6_230_yta-pay_type";
                var frames = page.Frames;

                foreach (var frame in frames)
                {
                    try
                    {
                        var selectElement = frame.Locator($"#{selectId}").First;
                        if (await selectElement.CountAsync() > 0)
                        {
                            await selectElement.SelectOptionAsync(new[] { paymentType });
                            Console.WriteLine($"      成功选择支付方式: {paymentType}");
                            return;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                Console.WriteLine($"      警告：未找到支付方式下拉框 {selectId}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      选择支付方式失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 点击下一步按钮
        /// </summary>
        private async Task ClickNextStepButton()
        {
            try
            {
                string btnname = "下一步";
                var frames = page.Frames;
                bool clicked = false;

                foreach (var frame in frames)
                {
                    try
                    {
                        var buttonElement = frame.Locator($"button[btnname='{btnname}']").First;
                        if (await buttonElement.CountAsync() > 0)
                        {
                            await buttonElement.ClickAsync();
                            Console.WriteLine($"      成功点击下一步按钮");
                            clicked = true;
                            
                            // 等待页面加载完成
                            await Task.Delay(2000);
                            break;
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (!clicked)
                {
                    Console.WriteLine($"      警告：未找到下一步按钮");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      点击下一步按钮失败: {ex.Message}");
            }
        }

        /// <summary>
        /// 从iframe表格中提取报销项数据
        /// </summary>
        private async Task<List<Dictionary<string, string>>> ExtractTableData()
        {
            try
            {
                Console.WriteLine("      开始提取表格数据...");
                var tableData = new List<Dictionary<string, string>>();

                // 等待页面完全加载
                await Task.Delay(3000);
                Console.WriteLine("      等待页面加载完成...");

                // 在iframe中查找表格
                var frames = page.Frames;
                Console.WriteLine($"      当前页面共有 {frames.Count} 个iframe");
                IFrame targetFrame = null;
                
                for (int i = 0; i < frames.Count; i++)
                {
                    var frame = frames[i];
                    try
                    {
                        Console.WriteLine($"      检查iframe {i}...");
                        
                        // 只查找特定的表格ID
                        var table = frame.Locator("#gridWF_YB6_2375");
                        var count = await table.CountAsync();
                        if (count > 0)
                        {
                            // 验证表格ID是否正确
                            var tableId = await table.First.GetAttributeAsync("id");
                            if (tableId == "gridWF_YB6_2375")
                            {
                                Console.WriteLine($"      在iframe {i} 中找到目标表格: #gridWF_YB6_2375");
                                
                                // 获取表格的详细信息
                                var tableClass = await table.First.GetAttributeAsync("class");
                                Console.WriteLine($"      表格ID: {tableId}, 类名: {tableClass}");
                                
                                targetFrame = frame;
                                break;
                            }
                            else
                            {
                                Console.WriteLine($"      在iframe {i} 中找到表格但ID不匹配: {tableId}");
                            }
                        }
                        
                        if (targetFrame != null) break;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      在iframe {i} 中查找表格时出错: {ex.Message}");
                    }
                }

                if (targetFrame == null)
                {
                    Console.WriteLine("      未找到目标表格 #gridWF_YB6_2375，跳过表格数据提取");
                    return new List<Dictionary<string, string>>(); // 返回空列表而不是抛出异常
                }

                // 获取所有数据行（跳过表头行）
                // 优先选择叶节点行（有span.cell-wrapperleaf的行）
                var leafRows = targetFrame.Locator("#gridWF_YB6_2375 tbody tr:not(.jqgfirstrow)");
                var allRows = leafRows;
                var rowCount = await allRows.CountAsync();
                Console.WriteLine($"      找到 {rowCount} 行数据");
                
                if (rowCount == 0)
                {
                    // 如果没找到数据行，尝试其他选择器
                    Console.WriteLine("      尝试其他行选择器...");
                    var altRows = targetFrame.Locator("#gridWF_YB6_2375 tbody tr");
                    var altRowCount = await altRows.CountAsync();
                    Console.WriteLine($"      使用备用选择器找到 {altRowCount} 行");
                    
                    if (altRowCount > 0)
                    {
                        allRows = altRows;
                        rowCount = altRowCount;
                    }
                }

                for (int i = 0; i < rowCount; i++)
                {
                    try
                    {
                        var row = allRows.Nth(i);
                        
                        // 检查是否为叶节点（有span.cell-wrapperleaf的行）
                        var leafElement = row.Locator("td:first-child span.cell-wrapperleaf");
                        var leafCount = await leafElement.CountAsync();
                        
                        if (leafCount > 0)
                        {
                            Console.WriteLine($"      正在处理第 {i + 1} 行（叶节点）...");
                            var rowData = await ExtractRowData(targetFrame, row, i + 1);
                            if (rowData != null)
                            {
                                tableData.Add(rowData);
                                Console.WriteLine($"      第 {i + 1} 行处理成功");
                            }
                            else
                            {
                                Console.WriteLine($"      第 {i + 1} 行提取失败，跳过该行");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"      第 {i + 1} 行不是叶节点，跳过");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"      提取第 {i + 1} 行数据时出错: {ex.Message}，跳过该行");
                    }
                }

                Console.WriteLine($"      成功提取 {tableData.Count} 行数据");
                return tableData;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      提取表格数据失败: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 提取单行数据
        /// </summary>
        private async Task<Dictionary<string, string>> ExtractRowData(IFrame frame, ILocator row, int rowIndex)
        {
            try
            {
                var rowData = new Dictionary<string, string>();

                // 提取报销项名称（优先获取span.cell-wrapperleaf，如果没有则获取span.cell-wrapper）
                string expenseName = "";
                try
                {
                    // 先尝试获取叶节点的名称
                    var leafElement = row.Locator("td:first-child span.cell-wrapperleaf");
                    var leafCount = await leafElement.CountAsync();
                    if (leafCount > 0)
                    {
                        expenseName = await leafElement.First.TextContentAsync();
                    }
                    else
                    {
                        // 如果没有叶节点，尝试获取父节点名称
                        var wrapperElement = row.Locator("td:first-child span.cell-wrapper");
                        var wrapperCount = await wrapperElement.CountAsync();
                        if (wrapperCount > 0)
                        {
                            expenseName = await wrapperElement.First.TextContentAsync();
                        }
                    }
                }
                catch (TimeoutException)
                {
                    Console.WriteLine($"      第 {rowIndex} 行：获取报销项名称超时");
                    return null;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      第 {rowIndex} 行：获取报销项名称出错: {ex.Message}");
                    return null;
                }
                
                if (string.IsNullOrEmpty(expenseName))
                {
                    Console.WriteLine($"      第 {rowIndex} 行：未找到报销项名称");
                    return null;
                }
                expenseName = expenseName.Trim();
                rowData["报销项"] = expenseName;

                // 提取金额输入框并获取余额信息
                var amountInput = row.Locator("td:nth-child(2) input.qinput").First;
                int inputCount = 0;
                try
                {
                    inputCount = await amountInput.CountAsync();
                }
                catch (TimeoutException)
                {
                    Console.WriteLine($"      第 {rowIndex} 行：获取金额输入框超时");
                    return null;
                }
                
                if (inputCount > 0)
                {
                    var balanceInfo = await GetBalanceInfoByHover(frame, amountInput, rowIndex);
                    rowData["余额"] = balanceInfo;
                }
                else
                {
                    Console.WriteLine($"      第 {rowIndex} 行：未找到金额输入框");
                    rowData["余额"] = "未找到金额输入框";
                }

                // 提取说明信息（第六个td）
                var descriptionElement = row.Locator("td:nth-child(6)").First;
                string description = "";
                try
                {
                    description = await descriptionElement.TextContentAsync();
                }
                catch (TimeoutException)
                {
                    Console.WriteLine($"      第 {rowIndex} 行：获取说明信息超时");
                    description = "";
                }
                rowData["说明"] = string.IsNullOrEmpty(description) ? "" : description.Trim();

                Console.WriteLine($"      第 {rowIndex} 行：{expenseName} - {rowData["余额"]}");
                return rowData;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      提取第 {rowIndex} 行数据时出错: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 通过鼠标悬停获取余额信息
        /// </summary>
        private async Task<string> GetBalanceInfoByHover(IFrame frame, ILocator inputElement, int rowIndex)
        {
            try
            {
                Console.WriteLine($"      第 {rowIndex} 行：开始悬停操作...");
                
                // 先尝试获取title属性（有些输入框可能直接在title中显示余额信息）
                var title = await inputElement.GetAttributeAsync("title");
                if (!string.IsNullOrEmpty(title) && !title.Trim().Equals(""))
                {
                    Console.WriteLine($"      第 {rowIndex} 行：从title获取: {title}");
                    return title.Trim();
                }
                
                // 如果title为空，尝试悬停操作
                var tipContent = frame.Locator("#tiptip_content");
                var tipHolder = frame.Locator("#tiptip_holder");
                
                // 检查提示框元素是否存在
                var tipContentCount = await tipContent.CountAsync();
                var tipHolderCount = await tipHolder.CountAsync();
                Console.WriteLine($"      第 {rowIndex} 行：提示框元素 - tipContent: {tipContentCount}, tipHolder: {tipHolderCount}");
                
                // 鼠标悬停在输入框上
                await inputElement.HoverAsync();
                Console.WriteLine($"      第 {rowIndex} 行：已执行悬停操作");
                await Task.Delay(2000); // 增加等待时间让提示框显示

                // 检查提示框是否显示
                if (tipHolderCount > 0)
                {
                    var tipDisplay = await tipHolder.GetAttributeAsync("style");
                    Console.WriteLine($"      第 {rowIndex} 行：提示框样式: {tipDisplay}");
                    
                    if (string.IsNullOrEmpty(tipDisplay) || !tipDisplay.Contains("display: none"))
                    {
                        // 获取提示框内容
                        if (tipContentCount > 0)
                        {
                            var balanceText = await tipContent.TextContentAsync();
                            if (!string.IsNullOrEmpty(balanceText))
                            {
                                var cleanBalance = balanceText.Trim();
                                Console.WriteLine($"      第 {rowIndex} 行悬停获取余额: {cleanBalance}");
                                return cleanBalance;
                            }
                        }
                    }
                }

                // 如果提示框没有显示，尝试其他方法
                Console.WriteLine($"      第 {rowIndex} 行：悬停未显示提示框，尝试其他方法");
                
                // 尝试点击输入框来触发提示框
                try
                {
                    await inputElement.ClickAsync();
                    await Task.Delay(1000);
                    
                    if (tipHolderCount > 0)
                    {
                        var tipDisplay = await tipHolder.GetAttributeAsync("style");
                        if (string.IsNullOrEmpty(tipDisplay) || !tipDisplay.Contains("display: none"))
                        {
                            if (tipContentCount > 0)
                            {
                                var balanceText = await tipContent.TextContentAsync();
                                if (!string.IsNullOrEmpty(balanceText))
                                {
                                    var cleanBalance = balanceText.Trim();
                                    Console.WriteLine($"      第 {rowIndex} 行点击获取余额: {cleanBalance}");
                                    return cleanBalance;
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"      第 {rowIndex} 行：点击操作失败: {ex.Message}");
                }
                if (!string.IsNullOrEmpty(title))
                {
                    Console.WriteLine($"      第 {rowIndex} 行：从title获取: {title}");
                    return title.Trim();
                }

                // 尝试获取输入框的value属性
                var value = await inputElement.GetAttributeAsync("value");
                if (!string.IsNullOrEmpty(value))
                {
                    Console.WriteLine($"      第 {rowIndex} 行：从value获取: {value}");
                    return $"输入框值: {value}";
                }

                Console.WriteLine($"      第 {rowIndex} 行：未获取到任何余额信息");
                return "未获取到余额信息";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"      第 {rowIndex} 行悬停获取余额时出错: {ex.Message}");
                return $"获取余额失败: {ex.Message}";
            }
        }

        /// <summary>
        /// 查找Excel文件
        /// </summary>
        private string FindExcelFile(string fileName)
        {
            // 从当前目录开始查找
            string[] searchPaths = {
                fileName,
                Path.Combine("..", fileName),
                Path.Combine("..", "..", fileName),
                Path.Combine("..", "..", "..", fileName),
                Path.Combine("..", "..", "..", "..", fileName)
            };

            foreach (string path in searchPaths)
            {
                if (File.Exists(path))
                {
                    return Path.GetFullPath(path);
                }
            }

            return null;
        }
    }

    /// <summary>
    /// 项目信息类
    /// </summary>
    public class ProjectInfo
    {
        public string ProjectNumber { get; set; }
        public string ResponsiblePerson { get; set; }
        public int ExcelRow { get; set; }
    }
}
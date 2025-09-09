using System;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Playwright;

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
                try { if (browser != null) await browser.CloseAsync(); } catch { }
                try { if (playwright != null) playwright.Dispose(); } catch { }
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
            Console.WriteLine("请输入登录信息：");
            Console.Write("用户名: ");
            string username = Console.ReadLine()?.Trim() ?? string.Empty;
            Console.Write("密码: ");
            string password = ReadPassword();
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
    }
}
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
        /// ��������������������в���ϵͳ
        /// </summary>
        public async Task InitializeAsync()
        {
            try
            {
                Console.WriteLine("=== ���в���ϵͳ�Զ��� ===");
                Console.WriteLine("�������������...");

                playwright = await Playwright.CreateAsync();
                browser = await playwright.Chromium.LaunchAsync(new BrowserTypeLaunchOptions
                {
                    Headless = false, // ��ʾ���������
                    SlowMo = 100 // ���������ٶȣ����ڹ۲�
                });

                page = await browser.NewPageAsync();
                page.SetDefaultTimeout(10000); // 10�볬ʱ

                Console.WriteLine("����������ɹ�");
                Console.WriteLine("���ڵ��������в���ϵͳ...");

                // ������Ŀ����ҳ
                await page.GotoAsync("https://www.kycw.uestc.edu.cn/WFManager/home.jsp");
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);

                Console.WriteLine("�ɹ����������в���ϵͳ");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"��ʼ��ʧ��: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// ִ�е�¼����
        /// </summary>
        public async Task<bool> LoginAsync()
        {
            try
            {
                Console.WriteLine("��ʼִ�е�¼����...");

                // �ȴ���¼������
                await page.WaitForSelectorAsync("#uid", new PageWaitForSelectorOptions { Timeout = 5000 });
                Console.WriteLine("��¼���������");

                // ��д�û���
                await page.FillAsync("#uid", "5070016");
                Console.WriteLine("����д�û���: 5070016");

                // ��д����
                await page.FillAsync("#pwd", "Kp5070016");
                Console.WriteLine("����д����: Kp5070016");

                // �ȴ��û�������֤��
                Console.WriteLine("��鿴��֤��ͼƬ�����·�������֤��:");
                Console.WriteLine("��֤�������ID: chkcode1");
                Console.WriteLine("��֤��ͼƬID: checkcodeImg");
                Console.WriteLine("��¼��ťID: zhLogin");
                Console.WriteLine("----------------------------------------");

                // ������֤�������
                await page.EvaluateAsync(@"
                    document.getElementById('chkcode1').style.border = '2px solid red';
                    document.getElementById('chkcode1').style.backgroundColor = '#fff3cd';
                    document.getElementById('chkcode1').focus();
                ");

                // �ȴ��û��ڿ���̨������֤��
                Console.Write("��������֤��: ");
                var captchaInput = Console.ReadLine()?.Trim();

                if (string.IsNullOrEmpty(captchaInput))
                {
                    Console.WriteLine("������֤�벻��Ϊ��");
                    return false;
                }

                Console.WriteLine($"�յ���֤��: {captchaInput}");

                // ����֤����䵽��ҳ�����
                await page.FillAsync("#chkcode1", captchaInput);
                Console.WriteLine("��֤������䵽��ҳ�����");

                // �Ƴ�������ʽ
                await page.EvaluateAsync(@"
                    document.getElementById('chkcode1').style.border = '';
                    document.getElementById('chkcode1').style.backgroundColor = '';
                ");

                Console.WriteLine("׼�������¼��ť...");

                // �ȴ���¼��ť����
                Console.WriteLine("�ȴ���¼��ť����...");
                try
                {
                    // ���ȳ��Եȴ���ť�ɼ�
                    await page.WaitForSelectorAsync("#zhLogin", new PageWaitForSelectorOptions { Timeout = 5000 });
                    Console.WriteLine("��¼��ť���ҵ�");

                    // ��鰴ť�Ƿ����
                    var buttonElement = page.Locator("#zhLogin");
                    var isDisabled = await buttonElement.GetAttributeAsync("disabled");

                    if (isDisabled != null)
                    {
                        Console.WriteLine("��¼��ť��ǰ�����ã��ȴ�����...");
                        // �ȴ���ť����
                        await page.WaitForSelectorAsync("#zhLogin:not([disabled])", new PageWaitForSelectorOptions { Timeout = 10000 });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"�ȴ���¼��ťʱ����: {ex.Message}");
                    Console.WriteLine("����ֱ�ӵ����¼��ť...");
                }

                // �����¼��ť
                Console.WriteLine("���ڵ����¼��ť...");
                await page.ClickAsync("#zhLogin");
                Console.WriteLine("�ѵ����¼��ť");

                // ��ӵ�����Ϣ
                Console.WriteLine("�ȴ�ҳ����Ӧ...");
                await Task.Delay(3000); // �ȴ�3��

                // ��鵱ǰURL
                var currentUrlAfterClick = page.Url;
                Console.WriteLine($"�����¼��ť���URL: {currentUrlAfterClick}");

                // �ȴ�����ʱ����ҳ����ȫ����
                await Task.Delay(5000);

                // ����Ƿ��¼�ɹ� - ʹ�ø����ɵ��ж�����
                try
                {
                    // ���URL�仯 - ���URL�����˱仯��ͨ����ʾ��¼�ɹ�
                    var finalUrl = page.Url;
                    Console.WriteLine($"����URL: {finalUrl}");

                    // ���URL������¼��صĹؼ��ʣ����ܻ��ڵ�¼ҳ��
                    if (finalUrl.Contains("login") || finalUrl.Contains("auth"))
                    {
                        Console.WriteLine("���ڵ�¼ҳ�棬��¼����ʧ��");
                        return false;
                    }

                    // ����Ƿ������ԵĴ�����Ϣ
                    var pageContent = await page.ContentAsync();
                    if (pageContent.Contains("�û������������") ||
                        pageContent.Contains("��֤�����") ||
                        pageContent.Contains("��¼ʧ��"))
                    {
                        Console.WriteLine("ҳ�������ȷ�Ĵ�����Ϣ����¼ʧ��");
                        return false;
                    }

                    // ����Ƿ��е�¼�ɹ��ı�ʶ
                    var successIndicators = new[] { "��ӭ", "����", "��¼�ɹ�", "��ҳ", "��ҳ��" };
                    var hasSuccessIndicator = false;

                    foreach (var indicator in successIndicators)
                    {
                        try
                        {
                            var locator = page.Locator($"text={indicator}");
                            if (await locator.CountAsync() > 0)
                            {
                                Console.WriteLine($"�ҵ��ɹ���ʶ: {indicator}");
                                hasSuccessIndicator = true;
                                break;
                            }
                        }
                        catch
                        {
                            // ���Ե���ѡ�����Ĵ���
                        }
                    }

                    if (hasSuccessIndicator)
                    {
                        Console.WriteLine("��¼�ɹ���");
                        isLoggedIn = true;
                        return true;
                    }

                    // ���û����ȷ�ĳɹ���ʶ����URL�Ѿ��仯��û�д�����Ϣ��Ҳ��Ϊ��¼�ɹ�
                    if (finalUrl != "https://www.kycw.uestc.edu.cn/WFManager/home.jsp")
                    {
                        Console.WriteLine("URL�ѱ仯���޴�����Ϣ����Ϊ��¼�ɹ�");
                        isLoggedIn = true;
                        return true;
                    }

                    Console.WriteLine("��¼״̬����ȷ��������ȷ������Ϊ��¼�ɹ�");
                    isLoggedIn = true;
                    return true;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"����¼״̬ʱ����: {ex.Message}");
                    // ��ʹ���������URL�Ѿ��仯��Ҳ��Ϊ���ܵ�¼�ɹ�
                    var currentUrl = page.Url;
                    if (currentUrl != "https://www.kycw.uestc.edu.cn/WFManager/home.jsp")
                    {
                        Console.WriteLine("������URL�ѱ仯����Ϊ��¼�ɹ�");
                        isLoggedIn = true;
                        return true;
                    }
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"��¼���̳���: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ������ָ��ҳ��
        /// </summary>
        public async Task NavigateToPageAsync(string url)
        {
            try
            {
                if (!isLoggedIn)
                {
                    Console.WriteLine("������δ��¼������ִ�е�¼");
                    return;
                }

                Console.WriteLine($"���ڵ�����: {url}");
                await page.GotoAsync(url);
                await page.WaitForLoadStateAsync(LoadState.NetworkIdle);
                Console.WriteLine("ҳ�浼�����");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"ҳ�浼��ʧ��: {ex.Message}");
            }
        }

        /// <summary>
        /// ������ݲ�ѯ����ʾ������ť
        /// </summary>
        public async Task<bool> ClickDataQueryAndPublicityButtonAsync()
        {
            try
            {
                if (!isLoggedIn)
                {
                    Console.WriteLine("������δ��¼������ִ�е�¼");
                    return false;
                }

                Console.WriteLine("�ȴ�3���ʼ���������ť...");
                await Task.Delay(3000);

                Console.WriteLine("���ڲ���'���ݲ�ѯ����ʾ'������ť...");

                // �����û��ṩ��HTML���룬ʹ�ø���ȷ��ѡ����
                var buttonSelectors = new[]
                {
                    // ��ȷƥ�䣺����"���ݲ�ѯ����ʾ"�ı�������
                    "a:has-text('���ݲ�ѯ����ʾ')",
                    // �����ض�onclick���Ե�����
                    "a[onclick*='��Ŀִ�������ѯ']",
                    "a[onclick*='���ݲ�ѯ����ʾ']",
                    // �����ض�span�ı�������
                    "a:has(span:text('���ݲ�ѯ����ʾ'))",
                    // ͨ���ı�ƥ��
                    "text=���ݲ�ѯ����ʾ",
                    // �����ض������ʽ������
                    "a[style*='cursor:pointer']:has-text('���ݲ�ѯ����ʾ')"
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
                            Console.WriteLine($"�ҵ���ť��ʹ��ѡ����: {selector}");
                            break;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"ѡ���� {selector} ����ʧ��: {ex.Message}");
                    }
                }

                if (buttonLocator == null)
                {
                    Console.WriteLine("δ�ҵ�'���ݲ�ѯ����ʾ'��ť�����Բ��Ұ������ı���Ԫ��...");

                    // �������а���"���ݲ�ѯ����ʾ"�ı���Ԫ��
                    var allElements = await page.Locator("*:has-text('���ݲ�ѯ����ʾ')").AllAsync();
                    if (allElements.Count > 0)
                    {
                        Console.WriteLine($"�ҵ� {allElements.Count} ������'���ݲ�ѯ����ʾ'�ı���Ԫ��");

                        // ���Ե����һ���ɵ����Ԫ��
                        for (int i = 0; i < allElements.Count; i++)
                        {
                            try
                            {
                                var element = allElements[i];
                                var tagName = await element.EvaluateAsync<string>("el => el.tagName.toLowerCase()");
                                var isClickable = await element.EvaluateAsync<bool>("el => !!(el.onclick || el.getAttribute('href') || el.tagName.toLowerCase() === 'a')");

                                Console.WriteLine($"Ԫ�� {i + 1}: ��ǩ={tagName}, �ɵ��={isClickable}");

                                if (isClickable)
                                {
                                    Console.WriteLine($"���Ե��Ԫ�� {i + 1}...");
                                    await element.ClickAsync();
                                    Console.WriteLine("������ť����ɹ���");
                                    return true;
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"���Ԫ�� {i + 1} ʧ��: {ex.Message}");
                            }
                        }
                    }

                    Console.WriteLine("�޷��ҵ��ɵ����'���ݲ�ѯ����ʾ'��ť");
                    return false;
                }

                // ����ҵ��İ�ť
                Console.WriteLine("���ڵ��'���ݲ�ѯ����ʾ'������ť...");
                await buttonLocator.ClickAsync();
                Console.WriteLine("������ť����ɹ���");

                // �ȴ�ҳ����Ӧ
                await Task.Delay(2000);

                // ���ҳ���Ƿ����仯
                var currentUrl = page.Url;
                Console.WriteLine($"���������ť���URL: {currentUrl}");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"���������ťʧ��: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// ����ҳ���е��ض���Ϣ
        /// </summary>
        public async Task<Dictionary<string, string>> ExtractPageInfoAsync()
        {
            try
            {
                if (!isLoggedIn)
                {
                    Console.WriteLine("������δ��¼������ִ�е�¼");
                    return new Dictionary<string, string>();
                }

                Console.WriteLine("��ʼ��ȡҳ����Ϣ...");
                var pageInfo = new Dictionary<string, string>();

                // ��ȡҳ�����
                var title = await page.TitleAsync();
                pageInfo["ҳ�����"] = title;
                Console.WriteLine($"ҳ�����: {title}");

                // ��ȡ��ǰURL
                var currentUrl = page.Url;
                pageInfo["��ǰURL"] = currentUrl;
                Console.WriteLine($"��ǰURL: {currentUrl}");

                // ���ұ����Ϣ
                var tables = await page.Locator("table").AllAsync();
                Console.WriteLine($"�ҵ� {tables.Count} �����");

                for (int i = 0; i < tables.Count; i++)
                {
                    try
                    {
                        var table = tables[i];
                        var tableInfo = await ExtractTableInfoAsync(table, i + 1);
                        foreach (var kvp in tableInfo)
                        {
                            pageInfo[$"���{i + 1}_{kvp.Key}"] = kvp.Value;
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"��ȡ��� {i + 1} ��Ϣʱ����: {ex.Message}");
                    }
                }

                // ����������Ҫ��Ϣ
                var importantElements = await page.Locator("h1, h2, h3, .title, .header, [class*='title'], [class*='header']").AllAsync();
                foreach (var element in importantElements)
                {
                    try
                    {
                        var text = await element.TextContentAsync();
                        if (!string.IsNullOrEmpty(text))
                        {
                            var cleanText = text.Trim();
                            if (cleanText.Length > 0 && cleanText.Length < 100) // ���Ƴ��ȱ������
                            {
                                pageInfo[$"����_{cleanText.Substring(0, Math.Min(20, cleanText.Length))}"] = cleanText;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"��ȡ������Ϣʱ����: {ex.Message}");
                    }
                }

                Console.WriteLine($"�ɹ���ȡ {pageInfo.Count} ��ҳ����Ϣ");
                return pageInfo;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"��ȡҳ����Ϣʧ��: {ex.Message}");
                return new Dictionary<string, string>();
            }
        }

        /// <summary>
        /// ��ȡ�����Ϣ
        /// </summary>
        private async Task<Dictionary<string, string>> ExtractTableInfoAsync(ILocator table, int tableIndex)
        {
            var tableInfo = new Dictionary<string, string>();

            try
            {
                // ��ȡ�������
                var rows = await table.Locator("tr").AllAsync();
                tableInfo[$"����"] = rows.Count.ToString();

                // ��ȡ����������Ե�һ��Ϊ׼��
                if (rows.Count > 0)
                {
                    var firstRowCells = await rows[0].Locator("td, th").AllAsync();
                    tableInfo[$"����"] = firstRowCells.Count.ToString();
                }

                // ��ȡ������ݣ�ǰ���У�
                var maxRowsToExtract = Math.Min(5, rows.Count); // �����ȡ5��
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
                        tableInfo[$"��{i + 1}"] = string.Join(" | ", rowData);
                    }
                }

                Console.WriteLine($"��� {tableIndex} ��Ϣ��ȡ���");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"��ȡ��� {tableIndex} ��Ϣʱ����: {ex.Message}");
            }

            return tableInfo;
        }

        /// <summary>
        /// �ȴ��û��������
        /// </summary>
        public async Task WaitForUserOperationAsync()
        {
            Console.WriteLine("�ȴ��û��������...");
            Console.WriteLine("���س�����������ִ��...");

            try
            {
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"�ȴ��û�����ʱ����: {ex.Message}");
            }
        }

        /// <summary>
        /// �ر������
        /// </summary>
        public async Task CloseAsync()
        {
            try
            {
                if (browser != null)
                {
                    await browser.CloseAsync();
                    Console.WriteLine("������ѹر�");
                }

                if (playwright != null)
                {
                    playwright.Dispose();
                    Console.WriteLine("Playwright��Դ���ͷ�");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"�ر���Դʱ����: {ex.Message}");
            }
        }

        /// <summary>
        /// ����¼״̬
        /// </summary>
        public bool IsLoggedIn => isLoggedIn;

        /// <summary>
        /// ��ȡ��ǰҳ����󣨹��ⲿʹ�ã�
        /// </summary>
        public IPage CurrentPage => page;
    }
}
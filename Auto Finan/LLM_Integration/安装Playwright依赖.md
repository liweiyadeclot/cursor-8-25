# 安装 Playwright 依赖

## 📋 前置要求

网关执行版本需要安装 Playwright 才能运行。

---

## 🔧 安装步骤

### 步骤 1：安装 Playwright Python 包

```bash
pip install playwright
```

### 步骤 2：安装浏览器

```bash
playwright install chromium
```

或者安装所有浏览器：

```bash
playwright install
```

---

## ✅ 验证安装

运行以下命令验证：

```bash
python -c "from playwright.sync_api import sync_playwright; print('✅ Playwright 已安装')"
```

---

## 🚀 启动服务

安装完成后，启动服务：

```bash
cd "Auto Finan\LLM_Integration"
start_mcp_gateway.bat
```

---

## ⚠️ 如果安装失败

### 问题 1：网络问题

如果下载浏览器失败，可以：

1. **使用镜像源**：
   ```bash
   set PLAYWRIGHT_DOWNLOAD_HOST=https://npmmirror.com/mirrors/playwright
   playwright install chromium
   ```

2. **手动下载**：从 Playwright 官网下载浏览器

### 问题 2：权限问题

确保有管理员权限，或使用用户目录安装。

---

## 📚 相关文档

- Playwright 官方文档：https://playwright.dev/python/


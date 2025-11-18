# Dify 连接本地服务问题解决方案

## ❌ 错误：Reached maximum retries (0) for URL

这个错误通常是因为 **Dify 无法访问本地服务**。

---

## 🔍 问题诊断

### 检查 1：服务是否运行

```bash
# 检查端口
netstat -an | findstr :8001

# 应该看到：TCP    0.0.0.0:8001   LISTENING
```

✅ **你的服务正在运行**（已确认）

---

### 检查 2：Dify 位置

**关键问题**：Dify 在哪里运行？

- **本地**：和你的电脑在同一台机器
- **远程服务器**：在另一台机器或云端

---

## ✅ 解决方案

### 方案 1：Dify 在本地

如果 Dify 和本地服务在同一台机器：

**配置**：
```
URL: http://localhost:8001/api/excel-to-prompt
```

**如果还是失败**，尝试：
```
URL: http://127.0.0.1:8001/api/excel-to-prompt
```

---

### 方案 2：Dify 在远程服务器（最常见）

如果 Dify 在远程服务器，需要：

#### 步骤 1：获取本地 IP 地址

```bash
# Windows
ipconfig

# 查找 IPv4 地址，例如：
# IPv4 地址 . . . . . . . . . . . . : 192.168.1.100
```

#### 步骤 2：配置防火墙

**Windows 防火墙**：

1. 打开"Windows Defender 防火墙"
2. 点击"高级设置"
3. 点击"入站规则" → "新建规则"
4. 选择"端口" → "下一步"
5. 选择"TCP"，输入端口 `8001` → "下一步"
6. 选择"允许连接" → "下一步"
7. 全部勾选 → "下一步"
8. 名称：`Dify Local Service` → "完成"

**或使用命令行**：

```powershell
# 以管理员身份运行 PowerShell
New-NetFirewallRule -DisplayName "Dify Local Service" -Direction Inbound -LocalPort 8001 -Protocol TCP -Action Allow
```

#### 步骤 3：修改 Dify URL

在 Dify HTTP 请求节点中：

**将**：
```
http://localhost:8001/api/excel-to-prompt
```

**改为**：
```
http://你的IP地址:8001/api/excel-to-prompt
```

例如：
```
http://192.168.1.100:8001/api/excel-to-prompt
```

#### 步骤 4：测试连接

在 Dify 服务器上测试：

```bash
# 测试健康检查
curl http://192.168.1.100:8001/health

# 应该返回：
# {"status":"healthy","service":"Dify Local Service"}
```

---

### 方案 3：使用内网穿透（推荐用于生产环境）

如果需要从外网访问，使用内网穿透：

#### 使用 ngrok

1. **下载 ngrok**：https://ngrok.com/

2. **启动本地服务**：
   ```bash
   start_dify_local_service.bat
   ```

3. **启动 ngrok**（新窗口）：
   ```bash
   ngrok http 8001
   ```

4. **获取公网 URL**：
   ```
   Forwarding: https://abc123.ngrok.io -> http://localhost:8001
   ```

5. **在 Dify 中使用**：
   ```
   https://abc123.ngrok.io/api/excel-to-prompt
   ```

**优点**：
- ✅ 可以从任何地方访问
- ✅ 自动 HTTPS
- ✅ 无需配置防火墙

**缺点**：
- ❌ 免费版有连接限制
- ❌ URL 每次启动会变化（付费版可固定）

---

### 方案 4：将服务部署到 Dify 服务器

如果 Dify 在服务器上，可以将服务也部署到服务器：

1. **上传文件到服务器**：
   - `dify_local_service.py`
   - `workflow_core.py`
   - `excel_to_nl.py`
   - 其他依赖文件

2. **在服务器上运行**：
   ```bash
   python dify_local_service.py --host 0.0.0.0 --port 8001
   ```

3. **在 Dify 中使用**：
   ```
   http://localhost:8001/api/excel-to-prompt
   ```

---

## 🔧 快速修复步骤

### 如果 Dify 在远程服务器

1. **获取本地 IP**：
   ```bash
   ipconfig
   # 记录 IPv4 地址，例如：192.168.1.100
   ```

2. **配置防火墙**（见上方）

3. **修改 Dify URL**：
   ```
   http://192.168.1.100:8001/api/excel-to-prompt
   ```

4. **测试**：
   ```bash
   # 在 Dify 服务器上
   curl http://192.168.1.100:8001/health
   ```

---

## 📝 Dify 配置示例

### HTTP 请求节点配置

**节点类型**：`HTTP 请求`

**配置**：

| 项目 | 值 |
|------|-----|
| 方法 | POST |
| URL | `http://192.168.1.100:8001/api/excel-to-prompt`<br>（替换为你的实际 IP） |
| 请求头 | `{"Content-Type": "application/json"}` |
| 请求体 | `{"excel_path": "{{#workflow.excel_path#}}", "sheet_name": "{{#workflow.sheet_name#}}", "serial": "{{#workflow.serial#}}"}` |
| 超时 | 60 秒 |

---

## 🐛 常见错误

### 错误 1：Connection refused

**原因**：服务未运行或端口错误

**解决**：
1. 启动服务：`start_dify_local_service.bat`
2. 检查端口：`netstat -an | findstr :8001`

---

### 错误 2：Timeout

**原因**：网络延迟或处理时间过长

**解决**：
1. 增加超时时间（在 Dify 中设置为 120 秒）
2. 检查网络连接

---

### 错误 3：无法解析主机

**原因**：IP 地址错误或网络不通

**解决**：
1. 验证 IP 地址：`ping 你的IP`
2. 检查网络连接
3. 确认防火墙配置

---

## ✅ 验证清单

- [ ] 本地服务正在运行
- [ ] 获取了正确的本地 IP 地址
- [ ] 配置了防火墙规则
- [ ] Dify URL 使用正确的 IP 地址
- [ ] 可以从 Dify 服务器访问本地服务（测试 curl）

---

## 💡 推荐方案

**根据你的情况选择**：

1. **Dify 在本地** → 使用 `http://localhost:8001`
2. **Dify 在局域网内** → 使用内网 IP（如 `http://192.168.1.100:8001`）
3. **Dify 在云端/外网** → 使用内网穿透（ngrok）

---

## 📚 相关文件

- `dify_local_service.py` - 本地服务
- `start_dify_local_service.bat` - 启动脚本
- `test_dify_local_service.py` - 测试脚本
- `Dify使用本地服务配置.md` - 详细配置


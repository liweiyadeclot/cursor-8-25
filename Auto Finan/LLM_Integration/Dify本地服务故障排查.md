# Dify 本地服务故障排查

## ❌ 错误：Reached maximum retries (0) for URL

这个错误表示 Dify 无法连接到本地服务。

---

## 🔍 故障排查步骤

### 步骤 1：检查服务是否运行

**方法 1：查看进程**

```bash
# Windows
tasklist | findstr python

# 或查看端口占用
netstat -an | findstr :8001
```

**方法 2：测试连接**

```bash
# 使用 curl 或 PowerShell
curl http://localhost:8001/health

# 或 PowerShell
Invoke-RestMethod -Uri "http://localhost:8001/health"
```

**方法 3：运行测试脚本**

```bash
cd "Auto Finan\LLM_Integration"
python test_dify_local_service.py
```

---

### 步骤 2：启动服务

如果服务未运行，启动它：

```bash
cd "Auto Finan\LLM_Integration"
start_dify_local_service.bat
```

应该看到：
```
🌐 服务地址: http://0.0.0.0:8001
📡 Excel 转提示词: http://0.0.0.0:8001/api/excel-to-prompt
```

---

### 步骤 3：检查端口占用

如果端口被占用，修改端口：

**方法 1：修改启动脚本**

编辑 `start_dify_local_service.bat`：
```batch
set SERVICE_PORT=8002  # 改为其他端口
```

**方法 2：命令行指定**

```bash
python dify_local_service.py --port 8002
```

然后在 Dify 中使用新端口。

---

### 步骤 4：检查防火墙

**Windows 防火墙**：

1. 打开"Windows Defender 防火墙"
2. 点击"高级设置"
3. 添加入站规则：
   - 端口：8001
   - 协议：TCP
   - 操作：允许连接

**或临时关闭防火墙测试**（仅用于测试）

---

### 步骤 5：检查网络配置

#### 如果 Dify 在本地

**URL 配置**：
```
http://localhost:8001/api/excel-to-prompt
```

#### 如果 Dify 在远程服务器

**1. 获取本地 IP**：
```bash
ipconfig
# 查找 IPv4 地址，例如：192.168.1.100
```

**2. 修改 Dify URL**：
```
http://192.168.1.100:8001/api/excel-to-prompt
```

**3. 确保网络可达**：
```bash
# 在 Dify 服务器上测试
ping 192.168.1.100
curl http://192.168.1.100:8001/health
```

---

## 🔧 常见问题

### 问题 1：服务启动失败

**错误**：`Address already in use`

**原因**：端口被占用

**解决**：
```bash
# 查找占用端口的进程
netstat -ano | findstr :8001

# 结束进程（替换 PID 为实际进程 ID）
taskkill /PID <PID> /F

# 或使用其他端口
python dify_local_service.py --port 8002
```

---

### 问题 2：连接被拒绝

**错误**：`Connection refused`

**原因**：
- 服务未运行
- 防火墙阻止
- 端口错误

**解决**：
1. 确保服务正在运行
2. 检查防火墙设置
3. 验证端口号

---

### 问题 3：超时

**错误**：`Timeout`

**原因**：
- Excel 文件较大
- 网络延迟

**解决**：
1. 增加超时时间（在 Dify HTTP 请求节点中）
2. 检查 Excel 文件大小
3. 优化网络连接

---

### 问题 4：Dify 在远程，无法访问本地服务

**解决方案**：

**方案 A：使用内网穿透**

1. **ngrok**：
   ```bash
   ngrok http 8001
   # 会生成公网 URL，例如：https://abc123.ngrok.io
   ```

2. **在 Dify 中使用**：
   ```
   https://abc123.ngrok.io/api/excel-to-prompt
   ```

**方案 B：VPN 连接**

- 将 Dify 服务器和本地机器连接到同一 VPN
- 使用内网 IP 地址

**方案 C：将服务部署到服务器**

- 将 `dify_local_service.py` 部署到 Dify 服务器
- 在服务器上运行服务

---

## ✅ 验证清单

- [ ] 服务正在运行（`start_dify_local_service.bat`）
- [ ] 端口 8001 未被占用
- [ ] 防火墙允许端口访问
- [ ] 可以访问健康检查端点（`http://localhost:8001/health`）
- [ ] Dify URL 配置正确
- [ ] 网络连接正常（如果 Dify 在远程）

---

## 🚀 快速测试

运行测试脚本：

```bash
cd "Auto Finan\LLM_Integration"
python test_dify_local_service.py
```

如果测试通过，说明服务正常，问题可能在 Dify 配置。

---

## 📝 Dify 配置检查

### 1. HTTP 请求节点配置

**URL**：
- 本地：`http://localhost:8001/api/excel-to-prompt`
- 远程：`http://你的IP:8001/api/excel-to-prompt`

**请求方法**：`POST`

**请求头**：
```json
{
  "Content-Type": "application/json"
}
```

**请求体**：
```json
{
  "excel_path": "{{#workflow.excel_path#}}",
  "sheet_name": "{{#workflow.sheet_name#}}",
  "serial": "{{#workflow.serial#}}"
}
```

**超时**：60 秒（或更长）

---

## 💡 调试技巧

### 1. 查看服务日志

服务启动后会显示请求日志，检查是否有错误。

### 2. 使用 Postman 或 curl 测试

```bash
curl -X POST http://localhost:8001/api/excel-to-prompt \
  -H "Content-Type: application/json" \
  -d "{\"excel_path\":\"C:\\\\Users\\\\FH\\\\source\\\\repos\\\\Auto Finan\\\\Auto Finan\\\\420财务050823.xlsx\",\"sheet_name\":\"3-报销\",\"serial\":\"1\"}"
```

### 3. 检查 Dify 日志

查看 Dify 的错误日志，了解具体错误信息。

---

## 📚 相关文件

- `dify_local_service.py` - 本地服务
- `start_dify_local_service.bat` - 启动脚本
- `test_dify_local_service.py` - 测试脚本
- `Dify使用本地服务配置.md` - 配置指南


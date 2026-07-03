---
name: ecm-crawl
description: >-
  爬取内网 OpenText Content Server（ECM，恒瑞 ecm.hengrui.com）指定目录下所有文件，
  导出文件名 + 节点链接 + 相对路径到 xlsx，供 SPAhub 等做文档树。当用户提到「爬取 ECM /
  OTCS / 内网文档目录 / 刷新 sharepoint_files.xlsx / 更新文档链接」，或从 SharePoint
  迁移到 ECM 时使用。核心难点是 ADFS 联合登录——直接用账号密码登录会失败，必须复用浏览器票据。
---

# 内网 ECM（OpenText Content Server）目录爬取

把 ECM 若干根目录节点下的所有文件（含子文件夹）导出为 xlsx，每个文件一行：
文件名 / ECM 节点链接 / 相对子目录。脚本在 `resources/sharepoint_file_list.py`，
输出 `resources/sharepoint_files.xlsx`（沿用旧名，SPAhub 直接读）。

## 最重要的一条：认证走「浏览器票据」，不要试图用密码登录

本 ECM 是 **ADFS 联合登录(SSO)**。已验证：OTCS 表单、OTDS、NTLM 直接用账号密码
**一定失败**（密码校验被联合到 AD FS，本地端点不认）。判断依据是 ADFS WS-Trust 会返回
`InvalidScope / ID3082`（作用域无效）而不是 `FailedAuthentication`——说明密码其实是对的，
只是没有可用的直连登录端点。

**唯一稳定的办法 = 复用浏览器已登录的会话票据（OTCSTicket）：**

1. 浏览器正常登录 ECM，打开任意 Smart View 页面。
2. `F12` → **Network** → 刷新 → 点一个发往 `/OTCS/cs.exe/api/...` 的请求 →
   **Request Headers** 里复制 `otcsticket` 的值（或 Application → Cookies →
   ecm.hengrui.com → 名为 `OTCSTicket` 的 cookie 值）。
3. 运行（**推荐 `--ticket` 参数，最不易踩坑**）：
   ```
   python resources/sharepoint_file_list.py --ticket <粘贴票据> --diagnose
   ```
   确认无误后去掉 `--diagnose` 正式跑。

票据有有效期，过期就重新复制一次。

### PowerShell 陷阱（高频踩坑）

PowerShell 里 **不要**用 `set ECM_TICKET=...`——那是 cmd 语法，在 PS 里只设成 PS 变量、
不是环境变量，Python `os.environ` 读不到，脚本会退回去问密码然后登录失败。正确写法：
- cmd：`set ECM_TICKET=xxx && python ...`
- PowerShell：`$env:ECM_TICKET="xxx"; python ...`
- 最省心：直接 `--ticket xxx` 参数，绕开环境变量。

## 配节点 ID：从 Smart View URL 直接读

每个目录的 URL 形如 `https://ecm.hengrui.com/OTCS/cs.exe/app/nodes/<node_id>`，
末尾数字就是 `node_id`。填进脚本 `TARGETS`：
```python
TARGETS = [
    ("sop",     31594212, "SOP"),
    ("wi",      31587503, "WI"),
    ("manuals", 31587312, "工具文件"),
]
```
⚠ 各 tab 的 node_id 必须彼此不同（历史上出过 sop/wi 复制粘贴成同一 ID 的 bug，
会导致两个 tab 内容一模一样）。改完先用 `--diagnose` 核对节点名和子项数是否对得上。

## REST API 要点（OTCS Content Server v2）

- 列子项：`GET /OTCS/cs.exe/api/v2/nodes/{id}/nodes?limit=200&page=N`
  - 认证头：`OTCSTicket: <ticket>`
  - 返回 `results[].data.properties`，取 `id` / `name` / `type` / `container`
  - **判断文件夹用 `container` 布尔值**（回退 `type==0`），不要只认 `type==0`，
    否则非 Folder 型容器会被当成文件、子目录漏爬
  - 分页看 `collection.paging.page_total`
- 单节点元数据：`GET .../api/v2/nodes/{id}`
- 文件浏览链接：`https://ecm.hengrui.com/OTCS/cs.exe/app/nodes/{id}`（Smart View 预览）

## 排错速查（`--diagnose` 会打印状态码 + 响应体）

- `401 Invalid username/password` 走密码登录 → 改用浏览器票据（方式 A）
- 爬取时 `401/403` → 票据过期/无权限；票据过期重新复制，个别子文件夹 403 属正常
- `404` → node_id 填错，回 Smart View URL 核对
- 爬到 0 个文件 → 脚本会**拒绝覆盖**旧 xlsx（防止清空可用数据），先跑 `--diagnose`

## 完成后

`sharepoint_files.xlsx`（sop / wi / manuals 三个 sheet）与 `weblink.xlsx` 同放在
`resources/` 下，下次启动 SPAhub 生效。这是离线刷新任务，ECM 增删文件后由维护人手动跑一次，
**不要**集成进 SPAhub GUI 触发（避免频繁网络爬取压 ECM 服务器）。

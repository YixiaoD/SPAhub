# -*- coding: utf-8 -*-
"""
ECM 文档目录爬取脚本（v5 — OpenText Content Server REST API，带诊断）
====================================================================

功能：
  遍历 SPAhub 关心的 ECM 目录节点下所有文件（含子文件夹），
  对每个文件提取：文件名 / NodeId / 相对子目录，并构造 ECM 节点浏览 URL。
  浏览器单击 URL → ECM Smart View 打开文件预览。

输出文件（脚本同目录）：sharepoint_files.xlsx，3 个 sheet：sop / wi / manuals
  每个 sheet 3 列：A=网页展示(文件名) B=网页链接(ECM URL) C=相对路径

------------------------------------------------------------------------
v5 相对 v4 的修复（针对“爬不到链接”问题）：
  1. sop / wi 之前误用同一个 node_id（复制粘贴错误），现改为在 TARGETS
     中独立配置，wi 的 node_id 必须由维护人填入正确值（见下方 !!! 标记）。
  2. 文件夹判定改用 properties.container（布尔），不再只认 type==0，
     避免非 Folder 型容器被当成文件、导致子目录漏爬。
  3. 关键：所有 HTTP 失败现在会打印【状态码 + 响应体】，认证失败会打印
     每种方式的服务器返回，方便一眼看出是 401(票据过期) / 403(无权限) /
     404(节点ID错) 还是认证根本没通过。之前这些信息全被吞掉了。
  4. 票据过期自动重认证一次再重试。
  5. 新增诊断模式：python sharepoint_file_list.py --diagnose
     只做认证 + 逐个根节点探测（打印节点名/类型/子项数），不写 xlsx，
     用来快速验证节点 ID、权限、认证是否正常。

使用：
    pip install openpyxl requests requests-ntlm
    # 正式爬取并覆写 xlsx：
    python sharepoint_file_list.py
    # 只诊断连通性 / 节点 / 权限（强烈建议第一次先跑这个）：
    python sharepoint_file_list.py --diagnose
  可用环境变量预置账号：set ECM_USER=... & set ECM_PASS=...
"""
import os
import sys
import getpass
import urllib3

try:
    from openpyxl import Workbook
except ImportError:
    print("请先安装依赖: pip install openpyxl")
    raise

try:
    import requests
except ImportError:
    print("请先安装依赖: pip install requests")
    raise

try:
    from requests_ntlm import HttpNtlmAuth
    HAS_NTLM = True
except ImportError:
    HAS_NTLM = False

# 禁用 SSL 警告（内网自签证书）
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================================
# 配置
# ============================================================================

ECM_BASE_URL = "https://ecm.hengrui.com"
ECM_API_BASE = ECM_BASE_URL + "/OTCS/cs.exe/api/v2"
ECM_AUTH_URL = ECM_BASE_URL + "/OTCS/cs.exe/api/v1/auth"
ECM_NODE_URL = ECM_BASE_URL + "/OTCS/cs.exe/app/nodes"

# (sheet_name, node_id, friendly_label)
# !!! 重要：sop 与 wi 必须是不同的 ECM 节点 ID。v4 里两者都写成了
# !!! 28329098（复制粘贴错误），请把 wi 改成 WI 目录真实的节点 ID。
TARGETS = [
    ("sop", 28329098, "SOP"),
    ("wi", 28329098, "WI"),   # <-- TODO: 替换为 WI 目录真实 node_id！
    ("manuals", 31587312, "工具文件"),
]

OUTPUT_FILENAME = "sharepoint_files.xlsx"

# 打印更多调试信息
DEBUG = True

# 单页大小
PAGE_SIZE = 200


def dbg(msg):
    if DEBUG:
        print(msg)


def _short(text, n=800):
    """截断响应体，避免刷屏"""
    text = text or ""
    return text if len(text) <= n else text[:n] + " ...(truncated)"


# ============================================================================
# 认证
# ============================================================================

def ecm_authenticate(session, username, password):
    """
    通过 ECM 认证，返回 OTCSTicket。
    依次尝试：OTCS 表单直连 → NTLM → OTDS → ADFS。
    每种方式失败都会打印服务器返回，方便定位。
    """
    # 方法1：OTCS 表单直接认证（api/v1/auth，最标准）——尝试多种用户名格式
    username_formats = [username]
    if "\\" not in username and "@" not in username:
        username_formats += [f"HENGRUI\\{username}", f"{username}@hengrui.com"]

    for uname in username_formats:
        print(f"  [OTCS 表单认证] 用户名={uname} ...")
        try:
            resp = session.post(
                ECM_AUTH_URL,
                data={"username": uname, "password": password},
                verify=False, timeout=30,
            )
            if resp.status_code == 200:
                ticket = (resp.json() or {}).get("ticket") or ""
                if ticket:
                    print(f"  ✔ OTCS 表单认证成功 (用户名={uname})")
                    return ticket
                print(f"  ✘ 返回 200 但无 ticket: {_short(resp.text)}")
            else:
                print(f"  ✘ HTTP {resp.status_code}: {_short(resp.text)}")
        except Exception as e:
            print(f"  ✘ 请求异常: {e}")

    # 方法2：NTLM（域账号，仅当服务器对该端点启用了集成 Windows 认证时有效）
    if HAS_NTLM:
        ntlm_user = username if "\\" in username else f"HENGRUI\\{username}"
        print(f"  [NTLM 认证] 用户名={ntlm_user} ...")
        try:
            resp = session.post(
                ECM_AUTH_URL, auth=HttpNtlmAuth(ntlm_user, password),
                verify=False, timeout=30,
            )
            if resp.status_code == 200:
                ticket = (resp.json() or {}).get("ticket") or ""
                if ticket:
                    print(f"  ✔ NTLM 认证成功 ({ntlm_user})")
                    return ticket
                print(f"  ✘ 返回 200 但无 ticket: {_short(resp.text)}")
            else:
                print(f"  ✘ HTTP {resp.status_code}: {_short(resp.text)}")
        except Exception as e:
            print(f"  ✘ 请求异常: {e}")

    # 方法3：OTDS 认证 → 换取 OTCS ticket
    otds_url = ECM_BASE_URL + "/otdsws/v1/authentication/credentials"
    print(f"  [OTDS 认证] {otds_url} ...")
    try:
        resp = session.post(
            otds_url,
            json={"user_name": username, "password": password},
            headers={"Content-Type": "application/json"},
            verify=False, timeout=30,
        )
        if resp.status_code == 200:
            j = resp.json() or {}
            token = j.get("ticket") or j.get("token") or ""
            if token:
                print("  OTDS 认证成功，尝试用 OTDSTicket 换取 OTCS ticket ...")
                otcs_resp = session.post(
                    ECM_AUTH_URL, data={"OTDSTicket": token},
                    verify=False, timeout=30,
                )
                if otcs_resp.status_code == 200:
                    ticket = (otcs_resp.json() or {}).get("ticket") or ""
                    if ticket:
                        print("  ✔ OTDS → OTCS ticket 成功")
                        return ticket
                    print(f"  ✘ 换票返回 200 但无 ticket: {_short(otcs_resp.text)}")
                else:
                    print(f"  ✘ 换票 HTTP {otcs_resp.status_code}: "
                          f"{_short(otcs_resp.text)}")
                # 有些环境 OTDS token 可直接作为票据使用
                print("  （回退：直接使用 OTDS token 作为票据）")
                return token
            print(f"  ✘ OTDS 返回 200 但无 token: {_short(resp.text)}")
        else:
            print(f"  ✘ OTDS HTTP {resp.status_code}: {_short(resp.text)}")
    except Exception as e:
        print(f"  ✘ OTDS 请求异常: {e}")

    # 方法4：ADFS WS-Trust（usernamemixed）→ SAML → 换 OTCS ticket
    print("  [ADFS WS-Trust 认证] ...")
    adfs_user = username if "\\" in username else f"HENGRUI\\{username}"
    adfs_wstrust_url = ("https://adfs.hengrui.com/adfs/services/trust/13"
                        "/usernamemixed")
    soap_envelope = f"""<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:a="http://www.w3.org/2005/08/addressing"
            xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
  <s:Header>
    <a:Action s:mustUnderstand="1">http://docs.oasis-open.org/ws-sx/ws-trust/200512/RST/Issue</a:Action>
    <a:To s:mustUnderstand="1">{adfs_wstrust_url}</a:To>
    <o:Security s:mustUnderstand="1"
       xmlns:o="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd">
      <o:UsernameToken>
        <o:Username>{adfs_user}</o:Username>
        <o:Password>{password}</o:Password>
      </o:UsernameToken>
    </o:Security>
  </s:Header>
  <s:Body>
    <trust:RequestSecurityToken xmlns:trust="http://docs.oasis-open.org/ws-sx/ws-trust/200512">
      <wsp:AppliesTo xmlns:wsp="http://schemas.xmlsoap.org/ws/2004/09/policy">
        <a:EndpointReference>
          <a:Address>{ECM_BASE_URL}</a:Address>
        </a:EndpointReference>
      </wsp:AppliesTo>
      <trust:RequestType>http://docs.oasis-open.org/ws-sx/ws-trust/200512/Issue</trust:RequestType>
      <trust:KeyType>http://docs.oasis-open.org/ws-sx/ws-trust/200512/Bearer</trust:KeyType>
    </trust:RequestSecurityToken>
  </s:Body>
</s:Envelope>"""
    try:
        resp = session.post(
            adfs_wstrust_url, data=soap_envelope.encode("utf-8"),
            headers={"Content-Type": "application/soap+xml; charset=utf-8"},
            verify=False, timeout=30,
        )
        if resp.status_code == 200 and "RequestedSecurityToken" in resp.text:
            import re
            match = re.search(
                r"<trust:RequestedSecurityToken>(.*?)</trust:RequestedSecurityToken>",
                resp.text, re.DOTALL,
            )
            if match:
                saml_token = match.group(1).strip()
                print("  ADFS token 获取成功，换取 OTCS ticket ...")
                otcs_resp = session.post(
                    ECM_AUTH_URL, data={"SAMLToken": saml_token},
                    verify=False, timeout=30,
                )
                if otcs_resp.status_code == 200:
                    ticket = (otcs_resp.json() or {}).get("ticket") or ""
                    if ticket:
                        print("  ✔ ADFS → OTCS ticket 成功")
                        return ticket
                print(f"  ✘ SAML→OTCS 换票失败 HTTP {otcs_resp.status_code}: "
                      f"{_short(otcs_resp.text)}")
        else:
            print(f"  ✘ ADFS WS-Trust HTTP {resp.status_code}: "
                  f"{_short(resp.text)}")
    except Exception as e:
        print(f"  ✘ ADFS 请求异常: {e}")

    raise RuntimeError(
        "所有认证方式均失败。请根据上面每种方式打印的服务器返回定位问题："
        "常见原因是账号/密码错、该账号无 REST API 权限、或 ECM 未启用对应认证端点。"
    )


# ============================================================================
# ECM 客户端（持有 session + ticket，支持票据过期自动重认证）
# ============================================================================

class EcmClient:
    def __init__(self, username, password):
        self.username = username
        self.password = password
        self.session = requests.Session()
        self.ticket = ecm_authenticate(self.session, username, password)

    def _headers(self):
        return {"OTCSTicket": self.ticket}

    def get(self, url, params=None, _retry=True):
        """带票据的 GET；遇 401/403 先重认证再重试一次。"""
        resp = self.session.get(
            url, headers=self._headers(), params=params,
            verify=False, timeout=60,
        )
        if resp.status_code in (401, 403) and _retry:
            print(f"  ! {resp.status_code} 疑似票据过期/无权限，重新认证后重试 ...")
            self.ticket = ecm_authenticate(self.session,
                                           self.username, self.password)
            return self.get(url, params=params, _retry=False)
        return resp

    def node_info(self, node_id):
        """取单个节点元数据（诊断用）。返回 properties dict 或 None。"""
        resp = self.get(f"{ECM_API_BASE}/nodes/{node_id}")
        if resp.status_code != 200:
            print(f"    ✘ 读取节点 {node_id} 失败 HTTP {resp.status_code}: "
                  f"{_short(resp.text, 400)}")
            return None
        data = resp.json() or {}
        # v2: {"results": {"data": {"properties": {...}}}}  (单节点)
        results = data.get("results") or {}
        if isinstance(results, list):
            results = results[0] if results else {}
        return results.get("data", {}).get("properties", {})


def build_ecm_url(node_id):
    """构造 ECM 节点浏览 URL（Smart View 预览）"""
    return f"{ECM_NODE_URL}/{node_id}"


def collect_files_from_ecm(client, node_id, base_path=""):
    """
    递归遍历 ECM 节点下所有文件，返回 [(name, url, sub_dir), ...]
    base_path 用于计算 sub_dir（根目录文件 sub_dir=""）。
    """
    out = []
    page = 1
    while True:
        resp = client.get(
            f"{ECM_API_BASE}/nodes/{node_id}/nodes",
            params={"limit": PAGE_SIZE, "page": page},
        )
        if resp.status_code != 200:
            # 不再静默：打印失败原因（节点ID/权限/票据）
            raise RuntimeError(
                f"列出节点 {node_id} 子项失败 HTTP {resp.status_code}: "
                f"{_short(resp.text, 400)}"
            )
        data = resp.json() or {}
        results = data.get("results") or []
        if not results:
            break

        for item in results:
            props = item.get("data", {}).get("properties", {})
            name = (props.get("name") or "").strip()
            child_id = props.get("id")
            if not name or child_id is None:
                continue

            # 关键修复：优先用 container 布尔判断是否为文件夹/容器；
            # 回退到 type==0（Folder）。
            is_container = props.get("container")
            if is_container is None:
                is_container = (props.get("type") == 0)

            if is_container:
                sub_path = f"{base_path}/{name}" if base_path else name
                out.extend(collect_files_from_ecm(client, child_id, sub_path))
            else:
                out.append((name, build_ecm_url(child_id), base_path))

        # 分页：优先用服务器返回的 page_total
        paging = data.get("collection", {}).get("paging", {})
        page_total = paging.get("page_total")
        if page_total is not None:
            if page >= page_total:
                break
        else:
            total = paging.get("total_count", 0)
            if page * PAGE_SIZE >= total:
                break
        page += 1

    return out


# ============================================================================
# xlsx 写入
# ============================================================================

def write_xlsx(output_path, data_by_sheet):
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_name, _node_id, _label in TARGETS:
        rows = data_by_sheet.get(sheet_name, [])
        ws = wb.create_sheet(title=sheet_name)
        ws.append(["网页展示", "网页链接", "相对路径"])
        ws.column_dimensions["A"].width = 60
        ws.column_dimensions["B"].width = 100
        ws.column_dimensions["C"].width = 30
        for name, url, sub_dir in sorted(
                rows, key=lambda x: (x[2] or "", x[0].lower())):
            ws.append([name, url, sub_dir])
    wb.save(output_path)


# ============================================================================
# 诊断模式
# ============================================================================

def run_diagnose(client):
    print("\n" + "=" * 60)
    print("诊断模式：逐个根节点探测（不写 xlsx）")
    print("=" * 60)
    seen = {}
    for sheet_name, node_id, label in TARGETS:
        print(f"\n[{sheet_name}] {label}  node_id={node_id}")
        if node_id in seen:
            print(f"  ⚠ 警告：与 [{seen[node_id]}] 使用了相同的 node_id，"
                  f"两个 tab 会得到完全相同的内容（很可能是配置错误）")
        seen[node_id] = sheet_name

        props = client.node_info(node_id)
        if props is None:
            continue
        print(f"  节点名: {props.get('name')}  type={props.get('type')} "
              f"container={props.get('container')}")
        # 探测第一页子项数量
        resp = client.get(f"{ECM_API_BASE}/nodes/{node_id}/nodes",
                          params={"limit": 5, "page": 1})
        if resp.status_code != 200:
            print(f"  ✘ 列子项失败 HTTP {resp.status_code}: "
                  f"{_short(resp.text, 400)}")
            continue
        data = resp.json() or {}
        paging = data.get("collection", {}).get("paging", {})
        print(f"  子项总数(total_count)={paging.get('total_count')}  "
              f"前几项：")
        for item in (data.get("results") or [])[:5]:
            p = item.get("data", {}).get("properties", {})
            print(f"    - {p.get('name')}  (id={p.get('id')} "
                  f"type={p.get('type')} container={p.get('container')})")
    print("\n诊断结束。若上面能看到正确的节点名和子项，说明认证/权限/节点 ID 均正常，")
    print("可去掉 --diagnose 正式运行。若某项报错，请按状态码定位（401票据/403权限/404节点）。")


# ============================================================================
# 主流程
# ============================================================================

def main():
    diagnose = "--diagnose" in sys.argv

    print("=" * 60)
    print("ECM 文档目录爬取工具 — SPAhub 数据源刷新  (v5)")
    print("=" * 60)

    username = os.environ.get("ECM_USER") or input(
        "请输入 ECM 用户名（如 DOMAIN\\username）: ")
    password = os.environ.get("ECM_PASS") or getpass.getpass("请输入密码: ")

    print("\n正在连接 ECM 并认证...")
    try:
        client = EcmClient(username, password)
        print("认证成功，已获取 OTCSTicket")
    except Exception as e:
        print(f"\n认证失败: {e}")
        sys.exit(1)

    if diagnose:
        run_diagnose(client)
        return

    data_by_sheet = {}
    total = 0
    for sheet_name, node_id, label in TARGETS:
        print(f"\n爬取目录: {label}  节点 ID: {node_id}")
        print(f"  URL: {build_ecm_url(node_id)}")
        try:
            items = collect_files_from_ecm(client, node_id)
            data_by_sheet[sheet_name] = items
            print(f"  找到 {len(items)} 个文件 → sheet '{sheet_name}'")
            from collections import Counter
            for sd, n in Counter(
                    s[2] or "(根目录)" for s in items).most_common():
                print(f"    └─ {sd}: {n}")
            total += len(items)
        except Exception as e:
            print(f"  ✘ 爬取失败: {e}")
            data_by_sheet[sheet_name] = []

    if total == 0:
        print("\n⚠ 一个文件都没爬到，未覆盖旧的 xlsx（避免清空可用数据）。")
        print("  请先用 python sharepoint_file_list.py --diagnose 排查。")
        sys.exit(2)

    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_path = os.path.join(script_dir, OUTPUT_FILENAME)
    tmp_path = output_path + ".tmp"
    write_xlsx(tmp_path, data_by_sheet)
    os.replace(tmp_path, output_path)

    print(f"\n{'=' * 60}")
    print(f"完成！共写入 {total} 个文件到\n  {output_path}")
    for sheet_name, _, label in TARGETS:
        print(f"  {sheet_name:10s} ({label}): "
              f"{len(data_by_sheet.get(sheet_name, []))} 行")
    print(f"{'=' * 60}")
    print("\n下一步：把 sharepoint_files.xlsx 与 weblink.xlsx 同放到 "
          "SPAhub/resources/ 下，下次启动 SPAhub 生效。")


if __name__ == "__main__":
    main()

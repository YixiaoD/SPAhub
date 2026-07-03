# -*- coding: utf-8 -*-
"""
ECM 文档目录爬取脚本（v6 — OpenText Content Server REST API）
============================================================

功能：
  遍历 SPAhub 关心的 ECM 目录节点下所有文件（含子文件夹），
  对每个文件提取：文件名 / NodeId / 相对子目录，并构造 ECM 节点浏览 URL。
  浏览器单击 URL → ECM Smart View 打开文件预览。

输出文件（脚本同目录）：sharepoint_files.xlsx，3 个 sheet：sop / wi / manuals
  每个 sheet 3 列：A=网页展示(文件名) B=网页链接(ECM URL) C=相对路径

------------------------------------------------------------------------
关于认证（重要！先读这段）
------------------------------------------------------------------------
诊断发现：本 ECM 走的是 ADFS 联合登录(SSO)——直接拿账号密码去 OTCS/OTDS
登录会被拒（密码校验被联合到了 AD FS，本地端点不认）。因此脚本支持三种方式，
按可靠度从高到低：

  【方式 A：直接用浏览器票据（最稳，强烈推荐）】
    1. 用浏览器正常登录 ECM（打开任意 Smart View 页面）。
    2. F12 打开开发者工具 → Network（网络）标签 → 刷新页面 →
       随便点一个发往 /OTCS/cs.exe/api/... 的请求 → 看 Request Headers，
       复制其中 otcsticket 的值；
       （或 Application → Cookies → ecm.hengrui.com → 复制名为
        OTCSTicket 的 cookie 值。）
    3. 运行：
         Windows:  set ECM_TICKET=粘贴票据 && python sharepoint_file_list.py
         或：      python sharepoint_file_list.py --ticket 粘贴票据
       脚本会直接用这个票据，跳过所有登录。票据有有效期，过期再复制一次即可。

  【方式 B：OTDS 账号密码（若你们 OTDS 支持直接密码认证，脚本会自动尝试）】
    set ECM_USER=liup71 && set ECM_PASS=你的密码 && python sharepoint_file_list.py
    （v6 已修正 OTDS 请求字段 userName，v5 之前写成了 user_name 会必失败。）

  【方式 C：ADFS/OTCS 表单（多为联合登录环境，通常不可用，仅兜底尝试）】

  先跑诊断确认认证与节点是否 OK：
    python sharepoint_file_list.py --diagnose
------------------------------------------------------------------------

依赖：pip install openpyxl requests
"""
import os
import sys
import re
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

# 禁用 SSL 警告（内网自签证书）
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# ============================================================================
# 配置
# ============================================================================

ECM_BASE_URL = "https://ecm.hengrui.com"
ECM_HOST = "ecm.hengrui.com"
ECM_API_BASE = ECM_BASE_URL + "/OTCS/cs.exe/api/v2"
ECM_AUTH_URL = ECM_BASE_URL + "/OTCS/cs.exe/api/v1/auth"
ECM_NODE_URL = ECM_BASE_URL + "/OTCS/cs.exe/app/nodes"
OTDS_CRED_URL = ECM_BASE_URL + "/otdsws/v1/authentication/credentials"

# (sheet_name, node_id, friendly_label)
# 节点 ID 由维护人从 ECM Smart View URL 中读取（.../app/nodes/<id>）。
TARGETS = [
    ("sop", 31594212, "SOP"),        # https://ecm.hengrui.com/OTCS/cs.exe/app/nodes/31594212
    ("wi", 31587503, "WI"),          # https://ecm.hengrui.com/OTCS/cs.exe/app/nodes/31587503
    ("manuals", 31587312, "工具文件"),  # 沿用原值，如有变更请一并更新
]

OUTPUT_FILENAME = "sharepoint_files.xlsx"
DEBUG = True
PAGE_SIZE = 200


def dbg(msg):
    if DEBUG:
        print(msg)


def _short(text, n=800):
    text = text or ""
    return text if len(text) <= n else text[:n] + " ...(truncated)"


# ============================================================================
# 认证
# ============================================================================

def _username_formats(username):
    formats = [username]
    if "\\" not in username and "@" not in username:
        formats += [f"HENGRUI\\{username}", f"{username}@hengrui.com"]
    return formats


def otcs_form_login(session, username, password):
    """OTCS 表单认证 api/v1/auth，返回 ticket 或 None。"""
    for uname in _username_formats(username):
        print(f"  [OTCS 表单认证] 用户名={uname} ...")
        try:
            resp = session.post(
                ECM_AUTH_URL, data={"username": uname, "password": password},
                verify=False, timeout=30,
            )
            if resp.status_code == 200:
                ticket = (resp.json() or {}).get("ticket") or ""
                if ticket:
                    print(f"  ✔ OTCS 表单认证成功 (用户名={uname})")
                    return ticket
                print(f"  ✘ 200 但无 ticket: {_short(resp.text)}")
            else:
                print(f"  ✘ HTTP {resp.status_code}: {_short(resp.text)}")
        except Exception as e:
            print(f"  ✘ 请求异常: {e}")
    return None


def otds_login(session, username, password):
    """
    OTDS 认证 → 换 OTCS ticket。
    v6 关键修复：请求体字段用 userName（驼峰），v5 之前写成 user_name，
    OTDS 收到空用户名必然返回 INVALID_CREDENTIALS。
    返回 OTCS ticket 或特殊值 "COOKIE"（表示已用 cookie 建立会话）或 None。
    """
    for uname in _username_formats(username):
        print(f"  [OTDS 认证] userName={uname} ...")
        try:
            resp = session.post(
                OTDS_CRED_URL,
                json={"userName": uname, "password": password},
                headers={"Content-Type": "application/json"},
                verify=False, timeout=30,
            )
        except Exception as e:
            print(f"  ✘ OTDS 请求异常: {e}")
            continue

        if resp.status_code != 200:
            print(f"  ✘ OTDS HTTP {resp.status_code}: {_short(resp.text)}")
            continue

        j = resp.json() or {}
        otds_ticket = j.get("ticket") or j.get("token") or ""
        if not otds_ticket:
            print(f"  ✘ OTDS 200 但无 ticket: {_short(resp.text)}")
            continue
        print(f"  OTDS 认证成功 (userName={uname})，换取 OTCS 会话 ...")

        # 换法1：把 OTDSTicket 作为表单参数交给 OTCS
        try:
            r = session.post(ECM_AUTH_URL, data={"OTDSTicket": otds_ticket},
                             verify=False, timeout=30)
            if r.status_code == 200:
                t = (r.json() or {}).get("ticket") or ""
                if t:
                    print("  ✔ OTDS → OTCS ticket 成功（表单换票）")
                    return t
        except Exception as e:
            print(f"  （表单换票异常，忽略）: {e}")

        # 换法2：把 OTDSTicket 塞进 cookie，让 OTCS SSO 直接放行
        session.cookies.set("OTDSTicket", otds_ticket, domain=ECM_HOST)
        probe = session.get(f"{ECM_API_BASE}/nodes/{TARGETS[0][1]}",
                            verify=False, timeout=30)
        if probe.status_code == 200:
            print("  ✔ OTDS cookie 会话可用（cookie 模式）")
            return "COOKIE"
        print(f"  ✘ OTDS cookie 探测失败 HTTP {probe.status_code}: "
              f"{_short(probe.text, 300)}")
    return None


def _discover_saml_entityids(session):
    """尝试从 OTDS / OTCS 元数据发现 SAML SP entityID，用作 ADFS AppliesTo。"""
    found = []
    candidates = [
        ECM_BASE_URL + "/otdsws/login?metadata",
        ECM_BASE_URL + "/otdsws/saml2/metadata",
        ECM_BASE_URL + "/OTCS/cs.exe?func=saml2.metadata",
        ECM_BASE_URL + "/OTCS/cs.exe?func=saml.spmetadata",
    ]
    for url in candidates:
        try:
            r = session.get(url, verify=False, timeout=15)
            if r.status_code == 200 and "entityID" in r.text:
                for m in re.findall(r'entityID="([^"]+)"', r.text):
                    if m not in found:
                        found.append(m)
                        print(f"  发现 SAML entityID: {m}")
        except Exception:
            pass
    return found


def adfs_login(session, username, password):
    """
    ADFS WS-Trust usernamemixed：密码通常在此被接受（诊断显示 InvalidScope
    而非 FailedAuthentication）。逐个 AppliesTo 候选尝试，拿到 SAML 后换 OTCS。
    返回 OTCS ticket / "COOKIE" / None。
    """
    adfs_user = username if "\\" in username else f"HENGRUI\\{username}"
    adfs_url = ("https://adfs.hengrui.com/adfs/services/trust/13/usernamemixed")

    applies_to_list = _discover_saml_entityids(session) + [
        ECM_BASE_URL + "/otdsws/login",
        ECM_BASE_URL + "/otdsws",
        ECM_BASE_URL + "/OTCS/cs.exe",
        ECM_BASE_URL + "/OTCS/cs.exe/",
        ECM_BASE_URL,
    ]

    for applies_to in applies_to_list:
        soap = f"""<?xml version="1.0" encoding="utf-8"?>
<s:Envelope xmlns:s="http://www.w3.org/2003/05/soap-envelope"
            xmlns:a="http://www.w3.org/2005/08/addressing"
            xmlns:u="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd">
  <s:Header>
    <a:Action s:mustUnderstand="1">http://docs.oasis-open.org/ws-sx/ws-trust/200512/RST/Issue</a:Action>
    <a:To s:mustUnderstand="1">{adfs_url}</a:To>
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
        <a:EndpointReference><a:Address>{applies_to}</a:Address></a:EndpointReference>
      </wsp:AppliesTo>
      <trust:RequestType>http://docs.oasis-open.org/ws-sx/ws-trust/200512/Issue</trust:RequestType>
      <trust:KeyType>http://docs.oasis-open.org/ws-sx/ws-trust/200512/Bearer</trust:KeyType>
    </trust:RequestSecurityToken>
  </s:Body>
</s:Envelope>"""
        print(f"  [ADFS WS-Trust] AppliesTo={applies_to} ...")
        try:
            resp = session.post(
                adfs_url, data=soap.encode("utf-8"),
                headers={"Content-Type": "application/soap+xml; charset=utf-8"},
                verify=False, timeout=30,
            )
        except Exception as e:
            print(f"  ✘ ADFS 请求异常: {e}")
            continue

        if "FailedAuthentication" in resp.text or "ID3242" in resp.text:
            print("  ✘ ADFS 拒绝：用户名或密码错误。")
            return None  # 密码错，换 AppliesTo 也没用
        if resp.status_code == 200 and "RequestedSecurityToken" in resp.text:
            m = re.search(
                r"<trust:RequestedSecurityToken>(.*?)</trust:RequestedSecurityToken>",
                resp.text, re.DOTALL)
            if m:
                saml = m.group(1).strip()
                print("  ✔ 拿到 SAML token，尝试换 OTCS ...")
                # 换法1：表单 SAMLToken
                try:
                    r = session.post(ECM_AUTH_URL, data={"SAMLToken": saml},
                                     verify=False, timeout=30)
                    if r.status_code == 200:
                        t = (r.json() or {}).get("ticket") or ""
                        if t:
                            print("  ✔ ADFS → OTCS ticket 成功")
                            return t
                    print(f"  ✘ SAML→OTCS 换票 HTTP {r.status_code}: "
                          f"{_short(r.text, 300)}")
                except Exception as e:
                    print(f"  ✘ 换票异常: {e}")
                print("  （SAML 换 OTCS 失败，多需管理员配好 ACS；建议改用方式 A 票据）")
                return None
        if "InvalidScope" in resp.text or "ID3082" in resp.text:
            print("  ✘ 该 AppliesTo 作用域无效，换下一个候选 ...")
            continue
        print(f"  ✘ ADFS 未知返回 HTTP {resp.status_code}: {_short(resp.text, 300)}")
    return None


def authenticate(session, username, password):
    """按 B(OTDS) → C(OTCS表单/ADFS) 顺序尝试，返回 ticket 字符串或 'COOKIE'。"""
    t = otds_login(session, username, password)
    if t:
        return t
    t = otcs_form_login(session, username, password)
    if t:
        return t
    t = adfs_login(session, username, password)
    if t:
        return t
    raise RuntimeError(
        "自动登录全部失败。本 ECM 多为 ADFS 联合登录，"
        "请改用【方式 A：浏览器票据】——见脚本顶部说明，或运行 "
        "python sharepoint_file_list.py --ticket <浏览器复制的OTCSTicket>"
    )


# ============================================================================
# ECM 客户端
# ============================================================================

class EcmClient:
    def __init__(self, ticket=None, username=None, password=None):
        self.session = requests.Session()
        self.username = username
        self.password = password
        if ticket:
            # 方式 A：直接使用浏览器票据
            self.ticket = ticket
            print("  使用手动提供的 OTCSTicket（方式 A）。")
        else:
            # 方式 B/C：自动登录
            self.ticket = authenticate(self.session, username, password)

    def _headers(self):
        # cookie 模式下 self.ticket=="COOKIE"，靠 session cookie 认证，不发 header
        if self.ticket and self.ticket != "COOKIE":
            return {"OTCSTicket": self.ticket}
        return {}

    def _reauth(self):
        if self.username and self.password:
            print("  ! 重新认证 ...")
            self.ticket = authenticate(self.session, self.username, self.password)
            return True
        print("  ! 手动票据已过期，请重新从浏览器复制 OTCSTicket 后再跑。")
        return False

    def get(self, url, params=None, _retry=True):
        resp = self.session.get(url, headers=self._headers(), params=params,
                                verify=False, timeout=60)
        if resp.status_code in (401, 403) and _retry and self._reauth():
            return self.get(url, params=params, _retry=False)
        return resp

    def node_info(self, node_id):
        resp = self.get(f"{ECM_API_BASE}/nodes/{node_id}")
        if resp.status_code != 200:
            print(f"    ✘ 读取节点 {node_id} 失败 HTTP {resp.status_code}: "
                  f"{_short(resp.text, 400)}")
            return None
        results = (resp.json() or {}).get("results") or {}
        if isinstance(results, list):
            results = results[0] if results else {}
        return results.get("data", {}).get("properties", {})


def build_ecm_url(node_id):
    return f"{ECM_NODE_URL}/{node_id}"


def collect_files_from_ecm(client, node_id, base_path=""):
    out = []
    page = 1
    while True:
        resp = client.get(f"{ECM_API_BASE}/nodes/{node_id}/nodes",
                          params={"limit": PAGE_SIZE, "page": page})
        if resp.status_code != 200:
            raise RuntimeError(
                f"列出节点 {node_id} 子项失败 HTTP {resp.status_code}: "
                f"{_short(resp.text, 400)}")
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
            is_container = props.get("container")
            if is_container is None:
                is_container = (props.get("type") == 0)
            if is_container:
                sub = f"{base_path}/{name}" if base_path else name
                out.extend(collect_files_from_ecm(client, child_id, sub))
            else:
                out.append((name, build_ecm_url(child_id), base_path))

        paging = data.get("collection", {}).get("paging", {})
        page_total = paging.get("page_total")
        if page_total is not None:
            if page >= page_total:
                break
        else:
            if page * PAGE_SIZE >= paging.get("total_count", 0):
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
# 诊断
# ============================================================================

def run_diagnose(client):
    print("\n" + "=" * 60)
    print("诊断模式：逐个根节点探测（不写 xlsx）")
    print("=" * 60)
    seen = {}
    for sheet_name, node_id, label in TARGETS:
        print(f"\n[{sheet_name}] {label}  node_id={node_id}")
        if node_id in seen:
            print(f"  ⚠ 与 [{seen[node_id]}] 用了相同 node_id，两个 tab 内容会一样"
                  f"（很可能是配置错误）")
        seen[node_id] = sheet_name
        props = client.node_info(node_id)
        if props is None:
            continue
        print(f"  节点名: {props.get('name')}  type={props.get('type')} "
              f"container={props.get('container')}")
        resp = client.get(f"{ECM_API_BASE}/nodes/{node_id}/nodes",
                          params={"limit": 5, "page": 1})
        if resp.status_code != 200:
            print(f"  ✘ 列子项失败 HTTP {resp.status_code}: "
                  f"{_short(resp.text, 400)}")
            continue
        data = resp.json() or {}
        paging = data.get("collection", {}).get("paging", {})
        print(f"  子项总数={paging.get('total_count')}  前几项：")
        for item in (data.get("results") or [])[:5]:
            p = item.get("data", {}).get("properties", {})
            print(f"    - {p.get('name')}  (id={p.get('id')} "
                  f"type={p.get('type')} container={p.get('container')})")
    print("\n诊断结束。若能看到正确节点名与子项，说明认证/权限/节点 ID 均正常。")


# ============================================================================
# 主流程
# ============================================================================

def _get_arg_ticket():
    if "--ticket" in sys.argv:
        i = sys.argv.index("--ticket")
        if i + 1 < len(sys.argv):
            return sys.argv[i + 1]
    return os.environ.get("ECM_TICKET")


def main():
    diagnose = "--diagnose" in sys.argv
    ticket = _get_arg_ticket()

    print("=" * 60)
    print("ECM 文档目录爬取工具 — SPAhub 数据源刷新  (v6)")
    print("=" * 60)

    if ticket:
        print("检测到 OTCSTicket（方式 A），跳过账号密码登录。")
        client = EcmClient(ticket=ticket)
    else:
        username = os.environ.get("ECM_USER") or input(
            "请输入 ECM 用户名（如 liup71）: ")
        password = os.environ.get("ECM_PASS") or getpass.getpass("请输入密码: ")
        print("\n正在连接 ECM 并认证...")
        try:
            client = EcmClient(username=username, password=password)
            print("认证成功。")
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
        print("\n⚠ 一个文件都没爬到，未覆盖旧 xlsx（避免清空可用数据）。")
        print("  请先 python sharepoint_file_list.py --diagnose 排查。")
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

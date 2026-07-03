# -*- coding: utf-8 -*-
"""
gui/shortcuts.py
================
Shortcuts 页：上方双击启动本地外部工具，下方一站式网页导航。

v12.10.54：
  - 本地工具按钮缩小（120×120 → 88×88，icon 64→44），spacing 30→18，
    给下方网页导航区让出版面
  - 新增 WebNavWidget：从 SPAhub/resources/weblink.xlsx 读取链接清单，
    支持搜索 + A-Z 字母索引（含中文拼音首字母）+ 单击在系统默认浏览器中打开。

v12.10.55：
  - 中文首字母改用 GB2312 区位法（_first_letter），零第三方依赖，对
    常用现代汉字 100% 准确；之前依赖 pypinyin，环境未装时全部退到 '#'。
  - 删除排序下拉 UI，固定 A-Z 排序。
  - 新增收藏（书签）功能：
      * 卡片右侧星星图标，hover 时显示，单击切换收藏态（黄/灰）
      * 收藏置顶在字母索引下方"⭐ 收藏"区块
      * 收藏数据持久化到 %LOCALAPPDATA%\\SPAhub\\weblinks_bookmarks.json
        （与 v12.10.47 的 logs 路径同根；Linux/Mac 走 ~/.SPAhub/）
  - 卡片改为自定义 LinkCard QWidget（QLabel + 星星 QToolButton），
    主体接收点击打开链接，星星单独接收点击切换收藏。

v12.10.56：
  - 修复卡片与主背景融为一体的视觉问题：改用"浅灰圆角分组容器 +
    白色无边框卡片"的双层结构 — 容器底色 #EBEEF5（比主背景 #F0F2F5
    深一档），白色卡片在其上自然跳出，无需边框线
  - 卡片间距 8 → 10px；容器内边距 10px；删除收藏区与全部区之间冗余
    的 QFrame.HLine 分隔线（双容器本身已形成视觉分组）
  - hover 时卡片才显示蓝边框 + 蓝底，强化点击反馈

v12.10.57：
  - 导航区改为 QTabWidget 4 tab：一站式导航 / 系统性工具手册 / WI 定稿版 / SOP
  - WebNavWidget 增加 sheet_name + simple_mode + title + hint 参数：
      * sheet_name：从多 sheet xlsx 按名取数（默认空 → 老逻辑）
      * simple_mode：True 时隐藏字母索引行 + 不分桶收藏，给 SOP/WI/手册
        这种小数据量 tab 用，UI 更聚焦（保留搜索 + 卡片列表）
  - Tab 1 数据：resources/weblink.xlsx（手工维护，39 条网页链接）
  - Tab 2/3/4 数据：resources/ecm_files.xlsx 的 manuals/wi/sop 三个
    sheet（由 resources/ecm_file_list.py 离线爬取生成）—
    不在 GUI 里给刷新入口，避免用户频繁触发 ECM 网络请求
  - _resolve_weblinks_path → _resolve_resource_path(filename) 通用化

v12.10.58：
  - simple_mode 改成单列文档树 + 子文件夹可折叠：
      * 数据多一列 sub_dir（来自 ecm_files.xlsx 第 3 列"相对路径"）
      * 根目录文件直接列出（单列 QVBoxLayout）
      * 子目录文件归到 CollapsibleSection（点击 ▾/▸ 切换展开收起）
      * 搜索时自动展开匹配的 section（避免文件躲在折叠里）
  - 新增 CollapsibleSection 类：标题栏（▾ 子目录名 (count)）+ 内容区
  - LinkCard 加 show_star 参数：simple_mode 下不创建星星 widget（更简洁）
  - _read_weblinks 返回三元组 (name, url, sub_dir)；旧版 2 列 xlsx 向后兼容
    （sub_dir 全部为空 → 单列卡片列表无折叠，行为等价 v12.10.57）
  - 搜索匹配范围扩展到 sub_dir（"popPK" 这种按子目录搜更直观）

  对应的 SharePoint 爬取脚本 v3：从 ServerRelativeUrl 直链 URL（下载行为）
  改成 WopiFrame.aspx 预览 URL（Office Online 在线预览），并额外输出
  sub_dir 列保留文件夹层级信息。

v12.10.59：
  - 删除 WebNavWidget 内部大标题（与 QTabWidget tab label 重复）；
    title 参数保留兼容但不再渲染，只保留 hint 副标题
  - 所有 4 个 tab 都支持收藏（之前 simple_mode 把 _favorites 强制设空）：
      * LinkCard 全部 show_star=True，卡片悬停可点 ☆ 收藏
      * 收藏数据共享同一份 weblinks_bookmarks.json（url 是文件唯一 key）
      * simple_mode 下收藏区强制平铺（不按 sub_dir 折叠），收藏数量少
        平铺更直观；下方"全部链接"区仍按 sub_dir 走 CollapsibleSection
  - SOP/WI/手册 tab 的 hint 文案改成与 Tab 1 一致的收藏提示

v12.10.60：
  - 所有 4 个 tab 布局完全统一：单列文档树 + 仅搜索（删除字母索引）
      * 一站式导航 也变成单列（不再 3 列网格），可滚动浏览
      * 一站式导航 字母索引行（"全部 C D E F ..."）整行删除
      * 4 tab 渲染走同一套 _add_grid_to_layout 代码路径
        （sub_dir 全空时 = 单列纯卡片列表；有 sub_dir 时 = 单列 +
         CollapsibleSection 折叠子目录）
  - WebNavWidget 删除 LINKS_PER_ROW / _filter_letter / _letter_buttons
    / _rebuild_letter_bar / _add_letter_btn / _on_letter_clicked /
    _set_active_letter — 总计 60+ 行字母索引代码被清除
  - simple_mode / title 参数标记 deprecated 保留兼容（外部传值无作用）
  - ShortcutWidget 构造 WebNavWidget 时不再传 simple_mode / title
  - hint 文案统一为"单击打开 · 鼠标悬停点击 ☆ 收藏(同步保存到本地)"

v12.10.61：
  - hint 文案去掉"(同步保存到本地)"括号说明 — 用户反馈"同步保存到本地"
    可能让人误以为有数据上传/同步动作。实际上收藏只是写入本地的
    weblinks_bookmarks.json，"同步"二字反而引起误解。
  - 最终文案："单击打开 · 鼠标悬停点击 ☆ 收藏"

v12.10.62：
  - 暂时隐藏"一站式导航"tab — Shortcuts 页只剩 3 个 SharePoint 文档库
    tab：系统性工具手册 / WI 定稿版 / SOP
  - addTab 代码块整段注释保留（不删除），WebNavWidget 构造样板原样保留，
    weblinks_path 也仍计算，未来恢复一站式导航只需取消注释 4 行
  - self.web_nav 占位置为 None（外部如有引用判 None 即可，
    本仓库内部并无引用）
  - resources/weblink.xlsx 文件保留不动（隐藏 tab 期间数据无需迁移）

v12.10.63：
  - 顶部工具按钮增加 Fork (Git GUI) + Pinnacle 21 Community (CDISC 校验)
    两个用户级本地工具：
      * 路径模板用 _local_appdata() 取 %LOCALAPPDATA% 自动适配当前
        机器的 Windows 域账号 — Fork:    %LOCALAPPDATA%\\Fork\\Fork.exe
                                Pinnacle: %LOCALAPPDATA%\\Programs\\Pinnacle 21 Community\\Pinnacle 21 Community.exe
      * exe 不存在时按钮显示 ⚠ 标识 + 虚线边框 + 暗灰色，hover 时左下角
        状态栏显示预期安装父目录（"请检查安装路径：C:\\Users\\xxx\\..."），
        且 tooltip 也设上；双击会弹完整路径提示而非真启动
  - IconLoaderThread 加 ico 缓存：
      * 缓存文件：%LOCALAPPDATA%\\SPAhub\\icon_cache.json
      * 启动时先读缓存，命中且 ico 文件还在 → 直接 emit，跳过 walk
      * 未命中时递归扫描（最大深度 3 层），按优先级选最佳 ico：
          0 = 与 exe 同名（如 Fork.ico）— 命中立即返回，不再扫
          1 = 包含 exe 名（如 fork_app.ico）
          10+depth = 其它（按深度排序）
      * 找到后写缓存（原子 tmp + rename），下次启动直接读
      * exe 不存在的条目跳过扫描，缓存里若有旧值清掉
  - 新增 _HoverHintFilter 类 — exe 不存在的按钮装上，hover 时通过
    主窗口状态栏告知找寻路径
  - _launch_app 加 exe 存在性预检查 — 不存在时弹 warning 对话框而非
    让 AppLaunchThread 抛 FileNotFoundError

v12.10.64：
  - IconLoaderThread 加"系统图标兜底"机制 — 解决 Fork 等 Electron 应用
    把图标内嵌在 PE 资源、没有外部 .ico 的情况：
      * 外部 .ico 扫描失败时，emit 新信号 use_system_icon(name, exe_path)
      * 主线程收到后调 QFileIconProvider 从 exe 资源提取系统图标挂上
      * 与 Windows 资源管理器显示该 exe 时用的图标完全一致
        （内部走同一个 SHGetFileInfo API）
  - 新增 ShortcutWidget.update_icon_from_exe 槽函数（主线程执行 Qt
    GUI 操作的限制要求）
  - 零硬编码 — 不为 Fork 单独维护 .ico 资源，所有未找到外部 .ico 的
    exe 都自动兜底（Pinnacle 21 / PDT / PFN / QCT 全覆盖）

v12.10.89：
  - 本地工具区新增"审阅记录提取"，固定排在第 4 位（QCT Tools 之后、
    Fork 之前）；exe 路径 Z:\\projects\\Z_PYTHON\\ex_comments\\审阅记录提取.exe。
  - 图标无需单独配置：IconLoaderThread 自动扫 ex_comments\\ 目录下唯一的
    .ico（按后缀匹配，文件名不限）并挂到按钮上；exe 不在时按钮显示 ⚠ 灰态。

v12.10.90：
  - PFN 不再写死 PFN.exe（该目录实际是 PFN_v4/v5/v6.exe 多版本迭代）。
    改为 _latest_exe_in_dir() 按 mtime 取目录顶层下最新的 .exe；构造时解析
    一次（用于按钮存在性/图标），_launch_app 启动前再扫一次取最新，覆盖
    "SPAhub 开着没关、期间又出新版本"的情况。图标仍由 IconLoaderThread
    扫 PFN\\ 目录命中 app_icon.ico。
  - 通用机制：self._latest_exe_dirs = {name: dir}，后续其它多版本工具按需登记即可。

设计要点：
  - 每个工具按钮独立监听 MouseButtonDblClick 事件（事件过滤器），避免
    QToolButton 内部 click 事件吞掉双击信号
  - 多个应用可并发启动：每次启动一个 AppLaunchThread，仅在线程
    真正完成（finished 信号）时才从列表移除
  - 主线程绝不调用 subprocess.Popen，所有重 IO 都在 QThread 中
  - 网页链接是 webbrowser.open，本身是非阻塞的（Windows 下走 ShellExecute），
    无需后台线程
  - 卡片池化：每个 url 对应一个 LinkCard 实例，过滤/重排只 reparent，
    不重建 widget — 39 条 / 数百条规模下都是常数级开销
"""
import os
import sys
import json
import subprocess
import webbrowser
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QToolButton,
    QLineEdit, QPushButton, QScrollArea, QFrame, QSizePolicy, QTabWidget
)
from PySide6.QtGui import QIcon
from PySide6.QtCore import Qt, QSize, QThread, Signal, QEvent, QObject


def _local_appdata() -> str:
    """
    返回当前 Windows 用户的 LocalAppData 目录（如 C:\\Users\\<USER>\\AppData\\Local）。

    优先级：
      1. 环境变量 %LOCALAPPDATA%（Windows 标准）
      2. ~/AppData/Local（Windows 兜底，~ 自动展开为 C:\\Users\\<当前用户>）
      3. ~/.local/share（Linux/Mac 兜底 — 仅用于在非 Windows 平台跑测试时
         能 import shortcuts.py 而不报错）

    v12.10.63：用于把 Fork / Pinnacle 21 这类用户级 exe 的路径模板里的
    "liup71" 自动替换为当前机器实际登录用户的域账号，不需要修改代码。
    """
    p = os.environ.get("LOCALAPPDATA")
    if p:
        return p
    if os.name == "nt":
        return os.path.join(os.path.expanduser("~"), "AppData", "Local")
    return os.path.join(os.path.expanduser("~"), ".local", "share")


def _latest_exe_in_dir(folder: str, fallback: str = "") -> str:
    """
    v12.10.90：返回 folder 顶层下"修改时间最新"的 .exe 绝对路径。

    用于 PFN 这类频繁迭代版本号（PFN_v4 / PFN_v5 / PFN_v6.exe ...）的工具：
    不写死文件名，每次按 mtime 取最新的 exe。只扫顶层（不递归），避免误命中
    子目录里的旧版本 / 安装器。

    找不到任何 .exe 时返回 fallback（默认空串 → 上层渲染成 ⚠ 不可用按钮）；
    任何 IO 异常（网络盘抖动、权限等）也回退 fallback，绝不抛出。
    """
    try:
        if not folder or not os.path.isdir(folder):
            return fallback
        newest, newest_mtime = fallback, -1.0
        with os.scandir(folder) as it:
            for e in it:
                try:
                    if e.is_file() and e.name.lower().endswith(".exe"):
                        m = e.stat().st_mtime
                        if m > newest_mtime:
                            newest, newest_mtime = e.path, m
                except OSError:
                    continue
        return newest
    except OSError:
        return fallback


def _icon_cache_path() -> str:
    """图标缓存 json 路径 — 与 weblinks_bookmarks.json 同目录"""
    base = _local_appdata()
    sub = os.path.join(base, "SPAhub")
    try:
        os.makedirs(sub, exist_ok=True)
    except OSError:
        pass
    return os.path.join(sub, "icon_cache.json")


# =============================================================================
# 图标加载后台线程
# =============================================================================
class IconLoaderThread(QThread):
    """
    后台扫描每个 exe 所在目录下的 .ico，发现后回传主线程挂到按钮上。

    v12.10.63：加缓存 + 加优先级匹配。
      - 启动时先读 icon_cache.json，命中且 ico 文件还在 → 直接 emit，
        跳过扫描（避免每次启动都 walk 目录树）
      - 缓存未命中：递归 walk（最大深度 3），优先选与 exe 同名的 ico；
        没同名时按"含 exe 名 > 深度浅"排序选最佳候选
      - exe 不存在的条目直接跳过（不写缓存，避免污染下次扫描）
      - 找到 ico 后写缓存（原子写：tmp + rename）
      - 缓存文件位置：%LOCALAPPDATA%\\SPAhub\\icon_cache.json

    v12.10.64：加"系统图标兜底"机制。
      - 外部 .ico 没找到时（如 Fork 把图标内嵌在 PE 资源中、没单独的
        .ico 文件），emit use_system_icon(name, exe_path)
      - 主线程收到信号后调 QFileIconProvider 从 exe 资源里提取
        系统图标，挂到按钮上。这样任何 Windows exe 都能有图标，不需要
        为每个 exe 单独维护一份 .ico 资源
      - QFileIconProvider 必须在主线程调用（Qt 限制），所以分两段处理：
        子线程扫描 → 没找到 → emit 信号 → 主线程拉系统图标
    """
    icon_found = Signal(str, str)        # (name, ico_path) — 外部 .ico 找到
    use_system_icon = Signal(str, str)   # (name, exe_path) — 外部没找到，让主线程拉系统图标

    # 递归扫描最大深度（相对 exe 所在目录）
    MAX_SCAN_DEPTH = 3

    def __init__(self, apps):
        super().__init__()
        self.apps = apps
        self._cache_path = _icon_cache_path()
        self._cache = self._load_cache()

    def _load_cache(self) -> dict:
        try:
            with open(self._cache_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                return data
        except (OSError, json.JSONDecodeError):
            pass
        return {}

    def _save_cache(self) -> None:
        try:
            tmp = self._cache_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self._cache, f, ensure_ascii=False, indent=2)
            os.replace(tmp, self._cache_path)
        except OSError:
            pass

    @classmethod
    def _scan_icon(cls, folder: str, exe_path: str):
        """
        递归 walk folder 找 .ico，返回最佳匹配 (None 表示没找到)。

        优先级（越小越优先）：
          0 - ico 文件名（去扩展名）与 exe 名完全一致
          1 - ico 文件名包含 exe 名
          10 + depth - 其它（同深度按 walk 顺序）
        """
        if not folder or not os.path.isdir(folder):
            return None
        exe_basename = os.path.splitext(os.path.basename(exe_path))[0].lower()
        candidates = []  # [(priority, full_path), ...]
        try:
            for root, dirs, files in os.walk(folder):
                # 深度控制：相对 folder 的层级
                rel = root[len(folder):].lstrip(os.sep)
                depth = rel.count(os.sep) + (1 if rel else 0)
                if depth >= cls.MAX_SCAN_DEPTH:
                    del dirs[:]  # 不再深入
                for f in files:
                    if not f.lower().endswith(".ico"):
                        continue
                    name = os.path.splitext(f)[0].lower()
                    if name == exe_basename:
                        priority = 0
                    elif exe_basename in name or name in exe_basename:
                        priority = 1
                    else:
                        priority = 10 + depth
                    candidates.append((priority, os.path.join(root, f)))
                    # 同名 ico 已是最佳，可提前退出
                    if priority == 0:
                        return candidates[-1][1]
        except OSError:
            return None
        if not candidates:
            return None
        candidates.sort(key=lambda x: x[0])
        return candidates[0][1]

    def run(self):
        cache_changed = False
        for name, exe_path in self.apps:
            # exe 不存在 → 跳过（按钮渲染时已显示警告，无需图标）
            if not exe_path or not os.path.isfile(exe_path):
                continue

            # 缓存命中 + ico 文件仍在 → 直接用
            cached_ico = self._cache.get(exe_path)
            if cached_ico and os.path.isfile(cached_ico):
                self.icon_found.emit(name, cached_ico)
                continue

            # 未命中或失效 → 扫描
            folder = os.path.dirname(exe_path)
            found_ico = self._scan_icon(folder, exe_path)
            if found_ico:
                self.icon_found.emit(name, found_ico)
                self._cache[exe_path] = found_ico
                cache_changed = True
            else:
                # v12.10.64：外部 .ico 没找到（如 Fork 把图标嵌在 PE 资源里），
                # 通知主线程用 QFileIconProvider 从 exe 资源提取系统图标兜底
                self.use_system_icon.emit(name, exe_path)
                # 缓存里如果有失效的旧值，清掉以便下次重扫
                if exe_path in self._cache:
                    self._cache.pop(exe_path, None)
                    cache_changed = True

        if cache_changed:
            self._save_cache()


# =============================================================================
# 外部程序启动后台线程
# =============================================================================
class AppLaunchThread(QThread):
    """
    在后台线程中拉起外部 exe。
    每个应用启动独立一个线程实例，多个应用可并发启动而互不阻塞。
    """
    status_signal = Signal(str)
    error_signal = Signal(str, str)

    def __init__(self, exe_path, app_name, parent=None):
        super().__init__(parent)
        self.exe_path = exe_path
        self.app_name = app_name

    def run(self):
        self.status_signal.emit(
            f"⏳ 正在启动 {self.app_name}(首次加载约 20-30 秒,请耐心等待)..."
        )
        try:
            cwd = os.path.dirname(self.exe_path)
            # creationflags=0x00000008 = DETACHED_PROCESS（不继承父进程控制台）
            subprocess.Popen([self.exe_path], cwd=cwd, shell=False, creationflags=0x00000008)
            self.status_signal.emit(
                f"✅ 已下发 {self.app_name} 启动命令,桌面窗口出现前请稍候..."
            )
        except FileNotFoundError:
            self.error_signal.emit("文件不存在", f"找不到 {self.app_name}:\n{self.exe_path}")
        except Exception as e:
            self.error_signal.emit("启动失败", f"启动 {self.app_name} 失败:{str(e)}\n路径:{self.exe_path}")


# =============================================================================
# 双击事件过滤器
# =============================================================================
class _DoubleClickFilter(QObject):
    """安装到每个 QToolButton 上：拦截 MouseButtonDblClick，触发 launch 回调"""
    def __init__(self, callback, parent=None):
        super().__init__(parent)
        self._callback = callback

    def eventFilter(self, obj, event):
        if event.type() == QEvent.MouseButtonDblClick and event.button() == Qt.LeftButton:
            self._callback()
            return True
        return False


# v12.10.63：exe 不存在的按钮 hover 时，通过主窗口状态栏告知用户找寻路径
class _HoverHintFilter(QObject):
    """
    安装到 QToolButton 上：鼠标进入时调用 enter_cb(hint_text)，离开时调用
    leave_cb()。用于 exe 不存在的工具按钮，hover 时在左下角状态栏提示
    用户的预期安装路径（让用户自己核实/修正）。
    不消费事件（return False），与 _DoubleClickFilter 共存兼容。
    """
    def __init__(self, hint_text, enter_cb, leave_cb, parent=None):
        super().__init__(parent)
        self._hint = hint_text
        self._enter_cb = enter_cb
        self._leave_cb = leave_cb

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Enter:
            try:
                self._enter_cb(self._hint)
            except Exception:
                pass
        elif event.type() == QEvent.Leave:
            try:
                self._leave_cb()
            except Exception:
                pass
        return False


# =============================================================================
# 拼音首字母（GB2312 区位法，零第三方依赖）
# =============================================================================
# 升序排列的 GB2312 拼音首字母分界字编码（每段起始字的 GB2312 字节序值）：
#   A=啊 0xB0A1, B=芭 0xB0C5, C=擦 0xB2C1, D=搭 0xB4EE, E=蛾 0xB6EA,
#   F=发 0xB7A2, G=噶 0xB8C1, H=哈 0xBAA1, J=击 0xBBF7, K=喀 0xBFA6,
#   L=垃 0xC0AC, M=妈 0xC2E8, N=拿 0xC4C3, O=哦 0xC5B6, P=啪 0xC5BE,
#   Q=期 0xC6DA, R=然 0xC8BB, S=撒 0xC8F6, T=塌 0xCBFA, W=挖 0xCDDA,
#   X=昔 0xCEF4, Y=压 0xD1B9, Z=匝 0xD4D1
# I / U / V 不存在以这些为声母起始的常用汉字，故跳过。
_PINYIN_BOUNDS = (
    (0xB0A1, 'A'), (0xB0C5, 'B'), (0xB2C1, 'C'), (0xB4EE, 'D'),
    (0xB6EA, 'E'), (0xB7A2, 'F'), (0xB8C1, 'G'), (0xBAA1, 'H'),
    (0xBBF7, 'J'), (0xBFA6, 'K'), (0xC0AC, 'L'), (0xC2E8, 'M'),
    (0xC4C3, 'N'), (0xC5B6, 'O'), (0xC5BE, 'P'), (0xC6DA, 'Q'),
    (0xC8BB, 'R'), (0xC8F6, 'S'), (0xCBFA, 'T'), (0xCDDA, 'W'),
    (0xCEF4, 'X'), (0xD1B9, 'Y'), (0xD4D1, 'Z'),
)


def _first_letter(name: str) -> str:
    """
    返回 name 的首字母索引（大写 A-Z，或 '#' 表示其它）。
    ASCII 字母直接走；中文按 GB2312 区位查表；其它（数字/符号/繁体/扩展汉字）→ '#'
    """
    if not name:
        return "#"
    ch = name.strip()[:1] if name.strip() else ""
    if not ch:
        return "#"
    if 'a' <= ch.lower() <= 'z':
        return ch.upper()
    try:
        b = ch.encode('gb2312')
    except UnicodeEncodeError:
        return "#"
    if len(b) != 2:
        return "#"
    code = b[0] * 256 + b[1]
    if not (0xB0A1 <= code <= 0xD7F9):
        return "#"
    result = "#"
    for bound, letter in _PINYIN_BOUNDS:
        if code >= bound:
            result = letter
        else:
            break
    return result


# =============================================================================
# 收藏持久化
# =============================================================================
def _bookmarks_path() -> str:
    """
    返回书签 JSON 文件路径。
    Windows: %LOCALAPPDATA%\\SPAhub\\weblinks_bookmarks.json
             （与 v12.10.47 的 logs 同根，跨账户都能写）
    其它系统：~/.SPAhub/weblinks_bookmarks.json
    """
    if sys.platform.startswith("win"):
        base = os.environ.get('LOCALAPPDATA') or os.path.expanduser('~')
        target_dir = os.path.join(base, 'SPAhub')
    else:
        target_dir = os.path.join(os.path.expanduser('~'), '.SPAhub')
    try:
        os.makedirs(target_dir, exist_ok=True)
    except OSError:
        pass
    return os.path.join(target_dir, 'weblinks_bookmarks.json')


def _load_bookmarks() -> set:
    """读取本地收藏的 URL 集合；任何异常都回退到空集"""
    path = _bookmarks_path()
    if not os.path.exists(path):
        return set()
    try:
        with open(path, 'r', encoding='utf-8') as f:
            data = json.load(f)
        favs = data.get('favorites') if isinstance(data, dict) else None
        if isinstance(favs, list):
            return {str(u) for u in favs if u}
    except (OSError, ValueError, UnicodeDecodeError):
        pass
    return set()


def _save_bookmarks(favs: set) -> bool:
    """写回收藏集合；先写 .tmp 再 rename 避免半截损坏"""
    path = _bookmarks_path()
    try:
        tmp = path + ".tmp"
        with open(tmp, 'w', encoding='utf-8') as f:
            json.dump({"favorites": sorted(favs)}, f, ensure_ascii=False, indent=2)
        os.replace(tmp, path)
        return True
    except OSError:
        return False


# =============================================================================
# 链接卡片（含星星收藏按钮）
# =============================================================================
class LinkCard(QWidget):
    """
    单条链接卡片：

      ┌──────────────────────────────────────────┬─────┐
      │  显示名称（QLabel，左对齐）               │  ☆  │
      └──────────────────────────────────────────┴─────┘

    交互：
      - 单击卡片体（非星星区域）→ sig_clicked.emit(url)
      - 单击星星 → 切换内部状态 + sig_favorite_toggled.emit(url)
      - hover 进入卡片 → 卡片底色 hover 态；未收藏的星星显形
      - 已收藏 → 星星恒为黄实心 ★（无论 hover 与否）
    """
    sig_clicked = Signal(str)            # url
    sig_favorite_toggled = Signal(str)   # url

    def __init__(self, name: str, url: str, is_favorite: bool,
                 show_star: bool = True, parent=None):
        """
        v12.10.58：show_star 控制星星是否可见。
          - True（默认）：原行为，hover 时灰星显形，已收藏时黄星常显
          - False：星星完全不渲染（给 simple_mode tab 用，SOP/WI/手册
            类文档库不需要"收藏置顶"语义，少一个视觉元素更聚焦）
        """
        super().__init__(parent)
        self.name = name
        self.url = url
        self._is_favorite = is_favorite
        self._hover = False
        self._show_star = bool(show_star)

        self.setCursor(Qt.PointingHandCursor)
        # v12.10.88：显式开启 styled background —— 否则普通 QWidget 子类的 QSS
        #            background/border 可能不渲染（左强调条/边框看不到的根因之一）。
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setMinimumHeight(38)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setToolTip(f"<b>{name}</b><br><span style='color:#909399'>{url}</span>")

        # v12.10.88：左强调条改用独立子控件（3px 蓝条），不再依赖 border-left。
        #   根因：旧版 Qt 对"不等宽 border（border:1px + border-left:3px）+ border-radius"
        #   组合会整条丢弃样式 → 白底/左条全看不见，透出灰容器。改成独立子控件后，
        #   卡片本体只用"等宽 border + radius"（各版本可靠），左条由子 QFrame 稳定绘制。
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.accent = QFrame()
        self.accent.setAttribute(Qt.WA_StyledBackground, True)
        self.accent.setFixedWidth(4)
        self.accent.setStyleSheet(
            "background:#4A6FA5; border:none;"
            "border-top-left-radius:4px; border-bottom-left-radius:4px;"
        )
        self.accent.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        outer.addWidget(self.accent)

        content = QWidget()
        content.setObjectName("LinkCardContent")
        content.setStyleSheet("#LinkCardContent { background: transparent; border: none; }")
        content.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        h = QHBoxLayout(content)
        h.setContentsMargins(10, 6, 6, 6)
        h.setSpacing(4)

        # 名称标签
        self.lbl = QLabel(name)
        self.lbl.setStyleSheet("color: #303133; font-size: 12px; background: transparent;")
        # 子控件不抢鼠标，hover / 点击全部落到 LinkCard 上
        self.lbl.setAttribute(Qt.WA_TransparentForMouseEvents, True)
        h.addWidget(self.lbl, 1)

        # 星星按钮（自己接收点击）
        # v12.10.58：show_star=False 时完全不创建星星 widget
        self.star = None
        if self._show_star:
            self.star = QToolButton(content)
            self.star.setCursor(Qt.PointingHandCursor)
            self.star.setFixedSize(24, 24)
            self.star.setToolTip("加入 / 取消收藏")
            self.star.setAutoRaise(True)
            self.star.clicked.connect(self._on_star_clicked)
            h.addWidget(self.star, 0)
        else:
            # 不显示星星 — 右边小 padding 平衡视觉
            h.setContentsMargins(10, 6, 10, 6)

        outer.addWidget(content, 1)

        self._refresh_star()
        self._refresh_card_style()

    # -------- 公共 API --------
    def set_favorite(self, fav: bool):
        if fav == self._is_favorite:
            return
        self._is_favorite = fav
        self._refresh_star()

    def is_favorite(self) -> bool:
        return self._is_favorite

    # -------- 鼠标事件 --------
    def enterEvent(self, e):
        self._hover = True
        self._refresh_card_style()
        self._refresh_star()
        super().enterEvent(e)

    def leaveEvent(self, e):
        self._hover = False
        self._refresh_card_style()
        self._refresh_star()
        super().leaveEvent(e)

    def mousePressEvent(self, e):
        if e.button() == Qt.LeftButton:
            # 落在星星上的点击交给星星处理，这里只处理卡片体
            # （star=None 时即 simple_mode，直接 emit）
            if self.star is not None and self.childAt(e.position().toPoint()) is self.star:
                return super().mousePressEvent(e)
            self.sig_clicked.emit(self.url)
            e.accept()
            return
        super().mousePressEvent(e)

    # -------- 内部刷新 --------
    def _on_star_clicked(self):
        # 翻转内部 + 通知上层；上层负责持久化与重排
        self._is_favorite = not self._is_favorite
        self._refresh_star()
        self.sig_favorite_toggled.emit(self.url)

    def _refresh_card_style(self):
        # v12.10.88：卡片本体只用"等宽 border + radius"（各 Qt 版本可靠渲染白底+边框）；
        #            左侧 3px 蓝强调条由独立子控件 self.accent 绘制（不依赖 border-left）。
        if self._hover:
            self.setStyleSheet("""
                LinkCard {
                    background: #EEF4FF;
                    border: 1px solid #4A6FA5;
                    border-radius: 4px;
                }
            """)
            self.lbl.setStyleSheet(
                "color: #4A6FA5; font-size: 12px; font-weight: bold; background: transparent; border: none;"
            )
        else:
            self.setStyleSheet("""
                LinkCard {
                    background: #FFFFFF;
                    border: 1px solid #E4E7ED;
                    border-radius: 4px;
                }
            """)
            self.lbl.setStyleSheet(
                "color: #303133; font-size: 12px; background: transparent; border: none;"
            )

    def _refresh_star(self):
        """
        星星显示规则：
          - 已收藏：始终黄实心 ★（#FFB300）
          - 未收藏 + hover：灰空心 ☆（#C0C4CC），hover 到星星变橙
          - 未收藏 + 非 hover：完全透明占位（保留布局尺寸避免 hover 抖动）
          - show_star=False：self.star is None，直接 return（v12.10.58）
        """
        if self.star is None:
            return
        if self._is_favorite:
            self.star.setText("★")
            self.star.setStyleSheet("""
                QToolButton { color: #FFB300; font-size: 16px;
                              background: transparent; border: none; }
                QToolButton:hover { color: #FF8F00; }
            """)
        elif self._hover:
            self.star.setText("☆")
            self.star.setStyleSheet("""
                QToolButton { color: #C0C4CC; font-size: 16px;
                              background: transparent; border: none; }
                QToolButton:hover { color: #FFB300; }
            """)
        else:
            self.star.setText("☆")
            self.star.setStyleSheet("""
                QToolButton { color: transparent; font-size: 16px;
                              background: transparent; border: none; }
            """)


# =============================================================================
# 折叠分组（v12.10.58）— simple_mode 下按 sub_dir 分组用
# =============================================================================
class CollapsibleSection(QFrame):
    """
    可折叠分组：标题栏（▾/▸ 子目录名 (count)）+ 内容区。

    标题栏点击切换展开/收起，箭头同步变换。内容区是一个 QVBoxLayout 容器，
    上层通过 .add_card() 把 LinkCard 加入；折叠时整个内容区 setVisible(False)。

    设计意图：让 SPAhub Shortcuts 的 SOP/WI/手册 tab 保留 SharePoint 的
    文件夹层级，根目录文件直接列、子目录文件归到可折叠 section 里。
    """
    def __init__(self, title: str, count: int, parent=None):
        super().__init__(parent)
        self._expanded = True
        self._title = title

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(4)

        # 标题按钮（▾ 子目录名 (count)）
        self.header = QPushButton(self._format_title())
        self.header.setCursor(Qt.PointingHandCursor)
        self.header.setFlat(True)
        self.header.setStyleSheet("""
            QPushButton {
                text-align: left; padding: 4px 8px;
                background: transparent; border: none;
                color: #606266; font-size: 13px; font-weight: bold;
            }
            QPushButton:hover { color: #409EFF; }
        """)
        self.header.clicked.connect(self.toggle)
        layout.addWidget(self.header)

        # 内容区（缩进显示子项）
        self.content = QWidget()
        self.content.setStyleSheet("background: transparent;")
        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(18, 0, 0, 0)  # 18px 左缩进体现层级
        self.content_layout.setSpacing(6)
        layout.addWidget(self.content)

    def _format_title(self) -> str:
        prefix = "▾" if self._expanded else "▸"
        # count 从内部 content_layout.count() 派生，避免每次添加都要传一次
        n = self.content_layout.count() if hasattr(self, "content_layout") else 0
        return f"{prefix}  {self._title}  ({n})" if n else f"{prefix}  {self._title}"

    def add_card(self, card):
        """把卡片加进内容区。卡片是从对象池来的，外层负责生命周期"""
        self.content_layout.addWidget(card)
        self.header.setText(self._format_title())

    def toggle(self, force_expanded=None):
        """切换展开/收起；force_expanded=True/False 强制设定"""
        if force_expanded is not None:
            self._expanded = bool(force_expanded)
        else:
            self._expanded = not self._expanded
        self.content.setVisible(self._expanded)
        self.header.setText(self._format_title())


# =============================================================================
# 一站式网页导航
# =============================================================================
class WebNavWidget(QWidget):
    """
    一站式网页导航。

    布局：
      搜索框 + 刷新
      A-Z 字母索引行（simple_mode 隐藏）
      ⭐ 收藏区（如有，simple_mode 不分桶）
      全部链接区（A-Z 排序）

    布局（v12.10.60 起 4 tab 完全统一）：
      hint 副标题
      搜索框 + 刷新按钮
      ⭐ 收藏区（如有，平铺单列）
      全部链接区（按 sub_dir 分桶；根目录文件直接列；子目录用 CollapsibleSection 折叠）

    所有 tab 都用同一套渲染逻辑 — 不再有"主导航 3 列网格 + 字母索引"
    与"simple_mode 单列文档树"两套分支。weblink.xlsx 没有 sub_dir 时
    自然就是单列纯卡片列表，行为完全等价于 ecm_files.xlsx 中没有
    子目录的 sheet（如 SOP）。

    数据：
      - self._all_links: List[(name, url, sub_dir)]，从 xlsx 读出
      - self._favorites: Set[url]，本地持久化的收藏集合
      - self._filter_text: 当前搜索关键字（lower）
      - self._card_pool: url → LinkCard，对象池避免反复创建销毁
    """

    def __init__(self, weblinks_xlsx_path: str, parent=None,
                 sheet_name: str = "",
                 simple_mode: bool = True,
                 title: str = "一站式导航",
                 hint: str = "单击打开 · 鼠标悬停点击 ☆ 收藏"):
        """
        weblinks_xlsx_path : xlsx 文件路径
        sheet_name         : 指定 sheet 名；空串 → 优先 "weblinks" sheet，
                             找不到用第一个 sheet
        simple_mode        : v12.10.60 起 deprecated — 所有 tab 都用统一
                             的单列文档树布局，此参数不再有任何作用，
                             保留只为兼容旧调用代码
        title              : v12.10.59 起 deprecated — 内部大标题已删除
                             （与 QTabWidget tab label 重复），保留参数兼容
        hint               : 副标题提示文字
        """
        super().__init__(parent)
        self._xlsx_path = weblinks_xlsx_path
        self._sheet_name = (sheet_name or "").strip()
        # v12.10.60：simple_mode 参数失效 — 所有 tab 都用单列文档树
        _ = simple_mode  # noqa: F841 — 显式标记 deprecated
        self._all_links = []
        # 所有 tab 都支持收藏 — 同一份持久化文件，跨 tab 共享 url 集合
        # （url 是文件唯一标识，不会跨数据源重复）
        self._favorites = _load_bookmarks()
        self._filter_text = ""
        self._card_pool = {}   # url → LinkCard，全程复用

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(8)

        # v12.10.59：删除内部大标题 — 与 QTabWidget tab label 重复
        # title 参数保留以兼容外部调用，但不再渲染
        # （hint 副标题仍保留，提示用户单击/悬停的交互方式）
        _ = title  # noqa: F841 — 显式标记保留参数
        hint_lbl = QLabel(hint)
        hint_lbl.setStyleSheet("color: #909399; font-size: 12px; padding-top: 2px;")
        root.addWidget(hint_lbl)

        # —— 搜索 + 刷新 行 ——
        ctrl_row = QHBoxLayout()
        ctrl_row.setSpacing(8)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 输入关键字筛选(名称或链接,不区分大小写)")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.setStyleSheet("""
            QLineEdit { border: 1px solid #DCDFE6; border-radius: 4px;
                        padding: 5px 10px; background: #FFFFFF; color: #303133;
                        font-size: 13px; }
            QLineEdit:focus { border-color: #409EFF; }
        """)
        self.search_edit.textChanged.connect(self._on_search_changed)
        ctrl_row.addWidget(self.search_edit, 1)

        self.btn_refresh = QPushButton("🔄 刷新")
        self.btn_refresh.setFixedWidth(80)
        self.btn_refresh.setCursor(Qt.PointingHandCursor)
        self.btn_refresh.setToolTip(
            f"重新读取 {os.path.basename(weblinks_xlsx_path)}"
            + (f"（sheet: {self._sheet_name}）" if self._sheet_name else "")
        )
        self.btn_refresh.clicked.connect(self.reload)
        ctrl_row.addWidget(self.btn_refresh)

        root.addLayout(ctrl_row)

        # v12.10.60：删除字母索引行 — 所有 tab 统一用搜索 + 文档树即可；
        # GB2312 拼音首字母排序的能力仍保留在内部（_first_letter 函数用于
        # 排序，让中文 / 英文按首字母混合自然排序）

        # —— 链接区（滚动） ——
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        # 滚动条样式与主窗口 _apply_global_style 一致
        scroll.setStyleSheet("""
            QScrollArea { background: transparent; }
            QScrollBar:vertical { border: none; background: #F5F7FA; width: 10px; border-radius: 5px; }
            QScrollBar::handle:vertical { background: #DCDFE6; border-radius: 5px; min-height: 20px; }
            QScrollBar::handle:vertical:hover { background: #C0C4CC; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }
        """)
        self._grid_host = QWidget()
        self._grid_host.setStyleSheet("background: transparent;")
        self._host_layout = QVBoxLayout(self._grid_host)
        self._host_layout.setContentsMargins(0, 4, 0, 0)
        self._host_layout.setSpacing(8)
        self._host_layout.setAlignment(Qt.AlignTop)
        scroll.setWidget(self._grid_host)
        root.addWidget(scroll, 1)

        # —— 空态提示 ——
        self.empty_label = QLabel("")
        self.empty_label.setAlignment(Qt.AlignCenter)
        self.empty_label.setStyleSheet("color: #C0C4CC; font-size: 13px; padding: 20px;")
        self.empty_label.setVisible(False)
        root.addWidget(self.empty_label)

        # 初次加载
        self.reload()

    # -------------------------------------------------------------------------
    # 数据加载
    # -------------------------------------------------------------------------
    def reload(self):
        """从 xlsx 重新读取并完整重建（覆盖卡片池）"""
        self._all_links = self._read_weblinks(self._xlsx_path, self._sheet_name)

        # 清空旧卡片池
        for c in self._card_pool.values():
            c.setParent(None)
            c.deleteLater()
        self._card_pool.clear()

        # 预创建所有卡片（每个 url 只一次）
        # v12.10.59：所有 tab 卡片都带星星 — SOP/WI/手册 也支持收藏
        for name, url, _sub_dir in self._all_links:
            card = LinkCard(name, url,
                            is_favorite=(url in self._favorites),
                            show_star=True)
            card.sig_clicked.connect(self._open_url)
            card.sig_favorite_toggled.connect(self._on_favorite_toggled)
            self._card_pool[url] = card

        self._refresh_grid()

    @staticmethod
    def _read_weblinks(xlsx_path: str, sheet_name: str = ""):
        """
        读取 xlsx，返回 [(name, url, sub_dir), ...]。所有异常 → 空列表。

        sheet_name 优先级（v12.10.57）：
          - 给定且存在 → 用该 sheet
          - 未给定 → 优先名为 "weblinks" 的 sheet，找不到用第一个 sheet
          - 给定但不存在 → 返回空（让 UI 显示"未找到"提示）

        列识别（v12.10.58）：
          - "网页展示" / "名称" / name/title/label → name 列
          - "网页链接" / url/link/href → url 列
          - "相对路径" / "路径" / "子目录" / path/folder/subdir → sub_dir 列
          - 旧版 xlsx（仅 2 列，无 sub_dir）→ sub_dir 全部为 ""，向后兼容
        """
        if not xlsx_path or not os.path.exists(xlsx_path):
            return []
        try:
            from openpyxl import load_workbook
            wb = load_workbook(xlsx_path, read_only=True, data_only=True)
            ws = None
            wanted = (sheet_name or "").strip()
            if wanted:
                for nm in wb.sheetnames:
                    if nm.strip().lower() == wanted.lower():
                        ws = wb[nm]
                        break
                if ws is None:
                    wb.close()
                    return []
            else:
                for nm in wb.sheetnames:
                    if nm.strip().lower() == "weblinks":
                        ws = wb[nm]
                        break
                if ws is None:
                    ws = wb[wb.sheetnames[0]]

            rows = list(ws.iter_rows(values_only=True))
            wb.close()
            if not rows:
                return []

            header = [str(c).strip() if c is not None else "" for c in rows[0]]
            name_idx, url_idx, sub_idx = 0, 1, -1
            for i, h in enumerate(header):
                hl = h.lower()
                if "链接" in h or hl in ("url", "link", "href"):
                    url_idx = i
                elif "展示" in h or "名称" in h or hl in ("name", "title", "label"):
                    name_idx = i
                elif ("相对路径" in h or "路径" in h or "子目录" in h
                      or hl in ("path", "folder", "subdir", "sub_dir", "subfolder")):
                    sub_idx = i

            out = []
            seen_urls = set()
            for r in rows[1:]:
                if r is None:
                    continue
                if max(name_idx, url_idx) >= len(r):
                    continue
                name = r[name_idx]
                url = r[url_idx]
                if not name or not url:
                    continue
                name_s = str(name).strip()
                url_s = str(url).strip()
                if not name_s or not url_s:
                    continue
                if not (url_s.startswith("http://") or url_s.startswith("https://")
                        or url_s.startswith("ftp://") or url_s.startswith("ftps://")):
                    url_s = "http://" + url_s
                if url_s in seen_urls:
                    continue
                seen_urls.add(url_s)
                sub_s = ""
                if 0 <= sub_idx < len(r) and r[sub_idx] is not None:
                    sub_s = str(r[sub_idx]).strip().strip("/")
                out.append((name_s, url_s, sub_s))
            return out
        except Exception:
            return []

    # -------------------------------------------------------------------------
    # 事件
    # -------------------------------------------------------------------------
    def _on_search_changed(self, text: str):
        self._filter_text = (text or "").strip().lower()
        self._refresh_grid()

    def _on_favorite_toggled(self, url: str):
        """卡片星星点击 → 翻转持久化集合 → 写盘 → 重排"""
        if url in self._favorites:
            self._favorites.discard(url)
        else:
            self._favorites.add(url)
        _save_bookmarks(self._favorites)
        self._refresh_grid()

    # -------------------------------------------------------------------------
    # 渲染
    # -------------------------------------------------------------------------
    def _refresh_grid(self):
        """过滤 → 拆分收藏/其它 → 各自排序 → 用池里卡片重排（不重建）

        items 元素为 (name, url, sub_dir) 三元组。
        """
        # 1. 过滤（仅搜索，v12.10.60 起字母索引已移除）
        items = list(self._all_links)
        if self._filter_text:
            kw = self._filter_text
            # 搜索也匹配 sub_dir 名称，方便 "popPK" 这类按子目录直接找
            items = [(n, u, s) for (n, u, s) in items
                     if kw in n.lower() or kw in u.lower() or kw in s.lower()]

        # 2. 分桶 + 排序（所有 tab 都按收藏 / 非收藏分桶）
        fav_items = [(n, u, s) for (n, u, s) in items if u in self._favorites]
        oth_items = [(n, u, s) for (n, u, s) in items if u not in self._favorites]
        # 收藏区不分子目录（平铺，按字母排序）
        fav_items.sort(key=lambda x: (_first_letter(x[0]), x[0].lower()))
        # 非收藏区按 (sub_dir A-Z, 文件名 A-Z) 排序；sub_dir 为空的（根目录）
        # 排在子目录前面，所有 tab 行为统一
        oth_items.sort(key=lambda x: (x[2], _first_letter(x[0]), x[0].lower()))

        # 3. 清空 host_layout（卡片用 setParent(None) 解绑保留，其它 deleteLater）
        self._clear_host_layout()

        # 4. 空态
        if not items:
            if not self._all_links:
                fname = os.path.basename(self._xlsx_path) or "weblink.xlsx"
                sheet_hint = f"（sheet: {self._sheet_name}）" if self._sheet_name else ""
                self.empty_label.setText(
                    f"📭 未找到 {fname}{sheet_hint}\n"
                    f"期望路径: SPAhub/resources/{fname}"
                )
            else:
                self.empty_label.setText("🔍 没有匹配的链接")
            self.empty_label.setVisible(True)
            return
        self.empty_label.setVisible(False)

        # 5. 收藏区
        if fav_items:
            fav_header = QLabel(f"⭐ 收藏 ({len(fav_items)})")
            fav_header.setStyleSheet(
                "color: #E6A23C; font-size: 13px; font-weight: bold; padding: 4px 0;"
            )
            self._host_layout.addWidget(fav_header)
            # 收藏区强制平铺 — 把 sub_dir 清空让 _add_grid_to_layout 走
            # "全部根目录直接列"分支。收藏数量本来就少，再分子目录折叠让
            # 用户多一次点击；平铺更符合"快速访问"的语义。
            fav_flat = [(n, u, "") for (n, u, _s) in fav_items]
            self._add_grid_to_layout(fav_flat)

        # 6. 全部链接区（保留 sub_dir，子目录自动用 CollapsibleSection 折叠）
        if oth_items:
            if fav_items:
                oth_header = QLabel("全部链接")
                oth_header.setStyleSheet(
                    "color: #909399; font-size: 13px; font-weight: bold; padding: 4px 0;"
                )
                self._host_layout.addWidget(oth_header)
            self._add_grid_to_layout(oth_items)

    def _add_grid_to_layout(self, items):
        """
        把一批 (name, url, sub_dir) 放到一个新容器里加入 host_layout，
        卡片从池里取。

        v12.10.60：4 个 tab 完全统一，永远走单列文档树渲染：
          - 根目录文件（sub_dir == ""）直接列在单列 QVBoxLayout 里
          - 子目录文件按 sub_dir 分桶，用 CollapsibleSection 折叠
          - sub_dir 全为空时（如 weblink.xlsx 一站式导航）自然就是纯
            单列卡片列表，无任何折叠 section
        """
        host = QWidget()
        # 浅灰底色 + 圆角 + 边框，让白色卡片在上面有对比层次
        host.setObjectName("LinkGridHost")
        host.setStyleSheet("""
            #LinkGridHost {
                background: #FFFFFF;
                border: 1px solid #E4E7ED;
                border-radius: 6px;
            }
        """)

        vlayout = QVBoxLayout(host)
        vlayout.setContentsMargins(10, 10, 10, 10)
        vlayout.setSpacing(6)
        vlayout.setAlignment(Qt.AlignTop)

        # 按 sub_dir 分桶（保留排序顺序）
        root_files = []
        sub_groups = {}  # sub_dir → [(name, url), ...]
        sub_order = []   # 维持首次出现顺序
        for name, url, sub_dir in items:
            if not sub_dir:
                root_files.append((name, url))
            else:
                if sub_dir not in sub_groups:
                    sub_groups[sub_dir] = []
                    sub_order.append(sub_dir)
                sub_groups[sub_dir].append((name, url))

        # 根目录文件直接列
        for name, url in root_files:
            card = self._card_pool.get(url)
            if card is not None:
                vlayout.addWidget(card)

        # 子目录用 CollapsibleSection（按子目录名拼音首字母 A-Z 稳定排序）
        for sub_dir in sorted(sub_order, key=lambda s: (_first_letter(s), s.lower())):
            section = CollapsibleSection(sub_dir, len(sub_groups[sub_dir]))
            for name, url in sub_groups[sub_dir]:
                card = self._card_pool.get(url)
                if card is not None:
                    section.add_card(card)
            # 搜索状态下强制展开 — 避免匹配的文件躲在折叠 section 里
            if self._filter_text:
                section.toggle(force_expanded=True)
            vlayout.addWidget(section)

        self._host_layout.addWidget(host)

    def _clear_host_layout(self):
        """
        清空 host_layout。
        卡片是池化对象，setParent(None) 解绑后保留；标题/分隔线/grid host 直接 deleteLater。
        CollapsibleSection 也是临时容器，setParent 把里面的卡片解绑出来后 deleteLater。
        """
        def _detach_cards_recursive(widget):
            """递归把所有 LinkCard 从 widget 的 layout 里 setParent(None)"""
            layout = widget.layout() if isinstance(widget, QWidget) else None
            if layout is None:
                return
            while layout.count():
                sub = layout.takeAt(0)
                sub_w = sub.widget()
                if sub_w is None:
                    continue
                if isinstance(sub_w, LinkCard):
                    sub_w.setParent(None)  # 池化保留
                else:
                    _detach_cards_recursive(sub_w)
                    sub_w.deleteLater()

        while self._host_layout.count():
            item = self._host_layout.takeAt(0)
            w = item.widget()
            if w is None:
                continue
            if isinstance(w, QWidget) and w.layout() is not None:
                _detach_cards_recursive(w)
                w.deleteLater()
            else:
                w.deleteLater()

    @staticmethod
    def _open_url(url: str):
        """系统默认浏览器打开（非阻塞）"""
        try:
            webbrowser.open(url, new=2)
        except Exception:
            try:
                if sys.platform.startswith("win"):
                    os.startfile(url)  # type: ignore[attr-defined]
            except Exception:
                pass


# =============================================================================
# Shortcuts 主界面
# =============================================================================
class ShortcutWidget(QWidget):
    """快捷工具页：上方双击图标启动外部工具，下方一站式网页导航"""

    # v12.10.54：按钮尺寸下调，给下方网页导航让位
    # v12.10.76：按钮略加高（88→96），让长名称（如 Pinnacle 21 Community/Enterprise）
    #            换两行后仍能完整显示，不再截断成 "Pinna...unity"。
    BTN_SIZE_W = 96
    BTN_SIZE_H = 100
    ICON_SIZE = 42
    BTN_SPACING = 18

    @staticmethod
    def _wrap_btn_label(name: str) -> str:
        """多词名称（含空格）一律在最后一个空格处折两行，避免 96px 按钮截断；
        单词名（PFN / Fork）原样返回。
        例：'PDT Manager'→'PDT\\nManager'、'Pinnacle 21 Community'→'Pinnacle 21\\nCommunity'。"""
        if " " in name.strip():
            head, _, tail = name.rstrip().rpartition(" ")
            return f"{head}\n{tail}"
        return name

    def __init__(self, gui):
        super().__init__()
        self.gui = gui
        self._launch_threads = []
        self._click_filters = []

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignTop)
        layout.setContentsMargins(30, 24, 30, 24)
        layout.setSpacing(10)

        # ======== 上：本地工具 ========
        title = QLabel("快捷入口")
        title.setStyleSheet("font-size: 20px; font-weight: bold; color: #303133;")
        layout.addWidget(title)

        hint = QLabel("双击图标启动")  # v12.10.75：删除"本地工具(支持同时打开多个)"后缀
        hint.setStyleSheet("color: #909399; font-size: 12px;")
        layout.addWidget(hint)

        self.grid = QGridLayout()
        self.grid.setSpacing(self.BTN_SPACING)
        self.grid.setAlignment(Qt.AlignLeft)
        layout.addLayout(self.grid)

        # v12.10.63：apps 列表加 Fork + Pinnacle 21（用户本地安装的工具）
        # 路径模板用 _local_appdata() 取 %LOCALAPPDATA%，自动适配当前
        # 机器的 Windows 域账号 — 不需要每台机器改代码
        # v12.10.75：每个条目升级为 (name, target, kind)：
        #   kind="exe" → target 是本地 exe 路径（双击启动，存在性预检查）
        #   kind="url" → target 是网址（双击用默认浏览器打开，图标用 resources 下指定 png）
        # 改动：原 "Pinnacle 21" 重命名为 "Pinnacle 21 Community"；新增
        #       "Pinnacle 21 Enterprise"（kind="url"，指向 certara 在线校验平台）。
        # v12.10.90：PFN 的 exe 频繁迭代版本号（PFN_v4/v5/v6.exe），不写死
        # 文件名，改为按 mtime 取该目录下最新的 .exe（见 _latest_exe_in_dir）。
        _pfn_dir = r"Z:\projects\Z_PYTHON\PFN"
        self.apps = [
            ("PDT Manager", r"Z:\projects\Z_PYTHON\PDT\PDTManager_V4.1.4.exe", "exe"),
            ("PFN", _latest_exe_in_dir(_pfn_dir), "exe"),
            ("QCT Tools", r"Z:\projects\Z_PYTHON\QCT_Tools\QCT_Tools.exe", "exe"),
            # v12.10.89：新增"审阅记录提取"（第 4 位）。图标走自动扫描 —
            # IconLoaderThread 会 walk 同目录 ex_comments\ 取唯一的 .ico（按后缀匹配）。
            ("审阅记录提取", r"Z:\projects\Z_PYTHON\ex_comments\审阅记录提取.exe", "exe"),
            ("Fork", os.path.join(_local_appdata(), "Fork", "Fork.exe"), "exe"),
            ("Pinnacle 21 Community", os.path.join(
                _local_appdata(), "Programs",
                "Pinnacle 21 Community", "Pinnacle 21 Community.exe"
            ), "exe"),
            ("Pinnacle 21 Enterprise",
             "https://hengrui.certara.net/login?returnPath=%2Foauth2%2Fauthorize%3F"
             "%26response_type%3Dcode%26response_mode%3Dquery%26scope%3Denterprise"
             "%26client_id%3Denterprise%26redirect_uri%3Dhttps%3A%2F%2Fhengrui.pinnacle21"
             ".certara.net%2Fauth%2Foauth2%26state%3D%2F",
             "url"),
        ]
        # url 类条目的图标：resources 下的 png（key=app 名）
        self._url_app_icons = {
            "Pinnacle 21 Enterprise": self._resolve_resource_path("pinnacle21_enterprise.png"),
        }

        # v12.10.90：name → 目录 映射。这些条目的 exe 在该目录下按 mtime 取最新 —
        # 启动时（_launch_app）会再扫一次，覆盖"SPAhub 开着没关、期间又出了新版本"
        # 的情况（构造时已解析过一次，用于按钮存在性判断与图标加载）。
        self._latest_exe_dirs = {
            "PFN": _pfn_dir,
        }

        # 正常按钮的样式
        _normal_style = """
            QToolButton { background-color: #FFFFFF; border: 1px solid #E4E7ED;
                          border-radius: 8px; color: #303133; font-weight: bold;
                          font-size: 12px; }
            QToolButton:hover { border: 1px solid #409EFF; background-color: #ECF5FF;
                                color: #409EFF; }
        """
        # v12.10.63：exe 不存在时的按钮样式 — 暗淡 + 虚线边框 + 橙色 hover
        # 让用户能直观看出该按钮不可用，hover 时状态栏告知找寻路径
        _missing_style = """
            QToolButton { background-color: #FAFAFA; border: 1px dashed #DCDFE6;
                          border-radius: 8px; color: #909399; font-weight: bold;
                          font-size: 12px; }
            QToolButton:hover { border: 1px dashed #E6A23C; background-color: #FDF6EC;
                                color: #E6A23C; }
        """

        self.buttons = {}
        for i, (name, path, kind) in enumerate(self.apps):
            is_url = (kind == "url")
            # url 类视为始终可用；exe 类按文件存在性判断
            exe_exists = True if is_url else (bool(path) and os.path.isfile(path))

            btn = QToolButton()
            # v12.10.76：长名称折两行显示（_wrap_btn_label），不存在时前加 ⚠
            disp = self._wrap_btn_label(name)
            btn.setText(disp if exe_exists else f"⚠ {self._wrap_btn_label(name)}")
            btn.setToolButtonStyle(Qt.ToolButtonTextUnderIcon)
            btn.setIcon(QIcon())
            btn.setIconSize(QSize(self.ICON_SIZE, self.ICON_SIZE))
            btn.setFixedSize(self.BTN_SIZE_W, self.BTN_SIZE_H)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(_normal_style if exe_exists else _missing_style)
            # 全名 tooltip（两行显示仍保证 hover 能看到完整原名）
            btn.setToolTip(name)

            # v12.10.75：url 类直接用 resources 下指定 png 作图标
            if is_url:
                ico_p = self._url_app_icons.get(name, "")
                if ico_p and os.path.isfile(ico_p):
                    btn.setIcon(QIcon(ico_p))
                btn.setToolTip(f"{name}（在默认浏览器打开）")

            # 双击启动（_launch_app 内部按 kind 分派 exe / url）
            flt = _DoubleClickFilter(callback=lambda n=name: self._launch_app(n), parent=btn)
            btn.installEventFilter(flt)
            self._click_filters.append(flt)

            # exe 不存在时挂 hover 提示过滤器（url 类不进此分支）
            if not exe_exists:
                hint_text = (
                    f"⚠ 未找到 {os.path.basename(path)}，请检查安装路径："
                    f"{os.path.dirname(path)}"
                )
                hover_flt = _HoverHintFilter(
                    hint_text=hint_text,
                    enter_cb=self.gui.update_status,
                    leave_cb=lambda: self.gui.update_status(""),
                    parent=btn,
                )
                btn.installEventFilter(hover_flt)
                self._click_filters.append(hover_flt)
                # tooltip 也设上（与状态栏互补，hover 略停后显示气泡提示）
                btn.setToolTip(hint_text)

            self.buttons[name] = btn
            self.grid.addWidget(btn, 0, i)

        # v12.10.75：IconLoaderThread 只扫 exe 类（url 类图标已直接设好）
        self.loader = IconLoaderThread([(n, p) for n, p, k in self.apps if k == "exe"])
        self.loader.icon_found.connect(self.update_icon)
        # v12.10.64：外部 .ico 没找到时（Fork 等内嵌图标的 exe），
        # 让主线程用 QFileIconProvider 从 exe 资源提取系统图标兜底
        self.loader.use_system_icon.connect(self.update_icon_from_exe)
        self.loader.start()

        # ======== 分割线 ========
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #EBEEF5;")
        layout.addSpacing(4)
        layout.addWidget(sep)
        layout.addSpacing(4)

        # ======== 下：导航 tab（v12.10.62：一站式导航 tab 暂隐藏）========
        # weblinks_path 保留计算（不去 _resolve_resource_path 的话恢复时要再加），
        # 仅在恢复一站式导航 tab 时才会被使用
        weblinks_path = self._resolve_resource_path("weblink.xlsx")  # noqa: F841
        ecm_files_path = self._resolve_resource_path("ecm_files.xlsx")

        self.tabs = QTabWidget()
        # tab 文字短，不需要滚动按钮 — 关掉避免 Qt 在某些状态下误判
        # 而把最后一个 tab 截断显示为 "▶" 翻页按钮
        self.tabs.setUsesScrollButtons(False)
        # 与 SPAhub 主窗口风格一致的轻量 tab 样式
        self.tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 1px solid #DCE0E8; border-radius: 6px;
                background: #FFFFFF; top: -1px;
            }
            QTabBar::tab {
                background: #EEF2F8; color: #606266;
                padding: 7px 20px; margin-right: 4px;
                border: 1px solid #DCE0E8;
                border-top-left-radius: 6px; border-top-right-radius: 6px;
                font-size: 13px;
            }
            QTabBar::tab:selected {
                background: #4A6FA5; color: #FFFFFF;
                border: 1px solid #4A6FA5; font-weight: bold;
            }
            QTabBar::tab:hover:!selected { background: #E2E9F6; color: #4A6FA5; }
        """)

        # v12.10.61：hint 文案精简 — 删除"(同步保存到本地)"说明
        # （用户反馈这个括号可能让人误解为有数据上传 / 同步动作，实际上
        # 收藏只是写入本地的 weblinks_bookmarks.json，与"同步"无关）
        unified_hint = "单击打开 · 鼠标悬停点击 ☆ 收藏"

        # v12.10.62：用户决定暂时隐藏"一站式导航"tab —— 整段代码保留，
        # 仅注释掉 addTab 与构造，未来要恢复直接取消注释即可。
        # 不要删除以下行，保留 weblinks_path / WebNavWidget 调用样板，
        # 便于一键恢复。
        #
        # self.web_nav = WebNavWidget(
        #     weblinks_path, parent=self,
        #     sheet_name="",   # 走默认逻辑（优先 weblinks sheet）
        #     hint=unified_hint,
        # )
        # self.tabs.addTab(self.web_nav, "一站式导航")
        self.web_nav = None  # 占位 — 外部如有引用判 None 即可

        # ECM 三个文档库 tab
        # 数据源：resources/ecm_files.xlsx，按 sheet 名分到对应 tab
        # 由 resources/ecm_file_list.py 离线刷新生成（不在 GUI 暴露刷新入口）
        # v12.10.75：tab 顺序倒序为 SOP → WI → 系统性工具手册（手册）
        sp_tabs_config = [
            ("sop", "SOP"),
            ("wi", "WI 定稿版"),
            ("manuals", "系统性工具手册"),
        ]
        self._sp_widgets = {}
        for sheet_name, tab_label in sp_tabs_config:
            w = WebNavWidget(
                ecm_files_path, parent=self,
                sheet_name=sheet_name,
                hint=unified_hint,
            )
            self.tabs.addTab(w, tab_label)
            self._sp_widgets[sheet_name] = w

        layout.addWidget(self.tabs, 1)

    # -------------------------------------------------------------------------
    @staticmethod
    def _resolve_resource_path(filename: str) -> str:
        """
        定位 SPAhub/resources/<filename>。以 __file__ 为锚，不依赖 cwd。
        找不到则返回 cwd 下的同名候选路径（让上层 _read_weblinks 走空态分支）。
        """
        here = os.path.dirname(os.path.abspath(__file__))                # .../SPAhub/gui
        spahub_root = os.path.abspath(os.path.join(here, os.pardir))     # .../SPAhub
        candidate = os.path.join(spahub_root, "resources", filename)
        if os.path.exists(candidate):
            return candidate
        return os.path.join(os.getcwd(), "resources", filename)

    @staticmethod
    def _resolve_weblinks_path() -> str:
        """向后兼容：保留旧 API 名，转调 _resolve_resource_path。"""
        return ShortcutWidget._resolve_resource_path("weblink.xlsx")

    # -------------------------------------------------------------------------
    def _launch_app(self, name):
        # v12.10.75：条目升级为 (name, target, kind)；按 kind 分派 exe / url
        entry = next(((p, k) for n, p, k in self.apps if n == name), None)
        if not entry:
            return
        target, kind = entry

        # url 类：默认浏览器打开（webbrowser.open 在 Windows 走 ShellExecute，非阻塞）
        if kind == "url":
            try:
                webbrowser.open(target)
                self.gui.update_status(f"已在浏览器打开：{name}")
            except Exception as e:
                self.gui._show_msg("打开失败", f"无法打开 {name}：\n{e}", "error")
            return

        # v12.10.90：动态 exe 目录（如 PFN）启动前再扫一次取最新 .exe，
        # 覆盖会话期间出新版本的情况；扫不到则沿用构造时解析的 target。
        if kind == "exe" and name in getattr(self, "_latest_exe_dirs", {}):
            target = _latest_exe_in_dir(self._latest_exe_dirs[name], fallback=target)

        # exe 类：不存在时不真启动，弹友好提示并把父目录展示给用户
        if not os.path.isfile(target):
            folder = os.path.dirname(target) or "(空路径)"
            self.gui._show_msg(
                "未找到程序",
                f"找不到 {name}.exe，请检查安装路径：\n\n{folder}\n\n"
                f"完整预期路径：\n{target}",
                "warning",
            )
            return
        thread = AppLaunchThread(target, name, parent=self)
        thread.status_signal.connect(self.gui.update_status)
        thread.error_signal.connect(lambda t, m: self.gui._show_msg(t, m, "error"))
        thread.finished.connect(lambda t=thread: self._on_thread_finished(t))
        self._launch_threads.append(thread)
        thread.start()

    def _on_thread_finished(self, thread):
        try:
            self._launch_threads.remove(thread)
        except ValueError:
            pass
        thread.deleteLater()

    def update_icon(self, name, ico_path):
        if name in self.buttons:
            self.buttons[name].setIcon(QIcon(ico_path))

    def update_icon_from_exe(self, name, exe_path):
        """
        v12.10.64：外部 .ico 没找到时的兜底 — 用 Qt QFileIconProvider 直接从
        exe 文件的 PE 资源中提取嵌入图标。

        覆盖场景：Fork、Pinnacle 21 这类 Electron / 现代应用，图标都内嵌在
        exe 资源里没有外部 .ico；Windows 资源管理器显示这些 exe 时用的就是
        这个嵌入图标，QFileIconProvider 内部走的是同一个系统 API
        (SHGetFileInfo)，所以拿到的图标与 Windows 默认显示完全一致。

        必须在主线程调用（Qt widget 系列限制），所以由 IconLoaderThread
        通过 use_system_icon signal 投递到主线程才执行。
        """
        if name not in self.buttons:
            return
        if not exe_path or not os.path.isfile(exe_path):
            return
        try:
            from PySide6.QtWidgets import QFileIconProvider
            from PySide6.QtCore import QFileInfo
            icon = QFileIconProvider().icon(QFileInfo(exe_path))
            if icon and not icon.isNull():
                self.buttons[name].setIcon(icon)
        except Exception:
            # 兜底失败也无所谓 — 按钮就保持无图标状态，文字仍然可见
            pass

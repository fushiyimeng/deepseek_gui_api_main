# -*- coding: utf-8 -*-
"""
PyQt5 聊天窗口框架（OpenAI 兼容模式）
—— 本地数据记录 + 按 token 计费 + 本地 tokenizer 预估 + 配置外置化

================================================================
一、配置文件（data/config.json）
================================================================
程序启动时读取，界面上可修改，也可直接用文本编辑器打开改。
结构：

    {
        "api_name":    "DeepSeek",
        "base_url":    "https://api.deepseek.com",
        "api_key":     "",
        "model":       "deepseek-flash",
        "balance":     0.0,
        "used_tokens": 0,
        "used_cost":   0.0
    }

    · api_name / base_url / api_key / model：调用接口所需
    · balance：账户余额（元），每次调用后自动按费用扣减
    · remain_tokens：剩余 token 数，每次调用后自动按用量扣减

================================================================
二、计费规则（依据本地时间戳自动判断）
================================================================
周一 ~ 周五：09:00~12:00、14:00~18:00 → 高峰；其余 → 空闲
周六 / 周日：全天 → 空闲
区间左闭右开：12:00、18:00 整点算空闲。

单价按「元 / 百万 token」计算，三类 token 单价不同：
    缓存命中输入 / 缓存未命中输入 / 输出

================================================================
三、本地数据落盘（默认目录 ./data）
================================================================
    data/
        config.json               配置（API、模型、余额、剩余token）
        usage.xlsx                每次调用的 token & 费用明细
        logs/app.log              运行日志
        chat/YYYYMMDDHHMM###.txt  聊天内容，按大小自动分卷
"""

import os
import sys
import json
import time
import logging
from datetime import datetime

from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QListWidget, QListWidgetItem, QTextEdit, QPushButton,
    QScrollArea, QFrame, QSplitter, QDialog, QFormLayout,
    QLineEdit, QDialogButtonBox, QDoubleSpinBox, QMessageBox,
    QMenu,
)

# Excel 依赖： pip install openpyxl
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font
    _HAS_OPENPYXL = True
except ImportError:
    _HAS_OPENPYXL = False


# ============================================================
# 【价格配置】单位：元 / 百万 token
# ------------------------------------------------------------
# 这部分是「计费算法」，不属于运行时参数，所以写在代码里。
# 未来加模型时在这里追加即可。
# ============================================================
MODEL_PRICING = {
    "deepseek-flash": {
        "idle": {"cache_hit": 0.02, "cache_miss": 1.0, "output": 4.0},
        "peak": {"cache_hit": 0.04, "cache_miss": 2.0, "output": 8.0},
    },
    "deepseek-v4-pro": {
        "idle": {"cache_hit": 0.15, "cache_miss": 4.5, "output": 13.5},
        "peak": {"cache_hit": 0.30, "cache_miss": 9.0, "output": 27.0},
    },
    # 兜底，防止 key 打错时程序崩溃（单价按 0 计，费用永远是 0）
    "_default": {
        "idle": {"cache_hit": 0.0, "cache_miss": 0.0, "output": 0.0},
        "peak": {"cache_hit": 0.0, "cache_miss": 0.0, "output": 0.0},
    },
}

PERIOD_PEAK  = "peak"
PERIOD_IDLE  = "idle"
PERIOD_LABEL = {PERIOD_PEAK: "高峰", PERIOD_IDLE: "空闲"}


# ============================================================
# 【本地 tokenizer 目录】
# ============================================================
TOKENIZER_DIR = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "deepseek_tokenizer"
)


# ============================================================
# 配置管理：读写 data/config.json
# ============================================================
class ConfigManager:
    """
    管理本地配置文件。任何字段都可以在界面上改、也可以手动打开文件改。

    - load()  读取；文件不存在或损坏时用默认值创建一份
    - save()  写回磁盘
    - get(k, default) / set(k, v)  便捷读写
    """

    # 默认配置。缺字段时用它补全，避免旧文件因少字段而崩。
    DEFAULTS = {
        "api_name":      "DeepSeek",
        "base_url":      "https://api.deepseek.com",
        "api_key":       "",
        "model":         "deepseek-flash",
        "balance":       0.0,
        "used_tokens":   0,        # 累计已用 token（每次调用后自动累加）
        "used_cost":     0.0,      # 累计已用费用（元）
    }

    def __init__(self, path):
        self.path = path
        self.data = dict(self.DEFAULTS)
        self.load()

    def load(self):
        if not os.path.exists(self.path):
            self.save()
            return
        try:
            with open(self.path, "r", encoding="utf-8") as f:
                disk = json.load(f)
            # 先用默认值垫底，再把磁盘上的所有字段覆盖上去（不再过滤）
            merged = dict(self.DEFAULTS)
            merged.update(disk)
            self.data = merged
        except Exception as e:
            logging.getLogger("chat_app").error(
                f"读取 config.json 失败，使用默认值：{e}"
            )
            self.data = dict(self.DEFAULTS)

    def save(self):
        """把内存配置写回磁盘。写完的文件可以直接用编辑器打开修改。"""
        try:
            os.makedirs(os.path.dirname(self.path), exist_ok=True)
            with open(self.path, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            logging.getLogger("chat_app").error(f"保存 config.json 失败：{e}")

    def get(self, key, default=None):
        return self.data.get(key, default)

    def set(self, key, value):
        self.data[key] = value


# ============================================================
# 设置对话框：在界面上编辑配置
# ============================================================
class SettingsDialog(QDialog):
    """
    编辑 API / 模型 / 余额 / 已用 tokens。
    点「确定」后把结果写回 ConfigManager 并保存。
    """

    def __init__(self, config: ConfigManager, parent=None):
        super().__init__(parent)
        self.config = config
        self.setWindowTitle("设置")
        self.setMinimumWidth(420)

        form = QFormLayout(self)
        form.setContentsMargins(16, 16, 16, 16)
        form.setSpacing(10)

        # ---- API 名称 ----
        self.ed_api_name = QLineEdit(config.get("api_name", ""))

        # ---- Base URL ----
        self.ed_base_url = QLineEdit(config.get("base_url", ""))

        # ---- API Key（用密码模式显示，防止旁人瞄到）----
        self.ed_api_key = QLineEdit(config.get("api_key", ""))
        self.ed_api_key.setEchoMode(QLineEdit.Password)

        # ---- 模型名 ----
        self.ed_model = QLineEdit(config.get("model", ""))

        # ---- 余额（元）----
        self.sp_balance = QDoubleSpinBox()
        self.sp_balance.setRange(0, 1e9)
        self.sp_balance.setDecimals(6)
        self.sp_balance.setValue(float(config.get("balance", 0.0)))

        # ---- 已用 tokens（自动累加，允许手动校准/清零）----
        self.sp_tokens = QDoubleSpinBox()
        self.sp_tokens.setRange(0, 1e12)
        self.sp_tokens.setDecimals(0)
        self.sp_tokens.setValue(float(config.get("used_tokens", 0)))

        form.addRow("API 名称：",   self.ed_api_name)
        form.addRow("Base URL：",  self.ed_base_url)
        form.addRow("API Key：",   self.ed_api_key)
        form.addRow("模型名：",     self.ed_model)
        form.addRow("余额(元)：",   self.sp_balance)
        form.addRow("已用 tokens：", self.sp_tokens)

        # ---- 按钮 ----
        btns = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btns.accepted.connect(self._on_ok)
        btns.rejected.connect(self.reject)
        form.addRow(btns)

    def _on_ok(self):
        self.config.set("api_name",      self.ed_api_name.text().strip())
        self.config.set("base_url",      self.ed_base_url.text().strip())
        self.config.set("api_key",       self.ed_api_key.text().strip())
        self.config.set("model",         self.ed_model.text().strip())
        self.config.set("balance",       float(self.sp_balance.value()))
        self.config.set("used_tokens", int(self.sp_tokens.value()))
        self.config.save()
        self.accept()


# ============================================================
# 本地 tokenizer（可选功能，加载失败自动降级）
# ============================================================
class LocalTokenizer:
    """
    对 transformers 的 AutoTokenizer 做一层薄封装。
    用途：预估输入 token 数，和接口返回的真实值做对比。
    加载失败不影响聊天主流程。
    """

    _instance = None

    @classmethod
    def instance(cls):
        if cls._instance is None:
            cls._instance = cls()
        return cls._instance

    def __init__(self):
        self.ok = False
        self.tokenizer = None
        self.logger = logging.getLogger("chat_app")
        self._load()

    def _load(self):
        try:
            import transformers
        except ImportError:
            self.logger.warning(
                "未安装 transformers，本地 token 预估不可用。"
                "请执行：pip install transformers"
            )
            return

        if not os.path.isdir(TOKENIZER_DIR):
            self.logger.warning(
                f"tokenizer 目录不存在：{TOKENIZER_DIR}，本地预估不可用。"
            )
            return

        try:
            self.tokenizer = transformers.AutoTokenizer.from_pretrained(
                TOKENIZER_DIR, trust_remote_code=True
            )
            self.ok = True
            self.logger.info(f"本地 tokenizer 加载成功：{TOKENIZER_DIR}")
        except Exception as e:
            self.logger.warning(f"加载本地 tokenizer 失败：{e}")

    def count_text(self, text: str) -> int:
        if not self.ok or not text:
            return 0
        try:
            return len(self.tokenizer.encode(text))
        except Exception:
            return 0

    def count_messages(self, messages: list) -> int:
        """按官方 chat_template 渲染整段对话后算 token。"""
        if not self.ok or not messages:
            return 0
        try:
            ids = self.tokenizer.apply_chat_template(
                messages, tokenize=True, add_generation_prompt=False,
            )
            return len(ids)
        except Exception:
            # 退而求其次：逐条编码相加
            total = 0
            for m in messages:
                total += self.count_text(m.get("content") or "")
            return total


# ============================================================
# 数据记录器：Excel 用量 + 日志 + 聊天文本
# ============================================================
class DataRecorder:
    """
    统一负责三件本地落盘的事：
        1) usage.xlsx          —— 每次调用的 token & 费用明细
        2) logs/app.log        —— 运行日志
        3) chat/*.txt          —— 聊天内容，按大小自动分卷
    """

    CHAT_FILE_MAX_BYTES = 1 * 1024 * 1024   # 1 MB

    EXCEL_HEADERS = [
        "日期", "时间", "时段", "会话ID", "模型",
        "输入命中", "输入未命中", "输入合计",
        "输出", "合计Tokens",
        "命中费用(元)", "未命中费用(元)", "输出费用(元)", "总费用(元)",
        "本地预估输入", "备注",
    ]

    def __init__(self, base_dir="data"):
        self.base_dir  = base_dir
        self.log_dir   = os.path.join(base_dir, "logs")
        self.chat_dir  = os.path.join(base_dir, "chat")
        self.xlsx_path = os.path.join(base_dir, "usage.xlsx")

        for d in (self.base_dir, self.log_dir, self.chat_dir):
            os.makedirs(d, exist_ok=True)

        self.logger = self._init_logger()
        self.wb, self.ws = self._init_excel()

        self._chat_fp        = None
        self._chat_path      = None
        self._chat_size      = 0
        self._chat_day       = None
        self._chat_seq_today = 0

    # --------------------------------------------------------
    def _init_logger(self):
        logger = logging.getLogger("chat_app")
        if logger.handlers:
            return logger
        logger.setLevel(logging.INFO)
        fmt = logging.Formatter(
            "[%(asctime)s] [%(levelname)s] %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
        fh = logging.FileHandler(
            os.path.join(self.log_dir, "app.log"), encoding="utf-8"
        )
        fh.setFormatter(fmt)
        logger.addHandler(fh)
        sh = logging.StreamHandler()
        sh.setFormatter(fmt)
        logger.addHandler(sh)
        return logger

    # --------------------------------------------------------
    def _init_excel(self):
        if not _HAS_OPENPYXL:
            self.logger.warning(
                "未安装 openpyxl，Excel 用量统计不可用。"
                "请执行：pip install openpyxl"
            )
            return None, None

        if os.path.exists(self.xlsx_path):
            try:
                wb = load_workbook(self.xlsx_path)
                ws = wb.active
                old_header = [c.value for c in ws[1]]
                if old_header == self.EXCEL_HEADERS:
                    return wb, ws

                wb.close()
                bak = self.xlsx_path.replace(
                    ".xlsx", f"_备份{datetime.now():%Y%m%d%H%M%S}.xlsx"
                )
                os.rename(self.xlsx_path, bak)
                self.logger.warning(
                    f"usage.xlsx 表头不一致，已备份为 {os.path.basename(bak)}"
                )
            except Exception as e:
                self.logger.error(f"打开 usage.xlsx 失败，将重建：{e}")

        wb = Workbook()
        ws = wb.active
        ws.title = "用量统计"
        ws.append(self.EXCEL_HEADERS)
        for c in ws[1]:
            c.font = Font(bold=True)
        try:
            wb.save(self.xlsx_path)
        except Exception as e:
            self.logger.error(f"初始化 usage.xlsx 失败：{e}")
        return wb, ws

    # --------------------------------------------------------
    @staticmethod
    def get_period(dt: datetime) -> str:
        """判断高峰/空闲。周一~五 9-12、14-18 为高峰；其余（含周末）为空闲。"""
        if dt.weekday() >= 5:
            return PERIOD_IDLE
        h = dt.hour
        if 9 <= h < 12:
            return PERIOD_PEAK
        if 14 <= h < 18:
            return PERIOD_PEAK
        return PERIOD_IDLE

    # --------------------------------------------------------
    @staticmethod
    def get_price(model: str, period: str) -> dict:
        table = MODEL_PRICING.get(model) or MODEL_PRICING["_default"]
        return table.get(period) or table.get(PERIOD_IDLE, {}) or {}

    @staticmethod
    def parse_usage(raw: dict) -> dict:
        """
        兼容 DeepSeek / OpenAI 风格的 usage 字段。
        DeepSeek: prompt_tokens / prompt_cache_hit_tokens /
                  prompt_cache_miss_tokens / completion_tokens
        OpenAI  : prompt_tokens / prompt_tokens_details.cached_tokens /
                  completion_tokens / total_tokens
        """
        raw = raw or {}

        completion = int(raw.get("completion_tokens")
                         or raw.get("output_tokens") or 0)

        hit = int(raw.get("prompt_cache_hit_tokens") or 0)
        if hit == 0:
            details = raw.get("prompt_tokens_details") or {}
            hit = int(details.get("cached_tokens") or 0)

        prompt_total = int(raw.get("prompt_tokens")
                           or raw.get("input_tokens") or 0)

        miss = int(raw.get("prompt_cache_miss_tokens") or 0)
        if miss == 0 and prompt_total > 0:
            miss = max(0, prompt_total - hit)

        if prompt_total == 0:
            prompt_total = hit + miss

        total = int(raw.get("total_tokens") or (prompt_total + completion))

        return {
            "prompt_tokens":     prompt_total,
            "cache_hit_tokens":  hit,
            "cache_miss_tokens": miss,
            "completion_tokens": completion,
            "total_tokens":      total,
        }

    def calc_cost(self, model: str, parsed: dict, period: str):
        p = self.get_price(model, period)
        hit_cost  = parsed["cache_hit_tokens"]  / 1_000_000.0 * p.get("cache_hit", 0.0)
        miss_cost = parsed["cache_miss_tokens"] / 1_000_000.0 * p.get("cache_miss", 0.0)
        out_cost  = parsed["completion_tokens"] / 1_000_000.0 * p.get("output", 0.0)
        return hit_cost, miss_cost, out_cost, hit_cost + miss_cost + out_cost

    # --------------------------------------------------------
    def record_usage(self, model, session_id, usage: dict,
                     local_estimate: int = 0, note=""):
        """
        记录一次调用：写日志 + 追加 Excel。
        返回 (归一化usage, 总费用元)。
        """
        now    = datetime.now()
        period = self.get_period(now)
        parsed = self.parse_usage(usage)
        hit_cost, miss_cost, out_cost, total_cost = self.calc_cost(model, parsed, period)

        self.logger.info(
            f"用量 | {now:%Y-%m-%d %H:%M:%S} | {PERIOD_LABEL[period]} | {model} | "
            f"in-hit={parsed['cache_hit_tokens']} "
            f"in-miss={parsed['cache_miss_tokens']} "
            f"out={parsed['completion_tokens']} "
            f"total={parsed['total_tokens']} | "
            f"本地预估={local_estimate} | "
            f"费用={total_cost:.6f}元 | sid={session_id}"
        )

        if not _HAS_OPENPYXL or self.ws is None:
            return parsed, total_cost

        row = [
            now.strftime("%Y-%m-%d"),
            now.strftime("%H:%M:%S"),
            PERIOD_LABEL[period],
            session_id or "",
            model,
            parsed["cache_hit_tokens"],
            parsed["cache_miss_tokens"],
            parsed["prompt_tokens"],
            parsed["completion_tokens"],
            parsed["total_tokens"],
            round(hit_cost,  6),
            round(miss_cost, 6),
            round(out_cost,  6),
            round(total_cost, 6),
            local_estimate,
            note,
        ]
        try:
            self.ws.append(row)
            self.wb.save(self.xlsx_path)
        except PermissionError:
            self.logger.warning(f"usage.xlsx 被占用，未写入：{row}")
        except Exception as e:
            self.logger.error(f"写入 usage.xlsx 失败：{e}")

        return parsed, total_cost

    # --------------------------------------------------------
    def append_chat(self, role, content, session_id=None):
        """把一条消息写入当前分卷文件，超限自动切新文件。"""
        try:
            line = self._format_line(role, content, session_id)
            data = line.encode("utf-8")

            need_new = (
                self._chat_fp is None
                or self._chat_size + len(data) > self.CHAT_FILE_MAX_BYTES
            )
            if need_new:
                self._rotate_chat_file()

            self._chat_fp.write(line)
            self._chat_fp.flush()
            self._chat_size += len(data)
        except Exception as e:
            self.logger.error(f"写入聊天文本失败：{e}")

    @staticmethod
    def _format_line(role, content, session_id) -> str:
        ts  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        tag = {"user": "用户", "assistant": "AI", "system": "系统"}.get(role, role)
        sid = f"[{session_id}] " if session_id else ""
        return f"[{ts}] {sid}[{tag}] {content}\n\n"

    def _rotate_chat_file(self):
        """关闭当前文件，按 YYYYMMDDHHMM + 3 位当日序号 开新文件。"""
        if self._chat_fp:
            try:
                self._chat_fp.close()
            except Exception:
                pass
            self._chat_fp = None

        now = datetime.now()
        day_key = now.strftime("%Y%m%d")
        if day_key != self._chat_day:
            self._chat_day = day_key
            self._chat_seq_today = 0

        self._chat_seq_today += 1
        fname = f"{now.strftime('%Y%m%d%H%M')}{self._chat_seq_today:03d}.txt"
        self._chat_path = os.path.join(self.chat_dir, fname)
        self._chat_fp   = open(self._chat_path, "a", encoding="utf-8")
        self._chat_size = os.path.getsize(self._chat_path)

        header = (
            f"===== 聊天记录 {now:%Y-%m-%d %H:%M:%S} =====\n"
            f"文件：{fname}\n"
            f"{'=' * 40}\n\n"
        )
        self._chat_fp.write(header)
        self._chat_size += len(header.encode("utf-8"))
        self.logger.info(f"新建聊天文本：{fname}")

    def close(self):
        if self._chat_fp:
            try:
                self._chat_fp.close()
            except Exception:
                pass
            self._chat_fp = None
        if _HAS_OPENPYXL and self.wb is not None:
            try:
                self.wb.save(self.xlsx_path)
            except Exception:
                pass

class SessionStore:
    """
    会话持久化：每个会话一个 JSON 文件，放在 data/sessions/ 下。
    文件名用 sid，例如 s1.json。

    文件结构：
        {
            "sid":     "s1",
            "title":   "新会话 1",
            "created": "2026-09-13 12:00:00",
            "updated": "2026-09-13 12:05:00",
            "messages": [
                {"role": "user",      "content": "..."},
                {"role": "assistant", "content": "..."}
            ]
        }
    """

    def __init__(self, base_dir):
        self.dir = os.path.join(base_dir, "sessions")
        os.makedirs(self.dir, exist_ok=True)
        self.logger = logging.getLogger("chat_app")

    def path_of(self, sid):
        return os.path.join(self.dir, f"{sid}.json")

    def load_all(self) -> dict:
        """扫描目录，返回 {sid: session}。缺字段自动补全，损坏文件跳过。"""
        sessions = {}
        for fname in os.listdir(self.dir):
            if not fname.endswith(".json"):
                continue
            sid = fname[:-5]
            try:
                with open(os.path.join(self.dir, fname), "r", encoding="utf-8") as f:
                    data = json.load(f)
                data.setdefault("sid", sid)
                data.setdefault("title", sid)
                data.setdefault("created", "")
                data.setdefault("updated", "")
                data.setdefault("messages", [])
                sessions[sid] = data
            except Exception as e:
                self.logger.error(f"加载会话文件失败 {fname}: {e}")
        return sessions

    def save(self, session: dict):
        """原子写入：先写 .tmp，再 replace，崩溃时不会把原文件写坏。"""
        if not session.get("sid"):
            return
        path = self.path_of(session["sid"])
        tmp  = path + ".tmp"
        try:
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(session, f, ensure_ascii=False, indent=2)
            os.replace(tmp, path)
        except Exception as e:
            self.logger.error(f"保存会话失败 {path}: {e}")

    def delete(self, sid):
        path = self.path_of(sid)
        try:
            if os.path.exists(path):
                os.remove(path)
        except Exception as e:
            self.logger.error(f"删除会话文件失败 {path}: {e}")

# ============================================================
# 输入框：Enter 发送，Shift+Enter 换行
# ============================================================
class ChatInput(QTextEdit):
    submitted = pyqtSignal()

    def keyPressEvent(self, event):
        if event.key() in (Qt.Key_Return, Qt.Key_Enter):
            if not (event.modifiers() & Qt.ShiftModifier):
                self.submitted.emit()
                return
        super().keyPressEvent(event)


# ============================================================
# 单条消息气泡
# ============================================================
class ChatMessage(QWidget):
    def __init__(self, text, is_user=False, parent=None):
        super().__init__(parent)
        self.is_user = is_user

        bubble = QFrame()
        bubble.setObjectName("bubble")
        bubble.setAttribute(Qt.WA_StyledBackground, True)

        if is_user:
            bubble.setStyleSheet(
                "#bubble { background-color: #c8e4ff; border-radius: 10px; }"
                "QLabel  { background: transparent; color: #14181d; }"
            )
        else:
            bubble.setStyleSheet(
                "#bubble { background-color: #f1f2f4; border-radius: 10px; }"
                "QLabel  { background: transparent; color: #14181d; }"
            )

        self.label = QLabel(text or "")
        self.label.setWordWrap(True)
        self.label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        self.label.setFont(QFont("Microsoft YaHei", 10))
        self.label.setMaximumWidth(640)

        bl = QVBoxLayout(bubble)
        bl.setContentsMargins(12, 8, 12, 8)
        bl.addWidget(self.label)

        h = QHBoxLayout(self)
        h.setContentsMargins(6, 3, 6, 3)
        if is_user:
            h.addStretch(1)
            h.addWidget(bubble)
        else:
            h.addWidget(bubble)
            h.addStretch(1)

    def set_text(self, text):
        self.label.setText(text)


# ============================================================
# 后台线程：调用 OpenAI 兼容接口
# ============================================================
class ApiWorker(QThread):
    finished_ok = pyqtSignal(str, dict)   # (完整文本, 原始usage)
    failed      = pyqtSignal(str)

    def __init__(self, messages, model, api_key, base_url,
                 session_id="", parent=None):
        super().__init__(parent)
        self.messages   = messages
        self.model      = model
        self.api_key    = api_key
        self.base_url   = base_url
        self.session_id = session_id
        self.local_estimate = 0      # 由主线程设置，方便日志对比

    def run(self):
        try:
            text, usage = self._call_api()
            self.finished_ok.emit(text, usage or {})
        except Exception as e:
            self.failed.emit(str(e))

    # --------------------------------------------------------
    def _call_api(self):
        """
        调用 DeepSeek 的 OpenAI 兼容接口。
        base_url 用 https://api.deepseek.com（不要带 /v1，官方就这样）。
        返回 (回复文本, 原始 usage 字典)。
        """
        from openai import OpenAI

        client = OpenAI(
            api_key  = self.api_key,
            base_url = self.base_url,
        )

        resp = client.chat.completions.create(
            model       = self.model,
            messages    = self.messages,
            temperature = 0.7,
            # stream    = True,          # 想做流式再打开
        )

        text  = resp.choices[0].message.content or ""
        usage = resp.usage.model_dump() if resp.usage else {}
        return text, usage

# ============================================================
# 主窗口
# ============================================================
class ChatWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("AI 聊天助手")
        self.resize(1180, 760)

        # ---------- 数据目录 ----------
        self.data_dir = os.path.join(os.getcwd(), "data")
        os.makedirs(self.data_dir, exist_ok=True)

        # ---------- 配置 & 记录器 ----------
        self.config = ConfigManager(os.path.join(self.data_dir, "config.json"))
        self.recorder = DataRecorder(base_dir=self.data_dir)
        self.recorder.logger.info("应用启动")

        # ---------- 运行时状态 ----------
        self.used_tokens = int(self.config.get("used_tokens", 0))
        self.used_cost   = 0.0
        # ---------- 会话持久化 ----------
        self.session_store = SessionStore(self.data_dir)
        self.sessions = self.session_store.load_all()
        self.current_sid = None

        # 从已有会话文件名推算 counter，避免新会话 id 冲突
        self._sid_counter = 0
        for sid in self.sessions:
            if sid.startswith("s") and sid[1:].isdigit():
                self._sid_counter = max(self._sid_counter, int(sid[1:]))

        self.worker = None

        # 本地 tokenizer
        self.tokenizer = LocalTokenizer.instance()

        # ---------- 预估防抖定时器 ----------
        # 输入停顿 300ms 后才真正算 token，避免每敲一个字都跑一遍模板渲染
        self._estimate_timer = QTimer(self)
        self._estimate_timer.setSingleShot(True)
        self._estimate_timer.setInterval(300)
        self._estimate_timer.timeout.connect(self._do_estimate)
        self._pending_messages = []

        # ---------- UI ----------
        self._build_ui()
        self._refresh_header()
        # 恢复左侧列表
        self._restore_sessions()

        # 一个会话都没有 → 新建一个；否则选中最后一个
        if not self.sessions:
            self._new_session()
        else:
            self.session_list.setCurrentRow(self.session_list.count() - 1)

        # 每分钟刷新「当前时段」
        self._period_timer = QTimer(self)
        self._period_timer.timeout.connect(self._refresh_header)
        self._period_timer.start(60 * 1000)

    # ==================== 界面搭建 ====================
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        root.addWidget(self._build_topbar())

        splitter = QSplitter(Qt.Horizontal)
        splitter.addWidget(self._build_sidebar())
        splitter.addWidget(self._build_chat_panel())
        splitter.setStretchFactor(0, 0)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([220, 960])
        splitter.setChildrenCollapsible(False)
        root.addWidget(splitter, 1)

    def _build_topbar(self):
        bar = QFrame()
        bar.setObjectName("topbar")
        bar.setFixedHeight(52)
        bar.setStyleSheet(
            "#topbar { background-color: #2b2f36; }"
            "QLabel  { color: #e8e8e8; font-size: 13px; }"
        )
        lay = QHBoxLayout(bar)
        lay.setContentsMargins(16, 0, 16, 0)
        lay.setSpacing(14)

        self.lbl_api     = QLabel()
        self.lbl_model   = QLabel()
        self.lbl_period  = QLabel()
        self.lbl_balance = QLabel()
        self.lbl_tokens  = QLabel()
        self.lbl_cost    = QLabel()

        self.lbl_api.setStyleSheet("color:#ffffff; font-weight:bold;")
        self.lbl_balance.setStyleSheet("color:#7fd18b;")
        self.lbl_cost.setStyleSheet("color:#f0b350;")
        self.lbl_period.setStyleSheet("color:#8fb6ff;")
        self.lbl_tokens.setStyleSheet("color:#e8e8e8;")

        # 右上角设置按钮
        btn_settings = QPushButton("设置")
        btn_settings.setFixedSize(56, 28)
        btn_settings.setStyleSheet(
            "QPushButton { background-color:#444a55; color:#e8e8e8; border:none; "
            "border-radius:6px; }"
            "QPushButton:hover { background-color:#5a6270; }"
        )
        btn_settings.clicked.connect(self._open_settings)

        lay.addWidget(self.lbl_api)
        lay.addWidget(self.lbl_model)
        lay.addWidget(self.lbl_period)
        lay.addStretch(1)
        lay.addWidget(self.lbl_balance)
        lay.addWidget(self.lbl_tokens)
        lay.addWidget(self.lbl_cost)
        lay.addWidget(btn_settings)
        return bar

    def _refresh_header(self):
        """刷新顶栏。API / 模型 / 余额 / 已用 token 都从 config 读。"""
        now_period = DataRecorder.get_period(datetime.now())
        self.lbl_api.setText(f"API：{self.config.get('api_name', '')}")
        self.lbl_model.setText(f"模型：{self.config.get('model', '')}")
        self.lbl_period.setText(f"当前时段：{PERIOD_LABEL[now_period]}")

        balance = float(self.config.get("balance", 0.0))
        used    = int(self.config.get("used_tokens", 0))
        used_cost = float(self.config.get("used_cost", 0.0))
        self.lbl_balance.setText(f"余额：{balance:.4f} 元")
        self.lbl_tokens.setText(f"已用 tokens：{used:,}")
        self.lbl_cost.setText(f"累计费用：{used_cost:.6f} 元")

    def _open_settings(self):
        dlg = SettingsDialog(self.config, self)
        if dlg.exec_() == QDialog.Accepted:
            self.recorder.logger.info("配置已更新")
            self._refresh_header()

    def _build_sidebar(self):
        w = QWidget()
        w.setMinimumWidth(180)
        w.setMaximumWidth(320)
        w.setStyleSheet("background-color: #f7f8fa;")

        lay = QVBoxLayout(w)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(8)

        title = QLabel("会话列表")
        title.setStyleSheet("color:#333; font-weight:bold; font-size:13px;")

        btn_new = QPushButton("+  新建会话")
        btn_new.setFixedHeight(32)
        btn_new.clicked.connect(self._new_session)

        # 先创建
        self.session_list = QListWidget()
        self.session_list.setStyleSheet(
            "QListWidget { border:1px solid #e0e0e0; border-radius:6px; background:#ffffff; }"
            "QListWidget::item { padding:8px; }"
            "QListWidget::item:selected { background:#d6e8ff; color:#000; }"
        )

        # 再设置右键菜单
        self.session_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.session_list.customContextMenuRequested.connect(self._on_session_menu)
        self.session_list.currentItemChanged.connect(self._on_session_changed)

        lay.addWidget(title)
        lay.addWidget(btn_new)
        lay.addWidget(self.session_list, 1)
        return w

    def _build_chat_panel(self):
        w = QWidget()
        lay = QVBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet("background-color: #ffffff;")

        self.chat_container = QWidget()
        self.chat_layout = QVBoxLayout(self.chat_container)
        self.chat_layout.setContentsMargins(12, 12, 12, 12)
        self.chat_layout.setSpacing(6)
        self.chat_layout.addStretch(1)
        self.scroll.setWidget(self.chat_container)

        lay.addWidget(self.scroll, 1)

        # 预估 token 提示条
        est_bar = QWidget()
        est_bar.setStyleSheet("background-color:#f7f8fa;")
        el = QHBoxLayout(est_bar)
        el.setContentsMargins(14, 4, 14, 0)
        el.setSpacing(8)

        self.lbl_estimate = QLabel("预计输入 tokens：—")
        self.lbl_estimate.setStyleSheet("color:#888; font-size:12px;")
        el.addWidget(self.lbl_estimate)
        el.addStretch(1)
        lay.addWidget(est_bar)

        # 底部输入区
        input_bar = QWidget()
        input_bar.setStyleSheet(
            "background-color:#f7f8fa; border-top:1px solid #e2e2e2;"
        )
        il = QHBoxLayout(input_bar)
        il.setContentsMargins(12, 6, 12, 12)
        il.setSpacing(10)

        self.input_edit = ChatInput()
        self.input_edit.setPlaceholderText("输入消息…  Enter 发送，Shift+Enter 换行")
        self.input_edit.setFixedHeight(88)
        self.input_edit.setStyleSheet(
            "QTextEdit { border:1px solid #dcdcdc; border-radius:8px; "
            "padding:6px; background:#ffffff; font-size:13px; }"
        )
        self.input_edit.submitted.connect(self._on_send)
        self.input_edit.textChanged.connect(self._on_input_changed)

        self.btn_send = QPushButton("发送")
        self.btn_send.setFixedSize(88, 88)
        self.btn_send.setStyleSheet(
            "QPushButton { background-color:#3b82f6; color:#fff; border:none; "
            "border-radius:8px; font-size:14px; }"
            "QPushButton:hover { background-color:#2f6fd8; }"
            "QPushButton:disabled { background-color:#a9c4ee; }"
        )
        self.btn_send.clicked.connect(self._on_send)

        il.addWidget(self.input_edit, 1)
        il.addWidget(self.btn_send)
        lay.addWidget(input_bar)
        return w

    # ==================== 会话管理 ====================
    def _new_session(self):
        self._sid_counter += 1
        sid = f"s{self._sid_counter}"
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        self.sessions[sid] = {
            "sid":      sid,
            "title":    f"新会话 {self._sid_counter}",
            "created":  now,
            "updated":  now,
            "messages": [],
        }
        self.session_store.save(self.sessions[sid])   # 立即落盘

        item = QListWidgetItem(self.sessions[sid]["title"])
        item.setData(Qt.UserRole, sid)
        self.session_list.addItem(item)
        self.session_list.setCurrentItem(item)
        self.recorder.logger.info(f"新建会话：{sid}")

    def _on_session_changed(self, current, previous):
        if current is None:
            return
        self.current_sid = current.data(Qt.UserRole)
        self._render_session()
        self._on_input_changed()

    def _render_session(self):
        while self.chat_layout.count() > 1:
            item = self.chat_layout.takeAt(0)
            widget = item.widget()
            if widget:
                widget.deleteLater()

        for msg in self.sessions.get(self.current_sid, {}).get("messages", []):
            self._append_message(
                msg["content"], msg["role"] == "user",
                store=False, persist=False,
            )
        self._scroll_to_bottom()

    def _update_session_title(self):
        sess = self.sessions.get(self.current_sid)
        if not sess or not sess["messages"]:
            return
        if sess["messages"][0]["role"] != "user":
            return
        title = sess["messages"][0]["content"][:15]
        sess["title"] = title
        item = self.session_list.currentItem()
        if item:
            item.setText(title)

    # ==================== 消息渲染 ====================
    def _append_message(self, text, is_user, store=True, persist=True):
        """
        store   : 是否写入内存会话历史
        persist : 是否写入本地聊天 txt（渲染历史时置 False）
        """
        if self.current_sid is None:
            return None

        role = "user" if is_user else "assistant"

        if store:
            sess = self.sessions[self.current_sid]
            sess["messages"].append({"role": role, "content": text})
            sess["updated"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            if is_user:
                self._update_session_title()
            self.session_store.save(sess)          # ← 每次变化立即落盘

        if persist:
            self.recorder.append_chat(role, text, session_id=self.current_sid)

        widget = ChatMessage(text, is_user)
        self.chat_layout.insertWidget(self.chat_layout.count() - 1, widget)
        self._scroll_to_bottom()
        return widget

    def _scroll_to_bottom(self):
        QTimer.singleShot(
            0,
            lambda: self.scroll.verticalScrollBar().setValue(
                self.scroll.verticalScrollBar().maximum()
            ),
        )

    # ==================== 本地预估 ====================
    def _on_input_changed(self):
        """输入变化时，把待发送 messages 暂存起来，由防抖定时器真正计算。"""
        if self.current_sid is None:
            return
        history = list(self.sessions[self.current_sid]["messages"])
        draft = self.input_edit.toPlainText()
        pending = history + ([{"role": "user", "content": draft}] if draft else [])
        self._pending_messages = pending
        self._estimate_timer.start()      # 300ms 无输入后再算

    def _do_estimate(self):
        if not self.tokenizer.ok:
            self.lbl_estimate.setText("预计输入 tokens：—（本地 tokenizer 未加载）")
            return
        pending = self._pending_messages or []
        n = self.tokenizer.count_messages(pending)
        self.lbl_estimate.setText(f"预计输入 tokens：{n}")

    # ==================== 发送逻辑 ====================
    def _on_send(self):
        if self.worker is not None and self.worker.isRunning():
            return
        if self.current_sid is None:
            return

        # 从配置里拿 key / base_url；为空就提示去设置里填
        api_key  = self.config.get("api_key", "").strip()
        base_url = self.config.get("base_url", "").strip()
        model    = self.config.get("model", "").strip()

        if not api_key or not base_url or not model:
            QMessageBox.warning(
                self, "配置不完整",
                "请先点击右上角「设置」，填写 Base URL、API Key 和模型名。"
            )
            return

        text = self.input_edit.toPlainText().strip()
        if not text:
            return

        self.input_edit.clear()
        self._append_message(text, is_user=True)
        self._set_busy(True)

        messages = list(self.sessions[self.current_sid]["messages"])
        # messages.insert(0, {"role": "system", "content": "你是一个有用的助手。"})

        # 本地预估
        local_estimate = self.tokenizer.count_messages(messages)

        self.worker = ApiWorker(
            messages, model, api_key, base_url,
            session_id=self.current_sid,
        )
        self.worker.local_estimate = local_estimate
        self.worker.finished_ok.connect(self._on_reply)
        self.worker.failed.connect(self._on_error)
        self.worker.start()

        self._pending_messages = messages
        self._do_estimate()

    def _on_reply(self, text, usage):
        self._append_message(text, is_user=False)
        self._set_busy(False)

        local_estimate = getattr(self.worker, "local_estimate", 0)
        self._update_usage(usage, local_estimate)

    def _on_error(self, message):
        self._set_busy(False)
        self.recorder.logger.error(f"API 请求失败：{message}")
        self._append_message(f"[请求失败] {message}", is_user=False,
                             store=False, persist=False)

    def _set_busy(self, busy):
        self.btn_send.setEnabled(not busy)
        self.btn_send.setText("思考中…" if busy else "发送")

    # ==================== 用量 & 费用 ====================
    def _update_usage(self, usage: dict, local_estimate: int = 0):
        """
        1) 写 Excel + 日志
        2) 从配置的余额里扣费
        3) 累计已用 token 写回 config.json
        """
        parsed, cost = self.recorder.record_usage(
            model=self.config.get("model", ""),
            session_id=self.current_sid,
            usage=usage or {},
            local_estimate=local_estimate,
        )

        self.used_tokens += parsed["total_tokens"]
        self.used_cost   += cost

        try:
            balance   = float(self.config.get("balance", 0.0))
            used      = int(self.config.get("used_tokens", 0))
            used_cost = float(self.config.get("used_cost", 0.0))

            new_balance   = max(0.0, balance - cost)
            new_used      = used + parsed["total_tokens"]
            new_used_cost = used_cost + cost

            self.config.set("balance",     round(new_balance, 6))
            self.config.set("used_tokens", new_used)
            self.config.set("used_cost",   round(new_used_cost, 6))
            self.config.save()
        except Exception as e:
            self.recorder.logger.error(f"更新余额/已用token失败：{e}")

        self._refresh_header()

    def update_balance(self, balance_text: str):
        """保留兼容：外部（比如查询接口）想覆盖余额时调用。"""
        try:
            self.config.set("balance", float(balance_text))
            self.config.save()
        except Exception:
            pass
        self._refresh_header()

    # ==================== 收尾 ====================
    def closeEvent(self, event):
        if self.worker is not None and self.worker.isRunning():
            self.worker.terminate()
            self.worker.wait(1000)
        self.config.save()
        self.recorder.logger.info("应用退出")
        self.recorder.close()
        super().closeEvent(event)

    def _restore_sessions(self):
        """启动时把磁盘上的会话灌进左侧列表（按 created 升序）。"""
        ordered = sorted(self.sessions.values(), key=lambda s: s.get("created", ""))
        for sess in ordered:
            item = QListWidgetItem(sess["title"])
            item.setData(Qt.UserRole, sess["sid"])
            self.session_list.addItem(item)

    def _on_session_menu(self, pos):
        item = self.session_list.itemAt(pos)
        if item is None:
            return
        sid = item.data(Qt.UserRole)
        title = self.sessions.get(sid, {}).get("title", sid)

        menu = QMenu(self)
        act_del = menu.addAction("删除此会话")
        if menu.exec_(self.session_list.mapToGlobal(pos)) != act_del:
            return

        if QMessageBox.question(
            self, "确认删除",
            f"确定删除会话「{title}」吗？\n会同时删除本地会话文件，不可恢复。",
        ) != QMessageBox.Yes:
            return

        self.session_store.delete(sid)
        self.sessions.pop(sid, None)
        self.session_list.takeItem(self.session_list.row(item))
        self.recorder.logger.info(f"删除会话：{sid}")

        # 删的是当前会话 → 手动把选中切走或新建
        if self.current_sid == sid:
            if self.session_list.count() > 0:
                self.session_list.setCurrentRow(0)
            else:
                self._new_session()


def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    app = QApplication(sys.argv)
    app.setFont(QFont("Microsoft YaHei", 9))
    window = ChatWindow()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
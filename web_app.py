#!/usr/bin/env python3
from __future__ import annotations

import cgi
import html
import os
import platform
import shutil
import socket
import subprocess
import sys
import time
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, quote, unquote, urlparse


WORKFLOW_DIR = Path(__file__).resolve().parent
PROJECT_DIR = WORKFLOW_DIR.parent
OUTPUT_DIR = WORKFLOW_DIR / "output"
INPUT_DIR = WORKFLOW_DIR / "input"
SUPPLIER_QUOTES_DIR = INPUT_DIR / "supplier_quotes"
STANDARD_DIR = INPUT_DIR / "standard"
STANDARD_FILE = STANDARD_DIR / "standard.xlsx"
RAW_DIR = INPUT_DIR / "raw"
RAW_FILE = RAW_DIR / "待处理采购清单.xlsx"
FEEDBACK_FILE = INPUT_DIR / "negotiation_feedback.xlsx"
DEFAULT_STANDARD = "/Users/houzhou/Library/Containers/com.tencent.xinWeChat/Data/Documents/xwechat_files/houge630973991_3046/temp/drag/永顺县芙蓉镇、青坪镇农贸市场人材机汇总表_询价清单(3).xlsx"


def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SUPPLIER_QUOTES_DIR.mkdir(parents=True, exist_ok=True)
    STANDARD_DIR.mkdir(parents=True, exist_ok=True)
    RAW_DIR.mkdir(parents=True, exist_ok=True)


def py() -> str:
    codex_python = Path("/Users/houzhou/.cache/codex-runtimes/codex-primary-runtime/dependencies/python/bin/python3")
    if codex_python.exists():
        return str(codex_python)
    return sys.executable


def safe_name(value: str) -> str:
    return "".join(ch for ch in value.strip() if ch not in "\\/:*?\"<>|") or "未命名"


def rel(path: Path) -> str:
    return str(path.resolve().relative_to(WORKFLOW_DIR.resolve()))


def latest_dir(keyword: str) -> Path | None:
    matches = [path for path in OUTPUT_DIR.glob(f"*{keyword}*") if path.is_dir()]
    return max(matches, key=lambda p: p.stat().st_mtime) if matches else None


def current_standard() -> Path:
    if STANDARD_FILE.exists():
        return STANDARD_FILE
    default_path = Path(DEFAULT_STANDARD)
    return default_path if default_path.exists() else STANDARD_FILE


def current_split_file() -> Path | None:
    raw_dir = latest_dir("由待处理清单生成拆分")
    standard_dir = latest_dir("按昨日样板标准拆分")
    candidates = [path for path in [raw_dir, standard_dir] if path]
    known_files = []
    if candidates:
        latest = max(candidates, key=lambda p: p.stat().st_mtime)
        known_files = [
            latest / "采购包拆分总表-由待处理清单生成.xlsx",
            latest / "采购包拆分总表-按昨日样板标准.xlsx",
        ]
    fallback_files = sorted(OUTPUT_DIR.rglob("采购包拆分总表-*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    for path in [*known_files, *fallback_files]:
        if path.exists():
            return path
    return None


def run_script(script: str, *args: str) -> str:
    command = [py(), str(WORKFLOW_DIR / "src" / script), *args]
    proc = subprocess.run(command, cwd=PROJECT_DIR, capture_output=True, text=True)
    output = (proc.stdout or "") + ("\n" + proc.stderr if proc.stderr else "")
    if proc.returncode != 0:
        raise RuntimeError(output or f"脚本运行失败：{script}")
    return output


def run_action(action: str) -> str:
    if action == "split":
        if RAW_FILE.exists():
            return run_script("process_raw_procurement_list.py", "--input", str(RAW_FILE), "--output-dir", str(OUTPUT_DIR))
        return run_script("build_split_from_standard.py", "--standard", str(current_standard()), "--output-dir", str(OUTPUT_DIR))
    if action == "invitations":
        split_file = current_split_file()
        if not split_file:
            if RAW_FILE.exists():
                run_script("process_raw_procurement_list.py", "--input", str(RAW_FILE), "--output-dir", str(OUTPUT_DIR))
            else:
                run_script("build_split_from_standard.py", "--standard", str(current_standard()), "--output-dir", str(OUTPUT_DIR))
            split_file = current_split_file()
        if not split_file:
            raise RuntimeError("没有找到采购包拆分总表。")
        return run_script("generate_invitations_from_split.py", "--standard", str(split_file), "--packages", "all", "--output-dir", str(OUTPUT_DIR))
    if action == "analyze":
        return run_script("analyze_supplier_quotes.py", "--output-dir", str(OUTPUT_DIR))
    if action == "finalize":
        return run_script("finalize_recommendation.py", "--output-dir", str(OUTPUT_DIR))
    if action == "full":
        out = []
        if RAW_FILE.exists():
            out.append(run_script("process_raw_procurement_list.py", "--input", str(RAW_FILE), "--output-dir", str(OUTPUT_DIR)))
        else:
            out.append(run_script("build_split_from_standard.py", "--standard", str(current_standard()), "--output-dir", str(OUTPUT_DIR)))
        split_file = current_split_file()
        if not split_file:
            raise RuntimeError("没有找到采购包拆分总表。")
        out.append(run_script("generate_invitations_from_split.py", "--standard", str(split_file), "--packages", "all", "--output-dir", str(OUTPUT_DIR)))
        out.append(run_script("analyze_supplier_quotes.py", "--output-dir", str(OUTPUT_DIR)))
        out.append(run_script("finalize_recommendation.py", "--output-dir", str(OUTPUT_DIR)))
        return "\n".join(out)
    raise RuntimeError(f"未知操作：{action}")


def open_browser(url: str) -> None:
    system = platform.system()
    if system == "Windows":
        os.startfile(url)  # type: ignore[attr-defined]
    elif system == "Darwin":
        subprocess.run(["open", url], check=False)
    else:
        subprocess.run(["xdg-open", url], check=False)


def page(title: str, body: str) -> bytes:
    html_text = f"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <style>
    body {{ margin:0; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Microsoft YaHei", sans-serif; background:#f6f7f9; color:#1f2933; }}
    header {{ background:#143b63; color:white; padding:22px 32px; }}
    main {{ max-width:1180px; margin:0 auto; padding:24px; }}
    h1 {{ margin:0; font-size:24px; }}
    h2 {{ font-size:18px; margin:0 0 14px; }}
    .sub {{ margin-top:6px; opacity:.86; }}
    .grid {{ display:grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap:16px; }}
    .card {{ background:white; border:1px solid #d9e2ec; border-radius:8px; padding:18px; box-shadow:0 1px 2px rgba(15,23,42,.04); }}
    .btn {{ display:inline-block; border:0; background:#1769aa; color:white; padding:10px 14px; border-radius:6px; font-size:14px; cursor:pointer; text-decoration:none; }}
    .btn.secondary {{ background:#52606d; }}
    .btn.warn {{ background:#b95000; }}
    .row {{ display:flex; flex-wrap:wrap; gap:10px; align-items:center; }}
    input, select {{ padding:9px; border:1px solid #bcccdc; border-radius:6px; background:white; }}
    table {{ width:100%; border-collapse:collapse; background:white; }}
    th, td {{ border:1px solid #d9e2ec; padding:8px; font-size:13px; vertical-align:top; }}
    th {{ background:#e7eef7; text-align:left; }}
    pre {{ white-space:pre-wrap; background:#102a43; color:#f0f4f8; padding:14px; border-radius:8px; max-height:360px; overflow:auto; }}
    .muted {{ color:#627d98; font-size:13px; }}
    .ok {{ color:#0b6b3a; }}
    .path {{ font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; font-size:12px; word-break:break-all; }}
  </style>
</head>
<body>
  <header>
    <h1>采购询价智能助手</h1>
    <div class="sub">本地网页原型：拆分采购包、生成询价包、分析报价、输出推荐结果</div>
  </header>
  <main>{body}</main>
</body>
</html>"""
    return html_text.encode("utf-8")


def file_link(path: Path, label: str | None = None) -> str:
    if not path.exists():
        return ""
    href = "/download?path=" + quote(rel(path))
    return f'<a href="{href}">{html.escape(label or path.name)}</a>'


def list_outputs() -> str:
    rows = []
    for path in sorted([p for p in OUTPUT_DIR.iterdir() if p.is_dir()], key=lambda p: p.stat().st_mtime, reverse=True)[:12]:
        files = [p for p in path.rglob("*") if p.is_file() and p.suffix.lower() in [".xlsx", ".docx", ".txt"]]
        links = " / ".join(file_link(f) for f in files[:4])
        if len(files) > 4:
            links += f" / 另有 {len(files)-4} 个文件"
        rows.append(f"<tr><td>{html.escape(path.name)}</td><td>{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(path.stat().st_mtime))}</td><td>{links}</td></tr>")
    if not rows:
        return "<p class='muted'>暂无输出结果。</p>"
    return "<table><tr><th>输出目录</th><th>时间</th><th>文件</th></tr>" + "".join(rows) + "</table>"


def home(message: str = "", log: str = "") -> bytes:
    standard = current_standard()
    standard_text = str(standard) if standard.exists() else "尚未上传样板标准文件"
    raw_text = str(RAW_FILE) if RAW_FILE.exists() else "尚未上传待处理采购清单"
    body = f"""
    {f'<div class="card"><strong class="ok">{html.escape(message)}</strong></div><br>' if message else ''}
    {f'<pre>{html.escape(log)}</pre><br>' if log else ''}
    <div class="grid">
      <section class="card">
        <h2>1. 上传待处理采购清单</h2>
        <p class="muted">这里上传未分类、待采购询价的清单 Excel。当前仅支持 .xlsx，旧版 .xls 请先在 Excel/WPS 中另存为 .xlsx。当前：<span class="path">{html.escape(raw_text)}</span></p>
        <form method="post" action="/upload_raw" enctype="multipart/form-data">
          <input type="file" name="file" accept=".xlsx" required>
          <button class="btn" type="submit">上传待处理清单</button>
        </form>
      </section>
      <section class="card">
        <h2>2. 生成采购包拆分总表</h2>
        <p class="muted">按内置规则识别询价大类、采购包、供应商类型，并标出需人工确认项。</p>
        <form method="post" action="/run"><input type="hidden" name="action" value="split"><button class="btn" type="submit">开始拆分</button></form>
      </section>
      <section class="card">
        <h2>3. 批量生成询价包</h2>
        <p class="muted">生成全部采购包的 Word 邀请包和 Excel 报价清单。</p>
        <form method="post" action="/run"><input type="hidden" name="action" value="invitations"><button class="btn" type="submit">生成询价包</button></form>
      </section>
      <section class="card">
        <h2>4. 上传供应商报价</h2>
        <p class="muted">选择采购包名称，上传供应商返回的报价 Excel。当前仅支持 .xlsx。</p>
        <form method="post" action="/upload_quote" enctype="multipart/form-data" class="row">
          <input name="package" placeholder="采购包名称，如 电线电缆包" required>
          <input type="file" name="file" accept=".xlsx" required>
          <button class="btn" type="submit">上传报价</button>
        </form>
      </section>
      <section class="card">
        <h2>5. 分析报价并出谈判策略</h2>
        <p class="muted">自动合并报价，检查漏项、异常项，输出谈判策略。</p>
        <form method="post" action="/run"><input type="hidden" name="action" value="analyze"><button class="btn" type="submit">分析报价</button></form>
      </section>
      <section class="card">
        <h2>6. 最终推荐</h2>
        <p class="muted">下载/填写谈判反馈表后，生成推荐合作结果。</p>
        <div class="row">
          {file_link(FEEDBACK_FILE, "下载谈判反馈表") or '<form method="post" action="/run"><input type="hidden" name="action" value="finalize"><button class="btn secondary" type="submit">先生成反馈表</button></form>'}
          <form method="post" action="/upload_feedback" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx" required>
            <button class="btn secondary" type="submit">上传已填反馈表</button>
          </form>
          <form method="post" action="/run"><input type="hidden" name="action" value="finalize"><button class="btn" type="submit">生成推荐结果</button></form>
        </div>
      </section>
    </div>
    <br>
    <section class="card">
        <h2>高级设置：上传分类标准样板</h2>
      <p class="muted">一般同事不用管。只有当你要替换后台分类标准时，才上传你昨天那类已经整理好的样板文件。当前仅支持 .xlsx。当前：<span class="path">{html.escape(standard_text)}</span></p>
      <form method="post" action="/upload_standard" enctype="multipart/form-data">
        <input type="file" name="file" accept=".xlsx" required>
        <button class="btn secondary" type="submit">上传分类标准样板</button>
      </form>
    </section>
    <br>
    <section class="card">
      <h2>一键完整闭环</h2>
      <p class="muted">适合资料已经放好时使用：拆分、生成询价包、分析报价、生成推荐结果一次跑完。</p>
      <form method="post" action="/run"><input type="hidden" name="action" value="full"><button class="btn warn" type="submit">一键运行完整闭环</button></form>
    </section>
    <br>
    <section class="card">
      <h2>最近输出结果</h2>
      {list_outputs()}
    </section>
    """
    return page("采购询价智能助手", body)


class Handler(BaseHTTPRequestHandler):
    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self.send_html(home())
        elif parsed.path == "/download":
            self.download(parsed.query)
        else:
            self.send_error(404)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        try:
            if parsed.path == "/run":
                length = int(self.headers.get("Content-Length", "0"))
                data = self.rfile.read(length).decode("utf-8")
                action = parse_qs(data).get("action", [""])[0]
                log = run_action(action)
                self.send_html(home("操作已完成", log))
            elif parsed.path == "/upload_standard":
                self.handle_upload(STANDARD_FILE)
                self.send_html(home("分类标准样板已上传"))
            elif parsed.path == "/upload_raw":
                self.handle_upload(RAW_FILE)
                self.send_html(home("待处理采购清单已上传"))
            elif parsed.path == "/upload_quote":
                self.handle_quote_upload()
                self.send_html(home("供应商报价已上传"))
            elif parsed.path == "/upload_feedback":
                self.handle_upload(FEEDBACK_FILE)
                self.send_html(home("谈判反馈表已上传"))
            else:
                self.send_error(404)
        except Exception as exc:
            self.send_html(home("操作失败", str(exc)))

    def multipart_form(self) -> cgi.FieldStorage:
        return cgi.FieldStorage(
            fp=self.rfile,
            headers=self.headers,
            environ={
                "REQUEST_METHOD": "POST",
                "CONTENT_TYPE": self.headers.get("Content-Type", ""),
                "CONTENT_LENGTH": self.headers.get("Content-Length", "0"),
            },
        )

    def uploaded_file(self, form: cgi.FieldStorage) -> cgi.FieldStorage:
        if "file" not in form:
            raise RuntimeError("未收到文件，请重新选择 Excel 文件后上传。")
        item = form["file"]
        if isinstance(item, list):
            item = item[0]
        if not getattr(item, "file", None) or not getattr(item, "filename", ""):
            raise RuntimeError("未收到有效文件，请确认上传的是 .xlsx 文件。")
        suffix = Path(str(item.filename)).suffix.lower()
        if suffix == ".xls":
            raise RuntimeError("当前原型暂不支持旧版 .xls 文件。请先用 Excel/WPS 打开该文件，另存为 .xlsx 后再上传。")
        if suffix != ".xlsx":
            raise RuntimeError("当前仅支持 .xlsx 文件，请选择 Excel 工作簿 .xlsx 后重新上传。")
        return item

    def handle_upload(self, target: Path) -> None:
        form = self.multipart_form()
        item = self.uploaded_file(form)
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("wb") as out:
            shutil.copyfileobj(item.file, out)

    def handle_quote_upload(self) -> None:
        form = self.multipart_form()
        package = safe_name(form.getfirst("package", "未命名采购包"))
        item = self.uploaded_file(form)
        filename = safe_name(getattr(item, "filename", "") or "供应商报价.xlsx")
        target = SUPPLIER_QUOTES_DIR / package / filename
        target.parent.mkdir(parents=True, exist_ok=True)
        with target.open("wb") as out:
            shutil.copyfileobj(item.file, out)

    def download(self, query: str) -> None:
        path_value = parse_qs(query).get("path", [""])[0]
        if not path_value:
            self.send_error(400)
            return
        target = (WORKFLOW_DIR / unquote(path_value)).resolve()
        if not str(target).startswith(str(WORKFLOW_DIR.resolve())) or not target.exists() or not target.is_file():
            self.send_error(404)
            return
        data = target.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", "application/octet-stream")
        self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(target.name)}")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def send_html(self, data: bytes) -> None:
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, format: str, *args: object) -> None:
        return


def local_ip() -> str:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"


def main() -> None:
    ensure_dirs()
    lan_mode = "--lan" in sys.argv
    cloud_mode = "--cloud" in sys.argv
    no_open = "--no-open" in sys.argv or cloud_mode
    host = "0.0.0.0" if lan_mode or cloud_mode else "127.0.0.1"
    port = int(os.environ.get("PORT", "8876"))
    server = ThreadingHTTPServer((host, port), Handler)
    url = f"http://127.0.0.1:{port}"
    print(f"采购询价智能助手已启动：{url}")
    if cloud_mode:
        print(f"云端模式已启用，监听地址：0.0.0.0:{port}")
    if lan_mode:
        print(f"同一局域网同事可尝试访问：http://{local_ip()}:{port}")
        print("注意：需要你的电脑和同事电脑在同一网络，且防火墙允许访问。")
    print("关闭此窗口即可停止服务。")
    if not no_open:
        try:
            open_browser(url)
        except Exception:
            pass
    server.serve_forever()


if __name__ == "__main__":
    main()

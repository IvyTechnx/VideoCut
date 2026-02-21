#!/usr/bin/env python3
"""
VideoCut - シンプルな動画カットエディタ for macOS
ブラウザベースのUIで動画をプレビューしながらカット編集できます。
"""

import http.server
import json
import os
import subprocess
import sys
import tempfile
import threading
import urllib.parse
import webbrowser
import mimetypes
import shutil

PORT = 0  # 自動で空きポートを使用
VIDEO_DIR = os.path.expanduser("~")  # デフォルトのディレクトリ


class VideoCutHandler(http.server.BaseHTTPRequestHandler):
    """HTTPリクエストハンドラー"""

    def log_message(self, format, *args):
        # ログを抑制（必要なら有効化）
        pass

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        query = urllib.parse.parse_qs(parsed.query)

        if path == "/":
            self._serve_html()
        elif path == "/api/browse":
            self._handle_browse(query)
        elif path == "/api/video-info":
            self._handle_video_info(query)
        elif path.startswith("/video/"):
            self._serve_video(path)
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)
        path = parsed.path
        content_len = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(content_len)

        if path == "/api/export":
            self._handle_export(json.loads(body))
        elif path == "/api/select-file":
            self._handle_select_file(json.loads(body))
        else:
            self.send_error(404)

    def _send_json(self, data, status=200):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    # --- ファイルブラウズ ---
    def _handle_browse(self, query):
        dir_path = query.get("dir", [VIDEO_DIR])[0]
        if not os.path.isdir(dir_path):
            dir_path = os.path.expanduser("~")

        video_exts = {".mp4", ".mov", ".avi", ".mkv", ".m4v", ".webm", ".ts", ".mts", ".m2ts"}
        items = []

        # 親ディレクトリ
        parent = os.path.dirname(dir_path)
        if parent != dir_path:
            items.append({"name": "..", "path": parent, "type": "dir"})

        try:
            entries = sorted(os.listdir(dir_path), key=str.lower)
        except PermissionError:
            self._send_json({"error": "アクセス権がありません", "dir": dir_path, "items": []})
            return

        for name in entries:
            if name.startswith("."):
                continue
            full = os.path.join(dir_path, name)
            if os.path.isdir(full):
                items.append({"name": name + "/", "path": full, "type": "dir"})
            else:
                ext = os.path.splitext(name)[1].lower()
                if ext in video_exts:
                    size = os.path.getsize(full)
                    items.append({
                        "name": name, "path": full, "type": "video",
                        "size": self._format_size(size)
                    })

        self._send_json({"dir": dir_path, "items": items})

    def _handle_select_file(self, data):
        """AppleScriptでファイル選択ダイアログを表示"""
        script = '''
        tell application "System Events"
            activate
            set theFile to choose file with prompt "動画ファイルを選択" of type {"public.movie", "public.mpeg-4", "com.apple.quicktime-movie", "public.avi"}
            return POSIX path of theFile
        end tell
        '''
        try:
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=60
            )
            if result.returncode == 0 and result.stdout.strip():
                file_path = result.stdout.strip()
                self._send_json({"path": file_path})
            else:
                self._send_json({"path": None, "cancelled": True})
        except Exception as e:
            self._send_json({"error": str(e)})

    # --- 動画情報 ---
    def _handle_video_info(self, query):
        path = query.get("path", [""])[0]
        if not os.path.isfile(path):
            self._send_json({"error": "ファイルが見つかりません"}, 404)
            return

        cmd = [
            "ffprobe", "-v", "quiet", "-print_format", "json",
            "-show_format", "-show_streams", path
        ]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=10)
            info = json.loads(result.stdout)
            # 動画ストリーム情報を抽出
            video_stream = None
            for s in info.get("streams", []):
                if s.get("codec_type") == "video":
                    video_stream = s
                    break

            fmt = info.get("format", {})
            duration = float(fmt.get("duration", 0))
            response = {
                "path": path,
                "filename": os.path.basename(path),
                "duration": duration,
                "size": self._format_size(int(fmt.get("size", 0))),
            }
            if video_stream:
                response["width"] = int(video_stream.get("width", 0))
                response["height"] = int(video_stream.get("height", 0))
                # fpsの計算
                r_fps = video_stream.get("r_frame_rate", "30/1")
                try:
                    num, den = r_fps.split("/")
                    response["fps"] = round(float(num) / float(den), 2)
                except (ValueError, ZeroDivisionError):
                    response["fps"] = 30.0
                response["codec"] = video_stream.get("codec_name", "unknown")

            self._send_json(response)
        except Exception as e:
            self._send_json({"error": str(e)}, 500)

    # --- 動画ストリーミング ---
    def _serve_video(self, path):
        # /video/<encoded_path>
        encoded = path[len("/video/"):]
        file_path = urllib.parse.unquote(encoded)
        if not os.path.isfile(file_path):
            self.send_error(404)
            return

        file_size = os.path.getsize(file_path)
        content_type = mimetypes.guess_type(file_path)[0] or "video/mp4"

        # Range対応（シーク可能にする）
        range_header = self.headers.get("Range")
        if range_header:
            ranges = range_header.replace("bytes=", "").split("-")
            start = int(ranges[0]) if ranges[0] else 0
            end = int(ranges[1]) if ranges[1] else file_size - 1
            length = end - start + 1

            self.send_response(206)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Range", f"bytes {start}-{end}/{file_size}")
            self.send_header("Content-Length", length)
            self.send_header("Accept-Ranges", "bytes")
            self.end_headers()

            with open(file_path, "rb") as f:
                f.seek(start)
                remaining = length
                while remaining > 0:
                    chunk = f.read(min(65536, remaining))
                    if not chunk:
                        break
                    self.wfile.write(chunk)
                    remaining -= len(chunk)
        else:
            self.send_response(200)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", file_size)
            self.send_header("Accept-Ranges", "bytes")
            self.end_headers()

            with open(file_path, "rb") as f:
                while True:
                    chunk = f.read(65536)
                    if not chunk:
                        break
                    self.wfile.write(chunk)

    # --- エクスポート ---
    def _handle_export(self, data):
        source = data["source"]
        segments = data["segments"]  # [{"in": seconds, "out": seconds}, ...]
        output_dir = data.get("output_dir", os.path.dirname(source))
        vflip = data.get("vflip", False)

        if not segments:
            self._send_json({"error": "セグメントが指定されていません"}, 400)
            return

        base, ext = os.path.splitext(os.path.basename(source))
        output_name = f"{base}_cut{ext}"
        output_path = os.path.join(output_dir, output_name)

        # ファイル名の重複回避
        counter = 1
        while os.path.exists(output_path):
            output_name = f"{base}_cut_{counter}{ext}"
            output_path = os.path.join(output_dir, output_name)
            counter += 1

        # vflip時はフィルタ付き再エンコード、それ以外は無劣化コピー
        def _codec_args():
            if vflip:
                return ["-vf", "vflip", "-c:v", "libx264", "-preset", "medium",
                        "-crf", "18", "-c:a", "copy"]
            else:
                return ["-c", "copy"]

        try:
            if len(segments) == 1:
                seg = segments[0]
                duration = seg["out"] - seg["in"]
                cmd = [
                    "ffmpeg", "-y",
                    "-ss", f"{seg['in']:.4f}",
                    "-i", source,
                    "-t", f"{duration:.4f}",
                ] + _codec_args() + [
                    "-avoid_negative_ts", "make_zero",
                    output_path
                ]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
                if result.returncode != 0:
                    self._send_json({
                        "error": f"ffmpegエラー: {result.stderr[-300:]}"
                    }, 500)
                    return
            else:
                # 複数セグメント: 個別処理後にconcat
                temp_dir = tempfile.mkdtemp()
                temp_files = []

                for i, seg in enumerate(segments):
                    duration = seg["out"] - seg["in"]
                    temp_file = os.path.join(temp_dir, f"seg_{i:04d}.ts")
                    temp_files.append(temp_file)
                    cmd = [
                        "ffmpeg", "-y",
                        "-ss", f"{seg['in']:.4f}",
                        "-i", source,
                        "-t", f"{duration:.4f}",
                    ]
                    if vflip:
                        cmd += ["-vf", "vflip", "-c:v", "libx264", "-preset", "medium",
                                "-crf", "18", "-c:a", "copy"]
                    else:
                        cmd += ["-c", "copy"]
                    cmd += [
                        "-avoid_negative_ts", "make_zero",
                        "-f", "mpegts",
                        temp_file
                    ]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
                    if result.returncode != 0:
                        shutil.rmtree(temp_dir, ignore_errors=True)
                        self._send_json({
                            "error": f"セグメント{i+1}のエラー: {result.stderr[-300:]}"
                        }, 500)
                        return

                concat_list = os.path.join(temp_dir, "concat.txt")
                with open(concat_list, "w") as f:
                    for tf in temp_files:
                        f.write(f"file '{tf}'\n")

                cmd = [
                    "ffmpeg", "-y",
                    "-f", "concat", "-safe", "0",
                    "-i", concat_list,
                    "-c", "copy",
                    output_path
                ]
                result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
                shutil.rmtree(temp_dir, ignore_errors=True)

                if result.returncode != 0:
                    self._send_json({
                        "error": f"結合エラー: {result.stderr[-300:]}"
                    }, 500)
                    return

            file_size = os.path.getsize(output_path)
            self._send_json({
                "success": True,
                "output": output_path,
                "size": self._format_size(file_size)
            })
        except subprocess.TimeoutExpired:
            self._send_json({"error": "処理がタイムアウトしました"}, 500)
        except Exception as e:
            self._send_json({"error": str(e)}, 500)

    # --- HTML UI ---
    def _serve_html(self):
        html = get_html()
        body = html.encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", len(body))
        self.end_headers()
        self.wfile.write(body)

    @staticmethod
    def _format_size(size):
        for unit in ["B", "KB", "MB", "GB"]:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"


def get_html():
    return '''<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>VideoCut</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
:root {
    --bg: #141414;
    --surface: #1e1e1e;
    --surface2: #252525;
    --border: #333;
    --text: #e0e0e0;
    --text2: #999;
    --accent: #4fc3f7;
    --green: #4caf50;
    --red: #f44336;
    --yellow: #ffb74d;
}
body {
    font-family: -apple-system, BlinkMacSystemFont, "Helvetica Neue", sans-serif;
    background: var(--bg);
    color: var(--text);
    height: 100vh;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    user-select: none;
}
/* ヘッダー */
.header {
    display: flex;
    align-items: center;
    padding: 8px 16px;
    background: var(--surface);
    border-bottom: 1px solid var(--border);
    gap: 12px;
}
.header h1 {
    font-size: 16px;
    font-weight: 600;
    color: var(--accent);
}
.header .filename {
    font-size: 13px;
    color: var(--text2);
    flex: 1;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
}
.btn {
    padding: 6px 14px;
    border: 1px solid var(--border);
    background: var(--surface2);
    color: var(--text);
    border-radius: 6px;
    font-size: 12px;
    cursor: pointer;
    transition: all 0.15s;
    white-space: nowrap;
}
.btn:hover { background: #333; border-color: #555; }
.btn:active { transform: scale(0.97); }
.btn.accent { background: #1a5276; border-color: var(--accent); color: var(--accent); }
.btn.accent:hover { background: #1e6a9a; }
.btn.green { background: #1b4d1e; border-color: var(--green); color: var(--green); }
.btn.green:hover { background: #256d29; }
.btn.red { background: #4a1515; border-color: var(--red); color: var(--red); }
.btn.red:hover { background: #6a1f1f; }

/* メイン */
.main { flex: 1; display: flex; flex-direction: column; overflow: hidden; }

/* プレビューエリア */
.preview-area {
    flex: 1;
    display: flex;
    align-items: center;
    justify-content: center;
    background: #000;
    position: relative;
    min-height: 200px;
}
.preview-area video {
    max-width: 100%;
    max-height: 100%;
}
.preview-area video.flipped {
    transform: scaleY(-1);
}
.btn.toggle-on {
    background: #1a5276;
    border-color: var(--accent);
    color: var(--accent);
    box-shadow: inset 0 0 8px rgba(79, 195, 247, 0.2);
}
.preview-area .placeholder {
    color: #444;
    font-size: 18px;
    text-align: center;
    line-height: 1.6;
}

/* タイムライン */
.timeline-section {
    background: var(--surface);
    border-top: 1px solid var(--border);
    padding: 8px 16px;
}
.time-display {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 6px;
}
.time-current {
    font-family: "Menlo", "SF Mono", monospace;
    font-size: 15px;
    color: var(--accent);
    font-weight: 600;
}
.time-info {
    font-size: 11px;
    color: var(--text2);
}

/* カスタムタイムライン */
.timeline-track {
    position: relative;
    height: 36px;
    background: #111;
    border-radius: 4px;
    cursor: pointer;
    overflow: visible;
    margin: 4px 0;
}
.timeline-track .track-bg {
    position: absolute;
    top: 8px;
    left: 0;
    right: 0;
    height: 20px;
    background: #222;
    border-radius: 3px;
}
.timeline-track .segment {
    position: absolute;
    top: 8px;
    height: 20px;
    background: rgba(76, 175, 80, 0.3);
    border: 1px solid rgba(76, 175, 80, 0.5);
    border-radius: 2px;
}
.timeline-track .current-segment {
    position: absolute;
    top: 8px;
    height: 20px;
    background: rgba(79, 195, 247, 0.2);
    border: 1px solid rgba(79, 195, 247, 0.4);
    border-radius: 2px;
}
.timeline-track .playhead {
    position: absolute;
    top: 0;
    width: 2px;
    height: 36px;
    background: var(--accent);
    transform: translateX(-1px);
    pointer-events: none;
    z-index: 10;
}
.timeline-track .playhead::before {
    content: '';
    position: absolute;
    top: 0;
    left: -5px;
    border-left: 6px solid transparent;
    border-right: 6px solid transparent;
    border-top: 8px solid var(--accent);
}
.timeline-track .marker-in,
.timeline-track .marker-out {
    position: absolute;
    top: 4px;
    width: 3px;
    height: 28px;
    transform: translateX(-1px);
    pointer-events: none;
    z-index: 5;
}
.timeline-track .marker-in { background: var(--green); }
.timeline-track .marker-out { background: var(--red); }
.timeline-track .marker-in::after,
.timeline-track .marker-out::after {
    position: absolute;
    bottom: -14px;
    font-size: 10px;
    font-weight: bold;
}
.timeline-track .marker-in::after { content: 'I'; color: var(--green); left: -1px; }
.timeline-track .marker-out::after { content: 'O'; color: var(--red); left: -2px; }

/* IN/OUT表示 */
.io-display {
    display: flex;
    gap: 16px;
    font-size: 12px;
    color: var(--text2);
    margin-top: 2px;
}
.io-display .in-time { color: var(--green); }
.io-display .out-time { color: var(--red); }
.io-display .sel-duration { margin-left: auto; color: var(--yellow); }

/* コントロール */
.controls {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    padding: 8px 16px;
    background: var(--surface);
    border-top: 1px solid var(--border);
    flex-wrap: wrap;
}
.controls .sep { width: 12px; }
.ctrl-btn {
    width: 40px;
    height: 32px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: 6px;
    color: var(--text);
    font-size: 14px;
    cursor: pointer;
    transition: all 0.15s;
}
.ctrl-btn:hover { background: #333; }
.ctrl-btn:active { transform: scale(0.95); }
.ctrl-btn.play { width: 50px; font-size: 18px; }

/* カットリスト */
.cutlist-section {
    background: var(--surface);
    border-top: 1px solid var(--border);
    max-height: 160px;
    display: flex;
    flex-direction: column;
}
.cutlist-header {
    display: flex;
    align-items: center;
    padding: 6px 16px;
    gap: 8px;
    font-size: 12px;
    color: var(--text2);
    border-bottom: 1px solid var(--border);
}
.cutlist-header .spacer { flex: 1; }
.cutlist-items {
    overflow-y: auto;
    flex: 1;
}
.cutlist-item {
    display: flex;
    align-items: center;
    padding: 4px 16px;
    font-family: "Menlo", monospace;
    font-size: 12px;
    gap: 8px;
    cursor: pointer;
    border-bottom: 1px solid #1a1a1a;
}
.cutlist-item:hover { background: #222; }
.cutlist-item .num { color: var(--text2); width: 24px; }
.cutlist-item .times { color: var(--text); flex: 1; }
.cutlist-item .dur { color: var(--text2); }
.cutlist-item .del-btn {
    color: #666;
    cursor: pointer;
    font-size: 14px;
    padding: 0 4px;
}
.cutlist-item .del-btn:hover { color: var(--red); }
.cutlist-empty {
    padding: 12px 16px;
    text-align: center;
    color: #444;
    font-size: 12px;
}

/* ステータスバー */
.statusbar {
    padding: 4px 16px;
    font-size: 11px;
    color: #666;
    background: var(--surface2);
    border-top: 1px solid var(--border);
}

/* キーボードショートカットヘルプ */
.shortcuts-hint {
    font-size: 11px;
    color: #555;
    display: flex;
    gap: 12px;
    flex-wrap: wrap;
}
.shortcuts-hint kbd {
    background: #2a2a2a;
    border: 1px solid #444;
    padding: 1px 5px;
    border-radius: 3px;
    font-family: "Menlo", monospace;
    font-size: 10px;
    color: #aaa;
}

/* モーダル */
.modal-overlay {
    display: none;
    position: fixed;
    top: 0; left: 0; right: 0; bottom: 0;
    background: rgba(0,0,0,0.7);
    z-index: 100;
    align-items: center;
    justify-content: center;
}
.modal-overlay.active { display: flex; }
.modal {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: 12px;
    padding: 24px;
    max-width: 500px;
    width: 90%;
}
.modal h2 { font-size: 16px; margin-bottom: 16px; }
.modal .export-info { font-size: 13px; color: var(--text2); margin-bottom: 16px; line-height: 1.6; }
.modal .btn-row { display: flex; gap: 8px; justify-content: flex-end; }
</style>
</head>
<body>

<!-- ヘッダー -->
<div class="header">
    <h1>VideoCut</h1>
    <span class="filename" id="filename">動画ファイルを開いてください</span>
    <button class="btn" onclick="openFile()">開く</button>
    <button class="btn" id="flipBtn" onclick="toggleFlip()" title="上下反転 (F)">上下反転</button>
    <button class="btn green" onclick="showExportModal()">エクスポート</button>
</div>

<!-- メイン -->
<div class="main">
    <!-- プレビュー -->
    <div class="preview-area" id="previewArea">
        <div class="placeholder" id="placeholder">
            動画ファイルを開いて編集を始めましょう<br>
            <span style="font-size:14px; color:#333">「開く」ボタンをクリック</span>
        </div>
        <video id="video" style="display:none" preload="auto"></video>
    </div>

    <!-- タイムライン -->
    <div class="timeline-section">
        <div class="time-display">
            <span class="time-current" id="timeCurrent">00:00:00.00</span>
            <span class="time-info" id="timeTotal">/ 00:00:00.00</span>
        </div>
        <div class="timeline-track" id="timeline">
            <div class="track-bg"></div>
            <div class="playhead" id="playhead" style="left:0"></div>
        </div>
        <div class="io-display">
            <span>IN: <span class="in-time" id="inTime">--:--:--.--</span></span>
            <span>OUT: <span class="out-time" id="outTime">--:--:--.--</span></span>
            <span class="sel-duration" id="selDuration"></span>
        </div>
    </div>

    <!-- コントロール -->
    <div class="controls">
        <div class="shortcuts-hint">
            <span><kbd>Space</kbd> 再生/停止</span>
            <span><kbd>←</kbd><kbd>→</kbd> 1秒</span>
            <span><kbd>Shift+←→</kbd> 10秒</span>
            <span><kbd>,</kbd><kbd>.</kbd> 1フレーム</span>
            <span><kbd>I</kbd> IN</span>
            <span><kbd>O</kbd> OUT</span>
            <span><kbd>A</kbd> 全選択</span>
            <span><kbd>Enter</kbd> 追加</span>
            <span><kbd>F</kbd> 上下反転</span>
        </div>
    </div>
    <div class="controls">
        <button class="ctrl-btn" onclick="seekTo(0)" title="先頭">⏮</button>
        <button class="ctrl-btn" onclick="stepTime(-10)" title="-10秒">⏪</button>
        <button class="ctrl-btn" onclick="stepTime(-1)" title="-1秒">◀</button>
        <button class="ctrl-btn play" onclick="togglePlay()" id="playBtn" title="再生/停止">▶</button>
        <button class="ctrl-btn" onclick="stepTime(1)" title="+1秒">▶</button>
        <button class="ctrl-btn" onclick="stepTime(10)" title="+10秒">⏩</button>
        <button class="ctrl-btn" onclick="seekTo(videoDuration)" title="末尾">⏭</button>
        <div class="sep"></div>
        <button class="btn accent" onclick="setInPoint()" title="INポイント設定 (I)">I - IN</button>
        <button class="btn accent" onclick="setOutPoint()" title="OUTポイント設定 (O)">O - OUT</button>
        <button class="btn" onclick="selectAll()" title="全選択 (A)">全選択</button>
        <button class="btn green" onclick="addToCutList()" title="カットリストに追加 (Enter)">＋ 追加</button>
    </div>

    <!-- カットリスト -->
    <div class="cutlist-section">
        <div class="cutlist-header">
            <span>カットリスト（保持するセグメント）</span>
            <span class="spacer"></span>
            <span id="cutlistSummary"></span>
            <button class="btn red" onclick="clearCutList()" style="font-size:11px; padding:3px 8px;">全削除</button>
        </div>
        <div class="cutlist-items" id="cutlistItems">
            <div class="cutlist-empty">セグメントなし — IN/OUTポイントを設定して追加してください</div>
        </div>
    </div>
</div>

<!-- ステータスバー -->
<div class="statusbar" id="statusbar">準備完了</div>

<!-- エクスポートモーダル -->
<div class="modal-overlay" id="exportModal">
    <div class="modal">
        <h2>エクスポート</h2>
        <div class="export-info" id="exportInfo"></div>
        <div class="btn-row">
            <button class="btn" onclick="closeExportModal()">キャンセル</button>
            <button class="btn green" onclick="doExport()">エクスポート</button>
        </div>
    </div>
</div>

<script>
// --- 状態 ---
let videoPath = null;
let videoInfo = null;
let videoDuration = 0;
let inPoint = null;
let outPoint = null;
let cutList = [];
const video = document.getElementById('video');
const timeline = document.getElementById('timeline');
const playhead = document.getElementById('playhead');
let isFlipped = false;
let isDragging = false;

// --- ファイルを開く ---
async function openFile() {
    setStatus('ファイル選択中...');
    try {
        const res = await fetch('/api/select-file', { method: 'POST', body: '{}' });
        const data = await res.json();
        if (data.path) {
            await loadVideo(data.path);
        } else {
            setStatus('準備完了');
        }
    } catch(e) {
        setStatus('エラー: ' + e.message);
    }
}

async function loadVideo(path) {
    setStatus('読み込み中...');
    try {
        const res = await fetch('/api/video-info?path=' + encodeURIComponent(path));
        videoInfo = await res.json();
        if (videoInfo.error) {
            setStatus('エラー: ' + videoInfo.error);
            return;
        }
        videoPath = path;
        videoDuration = videoInfo.duration;
        inPoint = null;
        outPoint = null;
        cutList = [];

        document.getElementById('filename').textContent =
            videoInfo.filename + ' (' + videoInfo.width + 'x' + videoInfo.height +
            ', ' + videoInfo.fps + 'fps, ' + videoInfo.codec + ', ' + videoInfo.size + ')';

        document.getElementById('placeholder').style.display = 'none';
        video.style.display = 'block';
        video.src = '/video/' + encodeURIComponent(path);
        video.load();

        updateIODisplay();
        refreshCutList();
        updateTimeTotal();
        setStatus('読み込み完了: ' + videoInfo.filename);
    } catch(e) {
        setStatus('エラー: ' + e.message);
    }
}

// --- 上下反転 ---
function toggleFlip() {
    isFlipped = !isFlipped;
    video.classList.toggle('flipped', isFlipped);
    document.getElementById('flipBtn').classList.toggle('toggle-on', isFlipped);
    setStatus(isFlipped ? '上下反転: ON（エクスポートにも反映されます）' : '上下反転: OFF');
}

// --- 再生制御 ---
function togglePlay() {
    if (!videoPath) return;
    if (video.paused) {
        video.play();
        document.getElementById('playBtn').textContent = '⏸';
    } else {
        video.pause();
        document.getElementById('playBtn').textContent = '▶';
    }
}

function stepTime(sec) {
    if (!videoPath) return;
    video.pause();
    document.getElementById('playBtn').textContent = '▶';
    video.currentTime = Math.max(0, Math.min(videoDuration, video.currentTime + sec));
}

function stepFrame(dir) {
    if (!videoPath || !videoInfo) return;
    video.pause();
    document.getElementById('playBtn').textContent = '▶';
    const frameDur = 1.0 / (videoInfo.fps || 30);
    video.currentTime = Math.max(0, Math.min(videoDuration, video.currentTime + dir * frameDur));
}

function seekTo(time) {
    if (!videoPath) return;
    video.pause();
    document.getElementById('playBtn').textContent = '▶';
    video.currentTime = Math.max(0, Math.min(videoDuration, time));
}

// --- タイムライン ---
video.addEventListener('timeupdate', updateTimeline);
video.addEventListener('pause', () => { document.getElementById('playBtn').textContent = '▶'; });
video.addEventListener('ended', () => { document.getElementById('playBtn').textContent = '▶'; });

function updateTimeline() {
    if (!videoPath || videoDuration === 0) return;
    const pct = (video.currentTime / videoDuration) * 100;
    playhead.style.left = pct + '%';
    document.getElementById('timeCurrent').textContent = formatTime(video.currentTime);
}

function updateTimeTotal() {
    document.getElementById('timeTotal').textContent = '/ ' + formatTime(videoDuration);
}

// タイムラインクリック&ドラッグ
timeline.addEventListener('mousedown', (e) => {
    isDragging = true;
    seekFromTimeline(e);
});
document.addEventListener('mousemove', (e) => {
    if (isDragging) seekFromTimeline(e);
});
document.addEventListener('mouseup', () => { isDragging = false; });

function seekFromTimeline(e) {
    if (!videoPath) return;
    const rect = timeline.getBoundingClientRect();
    let pct = (e.clientX - rect.left) / rect.width;
    pct = Math.max(0, Math.min(1, pct));
    video.currentTime = pct * videoDuration;
    updateTimeline();
}

// --- IN/OUT ---
function setInPoint() {
    if (!videoPath) return;
    inPoint = video.currentTime;
    if (outPoint !== null && outPoint <= inPoint) outPoint = null;
    updateIODisplay();
    drawTimelineMarkers();
    setStatus('INポイント: ' + formatTime(inPoint));
}

function setOutPoint() {
    if (!videoPath) return;
    outPoint = video.currentTime;
    if (inPoint !== null && inPoint >= outPoint) inPoint = null;
    updateIODisplay();
    drawTimelineMarkers();
    setStatus('OUTポイント: ' + formatTime(outPoint));
}

function selectAll() {
    if (!videoPath) return;
    inPoint = 0;
    outPoint = videoDuration;
    updateIODisplay();
    drawTimelineMarkers();
    setStatus('全選択: 0:00 → ' + formatTime(videoDuration));
}

function updateIODisplay() {
    document.getElementById('inTime').textContent = inPoint !== null ? formatTime(inPoint) : '--:--:--.--';
    document.getElementById('outTime').textContent = outPoint !== null ? formatTime(outPoint) : '--:--:--.--';
    if (inPoint !== null && outPoint !== null) {
        document.getElementById('selDuration').textContent = '選択区間: ' + formatTime(outPoint - inPoint);
    } else {
        document.getElementById('selDuration').textContent = '';
    }
}

function drawTimelineMarkers() {
    // 既存マーカーを削除
    timeline.querySelectorAll('.marker-in, .marker-out, .segment, .current-segment').forEach(el => el.remove());

    if (videoDuration === 0) return;

    // カットリストのセグメント
    cutList.forEach(seg => {
        const el = document.createElement('div');
        el.className = 'segment';
        el.style.left = (seg.in / videoDuration * 100) + '%';
        el.style.width = ((seg.out - seg.in) / videoDuration * 100) + '%';
        timeline.appendChild(el);
    });

    // 現在のIN-OUT範囲
    if (inPoint !== null && outPoint !== null) {
        const el = document.createElement('div');
        el.className = 'current-segment';
        el.style.left = (inPoint / videoDuration * 100) + '%';
        el.style.width = ((outPoint - inPoint) / videoDuration * 100) + '%';
        timeline.appendChild(el);
    }

    // INマーカー
    if (inPoint !== null) {
        const el = document.createElement('div');
        el.className = 'marker-in';
        el.style.left = (inPoint / videoDuration * 100) + '%';
        timeline.appendChild(el);
    }
    // OUTマーカー
    if (outPoint !== null) {
        const el = document.createElement('div');
        el.className = 'marker-out';
        el.style.left = (outPoint / videoDuration * 100) + '%';
        timeline.appendChild(el);
    }
}

// --- カットリスト ---
function addToCutList() {
    if (inPoint === null || outPoint === null) {
        alert('INポイントとOUTポイントを設定してください。\\nI キー: IN / O キー: OUT');
        return;
    }
    if (outPoint <= inPoint) {
        alert('OUTポイントはINポイントより後に設定してください。');
        return;
    }
    cutList.push({ in: inPoint, out: outPoint });
    cutList.sort((a, b) => a.in - b.in);
    inPoint = null;
    outPoint = null;
    updateIODisplay();
    refreshCutList();
    drawTimelineMarkers();
    setStatus('カットリストに追加 (合計 ' + cutList.length + ' セグメント)');
}

function removeCut(idx) {
    cutList.splice(idx, 1);
    refreshCutList();
    drawTimelineMarkers();
}

function clearCutList() {
    if (cutList.length === 0) return;
    if (confirm('カットリストをすべて削除しますか？')) {
        cutList = [];
        refreshCutList();
        drawTimelineMarkers();
    }
}

function refreshCutList() {
    const container = document.getElementById('cutlistItems');
    if (cutList.length === 0) {
        container.innerHTML = '<div class="cutlist-empty">セグメントなし — IN/OUTポイントを設定して追加してください</div>';
        document.getElementById('cutlistSummary').textContent = '';
        return;
    }
    let totalDur = 0;
    let html = '';
    cutList.forEach((seg, i) => {
        const dur = seg.out - seg.in;
        totalDur += dur;
        html += '<div class="cutlist-item" ondblclick="seekTo(' + seg.in + ')">' +
            '<span class="num">' + (i+1) + '.</span>' +
            '<span class="times">' + formatTime(seg.in) + ' → ' + formatTime(seg.out) + '</span>' +
            '<span class="dur">' + formatTime(dur) + '</span>' +
            '<span class="del-btn" onclick="event.stopPropagation();removeCut(' + i + ')">✕</span>' +
            '</div>';
    });
    container.innerHTML = html;
    document.getElementById('cutlistSummary').textContent =
        cutList.length + ' セグメント / 合計 ' + formatTime(totalDur);
}

// --- エクスポート ---
function showExportModal() {
    if (!videoPath) { alert('まず動画ファイルを開いてください。'); return; }
    let segments = cutList.length > 0 ? cutList : (inPoint !== null && outPoint !== null ? [{ in: inPoint, out: outPoint }] : []);
    if (segments.length === 0) {
        alert('エクスポートするセグメントがありません。\\nIN/OUTポイントを設定してカットリストに追加してください。');
        return;
    }
    let totalDur = segments.reduce((s, seg) => s + (seg.out - seg.in), 0);
    document.getElementById('exportInfo').innerHTML =
        '<strong>ソース:</strong> ' + videoInfo.filename + '<br>' +
        '<strong>セグメント数:</strong> ' + segments.length + '<br>' +
        '<strong>合計時間:</strong> ' + formatTime(totalDur) + '<br>' +
        '<strong>上下反転:</strong> ' + (isFlipped ? '<span style="color:var(--accent)">ON（再エンコードあり）</span>' : 'OFF') + '<br>' +
        '<strong>出力:</strong> ' + (isFlipped ? '再エンコード（H.264）' : '無劣化コピー（再エンコードなし）') + '<br>' +
        '<strong>保存先:</strong> ソースファイルと同じフォルダ';
    document.getElementById('exportModal').classList.add('active');
}

function closeExportModal() {
    document.getElementById('exportModal').classList.remove('active');
}

async function doExport() {
    closeExportModal();
    let segments = cutList.length > 0 ? cutList : [{ in: inPoint, out: outPoint }];
    setStatus('エクスポート中...');
    try {
        const res = await fetch('/api/export', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                source: videoPath,
                segments: segments,
                output_dir: videoPath.substring(0, videoPath.lastIndexOf('/')),
                vflip: isFlipped
            })
        });
        const data = await res.json();
        if (data.error) {
            alert('エクスポートエラー:\\n' + data.error);
            setStatus('エクスポート失敗');
        } else {
            setStatus('エクスポート完了: ' + data.output.split('/').pop() + ' (' + data.size + ')');
            alert('エクスポート完了!\\n' + data.output + '\\nサイズ: ' + data.size);
        }
    } catch(e) {
        alert('エラー: ' + e.message);
        setStatus('エクスポート失敗');
    }
}

// --- キーボード ---
document.addEventListener('keydown', (e) => {
    if (e.target.tagName === 'INPUT') return;
    switch(e.key) {
        case ' ': e.preventDefault(); togglePlay(); break;
        case 'ArrowLeft':
            e.preventDefault();
            stepTime(e.shiftKey ? -10 : -1);
            break;
        case 'ArrowRight':
            e.preventDefault();
            stepTime(e.shiftKey ? 10 : 1);
            break;
        case ',': e.preventDefault(); stepFrame(-1); break;
        case '.': e.preventDefault(); stepFrame(1); break;
        case 'a': case 'A': selectAll(); break;
        case 'f': case 'F': toggleFlip(); break;
        case 'i': case 'I': setInPoint(); break;
        case 'o': case 'O': setOutPoint(); break;
        case 'Enter': e.preventDefault(); addToCutList(); break;
    }
});

// --- ユーティリティ ---
function formatTime(sec) {
    if (sec < 0) sec = 0;
    const h = Math.floor(sec / 3600);
    const m = Math.floor((sec % 3600) / 60);
    const s = Math.floor(sec % 60);
    const f = Math.floor((sec % 1) * 100);
    return String(h).padStart(2,'0') + ':' +
           String(m).padStart(2,'0') + ':' +
           String(s).padStart(2,'0') + '.' +
           String(f).padStart(2,'0');
}

function setStatus(msg) {
    document.getElementById('statusbar').textContent = msg;
}
</script>
</body>
</html>'''


def find_ffmpeg():
    """ffmpeg / ffprobe のパスを探す。見つからなければ None を返す。"""
    search_paths = [
        "/opt/homebrew/bin",
        "/usr/local/bin",
        "/usr/bin",
        os.path.expanduser("~/bin"),
    ]
    # バンドル内の ffmpeg（将来用）
    if getattr(sys, "_MEIPASS", None):
        search_paths.insert(0, sys._MEIPASS)

    for d in search_paths:
        ff = os.path.join(d, "ffmpeg")
        fp = os.path.join(d, "ffprobe")
        if os.path.isfile(ff) and os.path.isfile(fp):
            return d
    # PATH にあるか
    if shutil.which("ffmpeg") and shutil.which("ffprobe"):
        return ""
    return None


def main():
    # ffmpeg チェック
    ffmpeg_dir = find_ffmpeg()
    if ffmpeg_dir is None:
        # macOS ダイアログで通知
        msg = ("VideoCut の実行には ffmpeg が必要です。\\n\\n"
               "Homebrew でインストールしてください:\\n"
               "  brew install ffmpeg")
        subprocess.run([
            "osascript", "-e",
            f'display dialog "{msg}" with title "VideoCut" '
            f'buttons {{"OK"}} default button "OK" with icon caution'
        ])
        sys.exit(1)

    # ffmpeg を PATH に追加（Homebrew 等のパスが通っていない場合の対策）
    if ffmpeg_dir:
        os.environ["PATH"] = ffmpeg_dir + ":" + os.environ.get("PATH", "")

    server = http.server.HTTPServer(("127.0.0.1", PORT), VideoCutHandler)
    actual_port = server.server_address[1]
    url = f"http://127.0.0.1:{actual_port}"
    print(f"VideoCut サーバー起動: {url}")
    print("ブラウザで上記URLを開いてください。終了するには Ctrl+C を押してください。")

    # ブラウザを自動で開く
    threading.Timer(0.5, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nサーバーを停止しました。")
        server.server_close()


if __name__ == "__main__":
    main()

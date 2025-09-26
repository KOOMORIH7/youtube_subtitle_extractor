import streamlit as st
import yt_dlp
import re
import os
from openpyxl import Workbook
from docx import Document
import tempfile

# 字幕抽出関数
def extract_subtitles(subtitle_file, keywords, use_censored):
    matches = []
    with open(subtitle_file, "r", encoding="utf-8") as f:
        blocks = f.read().split("\n\n")
    for block in blocks:
        lines = block.strip().split("\n")
        if len(lines) >= 3:
            timestamp = lines[1]
            text = " ".join(lines[2:])
            if keywords:
                for kw in keywords:
                    if kw in text:
                        matches.append((timestamp, text))
                        break
            elif use_censored:
                if re.search(r"\[\s*__\s*\]", text):
                    matches.append((timestamp, text))
    return matches

# 安全なファイル名作成
def safe_filename(name):
    return re.sub(r'[\\/*?:"<>|]', "_", name)

# 保存処理
def save_matches(matches, title, folder, fmt):
    safe_title = safe_filename(title)
    filepath = os.path.join(folder, f"{safe_title}.{fmt.lower()}")
    
    if fmt == "TXT":
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"=== {title} ===\n")
            for t, txt in matches:
                f.write(f"[{t}] {txt}\n")
            f.write(f"\n合計 {len(matches)} 件\n")
    elif fmt == "Word":
        doc = Document()
        doc.add_heading(title, level=1)
        for t, txt in matches:
            doc.add_paragraph(f"[{t}] {txt}")
        doc.add_paragraph(f"\n合計 {len(matches)} 件")
        doc.save(filepath)
    elif fmt == "Excel":
        wb = Workbook()
        ws = wb.active
        ws.append(["Timestamp", "Text"])
        for t, txt in matches:
            ws.append([t, txt])
        wb.save(filepath)
    return filepath

# ---------------- Streamlit UI ----------------
st.title("YouTube 字幕抽出アプリ")

url = st.text_input("YouTube動画URL")
folder = st.text_input("保存先フォルダ (例: C:/Users/...)")
fmt = st.selectbox("保存形式", ["TXT", "Word", "Excel"])
keywords_input = st.text_input("抽出キーワード（カンマ区切り）")
use_censored = st.checkbox("[ __ ]を抽出")

if st.button("抽出開始"):
    if not url or not folder or (not keywords_input and not use_censored):
        st.warning("URL・保存先・キーワードまたはチェックボックスは必須です")
    else:
        keywords = [kw.strip() for kw in keywords_input.split(",") if kw.strip()] if keywords_input else []
        st.info("📥 字幕抽出開始...")
        try:
            # 動画情報取得
            with yt_dlp.YoutubeDL({}) as ydl_info:
                info = ydl_info.extract_info(url, download=False)
            video_title = info.get("title", "untitled")
            st.write(f"🎬 動画タイトル: {video_title}")
            
            # 字幕ダウンロード
            ydl_opts = {
                "skip_download": True,
                "writesubtitles": True,
                "writeautomaticsub": True,
                "subtitleslangs": ["en"],
                "subtitlesformat": "srt",
                "outtmpl": "subtitle"
            }
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                ydl.download([url])
            
            subtitle_files = [f for f in os.listdir() if f.startswith("subtitle") and f.endswith(".srt")]
            if not subtitle_files:
                st.warning("⚠️ 字幕が見つかりませんでした")
            else:
                subtitle_file = subtitle_files[0]
                matches = extract_subtitles(subtitle_file, keywords, use_censored)
                os.remove(subtitle_file)
                
                saved_path = save_matches(matches, video_title, folder, fmt)
                st.success(f"✅ 保存完了: {saved_path} ({len(matches)} 件)")
        except Exception as e:
            st.error(f"❌ エラー: {e}")


# YouTube Subtitle Extractor

YouTube動画から字幕をダウンロード・解析し、特定のキーワードや伏字パターン `[ __ ]` を抽出できるツールです。  
GUI（tkinter）を利用しており、プログラミング初心者でも簡単に操作できます。  
現在自動字幕生成の精度の問題があり英語字幕のみに対応
GUIのデザインや細かいコード修正はAIで出力しました。
---

## 主な機能
- YouTube字幕ファイル（.srt / .vtt）を解析
- キーワード入力による一致部分の抽出
- `[ __ ]` パターンの自動検出（チェックボックスで切り替え可能）
- 結果を Word（.docx）や Excel（.xlsx）に保存

---

## 使い方
1. このリポジトリをクローン
   ```bash
   git clone https://github.com/KOOMORIH7/youtube_subtitle_extractor.git
   cd youtube_subtitle_extractor



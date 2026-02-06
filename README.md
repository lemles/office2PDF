# Office2PDF v4.5 (Office Ultimate App)

![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)
![Platform: Windows](https://img.shields.io/badge/Platform-Windows-blue.svg)

**事務作業の「面倒くさい」を解決する、実務特化型のOffice一括PDF変換・加工ツールです。**

Excel、Word、PowerPoint、画像をまとめてPDFに変換するだけでなく、「シートごとの分割」「ページ番号の挿入」「柔軟なファイル名ルール」などの定型業務を数クリックで自動化します。

---

## ⚠️ 保守・メンテナンスについて

**【重要】作者はプロの開発者ではありません。**
本ツールは個人の業務効率化のために作成したもので、作者自身が複雑なバグ修正や高度なコードの保守を行うことは困難です。

*   **エンジニアの方へ:** コードの改善案、バグ修正、機能追加などのプルリクエスト（Pull Request）を心よりお待ちしております。
*   **ユーザーの方へ:** 本ソフトは「現状有姿」での提供となります。あらかじめご了承ください。

---

## 🌟 主な機能

- **Excel変換の圧倒的柔軟性**
  - ブック内の各シートを個別のPDFとして一括出力。
  - 「横幅を1ページに収める」などの印刷設定を変換時に強制適用。
- **インテリジェントな命名ルール**
  - `{date}_{name}_{sheet}` のように、タグを組み合わせて出力ファイル名を自動決定。
- **高度な透かし(Watermark)・ページ番号**
  - タグを利用して「誰がいつ作ったか」を自動で透かしとして挿入。
  - 変換前に1ページ目だけで見た目を確認できるプレビュー機能。
- **セキュリティと最適化**
  - 一括パスワード設定、メタデータ（作成者情報）の削除、PDFファイルの軽量化。

---

## 🏷️ 利用可能なタグ一覧
「命名ルール」や「透かし文字」の設定で、以下のタグが自動置換されます。

| タグ | 説明・置換例 |
| :--- | :--- |
| `{name}` | 元のファイル名（拡張子なし） |
| `{sheet}` | Excelのシート名（Excel以外は空文字） |
| `{parent}` | 親フォルダの名前 |
| `{seq}` | 全体の通し番号 |
| `{fseq}` | 入力ファイルごとの連番 |
| `{pseq}` | PDF内のページ番号 |
| `{total}` | 処理する全ファイル数 |
| `{ptotal}` | そのファイルの総ページ数 |
| `{username}` | PCのログインユーザー名 |
| `{rand}` | 4桁のランダムな数字 |
| `{date:yyyy-mm-dd}` | 日付（例：2024-02-06） |
| `{date:HH:MM:SS}` | 現在時刻（例：14:30:05） |

---

## 🚀 使い方

### 動作環境
- Windows 10 / 11 (64bit)
- **Microsoft Office (Excel/Word/PowerPoint) がインストールされていること**

### 導入方法
1. [Releases](https://github.com/あなたのユーザー名/Office2PDF-Ultimate/releases) から最新の `Office2PDF.zip` をダウンロードします。
2. 解凍して、中の `Office2PDF.exe` を実行してください。（インストール不要）

## 🛠️ ソースコードからの実行・ビルド方法

エンジニアの方や、Python環境で直接動かしたい方向けの手順です。

### 1. 実行環境
- Python 3.9 以上 (Windows)
- Microsoft Office インストール済み

### 2. 必要ライブラリのインストール
以下のコマンドで依存ライブラリをインストールしてください。

```bash
pip install pywin32 pypdf Pillow reportlab tkinterdnd2

---

## ☕ カンパのお願い
本ソフトがお役に立ちましたら、開発継続の支援としてカンパをいただけると大変励みになります。

**[Amazonギフトカード（Eメールタイプ）で支援する](https://www.amazon.co.jp/Amazon-eGift-Card-Flower-Bouquet-Animated/dp/B004N3APDM/)**

- **金額:** 15円から任意に設定いただけます。
- **受取人メールアドレス:** `the.desert.of.the.real@gmail.com`
  - (上記アドレスをクリックするとコピーされます ※ブラウザによります)

---

## ⚖️ ライセンス
[MIT License](./LICENSE)

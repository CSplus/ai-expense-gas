# AI経費処理システム

> **開発コード名**：AI経費処理システム

Google Apps Script と Gemini API を利用した、中小企業向けのAI経費処理支援システムです。

本システムは、領収書やクレジットカード利用明細をAIが解析し、経費データとして管理するとともに、各種会計ソフトへ連携することを目的としています。

---

# プロジェクトの目的

中小企業の経費処理業務を効率化し、入力作業・転記作業・確認作業をAIで支援することを目的としています。

目標は、

* 領収書入力の自動化
* 経費データの一元管理
* 重複登録防止
* クレジットカード利用明細との照合
* 会計ソフトへのデータ連携

です。

---

# 主な機能

現在実装済み、または開発中の機能です。

* 領収書アップロード
* GeminiによるOCR解析
* 日付・金額・店舗名の抽出
* 勘定科目の自動判定
* 支払方法の判定
* Googleスプレッドシートへの登録
* クレジットカード明細取込
* 会計ソフト連携（開発中）

---

# システム構成

```text
スマートフォン
        │
        ▼
Google Apps Script
        │
        ▼
Gemini API
        │
        ▼
Google Drive
        │
        ▼
Google Sheets（共通経費データ）
        │
        ▼
会計ソフト連携
    ├── TKC FX2
    ├── 弥生会計
    ├── freee（予定）
    ├── マネーフォワード（予定）
    └── 汎用CSV
```

---

# 開発環境

* Google Apps Script
* Google Workspace
* Gemini API
* Visual Studio Code
* Node.js
* clasp
* Git
* GitHub

---

# 初回セットアップ

```bash
git clone https://github.com/CSplus/ai-expense-gas.git

cd ai-expense-gas

npm install

clasp login

npm run pull
```

Apps Scriptへ反映

```bash
npm run push
```

---

# ディレクトリ構成

```text
docs/
images/

Main.js
Gemini.js
Rule.js
Config.js
Sheet.js

README.md
package.json
```

---

# ドキュメント

| ファイル              | 内容        |
| ----------------- | --------- |
| 00_プロジェクト概要.md    | プロジェクト概要  |
| 01_要件定義.md        | システム要件    |
| 02_システム構成.md      | システム構成    |
| 03_スプレッドシート設計.md  | データ設計     |
| 04_Geminiプロンプト.md | AI判定ルール   |
| 05_開発ルール.md       | 開発ルール     |
| 06_変更履歴.md        | 変更履歴      |
| 07_今後の開発計画.md     | ロードマップ    |
| 08_運用マニュアル.md     | 運用方法      |
| 09_AIエージェント仕様.md  | AI設計      |
| 10_会計ソフト連携方針.md   | 会計ソフト共通設計 |
| 11_TKC_FX2連携仕様.md | TKC FX2仕様 |
| 12_弥生会計連携仕様.md    | 弥生会計仕様    |

---

# 開発方針

本システムは特定の会計ソフトに依存しない設計を採用しています。

AIが生成する経費データは共通形式で保持し、会計ソフトごとの差異は専用の連携モジュール（Exporter）が吸収します。

この設計により、将来的な会計ソフト追加にも柔軟に対応できる構成を目指します。

---

# ライセンス

Copyright © CSplus

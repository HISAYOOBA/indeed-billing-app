# Indeed請求明細ジェネレーター

## セットアップ手順

### 1. Google Cloud設定（サービスアカウント作成）

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新しいプロジェクトを作成（例：`indeed-billing`）
3. 「APIとサービス」→「ライブラリ」→「Google Drive API」を有効化
4. 「APIとサービス」→「認証情報」→「サービスアカウントを作成」
5. サービスアカウントのキーをJSON形式でダウンロード

### 2. Google Driveフォルダ設定

以下の2つのフォルダをGoogle Driveに作成し、  
**サービスアカウントのメールアドレスを「閲覧者」として共有**してください。

```
📁 Indeed請求データ/
    ├── Indeed_2026年1月.xlsx
    ├── Indeed_2026年2月.xlsx
    └── Indeed_2026年3月.xlsx

📁 キャンペーンパフォーマンス/
    ├── 月間キャンペーンパフォーマンス_data.csv（3月分）
    └── 月間キャンペーンパフォーマンス.csv（1・2月分）
```

フォルダIDはGoogle DriveのURLの末尾から確認できます。  
例：`https://drive.google.com/drive/folders/【ここがフォルダID】`

### 3. GitHubリポジトリ作成

1. [GitHub](https://github.com) にログイン
2. 新しいリポジトリを作成（例：`indeed-billing-app`）
3. 以下のファイルをアップロード：
   - `app.py`
   - `requirements.txt`
   - `README.md`

### 4. Streamlit Cloudにデプロイ

1. [share.streamlit.io](https://share.streamlit.io) にGitHubアカウントでログイン
2. 「New app」→ GitHubリポジトリを選択 → `app.py` を指定
3. 「Advanced settings」→「Secrets」に以下を設定：

```toml
GOOGLE_SERVICE_ACCOUNT = """
（ダウンロードしたJSONファイルの内容をそのまま貼り付け）
"""
```

4. 「Deploy」ボタンを押す

### 5. 使い方

1. アプリURLをブラウザで開く
2. サイドバーにGoogle DriveのフォルダIDを入力
3. 対象月・クライアント名を入力
4. 「請求明細Excelを生成」ボタンを押す
5. Excelをダウンロード

---

## ファイル構成

```
├── app.py              # メインアプリ
├── requirements.txt    # 必要なライブラリ
└── README.md           # このファイル
```

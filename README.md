# PowerPoint Keyword Detection and Replacement Tool

PowerPoint ファイル内の特定のキーワードを検出、削除、置換するWebツール。

## 機能

- **キーワード検出**: PowerPoint ファイルをアップロードして、含まれるキーワードを検出
- **キーワード削除**: 検出されたキーワードを削除
- **キーワード置換**: キーワードを別の文字列に置換
- **ファイルダウンロード**: 修正済みファイルをダウンロード

## セットアップ

### 必要な環境
- Python 3.8以上
- pip

### インストール

```bash
pip install -r requirements.txt
```

### 実行

```bash
python app.py
```

ブラウザで `http://localhost:5000` にアクセスします。

## プロジェクト構造

```
.
├── app.py                 # Flask アプリケーション
├── requirements.txt       # Python 依存関係
├── templates/
│   └── index.html        # HTML テンプレート
├── static/
│   └── style.css         # スタイルシート
│   └── script.js         # JavaScriptスクリプト
└── uploads/              # アップロードファイル保存ディレクトリ
```

## 使用方法

1. PowerPoint ファイルをアップロード
2. 検出または置換するキーワードを入力
3. 「検出」「削除」「置換」から操作を選択
4. 修正済みファイルをダウンロード


## ���

Created by Y.Atarashi with Github Copilot

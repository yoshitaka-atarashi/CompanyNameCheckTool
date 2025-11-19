# インストールガイド

このドキュメントでは、PowerPoint Keyword Detection Tool をインストールして実行する方法を説明します。

## 前提条件

- **Python**: 3.8 以上
- **pip**: Python パッケージマネージャー
- **git**: バージョン管理（オプション）

## インストール手順

### 1. リポジトリのクローン

```bash
git clone https://github.com/yourusername/CompanyNameCheckTool.git
cd CompanyNameCheckTool
```

### 2. 仮想環境の作成（推奨）

#### Windows

```bash
python -m venv .venv
.venv\Scripts\activate
```

#### macOS/Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
```

### 3. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 4. アプリケーションの実行

```bash
python app.py
```

出力例：
```
 * Running on http://127.0.0.1:5000
 * Debug mode: on
```

### 5. ブラウザでアクセス

ブラウザを開いて以下のアドレスにアクセスします：

```
http://localhost:5000
```

## トラブルシューティング

### `ModuleNotFoundError: No module named 'flask'`

`requirements.txt` が正しくインストールされていません。以下を実行してください：

```bash
pip install -r requirements.txt
```

### ポート 5000 が既に使用されている

別のアプリケーションがポート 5000 を使用しています。以下を実行してください：

```bash
python app.py --port 5001
```

### PowerPoint ファイルがアップロードできない

- ファイルサイズが 50MB 以下であることを確認してください
- ファイル形式が `.pptx` または `.ppt` であることを確認してください

## 開発環境のセットアップ

### 追加の開発ツールをインストール

```bash
pip install pytest pytest-cov flake8
```

### テストの実行

```bash
pytest
```

### コード品質チェック

```bash
flake8 app.py diagnose_pptx.py
```

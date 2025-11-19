# 開発者向けドキュメント

このドキュメントは、プロジェクトの開発に参加したい方向けです。

## プロジェクト構造

```
CompanyNameCheckTool/
├── app.py                    # Flask メインアプリケーション
├── diagnose_pptx.py          # PowerPoint ファイル診断ツール
├── requirements.txt          # Python 依存関係
├── static/                   # 静的ファイル
│   ├── style.css            # スタイルシート
│   └── script.js            # クライアント側 JavaScript
├── templates/               # HTML テンプレート
│   └── index.html           # メインページ
├── TestData/                # テストデータ
├── uploads/                 # アップロード済みファイル（一時保存）
├── .gitignore              # Git 除外ファイル
├── LICENSE                 # ライセンス
├── README.md               # プロジェクト概要
├── CHANGELOG.md            # 変更ログ
├── INSTALLATION.md         # インストールガイド
├── USAGE.md                # 使用ガイド
└── DEVELOPMENT.md          # このファイル
```

## 開発環境のセットアップ

### 1. リポジトリのクローン

```bash
git clone https://github.com/yourusername/CompanyNameCheckTool.git
cd CompanyNameCheckTool
```

### 2. 仮想環境の作成

```bash
python -m venv .venv
# Windows
.venv\Scripts\activate
# macOS/Linux
source .venv/bin/activate
```

### 3. 依存関係のインストール

```bash
pip install -r requirements.txt
```

### 4. 開発用追加パッケージのインストール

```bash
pip install pytest pytest-cov flake8 black
```

## コーディング規約

### スタイルガイド

- **Python**: PEP 8 に従ってください
- **インデント**: 4 スペース
- **行の長さ**: 最大 100 文字
- **命名規則**:
  - 関数・変数: `snake_case`
  - クラス: `PascalCase`
  - 定数: `UPPER_CASE`

### コードフォーマット

Black を使用してコードをフォーマットしてください：

```bash
black app.py diagnose_pptx.py
```

### リント

Flake8 を使用してコードチェックを行ってください：

```bash
flake8 app.py diagnose_pptx.py
```

## テスト

### テストの実行

```bash
pytest
```

### カバレッジレポート

```bash
pytest --cov=. --cov-report=html
```

## API エンドポイント

### GET `/`

- メインページを返します

### POST `/api/detect`

PowerPoint ファイル内のキーワードを検出します。

**リクエスト:**
```json
{
  "file": <FormData>,
  "keywords": ["keyword1", "keyword2"]
}
```

**レスポンス:**
```json
{
  "success": true,
  "results": [
    {
      "slide": 1,
      "shape": 0,
      "text": "Sample text",
      "keywords": ["keyword1"],
      "count": 1
    }
  ]
}
```

### POST `/api/delete`

PowerPoint ファイルからキーワードを削除します。

**リクエスト:**
```json
{
  "file": <FormData>,
  "keywords": ["keyword1", "keyword2"]
}
```

**レスポンス:**
```
<Binary PowerPoint file>
```

### POST `/api/replace`

PowerPoint ファイルのキーワードを置換します。

**リクエスト:**
```json
{
  "file": <FormData>,
  "keywords": ["keyword1"],
  "replacement": "new_value"
}
```

**レスポンス:**
```
<Binary PowerPoint file>
```

## 主要な関数

### `find_keywords_in_presentation(prs, keywords)`

プレゼンテーション内のキーワードを検出します。

**パラメータ:**
- `prs`: Presentation オブジェクト
- `keywords`: 検出するキーワードのリスト

**戻り値:**
- 検出結果のリスト

### `delete_keywords_in_presentation(prs, keywords)`

プレゼンテーション内のキーワードを削除します。

**パラメータ:**
- `prs`: Presentation オブジェクト
- `keywords`: 削除するキーワードのリスト

**戻り値:**
- 修正済み Presentation オブジェクト

### `replace_keywords_in_presentation(prs, keywords, replacement)`

プレゼンテーション内のキーワードを置換します。

**パラメータ:**
- `prs`: Presentation オブジェクト
- `keywords`: 置換するキーワードのリスト
- `replacement`: 置換先の文字列

**戻り値:**
- 修正済み Presentation オブジェクト

## コミットメッセージの規約

コミットメッセージは以下のフォーマットに従ってください：

```
<type>: <subject>

<body>

<footer>
```

### type

- `feat`: 新しい機能
- `fix`: バグ修正
- `docs`: ドキュメント関連
- `style`: コード規約、フォーマット
- `refactor`: 機能変更なしのコード整理
- `test`: テスト関連
- `chore`: ビルドプロセス、依存関係更新

### subject

- 命令文で記述（例：「機能を追加する」ではなく「機能を追加」）
- 最初の文字は大文字
- ピリオドで終わらない
- 50 文字以内

### body

- `subject` とは 1 行開ける
- 74 文字で改行
- 何を変更したのかと、なぜ変更したのかを説明

### footer

- 参照する Issue 番号を記述（例：`Closes #123`）

## リリース手順

1. テストが全て成功することを確認
2. バージョン番号を更新
3. CHANGELOG.md を更新
4. コミットしてプッシュ
5. Git タグを作成してプッシュ
6. GitHub Releases でリリースを作成

## 問題報告とサポート

問題が発生した場合は、以下の順序で対応してください：

1. GitHub Issues で既に報告されていないか確認
2. DEVELOPMENT.md と INSTALLATION.md を確認
3. 新しい Issue を作成して詳細を記述

## 参考資料

- [Flask ドキュメント](https://flask.palletsprojects.com/)
- [python-pptx ドキュメント](https://python-pptx.readthedocs.io/)
- [PEP 8 スタイルガイド](https://pep8-ja.readthedocs.io/)

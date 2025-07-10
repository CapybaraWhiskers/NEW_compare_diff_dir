# ディレクトリ比較ツール

このリポジトリには、OfficeドキュメントやPDFのテキスト変換に対応したGitのdiff機能を利用して、2つのディレクトリを比較するPythonツールが含まれています。

## 必要要件

- Python 3
- Git
- Pythonパッケージ: `docx2txt`, `openpyxl`, `xlrd`, `python-docx`, `python-pptx`, `PyPDF2`

依存パッケージのインストール:

```bash
pip install python-docx openpyxl xlrd python-pptx PyPDF2 docx2txt
```

## 使い方

```
python compare_dirs.py DIR1 DIR2 [--json]
```

ツールは、変更内容をステータスごとにグループ化して出力します:

1. 追加 (Added)
2. 削除 (Removed)
3. 名前変更 (Renamed)
4. 修正 (Modified)
5. 名前変更および修正 (RenamedAndModified)
6. 変更なし (Unchanged)

変更されたファイルについては、unified diff形式の差分が表示されます。`--json`オプションを付けるとJSON形式で出力されます。

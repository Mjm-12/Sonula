# Fuzz Face Analysis

Fuzz Face回路解析プロジェクト。schemdrawを使用した回路図作成が含まれます。

## 環境セットアップ

このプロジェクトはPythonパッケージ管理に[uv](https://docs.astral.sh/uv/)を使用しています。

### 1. uvのインストール

#### Windows (PowerShell)
```powershell
irm https://astral.sh/uv/install.ps1 | iex
```

#### macOS / Linux
```bash
curl -LsSf https://astral.sh/uv/install.sh | sh
```

### 2. Pythonと仮想環境のセットアップ

```bash
# プロジェクトディレクトリに移動
cd 2026-01-02__Fuzz-Face-Analysis

# Python 3.12をインストール（まだインストールされていない場合）
uv python install 3.12

# 仮想環境を作成して依存関係をインストール
uv sync
```

このコマンドにより、`.venv`ディレクトリに仮想環境が作成され、`pyproject.toml`に定義された依存関係が自動的にインストールされます。

### 3. スクリプトの実行

```bash
# uvを使って直接スクリプトを実行
uv run python Python/FuzzFace_fontcustom.py
```

または仮想環境をアクティベートして実行:

#### Windows (PowerShell)
```powershell
.\.venv\Scripts\Activate.ps1
python Python/FuzzFace_fontcustom.py
```

#### macOS / Linux
```bash
source .venv/bin/activate
python Python/FuzzFace_fontcustom.py
```

## 開発用依存関係

開発ツール（pytest, ruff）を含めてインストール:

```bash
uv sync --all-extras
```

## パッケージの追加

新しいパッケージを追加する場合:

```bash
# 通常の依存関係として追加
uv add パッケージ名

# 開発用依存関係として追加
uv add --dev パッケージ名
```

## プロジェクト構成

```
2026-01-02__Fuzz-Face-Analysis/
├── .gitignore           # Git除外設定
├── .python-version      # Pythonバージョン指定
├── .venv/               # 仮想環境（自動生成）
├── pyproject.toml       # プロジェクト設定・依存関係
├── uv.lock              # 依存関係ロックファイル（自動生成）
├── README.md            # このファイル
└── Python/
    ├── Fonts/           # フォントファイル
    └── FuzzFace_fontcustom.py  # 回路図生成スクリプト
```

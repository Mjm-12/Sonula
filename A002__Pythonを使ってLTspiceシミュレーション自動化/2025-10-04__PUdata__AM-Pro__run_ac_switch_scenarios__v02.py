# -*- coding: utf-8 -*-
"""LTspiceのACスイープを自動実行し、結果をCSVファイルにエクスポートする。

このスクリプトは、複数のピックアップスイッチの組み合わせに対してLTspiceシミュレーションを
自動実行し、周波数特性データを収集します。
"""
from __future__ import annotations

import subprocess
from datetime import datetime
from pathlib import Path
import re
from shutil import copy2
from typing import Callable, Optional

import numpy as np
import pandas as pd
from PyLTSpice import RawRead, SimCommander
from spicelib.editor.asc_editor import AscEditor
from spicelib.editor.spice_editor import SpiceEditor

BASE_DIR = Path(__file__).parent.resolve()  # スクリプトのあるディレクトリ

# ========================================================================
# ユーザー設定
# ========================================================================

# --- 基本設定 ---
PU_Name = "AM-Pro"  # ピックアップの名前（出力フォルダ名・ファイル名に使用される）
INPUT_PATH = BASE_DIR / "asc" / "AM-Pro_Analysis.asc"  # LTspiceの回路ファイルのパス
LTSPICE_EXE = r"C:\Users\[your-username]\AppData\Local\Programs\ADI\LTspice\LTspice.exe"  # LTspice実行ファイルのパス
                # ↑実行環境に合わせてファイルパスを設定してください
TARGET_NODE = "Amp-In"  # 測定対象のノード名（回路図上のネット名）
ANALYSIS_TEMPLATE = BASE_DIR / "Template" / "Analysis_Template.xlsm"  # グラフ描画用のExcelテンプレートファイルのパス

# --- スイッチ設定（電圧制御スイッチV2～V6の役割） ---
# V2: Neck PU（ネックピックアップ）
# V3: Middle PU（ミドルピックアップ）
# V4: Bridge PU（ブリッジピックアップ）
# V5: Tone1 + Volume-Lower（トーン1＋ボリューム下段 250k）
# V6: Tone2 + Volume-Upper（トーン2＋ボリューム上段 500k）
# 各スイッチは "5" で ON、"0" で OFF
# スイッチ設定の変更は下記の「スイッチの組み合わせ設定」セクション（406行目以降）で行う

# --- トーン/ボリューム設定 ---
# トーンノブの有効化/無効化とボリュームノブの有効化/無効化は104～133行目の tone_param_transform 関数内で設定されています
# ここを編集することで解析の順番や、掃引パラメータを変更できます
# 現在の設定:
#   - トーンノブ: 固定-->可変（.step param j を有効にする）
#   - ボリュームノブ: 可変-->固定（.step param k を無効にする）


# ========================================================================
# 内部設定（通常は変更不要）
# ========================================================================

PREFERRED_ENC = "cp932"  # 優先文字エンコーディング（日本語Windows環境）
FALLBACK_ENC = "utf-8"   # フォールバック文字エンコーディング

TextTransform = Callable[[str], str]  # テキスト変換関数の型定義


# ========================================================================
# ファイル入出力関数
# ========================================================================

def read_text_auto(path: Path) -> str:
    """ファイルを自動エンコーディング検出で読み込む。

    Args:
        path: 読み込むファイルのパス

    Returns:
        ファイルの内容（文字列）
    """
    try:
        return path.read_text(encoding=PREFERRED_ENC)
    except Exception:
        return path.read_text(encoding=FALLBACK_ENC, errors="ignore")


def write_text_cp932(path: Path, text: str) -> None:
    """ファイルをcp932エンコーディングで書き込む。

    Args:
        path: 書き込むファイルのパス
        text: 書き込む内容
    """
    path.write_text(text, encoding=PREFERRED_ENC, errors="replace")


def normalize_micro_symbols(text: str) -> str:
    """マイクロ記号（μ）を 'u' に正規化する。

    LTspiceとの互換性のため、各種マイクロ記号を通常の 'u' に変換する。

    Args:
        text: 正規化対象のテキスト

    Returns:
        正規化後のテキスト
    """
    return text.replace("\u00B5", "u").replace("\u03BC", "u")


def tone_param_transform(text: str) -> str:
    """トーンスイープ用にパラメータ定義を入れ替える。

    【重要】この関数でトーンノブ/ボリュームノブの有効/無効を切り替えます。

    現在の設定:
    - トーンノブ: 有効化（.step param j を有効にする）
    - ボリュームノブ: 無効化（.step param k を無効にする）

    変更したい場合は、以下の置換処理の「;」と「.」を入れ替えてください。
    例: トーンノブを無効化したい場合
        ".step param j 0 0.999 0.111" → ";step param j 0 0.999 0.111"

    Args:
        text: 変換対象のSPICEネットリストテキスト

    Returns:
        変換後のテキスト
    """
    # トーンノブの有効化（;を削除して.stepコマンドを有効にする）
    text = text.replace(";step param j 0 0.999 0.111", ".step param j 0 0.999 0.111")
    text = text.replace(";param Pt=(x**j-1)/(x-1)", ".param Pt=(x**j-1)/(x-1)")
    text = text.replace(".param Pt=1", ";param Pt=1")

    # ボリュームノブの無効化（.を;に変更して.stepコマンドを無効にする）
    text = text.replace(".step param k 0.111 0.999 0.111", ";step param k 0.111 0.999 0.111")
    text = text.replace(".param Px=(x**k-1)/(x-1)", ";param Px=(x**k-1)/(x-1)")
    text = text.replace(";param Px=1", ".param Px=1")

    return text


# ========================================================================
# 出力ディレクトリとファイル形式の検出
# ========================================================================

def make_outdir(base: Path) -> Path:
    """タイムスタンプ付きの出力ディレクトリを作成する。

    Args:
        base: ベースディレクトリのパス

    Returns:
        作成された出力ディレクトリのパス
    """
    stamp = datetime.now().strftime(f"%y-%m-%d__{PU_Name}__%H-%M-%S")
    outdir = base / stamp
    outdir.mkdir(parents=True, exist_ok=True)
    return outdir


def detect_format(path: Path) -> str:
    """ファイル形式を自動検出する。

    Args:
        path: 検出対象のファイルパス

    Returns:
        ファイル形式（"asc", "spice", "expresspcb", "unknown"）
    """
    ext = path.suffix.lower()
    if ext == ".asc":
        return "asc"
    text = read_text_auto(path)
    stripped = text.lstrip()
    if stripped.startswith("*") or stripped.startswith(".title") or stripped.startswith(".include"):
        return "spice"
    if "ExpressPCB Netlist" in text:
        return "expresspcb"
    return "unknown"


# ========================================================================
# エディタの保存処理（spicelib互換性対応）
# ========================================================================

def asc_save_compat(editor: AscEditor, path: Path) -> None:
    """AscEditorの複数のバージョンに対応した保存処理。

    spicelibのバージョンによってAPIが異なるため、利用可能なメソッドを検出して保存する。

    Args:
        editor: AscEditorオブジェクト
        path: 保存先のパス

    Raises:
        AttributeError: 対応する保存APIが見つからない場合
    """
    if hasattr(editor, "save") and callable(getattr(editor, "save")):
        editor.save()
        return
    if hasattr(editor, "save_netlist") and callable(getattr(editor, "save_netlist")):
        editor.save_netlist(str(path))
        return
    if hasattr(editor, "write_netlist") and callable(getattr(editor, "write_netlist")):
        editor.write_netlist(str(path))
        return
    if hasattr(editor, "save_as") and callable(getattr(editor, "save_as")):
        editor.save_as(str(path))
        return
    raise AttributeError("AscEditor does not expose a supported save API")


# ========================================================================
# エディタの準備とファイル変換
# ========================================================================

def prepare_editor(
    input_path: Path,
    work_dir: Path,
    text_transform: Optional[Callable[[str], str]] = None,
) -> tuple[object, Path, str, str]:
    """入力ファイルからエディタを準備する。

    ファイル形式を検出し、適切なエディタ（AscEditorまたはSpiceEditor）を作成する。
    必要に応じてテキスト変換（トーン/ボリューム切り替え等）を適用する。

    Args:
        input_path: 入力ファイル（.ascまたは.cirファイル）のパス
        work_dir: 作業ディレクトリのパス
        text_transform: テキスト変換関数（Noneの場合は変換なし）

    Returns:
        (editor, edited_file, kind, restore_text) のタプル
        - editor: AscEditorまたはSpiceEditorオブジェクト
        - edited_file: 編集されたファイルのパス
        - kind: ファイル形式（"asc" または "spice"）
        - restore_text: 元に戻すためのテキスト

    Raises:
        RuntimeError: サポートされていないファイル形式の場合
    """
    kind = detect_format(input_path)
    if kind == "expresspcb":
        raise RuntimeError(
            "ExpressPCB netlists are not supported. Export a SPICE netlist or use the .asc schematic."
        )
    if kind == "unknown":
        raise RuntimeError(
            "Could not detect the input file format. Please provide an .asc or SPICE netlist."
        )

    work_dir.mkdir(parents=True, exist_ok=True)

    # マイクロ記号の正規化
    normalized = normalize_micro_symbols(read_text_auto(input_path))
    # テキスト変換（トーン/ボリューム切り替え等）を適用
    transformed = text_transform(normalized) if text_transform else normalized

    if kind == "asc":
        # .ascファイルの場合
        edited = work_dir / input_path.name
        write_text_cp932(edited, transformed)
        editor = AscEditor(str(edited))
        restore_text = normalized
        return editor, edited, kind, restore_text

    # SPICEネットリストファイルの場合
    edited = work_dir / (input_path.stem + ".cir")
    restore_text = normalized
    transformed_output = transformed
    if not normalized.lstrip().startswith("*"):
        # SpiceEditor用にヘッダーを追加
        header = "* converted for SpiceEditor\r\n"
        restore_text = header + restore_text
        transformed_output = header + transformed_output
    write_text_cp932(edited, transformed_output)
    editor = SpiceEditor(str(edited))
    return editor, edited, kind, restore_text


# ========================================================================
# .wrdataディレクティブの削除（データエクスポート設定のクリーンアップ）
# ========================================================================

def remove_existing_directives(editor: object) -> None:
    """既存の.wrdataディレクティブを削除する。

    データエクスポートの競合を防ぐため、既存の.wrdataディレクティブを削除する。

    Args:
        editor: AscEditorまたはSpiceEditorオブジェクト
    """
    if hasattr(editor, "remove_Xinstruction"):
        editor.remove_Xinstruction(r"\.wrdata")
        return
    if hasattr(editor, "remove_instruction"):
        try:
            editor.remove_instruction(".wrdata")
        except Exception:
            pass


def remove_existing_wrdata_quietly(editor: object, file_path: Path) -> None:
    """既存の.wrdataディレクティブを静かに削除する。

    ファイルに.wrdataディレクティブが存在する場合のみ削除を試みる。
    spicelibの「not found」通知を回避するため、事前にファイルをチェックする。

    Args:
        editor: AscEditorまたはSpiceEditorオブジェクト
        file_path: チェック対象のファイルパス
    """
    try:
        text = read_text_auto(file_path)
    except Exception:
        text = ""

    # .wrdataディレクティブが存在しない場合はスキップ（大文字小文字を区別しない）
    if ".wrdata" not in text.lower():
        return

    # 利用可能なAPIで削除を試みる（非正規表現APIを優先）
    if hasattr(editor, "remove_instruction"):
        try:
            editor.remove_instruction(".wrdata")
            return
        except Exception:
            pass
    if hasattr(editor, "remove_Xinstruction"):
        try:
            editor.remove_Xinstruction(r"\.wrdata\b")
        except Exception:
            pass


def write_editor(editor: object, path: Path, kind: str) -> None:
    """エディタの内容をファイルに書き込む。

    Args:
        editor: AscEditorまたはSpiceEditorオブジェクト
        path: 保存先のパス
        kind: ファイル形式（"asc" または "spice"）
    """
    if kind == "asc":
        asc_save_compat(editor, path)
        refreshed = normalize_micro_symbols(read_text_auto(path))
        write_text_cp932(path, refreshed)
    else:
        if hasattr(editor, "write_netlist"):
            editor.write_netlist(str(path))
        else:
            asc_save_compat(editor, path)
        refreshed = normalize_micro_symbols(read_text_auto(path))
        write_text_cp932(path, refreshed)


# ========================================================================
# LTspiceのバッチ実行
# ========================================================================

def run_ltspice_batch(executable: str, input_file: Path) -> None:
    """LTspiceをバッチモードで実行する。

    複数のコマンドライン引数の組み合わせを試行して、最も互換性のある方法でLTspiceを起動する。

    Args:
        executable: LTspice実行ファイルのパス
        input_file: シミュレーション対象の回路ファイル

    Raises:
        FileNotFoundError: LTspice実行ファイルが見つからない場合
        RuntimeError: 全ての実行方法が失敗した場合
    """
    exe_path = Path(executable)
    if not exe_path.exists():
        raise FileNotFoundError(f"LTspice executable not found: {executable}")

    # 複数のコマンドライン引数の組み合わせを試行
    candidates = [
        [executable, "-b", str(input_file)],
        [executable, "-Run", "-b", str(input_file)],
        [executable, "-b", "-Run", str(input_file)],
    ]

    last_error = None
    for cmd in candidates:
        try:
            result = subprocess.run(
                cmd,
                cwd=str(input_file.parent),
                capture_output=True,
                text=True,
                check=False,
            )
            if result.returncode == 0:
                return
            last_error = (
                f"returncode={result.returncode}"
                f"\nstdout:\n{result.stdout}\nstderr:\n{result.stderr}"
            )
        except Exception as exc:
            last_error = str(exc)
    raise RuntimeError(
        "LTspice batch execution failed.\n"
        f"Tried: {candidates}\nLastError: {last_error}"
    )


def run_simulation(kind: str, editor_path: Path) -> None:
    """シミュレーションを実行する。

    ファイル形式に応じて、適切な方法でLTspiceシミュレーションを実行する。

    Args:
        kind: ファイル形式（"asc" または "spice"）
        editor_path: シミュレーション対象のファイルパス
    """
    if kind == "asc":
        # .ascファイルの場合は直接LTspiceをバッチ実行
        run_ltspice_batch(LTSPICE_EXE, editor_path)
        return

    # SPICEネットリストの場合はSimCommanderを使用
    sim = SimCommander(str(editor_path))
    try:
        if LTSPICE_EXE:
            sim.run(executable=LTSPICE_EXE)
        else:
            sim.run()
    except TypeError:
        # 古いバージョンのAPIに対応
        if LTSPICE_EXE:
            sim.run(ltspice_path=LTSPICE_EXE)
        else:
            sim.run()


# ========================================================================
# RAWファイルからのデータ抽出
# ========================================================================

def data_from_raw(raw_path: Path) -> pd.DataFrame:
    """RAWファイルから周波数特性データを抽出する。

    LTspiceのRAWファイルを読み込み、指定されたノードの周波数特性データ
    （周波数、ゲイン、位相）をDataFrameとして返す。

    Args:
        raw_path: RAWファイルのパス

    Returns:
        周波数特性データを含むDataFrame
        - frequency_Hz: 周波数 [Hz]
        - mag_dB: ゲイン [dB]
        - phase_deg: 位相 [度]
        - step_index: ステップインデックス
        - step_*: 各ステップパラメータ

    Raises:
        RuntimeError: 指定されたノードのトレースが見つからない場合
    """
    raw = RawRead(str(raw_path), verbose=False)

    # ノード名の大文字小文字を無視してトレースを検索
    target_lower = f"v({TARGET_NODE.lower()})"
    trace_name = next(
        (name for name in raw.get_trace_names() if name.lower() == target_lower),
        None,
    )
    if trace_name is None:
        available = ", ".join(raw.get_trace_names())
        raise RuntimeError(
            f"Trace for node '{TARGET_NODE}' not found in RAW file. Available traces: {available}"
        )

    trace = raw.get_trace(trace_name)
    plot = raw._plots[0] if raw._plots else None
    steps_info = plot.steps if plot and getattr(plot, "steps", None) else None

    frames: list[pd.DataFrame] = []
    steps = list(raw.get_steps())
    if not steps:
        steps = [0]

    # 各ステップのデータを抽出
    for step_idx in steps:
        freq = np.asarray(raw.get_axis(step_idx)).real.astype(float)
        wave = np.asarray(trace.get_wave(step_idx))
        mag = np.abs(wave)
        mag_db = 20.0 * np.log10(np.where(mag > 0.0, mag, np.finfo(float).tiny))
        phase = np.degrees(np.angle(wave))

        df = pd.DataFrame(
            {
                "frequency_Hz": freq,
                "mag_dB": mag_db,
                "phase_deg": phase,
                "step_index": step_idx,
            }
        )

        # ステップパラメータ情報を追加
        if steps_info and step_idx < len(steps_info):
            for key, value in steps_info[step_idx].items():
                df[f"step_{key}"] = value

        frames.append(df)

    # 全てのステップのデータを結合
    result = pd.concat(frames, ignore_index=True)
    ordered = ["frequency_Hz", "mag_dB", "phase_deg"]
    extras = [col for col in result.columns if col not in ordered]
    return result[ordered + extras]


# ========================================================================
# スイッチの組み合わせ設定とシミュレーション実行
# ========================================================================
# 【重要】電圧制御スイッチの組み合わせを変更するときは以下を書き換える（A1～3, B）

def run_case(
    input_path: Path,
    out_csv: Path,
    # A1. 制御したいスイッチ数に応じて引数を調整
    v2: str,  # Neck PU（ネックピックアップ）
    v3: str,  # Middle PU（ミドルピックアップ）
    v4: str,  # Bridge PU（ブリッジピックアップ）
    v5: str,  # Tone1 + Volume-Lower（トーン1＋ボリューム下段 250k）
    v6: str,  # Tone2 + Volume-Upper（トーン2＋ボリューム上段 500k）
    text_transform: Optional[TextTransform] = None,
) -> None:
    """1つのスイッチ組み合わせでシミュレーションを実行し、結果をCSVに保存する。

    Args:
        input_path: 入力ファイル（.ascまたは.cirファイル）のパス
        out_csv: 出力CSVファイルのパス
        v2: V2スイッチの設定値（"5"=ON, "0"=OFF）
        v3: V3スイッチの設定値（"5"=ON, "0"=OFF）
        v4: V4スイッチの設定値（"5"=ON, "0"=OFF）
        v5: V5スイッチの設定値（"5"=ON, "0"=OFF）
        v6: V6スイッチの設定値（"5"=ON, "0"=OFF）
        text_transform: テキスト変換関数（トーン/ボリューム切り替え等）
    """
    work_dir = out_csv.parent
    editor, edited_file, kind, restore_text = prepare_editor(
        input_path, work_dir, text_transform=text_transform
    )

    try:
        # A2. 制御したいスイッチ数に応じてset_component_value呼び出しを調整
        editor.set_component_value("V2", v2)
        editor.set_component_value("V3", v3)
        editor.set_component_value("V4", v4)
        editor.set_component_value("V5", v5)
        editor.set_component_value("V6", v6)

        # 既存の.wrdataディレクティブを削除
        remove_existing_wrdata_quietly(editor, edited_file)

        # エディタの内容を保存
        write_editor(editor, edited_file, kind)

        # シミュレーション実行
        run_simulation(kind, edited_file)

        # RAWファイルの確認
        raw_path = edited_file.with_suffix(".raw")
        if not raw_path.exists():
            log_path = edited_file.with_suffix(".log")
            if log_path.exists():
                tail = "\n".join(log_path.read_text(errors="ignore").splitlines()[-120:])
                raise RuntimeError(
                    "Simulation finished without producing a RAW file.\n"
                    f"Log: {log_path}\n---- log tail ----\n{tail}\n------------------",
                )
            raise FileNotFoundError(f"Expected RAW file not found: {raw_path}")

        # データを抽出してCSVに保存
        df = data_from_raw(raw_path)
        df.to_csv(out_csv, index=False, encoding=PREFERRED_ENC)
    finally:
        # ファイルを元の状態に戻す
        write_text_cp932(edited_file, restore_text)


def main() -> None:
    """メイン処理：全てのスイッチ組み合わせでシミュレーションを実行する。"""
    base_dir = Path(__file__).parent.resolve()
    if not INPUT_PATH.exists():
        raise FileNotFoundError(f"Input file not found: {INPUT_PATH}")

    if not ANALYSIS_TEMPLATE.exists():
        raise FileNotFoundError(f"Template workbook not found: {ANALYSIS_TEMPLATE}")

    # 出力ディレクトリを作成
    outdir = make_outdir(base_dir)
    dest_template = outdir / f"{PU_Name}_Analysis{ANALYSIS_TEMPLATE.suffix}"
    copy2(ANALYSIS_TEMPLATE, dest_template)

    # ========================================================================
    # B. スイッチの組み合わせ設定（ここを変更してスイッチパターンをカスタマイズ）
    # ========================================================================
    # 各組み合わせは (名前, スイッチ設定辞書) のタプル
    # スイッチ設定辞書: {"V2": "5"=ON/"0"=OFF, "V3": ..., "V4": ..., "V5": ..., "V6": ...}

    cases = [
        # Neck単独
        ("Neck",
            {"V2": "5", "V3": "0", "V4": "0", "V5": "0", "V6": "0"}),

        # Middle単独
        ("Middle",
            {"V2": "0", "V3": "5", "V4": "0", "V5": "0", "V6": "0"}),

        # Bridge単独
        ("Bridge",
            {"V2": "0", "V3": "0", "V4": "5", "V5": "0", "V6": "0"}),

        # Neck + Middle
        ("Neck-Middle",
            {"V2": "5", "V3": "5", "V4": "0", "V5": "5", "V6": "0"}),

        # Middle + Bridge
        ("Middle-Bridge",
            {"V2": "0", "V3": "5", "V4": "5", "V5": "0", "V6": "0"}),

        # Bridge + Neck
        ("Bridge-Neck",
            {"V2": "5", "V3": "0", "V4": "5", "V5": "0", "V6": "0"}),

        # Hum（ハムバッキング）
        ("Hum",
            {"V2": "0", "V3": "0", "V4": "5", "V5": "5", "V6": "5"}),

        # Neck + Hum
        ("Neck-Hum",
            {"V2": "5", "V3": "0", "V4": "5", "V5": "0", "V6": "5"}),

        # Middle + Hum
        ("Middle-Hum",
            {"V2": "0", "V3": "5", "V4": "5", "V5": "0", "V6": "5"}),
    ]

    # ボリューム/トーン切り替えバリエーション
    # ("Vol", None): ボリュームスイープのみ（トーンは固定）
    # ("Tone", tone_param_transform): トーンスイープのみ（ボリュームは固定）
    variants: list[tuple[str, Optional[TextTransform]]] = [
        ("Vol", None),
        ("Tone", tone_param_transform),
    ]

    # 全ての組み合わせでシミュレーションを実行
    for suffix, transform in variants:
        for name, values in cases:
            out_csv = outdir / f"{PU_Name}__{name}_{suffix}.csv"
            run_case(
                INPUT_PATH,
                out_csv,
                values["V2"],  # A3. 制御したいスイッチ数に応じて引数を調整
                values["V3"],
                values["V4"],
                values["V5"],
                values["V6"],
                text_transform=transform,
            )
            print(f"saved: {out_csv}")

    print("\nDone.")
    print(f"Output folder: {outdir}")


if __name__ == "__main__":
    main()

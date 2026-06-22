"""RF S11 Plot Tool — 通用化版本。

- 自动扫描 CSV 中所有 AntX(...) 天线
- pywebview 前端，支持多选天线
- 单图 + summary all 自适应布局
"""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import json
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass, field
from datetime import date, timedelta
from difflib import SequenceMatcher
from io import BytesIO
from pathlib import Path
from typing import Callable
from urllib.request import urlopen

import matplotlib

matplotlib.use("Agg")

import matplotlib.pyplot as plt
import numpy as np
from matplotlib.collections import LineCollection
from matplotlib import font_manager
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage


CSV_PATTERN = "*.csv"
DEFAULT_OUTPUT_SUBDIR = "output"
DEFAULT_SPC_ROOT = os.environ.get("RF_PLOT_TOOL_SPC_ROOT", "Z:\\").strip() or "Z:\\"
SPC_SNAPSHOT_SUBDIR = "RFPlotTool-spc-snapshots"
APP_VERSION = "0.0.9"
APP_DISPLAY_VERSION = "V0.0.9"
APP_UPDATE_EPOCH = 1
DEFAULT_UPDATE_MANIFEST_URL = "https://Andrew9896.github.io/RFPlotTool/version.json"
UPDATE_MANIFEST_URL = os.environ.get("RF_PLOT_TOOL_UPDATE_URL", DEFAULT_UPDATE_MANIFEST_URL).strip()
UPDATER_EXE_NAME = "updater.exe"
UPDATE_TIMEOUT_SECONDS = 20
UPDATE_PROGRESS_LOCK = threading.Lock()
UPDATE_PROGRESS: dict[str, object] = {
    "running": False,
    "stage": "idle",
    "percent": 0,
    "message": "",
    "error": "",
}
ANT_PATTERN = re.compile(
    r"^(S\d+):([^()]+?)\s*\(([^)]+)\)\s*Band:b(\d+)[\s:]+(?:Freq:\s*)?([0-9.]+)\s*M(?:Hz)?\s+Mag$"
)
ANT_KEY_PATTERN = re.compile(r"^(S\d+):(.+?)\(([^)]*)\)$")
RF_PREFIX_PATTERN = re.compile(r"^\s*(S\d+)\s*[:_\-\s]\s*(.+)$", re.IGNORECASE)
BAND_PATTERN = re.compile(r"(?:\bBand\s*[:=_\-\s]*b?\s*|\bB\s*)(\d+)\b", re.IGNORECASE)
FREQ_PATTERN = re.compile(
    r"\bFreq(?:uency)?\s*[:=_\-\s]*([0-9]+(?:\.[0-9]+)?)\s*([GMK]?Hz|[GMK])?\b",
    re.IGNORECASE,
)
COLON_FREQ_PATTERN = re.compile(
    r"[:_\-\s]+([0-9]+(?:\.[0-9]+)?)\s*([GMK]?Hz|[GMK])\b",
    re.IGNORECASE,
)
MAG_PATTERN = re.compile(r"\b(?:Mag(?:nitude)?|LogMag|dB|VSWR|Return\s*Loss|ReturnLoss|RL)\b", re.IGNORECASE)
HEX_COLOR_PATTERN = re.compile(r"^#[0-9a-fA-F]{6}$")
SERIAL_HEADER_KEYS = {"serialnumber", "serialno", "sn", "barcode", "unitserialnumber"}

SERIES_COLORS: list[tuple[str, str]] = [
    ("#16a34a", "#15803d"),
    ("#ef4444", "#dc2626"),
    ("#2563eb", "#1d4ed8"),
    ("#f97316", "#c2410c"),
    ("#9333ea", "#7e22ce"),
    ("#0891b2", "#0e7490"),
    ("#db2777", "#be185d"),
    ("#ca8a04", "#a16207"),
    ("#4b5563", "#374151"),
]

HIGHLIGHT_COLORS = [
    "#facc15",
    "#22d3ee",
    "#f472b6",
    "#fb923c",
    "#111827",
    "#a3e635",
]


@dataclass
class ChartLabels:
    overview_title: str = "S11 Comparison"
    group_labels: list[str] = field(default_factory=list)
    group_colors: list[str] = field(default_factory=list)
    upper_limit: str = "Upper Limit"
    lower_limit: str = "Lower Limit"


@dataclass
class ParsedData:
    """CSV 解析结果。

    antennas: 天线 key (Ant2(0:0:0)) -> [(列索引, band编号, 频率MHz)]
    segments: 由空行分隔的样本组，每个 segment 是若干行
    antenna_keys: 排好序的天线 key 列表
    """

    antennas: dict[str, list[tuple[int, int, float]]]
    upper_row: list[str]
    lower_row: list[str]
    segments: list[list[list[str]]]
    antenna_keys: list[str]


@dataclass
class CsvScanMetadata:
    antennas: dict[str, list[tuple[int, int, float]]]
    upper_row: list[str]
    lower_row: list[str]
    antenna_keys: list[str]
    segment_count: int
    samples: list[dict]


# ─────────────────────────── 文件辅助 ───────────────────────────


def app_base_dir() -> Path:
    """返回 exe/脚本 所在目录（不是 PyInstaller 临时解压目录）。

    - 打包后 (sys.frozen): exe 文件所在目录
    - 开发态: 当前脚本所在目录
    """
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def default_output_dir() -> Path:
    return app_base_dir() / DEFAULT_OUTPUT_SUBDIR


def find_default_csv() -> Path:
    files = sorted(app_base_dir().glob(CSV_PATTERN), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        raise FileNotFoundError("当前目录未找到任何 CSV 文件。")
    return files[0]


def resource_path(name: str) -> Path:
    """兼容 PyInstaller 打包后的资源访问（仅用于只读资源如 webui.html）。"""
    base = getattr(sys, "_MEIPASS", None)
    if base:
        return Path(base) / name
    return Path(__file__).resolve().parent / name


# ─────────────────────────── 字体配置 ───────────────────────────


# ─────────────────────────── 更新辅助 ───────────────────────────


def _version_numbers(version: str) -> list[int]:
    numbers = [int(part) for part in re.findall(r"\d+", str(version))]
    return numbers or [0]


def compare_versions(left: str, right: str) -> int:
    left_parts = _version_numbers(left)
    right_parts = _version_numbers(right)
    size = max(len(left_parts), len(right_parts))
    left_parts.extend([0] * (size - len(left_parts)))
    right_parts.extend([0] * (size - len(right_parts)))
    if left_parts < right_parts:
        return -1
    if left_parts > right_parts:
        return 1
    return 0


def normalize_update_manifest(
    manifest: dict,
    current_version: str,
    current_epoch: int = 0,
) -> dict:
    version = str(manifest.get("version", "")).strip()
    if not version:
        raise ValueError("更新清单缺少 version")
    app_version = str(manifest.get("app_version") or version).strip()
    display_version = str(manifest.get("display_version") or app_version).strip()
    try:
        epoch = int(manifest.get("epoch") or 0)
    except (TypeError, ValueError):
        raise ValueError("更新清单 epoch 必须是数字")

    has_update = epoch > current_epoch
    if epoch == current_epoch:
        has_update = compare_versions(app_version, current_version) > 0
    if not has_update:
        return {
            "has_update": False,
            "version": version,
            "app_version": app_version,
            "display_version": display_version,
            "epoch": epoch,
            "current_version": current_version,
            "notes": str(manifest.get("notes", "")).strip(),
        }

    url = str(manifest.get("url", "")).strip()
    mirrors = manifest.get("mirrors") or []
    if isinstance(mirrors, str):
        mirrors = [mirrors]
    download_urls: list[str] = []
    for candidate in [url, *mirrors]:
        candidate_url = str(candidate or "").strip()
        if candidate_url and candidate_url not in download_urls:
            download_urls.append(candidate_url)
    sha256 = str(manifest.get("sha256", "")).strip().lower()
    if not download_urls:
        raise ValueError("更新清单缺少 url")
    if not re.fullmatch(r"[0-9a-f]{64}", sha256):
        raise ValueError("更新清单 sha256 必须是 64 位十六进制字符串")

    return {
        "has_update": True,
        "version": version,
        "app_version": app_version,
        "display_version": display_version,
        "epoch": epoch,
        "current_version": current_version,
        "url": download_urls[0],
        "download_urls": download_urls,
        "sha256": sha256,
        "notes": str(manifest.get("notes", "")).strip(),
    }


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fetch_update_manifest(manifest_url: str) -> dict:
    with urlopen(manifest_url, timeout=UPDATE_TIMEOUT_SECONDS) as response:
        payload = response.read()
    return json.loads(payload.decode("utf-8-sig"))


def set_update_progress(stage: str, percent: int, message: str, error: str = "", running: bool = True) -> None:
    with UPDATE_PROGRESS_LOCK:
        UPDATE_PROGRESS.update({
            "running": running,
            "stage": stage,
            "percent": max(0, min(100, int(percent))),
            "message": message,
            "error": error,
        })


def get_update_progress_snapshot() -> dict:
    with UPDATE_PROGRESS_LOCK:
        return dict(UPDATE_PROGRESS)


def download_update_file(
    url: str,
    destination: Path,
    progress_callback: Callable[[int, int], None] | None = None,
) -> Path:
    destination.parent.mkdir(parents=True, exist_ok=True)
    with urlopen(url, timeout=UPDATE_TIMEOUT_SECONDS) as response:
        total = 0
        try:
            total = int(response.headers.get("Content-Length") or 0)
        except (TypeError, ValueError):
            total = 0
        downloaded = 0
        with destination.open("wb") as output:
            while True:
                chunk = response.read(1024 * 1024)
                if not chunk:
                    break
                output.write(chunk)
                downloaded += len(chunk)
                if progress_callback:
                    progress_callback(downloaded, total)
    return destination


def download_update_with_fallback(
    download_urls: list[str],
    destination: Path,
    expected_sha256: str,
    progress_callback: Callable[[int, int], None] | None = None,
) -> str:
    errors: list[str] = []
    for url in download_urls:
        try:
            download_update_file(url, destination, progress_callback)
            actual_sha256 = file_sha256(destination)
            if actual_sha256 == expected_sha256:
                return url
            destination.unlink(missing_ok=True)
            errors.append(f"{url}: sha256 mismatch {actual_sha256}")
        except Exception as exc:
            destination.unlink(missing_ok=True)
            errors.append(f"{url}: {exc}")
    raise RuntimeError("所有更新下载地址均失败: " + " | ".join(errors))


def ensure_updater_executable(runtime_dir: Path) -> Path:
    runtime_dir.mkdir(parents=True, exist_ok=True)
    target = runtime_dir / UPDATER_EXE_NAME
    bundled = resource_path(UPDATER_EXE_NAME)
    if not bundled.exists():
        raise FileNotFoundError(f"缺少更新器: {bundled}")
    target.write_bytes(bundled.read_bytes())
    return target


def launch_updater(updater_path: Path, source_path: Path, target_path: Path) -> None:
    if not updater_path.exists():
        raise FileNotFoundError(f"缺少更新器: {updater_path}")
    if not source_path.exists():
        raise FileNotFoundError(f"缺少新版程序: {source_path}")

    command = [
        str(updater_path),
        "--pid",
        str(os.getpid()),
        "--source",
        str(source_path),
        "--target",
        str(target_path),
        "--restart",
    ]
    creationflags = subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0
    subprocess.Popen(command, cwd=str(target_path.parent), creationflags=creationflags)


def configure_fonts() -> None:
    candidates = ["Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Arial Unicode MS", "DejaVu Sans"]
    installed = {font.name for font in font_manager.fontManager.ttflist}
    for name in candidates:
        if name in installed:
            plt.rcParams["font.sans-serif"] = [name, "DejaVu Sans"]
            break
    plt.rcParams["axes.unicode_minus"] = False


# ─────────────────────────── 标签辅助 ───────────────────────────


def default_group_labels(count: int) -> list[str]:
    return [f"Group {idx}" for idx in range(1, count + 1)]


def labels_for_segment_count(labels: ChartLabels, segment_count: int) -> ChartLabels:
    fallback = default_group_labels(segment_count)
    group_labels = [
        (labels.group_labels[idx].strip() or fallback[idx])
        if idx < len(labels.group_labels)
        else fallback[idx]
        for idx in range(segment_count)
    ]
    return ChartLabels(
        overview_title=labels.overview_title or "S11 Comparison",
        group_labels=group_labels,
        group_colors=labels.group_colors[:segment_count],
        upper_limit=labels.upper_limit or "Upper Limit",
        lower_limit=labels.lower_limit or "Lower Limit",
    )


def normalize_group_colors(values: list[str]) -> list[str]:
    return [value.strip() for value in values if HEX_COLOR_PATTERN.fullmatch(value.strip())]


def colors_for_group(labels: ChartLabels, group_idx: int) -> tuple[str, str]:
    if group_idx < len(labels.group_colors):
        color = labels.group_colors[group_idx]
        if HEX_COLOR_PATTERN.fullmatch(color):
            return (color, color)
    return SERIES_COLORS[group_idx % len(SERIES_COLORS)]


def antenna_sort_key(key: str) -> tuple[str, str, tuple[int, ...]]:
    """按 S参数 → 天线名 → sub_id 数字序排列。"""
    match = ANT_KEY_PATTERN.match(key)
    if not match:
        return ("ZZZ", key, ())
    s_param = match.group(1)    # S11, S21
    ant_name = match.group(2)   # Ant2, ANT4_IL
    sub_id = match.group(3)     # 0:0:0
    try:
        sub = tuple(int(part) for part in sub_id.split(":"))
    except ValueError:
        sub = ()
    return (s_param, ant_name, sub)


def safe_filename(key: str) -> str:
    return re.sub(r"[^\w\-]+", "_", key).strip("_")


def excel_sheet_name(image_path: Path, used_names: set[str]) -> str:
    if image_path.stem == "summary_all_compare":
        base = "Summary"
    else:
        base = image_path.stem.removesuffix("_compare")
    base = re.sub(r"[\[\]\:\*\?\/\\]+", "_", base).strip() or "Chart"
    base = base[:31]
    name = base
    suffix = 2
    while name in used_names:
        suffix_text = f"_{suffix}"
        name = f"{base[:31 - len(suffix_text)]}{suffix_text}"
        suffix += 1
    used_names.add(name)
    return name


# ─────────────────────────── CSV 解析 ───────────────────────────


def ensure_csv_field_size_limit() -> None:
    limit = sys.maxsize
    while True:
        try:
            csv.field_size_limit(limit)
            return
        except OverflowError:
            limit //= 10


def detect_csv_delimiter(csv_path: Path) -> str:
    with csv_path.open("r", newline="", encoding="utf-8-sig", errors="replace") as file:
        sample = file.read(65536)
    return "\t" if sample.count("\t") > sample.count(",") else ","


def iter_csv_rows(csv_path: Path):
    ensure_csv_field_size_limit()
    delimiter = detect_csv_delimiter(csv_path)
    with csv_path.open("r", newline="", encoding="utf-8-sig") as file:
        for row in csv.reader(file, delimiter=delimiter):
            yield row


def normalized_header_key(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def is_serial_header(value: str) -> bool:
    return normalized_header_key(value) in SERIAL_HEADER_KEYS


def normalize_column_name(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip())


def frequency_to_mhz(value: float, unit: str | None) -> float:
    unit_key = (unit or "MHz").strip().lower()
    if unit_key in {"g", "ghz"}:
        return value * 1000.0
    if unit_key in {"k", "khz"}:
        return value / 1000.0
    return value


def parse_antenna_column(name: str) -> tuple[str, str, str, int, float] | None:
    column = normalize_column_name(name)
    match_column = re.sub(r"[_]+", " ", column)
    if not column or not MAG_PATTERN.search(match_column):
        return None

    prefix_match = RF_PREFIX_PATTERN.match(match_column)
    if not prefix_match:
        return None
    s_param = prefix_match.group(1).upper()
    rest = prefix_match.group(2).strip()

    band_match = BAND_PATTERN.search(rest)
    if not band_match:
        return None
    band = int(band_match.group(1))

    label = rest[:band_match.start()].strip(" :-_")
    if not label:
        return None
    sub_id = "0"
    sub_match = re.search(r"\(([^)]*)\)\s*$", label)
    if not sub_match:
        sub_match = re.search(r"\[([^\]]*)\]\s*$", label)
    if sub_match:
        sub_id = sub_match.group(1).strip() or "0"
        label = label[:sub_match.start()].strip(" :-_")
    else:
        trailing_sub = re.search(
            r"(.+?)[_\s\-]+(?:port|ch|channel|sub)\s*[:=_\-]?\s*(\d+)\s*$",
            label,
            re.IGNORECASE,
        )
        if not trailing_sub:
            trailing_sub = re.search(r"(.+?)[_\s]+(\d+)\s*$", label)
        if trailing_sub:
            label = trailing_sub.group(1).strip(" :-_")
            sub_id = trailing_sub.group(2)
    if not label:
        return None

    tail = rest[band_match.end():]
    freq_match = FREQ_PATTERN.search(rest) or COLON_FREQ_PATTERN.search(tail)
    freq = frequency_to_mhz(float(freq_match.group(1)), freq_match.group(2) if freq_match else None) if freq_match else float(band)
    return s_param, label, sub_id, band, freq


def looks_like_rf_column(name: str) -> bool:
    column = normalize_column_name(name)
    return bool(re.search(r"\bS\d+\b", column, re.IGNORECASE) and MAG_PATTERN.search(column))


def unrecognized_rf_examples(header: list[str], limit: int = 5) -> list[str]:
    examples = []
    for name in header:
        if looks_like_rf_column(name) and parse_antenna_column(name) is None:
            examples.append(name)
        if len(examples) >= limit:
            break
    return examples


def read_csv_header(reader) -> tuple[list[str], list[str], list[str]]:
    for row in reader:
        if row and is_serial_header(row[0]):
            try:
                upper_row = next(reader)
                lower_row = next(reader)
                next(reader)
            except StopIteration as exc:
                raise ValueError("CSV structure is incomplete: missing limit or unit rows.") from exc
            return row, upper_row, lower_row
    raise ValueError("Header row with Serial Number/SN/Barcode was not found.")


def build_antennas(header: list[str], selected_keys: list[str] | None = None) -> dict[str, list[tuple[int, int, float]]]:
    antennas: dict[str, list[tuple[int, int, float]]] = {}
    for idx, name in enumerate(header):
        parsed = parse_antenna_column(name)
        if not parsed:
            continue
        s_param, ant_name, sub_id, band, freq = parsed
        key = f"{s_param}:{ant_name}({sub_id})"
        antennas.setdefault(key, []).append((idx, band, freq))

    if not antennas:
        examples = unrecognized_rf_examples(header)
        if examples:
            raise ValueError(
                "No recognized RF Band/Mag columns were found in CSV. "
                "疑似 RF 列无法识别: " + " | ".join(examples)
            )
        raise ValueError("No SXX/S21 Band/Mag columns were found in CSV.")
    if selected_keys:
        selected = set(selected_keys)
        antennas = {key: value for key, value in antennas.items() if key in selected}
        missing = [key for key in selected_keys if key not in antennas]
        if missing:
            raise ValueError(f"Selected antennas do not exist in CSV: {missing}")
    return antennas


def compact_csv_shape(
    antennas: dict[str, list[tuple[int, int, float]]],
    antenna_keys: list[str],
) -> tuple[list[int] | None, dict[str, list[tuple[int, int, float]]]]:
    if not antenna_keys:
        return None, antennas

    keep_indices = [0]
    seen = {0}
    for key in antenna_keys:
        for idx, _, _ in antennas[key]:
            if idx not in seen:
                keep_indices.append(idx)
                seen.add(idx)
    index_map = {old_idx: new_idx for new_idx, old_idx in enumerate(keep_indices)}
    compact_antennas = {
        key: [(index_map[idx], band, freq) for idx, band, freq in values]
        for key, values in antennas.items()
    }
    return keep_indices, compact_antennas


def compact_row(row: list[str], keep_indices: list[int] | None) -> list[str]:
    if keep_indices is None:
        return row
    return [row[idx] if idx < len(row) else "" for idx in keep_indices]


def parse_csv(csv_path: Path, selected_keys: list[str] | None = None) -> ParsedData:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file does not exist: {csv_path}")

    reader = iter_csv_rows(csv_path)
    try:
        header, upper_row, lower_row = read_csv_header(reader)
        antennas = build_antennas(header, selected_keys)
        antenna_keys = sorted(antennas, key=antenna_sort_key)
        keep_indices, antennas = compact_csv_shape(antennas, antenna_keys if selected_keys else [])
        upper_row = compact_row(upper_row, keep_indices)
        lower_row = compact_row(lower_row, keep_indices)

        segments: list[list[list[str]]] = []
        current: list[list[str]] = []
        for row in reader:
            if not any(cell.strip() for cell in row):
                if current:
                    segments.append(current)
                    current = []
                continue
            current.append(compact_row(row, keep_indices))
        if current:
            segments.append(current)

        if not segments:
            raise ValueError("No sample rows were found after the Measurement Unit row.")
    finally:
        reader.close()

    return ParsedData(
        antennas=antennas,
        upper_row=upper_row,
        lower_row=lower_row,
        segments=segments,
        antenna_keys=antenna_keys,
    )


def antenna_signature(data: ParsedData) -> list[tuple[str, tuple[tuple[int, float], ...]]]:
    return [
        (key, tuple((band, freq) for _, band, freq in sorted(data.antennas[key], key=lambda item: item[1])))
        for key in data.antenna_keys
    ]


def limit_signature(data: ParsedData) -> tuple[tuple[str, ...], tuple[str, ...]]:
    cols = [
        idx
        for key in data.antenna_keys
        for idx, _, _ in sorted(data.antennas[key], key=lambda item: item[1])
    ]
    upper = tuple(data.upper_row[idx].strip() if idx < len(data.upper_row) else "" for idx in cols)
    lower = tuple(data.lower_row[idx].strip() if idx < len(data.lower_row) else "" for idx in cols)
    return upper, lower


def parse_csv_group_files(csv_paths: list[Path], selected_keys: list[str] | None = None) -> ParsedData:
    paths = [Path(path) for path in csv_paths if str(path).strip()]
    if not paths:
        raise ValueError("CSV 分组模式下请至少选择一个 CSV 文件。")

    parsed = [parse_csv(path, selected_keys=selected_keys) for path in paths]
    reference = parsed[0]
    ref_antennas = antenna_signature(reference)
    ref_limits = limit_signature(reference)
    segments: list[list[list[str]]] = []

    for idx, data in enumerate(parsed):
        if antenna_signature(data) != ref_antennas:
            raise ValueError(f"CSV 分组文件表头 item 不一致: {paths[idx]}")
        if limit_signature(data) != ref_limits:
            raise ValueError(f"CSV 分组文件 Upper/Lower Limit 不一致: {paths[idx]}")
        segments.append([row for segment in data.segments for row in segment])

    return ParsedData(
        antennas=reference.antennas,
        upper_row=reference.upper_row,
        lower_row=reference.lower_row,
        segments=segments,
        antenna_keys=reference.antenna_keys,
    )


def validate_selection(data: ParsedData, selected: list[str]) -> None:
    """对选中的天线做数据完整性校验，避免绘图阶段出现 ValueError。"""
    invalid = [k for k in selected if k not in data.antennas]
    if invalid:
        raise ValueError(f"选中的天线在 CSV 中不存在: {invalid}")

    for key in selected:
        cols = [idx for idx, _, _ in data.antennas[key]]
        for seg_idx, segment in enumerate(data.segments, start=1):
            for row_idx, row in enumerate(segment, start=1):
                missing = [idx for idx in cols if idx >= len(row) or not row[idx].strip()]
                if missing:
                    raise ValueError(
                        f"{key} 在 segment {seg_idx} 第 {row_idx} 行存在缺失数据。"
                    )


# ─────────────────────────── 绘图 ───────────────────────────


def matrix_for(rows: list[list[str]], items: list[tuple[int, int, float]]) -> np.ndarray:
    matrix = np.empty((len(rows), len(items)), dtype=float)
    for row_idx, row in enumerate(rows):
        for col_idx, (idx, _, _) in enumerate(items):
            matrix[row_idx, col_idx] = float(row[idx])
    return matrix


def sample_id_for(segment_idx: int, row_idx: int) -> str:
    return f"s{segment_idx}:r{row_idx}"


def sample_options(data: ParsedData) -> list[dict]:
    samples = []
    for segment_idx, segment in enumerate(data.segments):
        for row_idx, row in enumerate(segment):
            serial_number = row[0].strip() if row else ""
            base_label = f"Group {segment_idx + 1} / Row {row_idx + 1}"
            label = f"{base_label} / {serial_number}" if serial_number else base_label
            samples.append({
                "id": sample_id_for(segment_idx, row_idx),
                "label": label,
                "serial_number": serial_number,
                "group_index": segment_idx + 1,
                "row_index": row_idx + 1,
            })
    return samples


def scan_csv_metadata(csv_path: Path) -> CsvScanMetadata:
    if not csv_path.exists():
        raise FileNotFoundError(f"CSV file does not exist: {csv_path}")

    reader = iter_csv_rows(csv_path)
    try:
        header, upper_row, lower_row = read_csv_header(reader)
        antennas = build_antennas(header)
        antenna_keys = sorted(antennas, key=antenna_sort_key)

        samples = []
        segment_idx = 0
        row_idx = 0
        in_segment = False
        segment_count = 0
        for row in reader:
            if not any(cell.strip() for cell in row):
                if in_segment:
                    segment_count += 1
                    segment_idx += 1
                    row_idx = 0
                    in_segment = False
                continue

            in_segment = True
            serial_number = row[0].strip() if row else ""
            base_label = f"Group {segment_idx + 1} / Row {row_idx + 1}"
            samples.append({
                "id": sample_id_for(segment_idx, row_idx),
                "label": f"{base_label} / {serial_number}" if serial_number else base_label,
                "serial_number": serial_number,
                "group_index": segment_idx + 1,
                "row_index": row_idx + 1,
            })
            row_idx += 1

        if in_segment:
            segment_count += 1
        if segment_count == 0:
            raise ValueError("No sample rows were found after the Measurement Unit row.")
    finally:
        reader.close()

    return CsvScanMetadata(
        antennas=antennas,
        upper_row=upper_row,
        lower_row=lower_row,
        antenna_keys=antenna_keys,
        segment_count=segment_count,
        samples=samples,
    )


def directory_result(path: Path) -> dict:
    stat = path.stat()
    return {
        "name": path.name,
        "path": str(path),
        "modified": stat.st_mtime,
    }


def normalized_resource_code(value: str) -> str:
    return re.sub(r"\s+", "", str(value or "")).upper()


def spc_resource_similarity(query: str, candidate: str) -> float:
    query_code = normalized_resource_code(query)
    candidate_code = normalized_resource_code(candidate)
    if not query_code or not candidate_code:
        return 0.0
    if query_code == candidate_code:
        return 2.0
    if query_code in candidate_code:
        return 1.0 + len(query_code) / max(len(candidate_code), 1)
    if candidate_code in query_code:
        return 1.0 + len(candidate_code) / max(len(query_code), 1)
    return SequenceMatcher(None, query_code, candidate_code).ratio()


def search_spc_resources(spc_root: Path | str, query: str, limit: int = 50) -> list[dict]:
    root = Path(str(spc_root or DEFAULT_SPC_ROOT).strip())
    if not root.exists():
        raise FileNotFoundError(f"SPC root does not exist: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"SPC root is not a directory: {root}")

    resource_code = normalized_resource_code(query)
    if not resource_code:
        return []
    resource = root / resource_code
    if resource.exists() and resource.is_dir():
        return [directory_result(resource)]
    return []


def suggest_spc_resources(spc_root: Path | str, query: str, limit: int = 5) -> list[dict]:
    root = Path(str(spc_root or DEFAULT_SPC_ROOT).strip())
    if not root.exists():
        raise FileNotFoundError(f"SPC root does not exist: {root}")
    if not root.is_dir():
        raise NotADirectoryError(f"SPC root is not a directory: {root}")

    query_code = normalized_resource_code(query)
    if not query_code:
        return []

    scored: list[tuple[float, str, Path]] = []
    with os.scandir(root) as entries:
        for entry in entries:
            try:
                if not entry.is_dir():
                    continue
            except OSError:
                continue
            score = spc_resource_similarity(query_code, entry.name)
            if score >= 0.58:
                scored.append((score, entry.name, Path(entry.path)))

    scored.sort(key=lambda item: (-item[0], item[1]))
    return [directory_result(path) for _, _, path in scored[:limit]]


def list_spc_dates(
    resource_path: Path | str,
    limit: int | None = None,
    recent_days: int = 15,
    include_all: bool = False,
) -> list[dict]:
    resource = Path(str(resource_path or "").strip())
    if not resource.exists():
        raise FileNotFoundError(f"SPC resource path does not exist: {resource}")
    if not resource.is_dir():
        raise NotADirectoryError(f"SPC resource path is not a directory: {resource}")

    if not include_all:
        today = date.today()
        days = max(1, int(recent_days or 15))
        dates = []
        for offset in range(days):
            child = resource / (today - timedelta(days=offset)).strftime("%Y%m%d")
            if child.is_dir():
                dates.append(directory_result(child))
        return dates

    dates = [
        directory_result(child)
        for child in resource.iterdir()
        if child.is_dir() and re.fullmatch(r"\d{8}", child.name)
    ]
    dates = sorted(dates, key=lambda item: item["name"], reverse=True)
    return dates[:limit] if limit is not None else dates


def find_largest_spc_csv(date_path: Path | str) -> Path:
    date_dir = Path(str(date_path or "").strip())
    if not date_dir.exists():
        raise FileNotFoundError(f"SPC date path does not exist: {date_dir}")
    if not date_dir.is_dir():
        raise NotADirectoryError(f"SPC date path is not a directory: {date_dir}")

    largest: tuple[int, float, Path] | None = None
    for path in date_dir.rglob("*"):
        try:
            if not path.is_file() or path.suffix.lower() != ".csv":
                continue
            stat = path.stat()
        except OSError:
            continue
        candidate = (int(stat.st_size), float(stat.st_mtime), path)
        if largest is None or candidate[:2] > largest[:2]:
            largest = candidate

    if largest is None:
        raise FileNotFoundError(f"No CSV file was found under SPC date path: {date_dir}")
    return largest[2]


def copy_spc_csv_for_read(source_csv: Path | str) -> Path:
    source = Path(str(source_csv or "").strip())
    if not source.exists():
        raise FileNotFoundError(f"SPC CSV does not exist: {source}")
    if not source.is_file():
        raise FileNotFoundError(f"SPC CSV is not a file: {source}")

    snapshot_dir = Path(tempfile.gettempdir()) / SPC_SNAPSHOT_SUBDIR
    snapshot_dir.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d%H%M%S")
    safe_name = safe_filename(source.stem) or "spc_csv"
    target = snapshot_dir / f"{safe_name}_{stamp}_{threading.get_ident()}.csv"
    with source.open("rb") as src, target.open("wb") as dst:
        shutil.copyfileobj(src, dst, length=1024 * 1024)
    return target


def resolve_spc_csv_snapshot(payload: dict) -> tuple[Path, Path]:
    snapshot_value = str(payload.get("spc_snapshot_csv") or "").strip()
    source_value = str(payload.get("spc_source_csv") or "").strip()
    if snapshot_value:
        snapshot = Path(snapshot_value)
        if not snapshot.exists():
            raise FileNotFoundError(f"SPC CSV snapshot does not exist: {snapshot}")
        source = Path(source_value) if source_value else snapshot
        return source, snapshot

    date_path = Path(str(payload.get("spc_date_path") or "").strip())
    source = find_largest_spc_csv(date_path)
    return source, copy_spc_csv_for_read(source)


def parse_sample_id(value: str) -> tuple[int, int] | None:
    match = re.fullmatch(r"s(\d+):r(\d+)", str(value or "").strip())
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def normalize_highlight_sample_ids(payload: dict) -> list[str]:
    values = payload.get("highlight_sample_ids")
    if values is None:
        values = [payload.get("highlight_sample_id")]
    if isinstance(values, str):
        values = [values]

    sample_ids: list[str] = []
    seen: set[str] = set()
    for value in values or []:
        sample_id = str(value or "").strip()
        if sample_id and sample_id not in seen:
            sample_ids.append(sample_id)
            seen.add(sample_id)
    return sample_ids


def draw_chart(
    ax,
    antenna_key: str,
    items: list[tuple[int, int, float]],
    data: ParsedData,
    labels: ChartLabels,
    legend_fontsize: int = 8,
    highlight_sample_id: str = "",
    highlight_sample_ids: list[str] | None = None,
) -> None:
    items = sorted(items, key=lambda item: item[1])
    cols = [idx for idx, _, _ in items]
    freqs = np.array([freq for _, _, freq in items], dtype=float)
    upper = np.array([float(data.upper_row[idx]) for idx in cols], dtype=float)
    lower = np.array([float(data.lower_row[idx]) for idx in cols], dtype=float)
    highlight_ids = list(highlight_sample_ids or [])
    if highlight_sample_id:
        highlight_ids.append(highlight_sample_id)
    highlight_positions = [
        (sample_id, parsed)
        for sample_id in highlight_ids
        if (parsed := parse_sample_id(sample_id)) is not None
    ]

    matrices: list[np.ndarray] = []
    for idx, segment in enumerate(data.segments):
        thin_color, mean_color = colors_for_group(labels, idx)
        label = labels.group_labels[idx] if idx < len(labels.group_labels) else f"Group {idx + 1}"
        matrix = matrix_for(segment, items)
        matrices.append(matrix)
        if len(matrix):
            line_segments = [np.column_stack((freqs, values)) for values in matrix]
            ax.add_collection(LineCollection(line_segments, colors=thin_color, alpha=0.18, linewidths=0.75))
        ax.plot(
            freqs,
            matrix.mean(axis=0),
            color=mean_color,
            linewidth=2.2,
            label=f"{label} mean (n={len(segment)})",
        )
        for highlight_idx, (_, highlight) in enumerate(highlight_positions):
            if highlight[0] == idx and highlight[1] < len(segment):
                row_idx = highlight[1]
                row = segment[row_idx]
                serial_number = row[0].strip() if row else ""
                sample_label = serial_number or f"Group {idx + 1} Row {row_idx + 1}"
                label_prefix = "Problem sample" if len(highlight_positions) == 1 else f"Problem sample {highlight_idx + 1}"
                ax.plot(
                    freqs,
                    matrix[row_idx],
                    color=HIGHLIGHT_COLORS[highlight_idx % len(HIGHLIGHT_COLORS)],
                    linewidth=3.2,
                    label=f"{label_prefix}: {sample_label}",
                    zorder=6 + highlight_idx,
                )

    ax.plot(freqs, upper, color="#dc2626", linestyle="--", linewidth=1.45, label=labels.upper_limit)
    ax.plot(freqs, lower, color="#dc2626", linestyle="--", linewidth=1.45, label=labels.lower_limit)

    all_values = np.concatenate([m.ravel() for m in matrices])
    y_min = min(float(lower.min()), float(all_values.min()))
    y_max = max(float(upper.max()), float(all_values.max()))
    pad = max((y_max - y_min) * 0.08, 0.02)

    ax.set_title(antenna_key, fontsize=11)
    ax.set_xlabel("Frequency (MHz)", fontsize=9)
    ax.set_ylabel("Mag", fontsize=9)
    ax.set_xlim(float(freqs.min()), float(freqs.max()))
    ax.set_ylim(y_min - pad, y_max + pad)
    ax.grid(True, alpha=0.28)
    ax.tick_params(axis="both", labelsize=8)
    ax.legend(loc="best", frameon=True, fontsize=legend_fontsize)


def chart_labels_from_payload(payload: dict) -> ChartLabels:
    return ChartLabels(
        overview_title=(payload.get("overview_title") or "").strip() or "S11 Comparison",
        group_labels=[s.strip() for s in (payload.get("group_labels") or [])],
        group_colors=normalize_group_colors(list(payload.get("group_colors") or [])),
        upper_limit=(payload.get("upper_limit") or "").strip() or "Upper Limit",
        lower_limit=(payload.get("lower_limit") or "").strip() or "Lower Limit",
    )


def render_preview_png(csv_path: Path, antenna_key: str, labels: ChartLabels) -> bytes:
    if not antenna_key:
        raise ValueError("请选择一个天线进行预览。")

    data = parse_csv(csv_path)
    return render_preview_png_for_data(data, antenna_key, labels)


def render_preview_png_for_data(
    data: ParsedData,
    antenna_key: str,
    labels: ChartLabels,
    highlight_sample_id: str = "",
    highlight_sample_ids: list[str] | None = None,
) -> bytes:
    configure_fonts()
    validate_selection(data, [antenna_key])
    labels = labels_for_segment_count(labels, len(data.segments))
    fig, ax = plt.subplots(figsize=(11, 5.8), dpi=120)
    try:
        draw_chart(
            ax,
            antenna_key,
            data.antennas[antenna_key],
            data,
            labels,
            legend_fontsize=8,
            highlight_sample_id=highlight_sample_id,
            highlight_sample_ids=highlight_sample_ids,
        )
        fig.tight_layout()
        buffer = BytesIO()
        fig.savefig(buffer, format="png")
        return buffer.getvalue()
    finally:
        plt.close(fig)


def summary_layout(count: int) -> tuple[int, int]:
    """根据天线数量计算 summary all 的网格布局。"""
    if count <= 0:
        return (1, 1)
    cols = math.ceil(math.sqrt(count))
    rows = math.ceil(count / cols)
    return (rows, cols)


def draw_summary_figure(
    data: ParsedData,
    selected_keys: list[str],
    labels: ChartLabels,
    highlight_sample_ids: list[str] | None = None,
) -> plt.Figure:
    if not selected_keys:
        raise ValueError("请至少勾选一个天线。")

    configure_fonts()
    validate_selection(data, selected_keys)
    labels = labels_for_segment_count(labels, len(data.segments))
    rows, cols = summary_layout(len(selected_keys))
    fig_w = max(8.0, cols * 7.0)
    fig_h = max(6.0, rows * 5.0)
    fig, axes = plt.subplots(rows, cols, figsize=(fig_w, fig_h), dpi=140, squeeze=False)
    fig.suptitle(labels.overview_title, fontsize=18, y=0.99)
    flat = list(axes.ravel())
    for ax, key in zip(flat, selected_keys):
        draw_chart(
            ax,
            key,
            data.antennas[key],
            data,
            labels,
            legend_fontsize=6,
            highlight_sample_ids=highlight_sample_ids,
        )
    for ax in flat[len(selected_keys) :]:
        ax.axis("off")
    fig.tight_layout(rect=(0, 0, 1, 0.96))
    return fig


def render_summary_preview_png_for_data(
    data: ParsedData,
    selected_keys: list[str],
    labels: ChartLabels,
    highlight_sample_ids: list[str] | None = None,
) -> bytes:
    fig = draw_summary_figure(data, selected_keys, labels, highlight_sample_ids)
    try:
        buffer = BytesIO()
        fig.savefig(buffer, format="png")
        return buffer.getvalue()
    finally:
        plt.close(fig)


def generate_charts(
    csv_path: Path,
    output_dir: Path,
    selected_keys: list[str],
    labels: ChartLabels,
    highlight_sample_ids: list[str] | None = None,
) -> list[Path]:
    data = parse_csv(csv_path, selected_keys=selected_keys)
    return generate_charts_for_data(data, output_dir, selected_keys, labels, highlight_sample_ids)


def generate_charts_for_data(
    data: ParsedData,
    output_dir: Path,
    selected_keys: list[str],
    labels: ChartLabels,
    highlight_sample_ids: list[str] | None = None,
) -> list[Path]:
    if not selected_keys:
        raise ValueError("请至少勾选一个天线。")

    configure_fonts()
    output_dir.mkdir(parents=True, exist_ok=True)
    validate_selection(data, selected_keys)
    labels = labels_for_segment_count(labels, len(data.segments))

    outputs: list[Path] = []

    for key in selected_keys:
        fig, ax = plt.subplots(figsize=(14, 7.5), dpi=160)
        draw_chart(
            ax,
            key,
            data.antennas[key],
            data,
            labels,
            legend_fontsize=9,
            highlight_sample_ids=highlight_sample_ids,
        )
        fig.tight_layout()
        out_path = output_dir / f"{safe_filename(key)}_compare.png"
        fig.savefig(out_path)
        plt.close(fig)
        outputs.append(out_path)

    fig = draw_summary_figure(data, selected_keys, labels, highlight_sample_ids)
    summary_path = output_dir / "summary_all_compare.png"
    fig.savefig(summary_path)
    plt.close(fig)
    outputs.insert(0, summary_path)

    return outputs


def create_excel_report(image_paths: list[Path], excel_path: Path) -> Path:
    if not image_paths:
        raise ValueError("没有可写入 Excel 的图表文件。")

    excel_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    used_names: set[str] = set()
    first_sheet = workbook.active

    for idx, image_path in enumerate(image_paths):
        sheet_name = excel_sheet_name(image_path, used_names)
        worksheet = first_sheet if idx == 0 else workbook.create_sheet()
        worksheet.title = sheet_name
        worksheet["A1"] = image_path.name
        image = ExcelImage(str(image_path))
        if image.width > 1200:
            ratio = 1200 / image.width
            image.width = 1200
            image.height = int(image.height * ratio)
        worksheet.add_image(image, "A3")
        worksheet.column_dimensions["A"].width = 24

    workbook.save(excel_path)
    workbook.close()
    return excel_path


# ─────────────────────────── pywebview Api ───────────────────────────


class Api:
    """暴露给前端 JS 的接口。所有方法返回 dict。"""

    def __init__(self) -> None:
        self._window = None
        self._update_thread: threading.Thread | None = None

    def bind_window(self, window) -> None:
        self._window = window

    # ---- 文件选择 ----
    def pick_csv(self) -> dict:
        if not self._window:
            return {"ok": False, "error": "窗口未初始化"}
        try:
            import webview
            result = self._window.create_file_dialog(
                dialog_type=webview.OPEN_DIALOG,
                file_types=("CSV files (*.csv)", "All files (*.*)"),
            )
            path = result[0] if result else ""
            return {"ok": True, "path": path}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def pick_dir(self) -> dict:
        if not self._window:
            return {"ok": False, "error": "窗口未初始化"}
        try:
            import webview
            result = self._window.create_file_dialog(dialog_type=webview.FOLDER_DIALOG)
            path = result[0] if result else ""
            return {"ok": True, "path": path}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    # ---- 默认值 ----
    def get_defaults(self) -> dict:
        return {
            "ok": True,
            "output_dir": str(default_output_dir()),
            "spc_root": DEFAULT_SPC_ROOT,
            "version": APP_VERSION,
            "display_version": APP_DISPLAY_VERSION,
        }

    def search_spc_resources(self, payload: dict) -> dict:
        try:
            root = payload.get("spc_root") or DEFAULT_SPC_ROOT
            query = payload.get("query") or ""
            resources = search_spc_resources(
                root,
                query,
            )
            suggestions = [] if resources else suggest_spc_resources(root, query)
            return {"ok": True, "resources": resources, "suggestions": suggestions}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def list_spc_dates(self, payload: dict) -> dict:
        try:
            include_all = bool(payload.get("include_all"))
            dates = list_spc_dates(
                payload.get("spc_resource_path") or "",
                include_all=include_all,
            )
            return {"ok": True, "dates": dates}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    # ---- 扫描 CSV ----
    # ---- 在线更新 ----
    def check_update(self, payload: dict | None = None) -> dict:
        payload = payload or {}
        if not UPDATE_MANIFEST_URL:
            return {"ok": False, "error": "未配置更新地址，请先设置 UPDATE_MANIFEST_URL"}

        try:
            manifest = normalize_update_manifest(fetch_update_manifest(UPDATE_MANIFEST_URL), APP_VERSION, APP_UPDATE_EPOCH)
            if not manifest["has_update"]:
                return {
                    "ok": True,
                    "has_update": False,
                    "version": manifest["version"],
                    "display_version": manifest.get("display_version", manifest["version"]),
                    "current_version": APP_VERSION,
                    "current_display_version": APP_DISPLAY_VERSION,
                    "message": "当前已是最新版本",
                }
            return {
                "ok": True,
                "has_update": True,
                "version": manifest["version"],
                "display_version": manifest.get("display_version", manifest["version"]),
                "current_version": APP_VERSION,
                "current_display_version": APP_DISPLAY_VERSION,
                "notes": manifest.get("notes", ""),
                "download_urls": manifest["download_urls"],
                "sha256": manifest["sha256"],
                "message": "发现新版本",
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

    def _install_update_worker(self, manifest: dict) -> None:
        try:
            download_dir = Path(tempfile.gettempdir()) / "RFPlotTool-updates"
            source_path = download_dir / f"RFPlotTool-{manifest['version']}.exe"

            def report_download(downloaded: int, total: int) -> None:
                if total > 0:
                    percent = min(80, max(1, int(downloaded * 80 / total)))
                    message = f"正在下载更新 {downloaded / 1024 / 1024:.1f}/{total / 1024 / 1024:.1f} MB"
                else:
                    percent = 10
                    message = f"正在下载更新 {downloaded / 1024 / 1024:.1f} MB"
                set_update_progress("download", percent, message)

            set_update_progress("download", 1, "正在连接下载服务器...")
            selected_url = download_update_with_fallback(
                manifest["download_urls"],
                source_path,
                manifest["sha256"],
                report_download,
            )
            set_update_progress("verify", 85, "正在校验更新文件...")

            app_dir = app_base_dir()
            target_path = Path(sys.executable).resolve() if getattr(sys, "frozen", False) else app_dir / "RFPlotTool.exe"
            set_update_progress("prepare", 92, "正在准备安装器...")
            updater_path = ensure_updater_executable(download_dir)
            set_update_progress("launch", 98, "正在启动安装器...")
            launch_updater(updater_path, source_path, target_path)
            set_update_progress(
                "launch",
                100,
                f"更新已下载，程序将退出并安装新版（来源：{selected_url}）",
                running=False,
            )
            if self._window:
                self._window.destroy()
        except Exception as exc:
            set_update_progress("error", 100, "更新失败", str(exc), running=False)

    def install_update(self, payload: dict | None = None) -> dict:
        payload = payload or {}
        if not payload.get("download_urls") or not payload.get("sha256"):
            checked = self.check_update()
            if not checked.get("ok") or not checked.get("has_update"):
                return checked
            payload = checked

        with UPDATE_PROGRESS_LOCK:
            if bool(UPDATE_PROGRESS.get("running")):
                return {"ok": False, "error": "已有更新任务正在运行"}
            UPDATE_PROGRESS.update({
                "running": True,
                "stage": "start",
                "percent": 0,
                "message": "准备开始更新...",
                "error": "",
            })

        self._update_thread = threading.Thread(target=self._install_update_worker, args=(payload,), daemon=True)
        self._update_thread.start()
        return {"ok": True, "started": True, "message": "更新已开始"}

    def update_progress(self) -> dict:
        return {"ok": True, **get_update_progress_snapshot()}

    def data_from_payload(self, payload: dict, selected_keys: list[str] | None = None) -> tuple[ParsedData, Path]:
        if payload.get("group_mode") == "csv_group":
            csv_paths = [Path(path) for path in (payload.get("csv_paths") or []) if str(path).strip()]
            data = parse_csv_group_files(csv_paths, selected_keys=selected_keys)
            return data, csv_paths[0]
        if payload.get("group_mode") == "spc":
            source_csv, snapshot_csv = resolve_spc_csv_snapshot(payload)
            data = parse_csv(snapshot_csv, selected_keys=selected_keys)
            return data, source_csv
        csv_path = Path(payload.get("csv_path", "").strip())
        return parse_csv(csv_path, selected_keys=selected_keys), csv_path

    def scan_csv(self, csv_path: str | dict) -> dict:
        try:
            if isinstance(csv_path, dict):
                if csv_path.get("group_mode") == "csv_group":
                    data, _ = self.data_from_payload(csv_path)
                    metadata = None
                    source_csv = None
                    snapshot_csv = None
                elif csv_path.get("group_mode") == "spc":
                    source_csv, snapshot_csv = resolve_spc_csv_snapshot(csv_path)
                    metadata = scan_csv_metadata(snapshot_csv)
                    data = None
                else:
                    metadata = scan_csv_metadata(Path(csv_path.get("csv_path", "").strip()))
                    data = None
                    source_csv = None
                    snapshot_csv = None
            else:
                metadata = scan_csv_metadata(Path(csv_path))
                data = None
                source_csv = None
                snapshot_csv = None
        except Exception as exc:
            return {"ok": False, "error": str(exc)}

        antennas = []
        antenna_keys = metadata.antenna_keys if metadata else data.antenna_keys
        antenna_map = metadata.antennas if metadata else data.antennas
        for key in antenna_keys:
            match = ANT_KEY_PATTERN.match(key)
            s_param = match.group(1) if match else ""
            ant_name = match.group(2) if match else key
            sub_id = match.group(3) if match else ""
            antennas.append({
                "key": key,
                "s_param": s_param,
                "ant_name": ant_name,
                "sub_id": sub_id,
                "band_count": len(antenna_map[key]),
            })
        result = {
            "ok": True,
            "segment_count": metadata.segment_count if metadata else len(data.segments),
            "antennas": antennas,
            "samples": metadata.samples if metadata else sample_options(data),
        }
        if source_csv and snapshot_csv:
            result["source_csv"] = str(source_csv)
            result["snapshot_csv"] = str(snapshot_csv)
        return result

    # ---- 生成图表 ----
    def preview(self, payload: dict) -> dict:
        try:
            antenna_key = (payload.get("antenna_key") or "").strip()
            data, _ = self.data_from_payload(payload, selected_keys=[antenna_key] if antenna_key else None)
            image_bytes = render_preview_png_for_data(
                data,
                antenna_key,
                chart_labels_from_payload(payload),
                highlight_sample_ids=normalize_highlight_sample_ids(payload),
            )
            encoded = base64.b64encode(image_bytes).decode("ascii")
            return {
                "ok": True,
                "antenna": antenna_key,
                "image": f"data:image/png;base64,{encoded}",
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
        finally:
            plt.close("all")
            import gc
            gc.collect()

    def summary_preview(self, payload: dict) -> dict:
        try:
            selected = list(payload.get("selected_antennas") or [])
            if not selected:
                raise ValueError("请至少勾选一个天线。")
            data, _ = self.data_from_payload(payload, selected_keys=selected or None)
            image_bytes = render_summary_preview_png_for_data(
                data,
                selected,
                chart_labels_from_payload(payload),
                highlight_sample_ids=normalize_highlight_sample_ids(payload),
            )
            encoded = base64.b64encode(image_bytes).decode("ascii")
            return {
                "ok": True,
                "antenna": "View All",
                "image": f"data:image/png;base64,{encoded}",
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
        finally:
            plt.close("all")
            import gc
            gc.collect()

    def generate(self, payload: dict) -> dict:
        try:
            selected = list(payload.get("selected_antennas") or [])
            data, source_path = self.data_from_payload(payload, selected_keys=selected or None)
            base_output = Path(payload.get("output_dir", "").strip() or default_output_dir())
            # 按 CSV 文件名创建子文件夹
            csv_subfolder = source_path.stem if payload.get("group_mode") != "csv_group" else f"{source_path.stem}_csv_groups"
            output_dir = base_output / csv_subfolder
            labels = chart_labels_from_payload(payload)
            outputs = generate_charts_for_data(
                data,
                output_dir,
                selected,
                labels,
                highlight_sample_ids=normalize_highlight_sample_ids(payload),
            )
            excel_path = create_excel_report(outputs, output_dir / "charts.xlsx")
            outputs.append(excel_path)
            return {
                "ok": True,
                "output_dir": str(output_dir.resolve()),
                "files": [p.name for p in outputs],
            }
        except Exception as exc:
            return {"ok": False, "error": str(exc)}
        finally:
            # 彻底释放 matplotlib 资源，防止状态残留
            plt.close("all")
            import gc
            gc.collect()

    def open_dir(self, path: str) -> dict:
        try:
            target = Path(path).resolve()
            if not target.exists():
                return {"ok": False, "error": f"路径不存在: {target}"}
            if sys.platform.startswith("win"):
                subprocess.Popen(["explorer", str(target)])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target)])
            return {"ok": True}
        except Exception as exc:
            return {"ok": False, "error": str(exc)}


# ─────────────────────────── WebView2 检测 ───────────────────────────


def _check_webview2() -> bool:
    """检测系统是否安装了 Edge WebView2 Runtime。"""
    import winreg

    registry_paths = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
        (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}"),
    ]
    for hive, key_path in registry_paths:
        try:
            with winreg.OpenKey(hive, key_path) as key:
                version, _ = winreg.QueryValueEx(key, "pv")
                if version and version != "0.0.0.0":
                    return True
        except OSError:
            continue
    return False


def _prompt_install_webview2() -> None:
    """弹窗提示用户安装 WebView2 Runtime。"""
    import ctypes

    msg = (
        "本程序需要 Microsoft Edge WebView2 Runtime 才能正常显示界面。\n\n"
        "点击「确定」将打开下载页面，安装后重新启动程序即可。\n\n"
        "下载地址：\n"
        "https://developer.microsoft.com/en-us/microsoft-edge/webview2/"
    )
    result = ctypes.windll.user32.MessageBoxW(
        None, msg, "RF Plot Tool - 缺少运行时组件", 0x31  # MB_OKCANCEL | MB_ICONWARNING
    )
    if result == 1:  # IDOK
        import webbrowser
        webbrowser.open("https://developer.microsoft.com/en-us/microsoft-edge/webview2/")


# ─────────────────────────── 入口 ───────────────────────────


def run_gui() -> int:
    if sys.platform.startswith("win") and not _check_webview2():
        _prompt_install_webview2()
        return 1

    import webview

    api = Api()
    html_path = resource_path("webui.html")
    if not html_path.exists():
        print(f"前端文件缺失: {html_path}", file=sys.stderr)
        return 1

    window = webview.create_window(
        title="RF Plot Tool",
        url=html_path.as_uri(),
        js_api=api,
        width=1180,
        height=820,
        min_size=(960, 680),
        background_color="#1a1d2e",
        confirm_close=True,
    )
    api.bind_window(window)
    webview.start(debug=False)
    return 0


def run_cli(args: argparse.Namespace) -> int:
    csv_path = args.csv or find_default_csv()
    data = parse_csv(csv_path)
    selected = args.antennas or data.antenna_keys
    output_dir = args.output_dir or default_output_dir()
    outputs = generate_charts(csv_path, output_dir, selected, ChartLabels())
    print("generated")
    for output in outputs:
        print(output.as_posix())
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description="RF S11 对比图生成工具")
    parser.add_argument("--cli", action="store_true", help="不启动 GUI，直接生成所有天线图。")
    parser.add_argument("--csv", type=Path, default=None, help="CSV 文件路径")
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help="输出目录，默认是 exe/脚本 所在目录下的 output/",
    )
    parser.add_argument(
        "--antennas",
        nargs="*",
        default=None,
        help="指定要绘制的天线 key，如 'Ant2(0:0:0)'，默认全部。",
    )
    args = parser.parse_args()

    if args.cli:
        return run_cli(args)
    return run_gui()


if __name__ == "__main__":
    sys.exit(main())

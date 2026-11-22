# report_generator.py
import os
import io
import re
import difflib
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
from docx import Document
from docx.shared import Inches

import config


class ReportAutomationError(Exception):
    """Custom exception to pinpoint missing files, sheets, or columns."""
    pass


# ==========================================
# AUTO-PERSIST NEW ALIAS INTO config.py
# ==========================================
def persist_alias_to_config(config_section: str, key_name: str, new_alias: str) -> bool:
    """
    Safely appends a newly selected column alias into config.py so future runs
    recognize it automatically without prompting again.
    """
    config_path = os.path.abspath(config.__file__)
    if not os.path.exists(config_path):
        return False

    try:
        with open(config_path, "r", encoding="utf-8") as f:
            content = f.read()

        # Match `"key_name": [ ... ]` or `"key_name": "single_val"`
        list_pattern = re.compile(
            rf'("{re.escape(key_name)}"\s*:\s*\[)([^\]]*)(\])',
            re.MULTILINE
        )
        match = list_pattern.search(content)
        if match:
            inner = match.group(2).strip()
            escaped_alias = new_alias.replace('"', '\\"')
            if f'"{escaped_alias}"' in inner:
                return True
            updated_inner = f'{inner}, "{escaped_alias}"' if inner else f'"{escaped_alias}"'
            new_content = content[:match.start()] + match.group(1) + updated_inner + match.group(3) + content[match.end():]
            with open(config_path, "w", encoding="utf-8") as f:
                f.write(new_content)
            return True

        str_pattern = re.compile(
            rf'("{re.escape(key_name)}"\s*:\s*)("[^"]*")',
            re.MULTILINE
        )
        str_match = str_pattern.search(content)
        if str_match:
            existing_val = str_match.group(2)
            escaped_alias = new_alias.replace('"', '\\"')
            replacement = f'{str_match.group(1)}[{existing_val}, "{escaped_alias}"]'
            new_content = content[:str_match.start()] + replacement + content[str_match.end():]
            with open(config_path, "w", encoding="utf-8") as f:
                f.write(new_content)
            return True

    except Exception:
        return False
    return False


# ==========================================
# INTERACTIVE COLUMN PICKER DIALOG (FALLBACK)
# ==========================================
def prompt_user_for_column(
    parent_root,
    file_label: str,
    sheet_name: str,
    metric_key: str,
    aliases_tried: list,
    available_cols: list
) -> tuple:
    """
    Opens a modal dialog allowing the user to pick the matching column from the sheet
    and optionally save that alias into config.py.
    Returns (selected_column_name, should_save_to_config) or (None, False) if cancelled.
    """
    dialog = tk.Toplevel(parent_root)
    dialog.title("Column Mapping Required")
    dialog.geometry("540x310")
    dialog.resizable(False, False)
    dialog.transient(parent_root)
    dialog.grab_set()

    result = {"col": None, "save": False}

    # Filter out blank 'Unnamed: X' columns by default for cleaner selection
    clean_cols = [c for c in available_cols if not str(c).startswith("Unnamed:")]
    display_cols = clean_cols if clean_cols else available_cols

    tk.Label(
        dialog,
        text="Could Not Automatically Match Column Header",
        font=("Segoe UI", 11, "bold"),
        fg="#B25900"
    ).pack(anchor="w", padx=16, pady=(14, 4))

    info_text = (
        f"File: {file_label}\n"
        f"Sheet: '{sheet_name}'\n"
        f"Metric Key: '{metric_key}' (Tried aliases: {aliases_tried})\n\n"
        f"Please select the correct column from this sheet:"
    )
    tk.Label(dialog, text=info_text, justify="left", font=("Segoe UI", 9)).pack(anchor="w", padx=16, pady=4)

    selected_var = tk.StringVar(value=display_cols[0] if display_cols else "")
    combo = ttk.Combobox(dialog, textvariable=selected_var, values=display_cols, width=58, state="readonly")
    combo.pack(padx=16, pady=6)

    show_all_var = tk.BooleanVar(value=False)
    def toggle_unnamed():
        combo["values"] = available_cols if show_all_var.get() else display_cols

    tk.Checkbutton(
        dialog,
        text="Show 'Unnamed' columns",
        variable=show_all_var,
        command=toggle_unnamed,
        font=("Segoe UI", 8)
    ).pack(anchor="w", padx=16)

    save_var = tk.BooleanVar(value=True)
    tk.Checkbutton(
        dialog,
        text="Remember this column name in config.py for future runs",
        variable=save_var,
        font=("Segoe UI", 9, "bold"),
        fg="#0F5C7E"
    ).pack(anchor="w", padx=16, pady=(6, 10))

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=6)

    def on_confirm():
        result["col"] = selected_var.get()
        result["save"] = save_var.get()
        dialog.destroy()

    def on_cancel():
        dialog.destroy()

    tk.Button(btn_frame, text="Confirm Selection", command=on_confirm, bg="#0F5C7E", fg="white", font=("Segoe UI", 9, "bold"), padx=14, pady=4).pack(side="left", padx=8)
    tk.Button(btn_frame, text="Cancel", command=on_cancel, padx=14, pady=4).pack(side="left", padx=8)

    parent_root.wait_window(dialog)
    return result["col"], result["save"]


# ==========================================
# RESILIENT SHEET & 4-TIER COLUMN RESOLVERS
# ==========================================
def match_default_sheet(available_sheets: list, target_name: str) -> str:
    if target_name in available_sheets:
        return target_name
    target_clean = target_name.strip().lower()
    for s in available_sheets:
        if s.strip().lower() == target_clean:
            return s
    close = difflib.get_close_matches(target_name, available_sheets, n=1, cutoff=0.6)
    if close:
        return close[0]
    return available_sheets[0] if available_sheets else ""


def resolve_columns(
    df: pd.DataFrame,
    required_items: dict,
    sheet_name: str,
    file_label: str,
    warnings_list: list,
    parent_root=None,
    config_section: str = "DT_CONFIG"
) -> dict:
    """
    Resolves {logical_key: single_str_or_list_of_aliases} against actual DataFrame columns:
      Tier 1: Exact or case/whitespace-insensitive match against any alias
      Tier 2: Substring / SQL LIKE '%alias%' match against any alias (logs warning)
      Tier 3: Fuzzy similarity match via difflib (logs warning)
      Tier 4: Interactive GUI column picker + optional auto-save to config.py
    """
    match_cfg = getattr(config, "COLUMN_MATCHING", {})
    enable_fuzzy = match_cfg.get("enable_fuzzy", True)
    cutoff = match_cfg.get("similarity_cutoff", 0.65)
    prompt_on_missing = match_cfg.get("prompt_on_missing", True)

    df.columns = [str(c) for c in df.columns]
    actual_cols = list(df.columns)
    norm_lookup = {c.strip().lower(): c for c in actual_cols}

    resolved_map = {}

    for logical_key, alias_spec in required_items.items():
        aliases = alias_spec if isinstance(alias_spec, list) else [alias_spec]
        matched_col = None

        # Tier 1: Exact / case-insensitive match across all aliases
        for alias in aliases:
            if alias in actual_cols:
                matched_col = alias
                break
            alias_norm = alias.strip().lower()
            if alias_norm in norm_lookup:
                matched_col = norm_lookup[alias_norm]
                break

        if matched_col:
            resolved_map[logical_key] = matched_col
            continue

        # Tier 2: Substring / SQL LIKE '%alias%' across all aliases
        if enable_fuzzy:
            for alias in aliases:
                alias_norm = alias.strip().lower()
                sub_hits = [c for c in actual_cols if alias_norm in c.strip().lower()]
                if sub_hits:
                    matched_col = min(sub_hits, key=len)
                    warnings_list.append(
                        f"[{file_label} | {sheet_name}] '{logical_key}' matched via LIKE '{alias}' -> '{matched_col}'"
                    )
                    break

        if matched_col:
            resolved_map[logical_key] = matched_col
            continue

        # Tier 3: Fuzzy string similarity across all aliases
        if enable_fuzzy:
            for alias in aliases:
                close_hits = difflib.get_close_matches(alias, actual_cols, n=1, cutoff=cutoff)
                if not close_hits:
                    norm_close = difflib.get_close_matches(alias.strip().lower(), list(norm_lookup.keys()), n=1, cutoff=cutoff)
                    if norm_close:
                        close_hits = [norm_lookup[norm_close[0]]]
                if close_hits:
                    matched_col = close_hits[0]
                    warnings_list.append(
                        f"[{file_label} | {sheet_name}] '{logical_key}' matched via fuzzy '{alias}' -> '{matched_col}'"
                    )
                    break

        if matched_col:
            resolved_map[logical_key] = matched_col
            continue

        # Tier 4: Interactive User Selection Popup
        if prompt_on_missing and parent_root is not None:
            chosen_col, should_save = prompt_user_for_column(
                parent_root, file_label, sheet_name, logical_key, aliases, actual_cols
            )
            if chosen_col:
                resolved_map[logical_key] = chosen_col
                # Update in-memory config immediately for current session
                if isinstance(alias_spec, list):
                    alias_spec.append(chosen_col)
                saved_msg = ""
                if should_save and persist_alias_to_config(config_section, logical_key, chosen_col):
                    saved_msg = " (saved to config.py)"
                warnings_list.append(
                    f"[{file_label} | {sheet_name}] '{logical_key}' manually mapped -> '{chosen_col}'{saved_msg}"
                )
                continue

        raise ReportAutomationError(
            f"[{file_label} -> Sheet '{sheet_name}'] Could not find column for '{logical_key}' "
            f"(tried: {aliases}).\nAvailable columns found: {actual_cols}"
        )

    return resolved_map


# ==========================================
# 1. CHART GENERATION ENGINE (WEEKLY OSS)
# ==========================================
def generate_weekly_charts(weekly_path: str, selected_sheet: str, warnings_list: list, parent_root=None) -> dict:
    cfg = config.WEEKLY_CONFIG
    style = config.CHART_STYLE
    file_label = os.path.basename(weekly_path)

    xls = pd.ExcelFile(weekly_path)
    if selected_sheet not in xls.sheet_names:
        raise ReportAutomationError(f"[{file_label}] Selected sheet '{selected_sheet}' not found.")

    df = pd.read_excel(xls, sheet_name=selected_sheet)

    aliases_cfg = cfg.get("column_aliases", {})
    needed_items = {
        "date_col": cfg["date_col"],
        "network_col": cfg["network_col"]
    }
    for _, (col_key, _, _, _) in cfg["charts"].items():
        needed_items[col_key] = aliases_cfg.get(col_key, col_key)

    col_map = resolve_columns(
        df, needed_items, selected_sheet, file_label, warnings_list, parent_root, "WEEKLY_CONFIG"
    )
    date_col = col_map["date_col"]
    net_col = col_map["network_col"]

    df[date_col] = pd.to_datetime(df[date_col], errors="coerce")
    df = df.dropna(subset=[date_col]).sort_values(by=date_col)
    df[net_col] = df[net_col].astype(str).str.strip()

    chart_images = {}

    for tag, (metric_key, agg_func, is_pct, net_filter) in cfg["charts"].items():
        metric_col = col_map[metric_key]
        sub = df[df[net_col] == net_filter].copy()
        if sub.empty:
            raise ReportAutomationError(f"[{file_label} -> Sheet '{selected_sheet}'] No rows found for Network = '{net_filter}'.")

        sub[metric_col] = pd.to_numeric(sub[metric_col], errors="coerce")
        grouped = (
            sub.groupby(date_col, as_index=False)[metric_col]
            .agg(agg_func)
            .sort_values(by=date_col)
        )

        x_labels = [f"{d.month}/{d.day}/{d.year}" for d in grouped[date_col]]
        y_vals = grouped[metric_col].values

        metric_pretty_name = cfg["metric_titles"].get(metric_key, metric_key)
        chart_title = f"{net_filter} - {metric_pretty_name}"

        fig, ax = plt.subplots(figsize=(style["fig_width"], style["fig_height"]), dpi=style["dpi"])
        fig.patch.set_facecolor("white")
        fig.patch.set_edgecolor(style["border_color"])
        fig.patch.set_linewidth(style["border_width"])
        ax.set_facecolor("white")

        ax.plot(
            range(len(x_labels)),
            y_vals,
            color=style["line_color"],
            linewidth=style["line_width"],
            label=net_filter
        )

        # <-- REMOVES SIDE GAPS BY LOCKING X-AXIS BOUNDS FLUSH TO EDGES -->
        ax.set_xlim(0, len(x_labels) - 1)

        ax.set_title(chart_title, color=style["title_color"], fontsize=style["title_size"], fontweight="bold", pad=14)
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, -0.18),
            frameon=False,
            prop={"weight": "bold", "size": 10}
        )

        ax.set_title(chart_title, color=style["title_color"], fontsize=style["title_size"], fontweight="bold", pad=14)
        # Repositioned legend to bottom center with horizontal orientation
        ax.legend(
            loc="upper center",
            bbox_to_anchor=(0.5, -0.22),
            ncol=1,
            frameon=False,
            prop={"weight": "bold", "size": 10}
        )

        interval = style["x_tick_interval"]
        tick_positions = list(range(0, len(x_labels), interval))
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(
            [x_labels[i] for i in tick_positions],
            rotation=style["x_label_rotation"],
            ha="right",
            fontweight="bold",
            fontsize=8.5,
            color="#333333"
        )
        ax.tick_params(axis="x", which="major", direction="out", length=4, color="#888888")

        ax.tick_params(axis="y", which="both", length=0)
        if is_pct:
            ax.yaxis.set_major_locator(ticker.MaxNLocator(nbins=style["y_axis_bins"]))
            ax.yaxis.set_major_formatter(ticker.FormatStrFormatter("%.2f%%"))
            ax.set_ylim(bottom=0)
        else:
            if metric_key == "Avail_Active":
                min_v = float(np.nanmin(y_vals)) if len(y_vals) > 0 else 95.0
                y_bottom = min(95.0, np.floor(min_v * 2) / 2.0)
                ax.set_ylim(bottom=y_bottom, top=100.5)
                ax.set_yticks(np.arange(y_bottom, 100.51, 0.5))
            else:
                ax.yaxis.set_major_locator(ticker.MaxNLocator(nbins=style["y_axis_bins"], integer=True))
                ax.yaxis.set_major_formatter(ticker.StrMethodFormatter("{x:,.0f}"))

        for label in ax.get_yticklabels():
            label.set_fontweight("bold")
            label.set_fontsize(9)
            label.set_color("#444444")

        ax.yaxis.grid(True, linestyle="-", linewidth=0.7, color="#D9D9D9")
        ax.xaxis.grid(False)
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_visible(False)
        ax.spines["bottom"].set_color("#CCCCCC")

        plt.tight_layout()
        buf = io.BytesIO()
        plt.savefig(
            buf,
            format="png",
            dpi=style["dpi"],
            bbox_inches="tight",
            pad_inches=style["border_pad_inches"],
            facecolor=fig.get_facecolor(),
            edgecolor=fig.get_edgecolor()
        )
        plt.close(fig)
        buf.seek(0)
        chart_images[tag] = buf

    return chart_images


# ==========================================
# 2. DRIVE TEST METRICS EXTRACTION
# ==========================================
def extract_dt_metrics(dt_path: str, selected_sheets: dict, warnings_list: list, parent_root=None) -> dict:
    cfg = config.DT_CONFIG
    cols = cfg["columns"]
    thresh = cfg["thresholds"]
    file_label = os.path.basename(dt_path)

    xls = pd.ExcelFile(dt_path)
    for key, s_name in selected_sheets.items():
        if s_name not in xls.sheet_names:
            raise ReportAutomationError(f"[{file_label}] Selected sheet '{s_name}' not found.")

    # --- A. CDR & CBR ---
    s_cdr = selected_sheets["cdr_cbr"]
    df_cdr = pd.read_excel(xls, sheet_name=s_cdr)
    cdr_map = resolve_columns(
        df_cdr,
        {k: cols[k] for k in ["attempt", "setup", "established", "end"]},
        s_cdr, file_label, warnings_list, parent_root, "DT_CONFIG"
    )
    attempts = int(df_cdr[cdr_map["attempt"]].notna().sum())
    setups = int(df_cdr[cdr_map["setup"]].notna().sum())
    established = int(df_cdr[cdr_map["established"]].notna().sum())
    ended = int(df_cdr[cdr_map["end"]].notna().sum())

    blocks = max(0, attempts - setups)
    drops = max(0, established - ended)
    blocks_pct = (blocks / attempts * 100.0) if attempts > 0 else 0.0
    drops_pct = (drops / established * 100.0) if established > 0 else 0.0

    # --- B. 2G Idle - received level ---
    s_rx2 = selected_sheets["rx_2g"]
    df_rx2 = pd.read_excel(xls, sheet_name=s_rx2)
    rx2_map = resolve_columns(df_rx2, {"rx_2g_val": cols["rx_2g_val"]}, s_rx2, file_label, warnings_list, parent_root, "DT_CONFIG")
    rx2_series = pd.to_numeric(df_rx2[rx2_map["rx_2g_val"]], errors="coerce").dropna()
    total_2g = int(len(rx2_series))
    bad_2g = int((rx2_series < thresh["rx_min_dbm"]).sum())
    good_2g = total_2g - bad_2g
    avg_2g = rx2_series.mean() if total_2g > 0 else 0.0

    # --- C. 3G Idle - received level ---
    s_rx3 = selected_sheets["rx_3g"]
    df_rx3 = pd.read_excel(xls, sheet_name=s_rx3)
    rx3_map = resolve_columns(df_rx3, {"rx_3g_val": cols["rx_3g_val"]}, s_rx3, file_label, warnings_list, parent_root, "DT_CONFIG")
    rx3_series = pd.to_numeric(df_rx3[rx3_map["rx_3g_val"]], errors="coerce").dropna()
    total_3g = int(len(rx3_series))
    bad_3g = int((rx3_series < thresh["rx_min_dbm"]).sum())
    good_3g = total_3g - bad_3g
    avg_3g = rx3_series.mean() if total_3g > 0 else 0.0

    # --- D. Data Export (Session Block Rate %) ---
    s_data = selected_sheets["data_exp"]
    df_data = pd.read_excel(xls, sheet_name=s_data)
    data_map = resolve_columns(
        df_data,
        {k: cols[k] for k in ["http_start", "http_end", "post_start", "post_end"]},
        s_data, file_label, warnings_list, parent_root, "DT_CONFIG"
    )
    total_starts = int(df_data[data_map["http_start"]].notna().sum() + df_data[data_map["post_start"]].notna().sum())
    total_ends = int(df_data[data_map["http_end"]].notna().sum() + df_data[data_map["post_end"]].notna().sum())
    session_blocks = max(0, total_starts - total_ends)
    session_block_pct = (session_blocks / total_starts * 100.0) if total_starts > 0 else 0.0

    # --- E. Session Setup Time ---
    s_setup = selected_sheets["setup_time"]
    df_setup = pd.read_excel(xls, sheet_name=s_setup)
    setup_map = resolve_columns(df_setup, {"setup_val": cols["setup_val"]}, s_setup, file_label, warnings_list, parent_root, "DT_CONFIG")
    setup_series = pd.to_numeric(df_setup[setup_map["setup_val"]], errors="coerce").dropna()
    avg_setup = setup_series.mean() if len(setup_series) > 0 else 0.0

    # --- F. Polqa ---
    s_polqa = selected_sheets["polqa"]
    df_polqa = pd.read_excel(xls, sheet_name=s_polqa)
    polqa_map = resolve_columns(df_polqa, {"polqa_val": cols["polqa_val"]}, s_polqa, file_label, warnings_list, parent_root, "DT_CONFIG")
    polqa_series = pd.to_numeric(df_polqa[polqa_map["polqa_val"]], errors="coerce").dropna()
    total_polqa = int(len(polqa_series))
    bad_polqa = int((polqa_series < thresh["polqa_min_mos"]).sum())
    bad_polqa_pct = (bad_polqa / total_polqa * 100.0) if total_polqa > 0 else 0.0

    calc_values = {
        "attempts": f"{attempts}",
        "blocks": f"{blocks}",
        "blocks_pct": f"{blocks_pct:.2f}%",
        "drops": f"{drops}",
        "drops_pct": f"{drops_pct:.2f}%",
        "avg_2g": f"{float(avg_2g):.2f}",
        "total_2g": f"{total_2g}",
        "bad_2g": f"{bad_2g}",
        "good_2g": f"{good_2g}",
        "avg_3g": f"{float(avg_3g):.2f}",
        "total_3g": f"{total_3g}",
        "bad_3g": f"{bad_3g}",
        "good_3g": f"{good_3g}",
        "session_block_pct": f"{session_block_pct:.2f}%",
        "avg_setup": f"{float(avg_setup):.2f}",
        "bad_polqa_pct": f"{bad_polqa_pct:.2f}%"
    }

    return {
        tag: calc_values[calc_key]
        for tag, calc_key in cfg["placeholders"].items()
    }


# ==========================================
# 3. WORD DOCUMENT INJECTION
# ==========================================
def replace_in_paragraph(paragraph, text_map: dict, image_map: dict, img_width: float):
    full_text = paragraph.text
    if "{{" not in full_text:
        return

    for img_tag, img_buf in image_map.items():
        if img_tag in full_text:
            for run in paragraph.runs:
                run.text = ""
            img_buf.seek(0)
            new_run = paragraph.add_run()
            new_run.add_picture(img_buf, width=Inches(img_width))
            return

    for tag, val in text_map.items():
        if tag in paragraph.text:
            replaced_in_run = False
            for run in paragraph.runs:
                if tag in run.text:
                    run.text = run.text.replace(tag, str(val))
                    replaced_in_run = True

            if not replaced_in_run and tag in paragraph.text:
                updated_text = paragraph.text.replace(tag, str(val))
                if paragraph.runs:
                    paragraph.runs[0].text = updated_text
                    for r in paragraph.runs[1:]:
                        r.text = ""


def populate_word_template(word_path: str, output_path: str, text_map: dict, image_map: dict):
    doc = Document(word_path)
    img_width = config.CHART_STYLE["word_image_width_inches"]

    for p in doc.paragraphs:
        replace_in_paragraph(p, text_map, image_map, img_width)

    def process_tables(tables):
        for table in tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        replace_in_paragraph(p, text_map, image_map, img_width)
                    if cell.tables:
                        process_tables(cell.tables)

    process_tables(doc.tables)
    doc.save(output_path)


# ==========================================
# 4. DESKTOP GUI APPLICATION
# ==========================================
class TelecomReportGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Telecom Performance Report Generator")
        self.root.geometry("780x770")
        self.root.resizable(False, False)

        self.weekly_var = tk.StringVar()
        self.dt_var = tk.StringVar()
        self.word_var = tk.StringVar()

        self.weekly_sheet_var = tk.StringVar(value=config.WEEKLY_CONFIG["sheet_name"])
        self.dt_sheet_vars = {
            key: tk.StringVar(value=default_name)
            for key, default_name in config.DT_CONFIG["sheets"].items()
        }

        pad = {"padx": 8, "pady": 4}

        file_frame = tk.LabelFrame(root, text="Step 1: Select Input & Template Files", font=("Segoe UI", 10, "bold"), padx=10, pady=6)
        file_frame.pack(fill="x", padx=14, pady=(10, 4))

        tk.Label(file_frame, text="Weekly OSS Report (.xlsx):", font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", **pad)
        tk.Entry(file_frame, textvariable=self.weekly_var, width=54).grid(row=0, column=1, **pad)
        tk.Button(file_frame, text="Browse...", command=self.browse_weekly).grid(row=0, column=2, **pad)

        tk.Label(file_frame, text="Drive Test Report (.xlsx):", font=("Segoe UI", 9, "bold")).grid(row=1, column=0, sticky="w", **pad)
        tk.Entry(file_frame, textvariable=self.dt_var, width=54).grid(row=1, column=1, **pad)
        tk.Button(file_frame, text="Browse...", command=self.browse_dt).grid(row=1, column=2, **pad)

        tk.Label(file_frame, text="Word Template (.docx):", font=("Segoe UI", 9, "bold")).grid(row=2, column=0, sticky="w", **pad)
        tk.Entry(file_frame, textvariable=self.word_var, width=54).grid(row=2, column=1, **pad)
        tk.Button(file_frame, text="Browse...", command=self.browse_word).grid(row=2, column=2, **pad)

        sheet_frame = tk.LabelFrame(root, text="Step 2: Target Excel Sheets", font=("Segoe UI", 10, "bold"), padx=10, pady=6)
        sheet_frame.pack(fill="x", padx=14, pady=4)

        tk.Label(sheet_frame, text="Weekly OSS (12 Charts):", font=("Segoe UI", 9, "bold"), fg="#0F5C7E").grid(row=0, column=0, sticky="w", **pad)
        self.weekly_combo = ttk.Combobox(sheet_frame, textvariable=self.weekly_sheet_var, width=44, state="readonly")
        self.weekly_combo.grid(row=0, column=1, sticky="w", **pad)

        dt_labels = [
            ("cdr_cbr",    "DT - CDR & CBR Sheet:"),
            ("rx_2g",      "DT - 2G Rx Level Sheet:"),
            ("rx_3g",      "DT - 3G Rx Level Sheet:"),
            ("data_exp",   "DT - Data Export Sheet:"),
            ("setup_time", "DT - Session Setup Time Sheet:"),
            ("polqa",      "DT - POLQA Voice Sheet:")
        ]

        self.dt_combos = {}
        for idx, (key, label_text) in enumerate(dt_labels, start=1):
            tk.Label(sheet_frame, text=label_text, font=("Segoe UI", 9)).grid(row=idx, column=0, sticky="w", **pad)
            cb = ttk.Combobox(sheet_frame, textvariable=self.dt_sheet_vars[key], width=44, state="readonly")
            cb.grid(row=idx, column=1, sticky="w", **pad)
            self.dt_combos[key] = cb

        preview_frame = tk.LabelFrame(root, text="Step 3: Live Extracted Metrics & Column Match Warnings", font=("Segoe UI", 10, "bold"), padx=10, pady=6)
        preview_frame.pack(fill="x", padx=14, pady=4)

        self.preview_text = tk.Text(preview_frame, height=7, width=90, font=("Consolas", 9), bg="#F8F9FA", fg="#333333")
        self.preview_text.pack(fill="x", padx=2, pady=2)
        self.preview_text.insert(tk.END, "Browse a Drive Test report above to live-preview extracted KPI values here...")
        self.preview_text.config(state="disabled")

        self.status_label = tk.Label(root, text="Ready. Select files to begin.", fg="#555555", font=("Segoe UI", 9))
        self.status_label.pack(pady=(4, 2))

        self.run_btn = tk.Button(
            root,
            text="Generate Updated Word Report",
            command=self.run_pipeline,
            bg="#0F5C7E",
            fg="white",
            font=("Segoe UI", 11, "bold"),
            padx=20,
            pady=6
        )
        self.run_btn.pack(pady=(2, 10))

    def browse_weekly(self):
        path = filedialog.askopenfilename(title="Select Weekly OSS Excel File", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            xls = pd.ExcelFile(path)
            sheets = xls.sheet_names
            self.weekly_var.set(path)
            self.weekly_combo["values"] = sheets
            matched = match_default_sheet(sheets, config.WEEKLY_CONFIG["sheet_name"])
            self.weekly_sheet_var.set(matched)
        except Exception as e:
            messagebox.showerror("File Read Error", f"Could not read sheets:\n{e}")

    def browse_dt(self):
        path = filedialog.askopenfilename(title="Select Drive Test Excel File", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not path:
            return
        try:
            xls = pd.ExcelFile(path)
            sheets = xls.sheet_names
            self.dt_var.set(path)
            for key, cb in self.dt_combos.items():
                cb["values"] = sheets
                default_target = config.DT_CONFIG["sheets"][key]
                matched = match_default_sheet(sheets, default_target)
                self.dt_sheet_vars[key].set(matched)

            selected_sheets = {k: v.get() for k, v in self.dt_sheet_vars.items()}
            warnings_list = []
            metrics = extract_dt_metrics(path, selected_sheets, warnings_list, parent_root=self.root)

            preview_str = (
                f"• Call Attempts: {metrics['{{DT_ATTEMPTS}}']}   |   Blocks: {metrics['{{DT_BLOCKS}}']} ({metrics['{{DT_BLOCKS_%}}']})   |   Drops: {metrics['{{DT_DROPS}}']} ({metrics['{{DT_DROPS_%}}']})\n"
                f"• 2G Signal Avg: {metrics['{{DT_SIGNAL_AVG_2G}}']} dBm (Total: {metrics['{{DT_TOTAL_SAMPLES_2G}}']}, Bad < -90dBm: {metrics['{{DT_BAD_SAMPLES_2G}}']})\n"
                f"• 3G Signal Avg: {metrics['{{DT_SIGNAL_AVG_3G}}']} dBm (Total: {metrics['{{DT_TOTAL_SAMPLES_3G}}']}, Bad < -90dBm: {metrics['{{DT_BAD_SAMPLES_3G}}']})\n"
                f"• Data Session Block Rate: {metrics['{{DT_SESSION_BLOCK_%}}']}   |   Session Setup Avg: {metrics['{{DT_SESSION_SETUP_TIME_AVG}}']}s\n"
                f"• POLQA Voice Samples < 2.8 MOS: {metrics['{{DT_VOICE_SAMPLES_BELOW_2.8_%}}']}"
            )
            if warnings_list:
                unique_warnings = list(dict.fromkeys(warnings_list))
                preview_str += "\n\n⚠ COLUMN FALLBACK NOTICES:\n" + "\n".join(f"  - {w}" for w in unique_warnings)

            self.preview_text.config(state="normal")
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert(tk.END, preview_str)
            self.preview_text.config(state="disabled")

            if warnings_list:
                self.status_label.config(text=f"Loaded with {len(set(warnings_list))} column fallback notice(s).", fg="#B25900")
            else:
                self.status_label.config(text="Drive Test report loaded and live-previewed successfully.", fg="#0F5C7E")

        except ReportAutomationError as e:
            messagebox.showerror("Data / Configuration Error", str(e))
        except Exception as e:
            messagebox.showerror("File Read Error", f"Could not read Drive Test sheets:\n{e}")

    def browse_word(self):
        path = filedialog.askopenfilename(title="Select Word Template File", filetypes=[("Word Documents", "*.docx")])
        if path:
            self.word_var.set(path)

    def run_pipeline(self):
        weekly_path = self.weekly_var.get().strip()
        dt_path = self.dt_var.get().strip()
        word_path = self.word_var.get().strip()

        if not all([weekly_path, dt_path, word_path]):
            messagebox.showwarning("Missing Files", "Please select all three files before running.")
            return

        selected_weekly_sheet = self.weekly_sheet_var.get()
        selected_dt_sheets = {k: v.get() for k, v in self.dt_sheet_vars.items()}

        base_dir = os.path.dirname(word_path)
        base_name = os.path.splitext(os.path.basename(word_path))[0]
        output_path = os.path.join(base_dir, f"{base_name}_GENERATED.docx")

        warnings_list = []

        try:
            self.status_label.config(text="Generating 12 styled OSS charts...", fg="#0F5C7E")
            self.root.update_idletasks()
            image_map = generate_weekly_charts(weekly_path, selected_weekly_sheet, warnings_list, parent_root=self.root)

            self.status_label.config(text="Extracting Drive Test metrics...", fg="#0F5C7E")
            self.root.update_idletasks()
            text_map = extract_dt_metrics(dt_path, selected_dt_sheets, warnings_list, parent_root=self.root)

            self.status_label.config(text="Populating Word template...", fg="#0F5C7E")
            self.root.update_idletasks()
            populate_word_template(word_path, output_path, text_map, image_map)

            unique_warnings = list(dict.fromkeys(warnings_list))
            warn_on_fallback = getattr(config, "COLUMN_MATCHING", {}).get("warn_on_fallback", True)

            if unique_warnings and warn_on_fallback:
                self.status_label.config(text="Completed with column fallback notices.", fg="#B25900")
                warn_details = "\n".join(f"• {w}" for w in unique_warnings)
                messagebox.showwarning(
                    "Report Complete (With Column Match Notices)",
                    f"Report generated and saved to:\n{output_path}\n\n"
                    f"Fallback column matching was used for {len(unique_warnings)} field(s):\n\n{warn_details}"
                )
            else:
                self.status_label.config(text="Success! Report generated.", fg="green")
                messagebox.showinfo("Report Complete", f"Report successfully generated and saved to:\n\n{output_path}")

        except ReportAutomationError as e:
            self.status_label.config(text="Failed: Data validation error.", fg="red")
            messagebox.showerror("Data / Configuration Error", str(e))
        except Exception as e:
            self.status_label.config(text="Failed: Unexpected error.", fg="red")
            messagebox.showerror("Unexpected Error", f"{str(e)}\n\nDetails:\n{traceback.format_exc()}")


if __name__ == "__main__":
    root = tk.Tk()
    app = TelecomReportGUI(root)
    root.mainloop()
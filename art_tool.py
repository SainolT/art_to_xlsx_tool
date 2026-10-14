# -*- coding: utf-8 -*-
"""美术资源配置工具（窗口版）

扫描美术资源库（一个根目录，递归收录全部文件类型），按名称拼音
（全拼/首字母/包含）批量匹配配置表中的名称字段，把资源标识写入资源字段。
写入格式由用户每次操作选择：
  · 文件名（不含扩展名），如 icon_hb_shuidi
  · 相对路径+文件名（/ 分隔，不含扩展名），如 chashu/shengzhang_texiao
    路径起始点可选资源库的某个父目录（此时路径包含资源库自身目录名，
    如 资源库在 D:\\art\\图标库，起始点选 D:\\art，写入 图标库/xxx）
也提供已填资源值的存在性校验（只读，可选 按文件名 / 按相对路径 两种校验方式）。

防护：先预览后执行；写入前自动备份；原子保存；表头永不修改；
公式单元格不读不写；重名/多匹配不自动写，列出由人工选定。

运行: python art_tool.py
"""
import json
import os
import re
import shutil
import sys
import tkinter as tk
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath
from tkinter import ttk, scrolledtext, messagebox, filedialog
import xml.etree.ElementTree as ET

import openpyxl
from openpyxl.utils import get_column_letter

try:
    from pypinyin import pinyin, Style
    try:
        from pypinyin_dict.pinyin_data import kxhc1983
        kxhc1983.load()  # 换成《现代汉语词典》读音库，过滤生僻读音（否则 柠→chǔ/zhù 这类会误配）
    except ImportError:
        pass
except ImportError:  # 没装 pypinyin 时仍可启动，中文名只按原名匹配
    pinyin = None
    Style = None

CONFIG_FILE = "工具配置.json"
CACHE_FILE = "资源扫描缓存.json"
BACKUP_DIR = "资源写入备份"

FMT_NAME = "name"  # 写入/校验：文件名（不含扩展名）
FMT_PATH = "path"  # 写入/校验：相对路径+文件名（不含扩展名）
FMT_NAMES = {FMT_NAME: "文件名", FMT_PATH: "相对路径+文件名"}

ROOT_LABEL = "（资源库根目录）"


# ---------- 配置 ----------

def app_dir():
    """脚本/exe 自身所在目录。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def load_config(base_dir):
    """读取工具配置，没有或损坏时返回空 dict。"""
    cfg = Path(base_dir) / CONFIG_FILE
    if not cfg.exists():
        return {}
    try:
        return json.loads(cfg.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_config(base_dir, **kw):
    """更新工具配置（只覆盖传入的键）。"""
    data = load_config(base_dir)
    data.update(kw)
    (Path(base_dir) / CONFIG_FILE).write_text(
        json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _config_path(base_dir, value):
    """配置里的路径: 相对路径相对 app 目录解析；不存在时返回 None。"""
    if not value:
        return None
    p = Path(value)
    if not p.is_absolute():
        p = Path(base_dir) / p
    return p


def get_config_dir(base_dir, key):
    """读取配置里的目录项，存在才返回 Path，否则 None。"""
    p = _config_path(base_dir, load_config(base_dir).get(key, ""))
    return p if p and p.exists() else None


def save_config_dir(base_dir, key, path):
    try:
        s = str(Path(path).resolve().relative_to(Path(base_dir).resolve()))
    except ValueError:
        s = str(path)
    save_config(base_dir, **{key: s})


# ---------- 资源库扫描 ----------

def norm_stem(name):
    """文件名词干规范化：去扩展名、小写、去 _ - 和空白，用于索引键。"""
    stem = Path(str(name)).stem
    return re.sub(r"[_\-\s]+", "", stem.lower())


def scan_resources(root, cache_dir=None):
    """递归扫描资源根目录，收录全部文件类型。
    返回 {规范化词干: [相对路径(POSIX), ...]}（同名多路径全部保留）。
    cache_dir 提供时按 路径+mtime+size 缓存，根目录无变化则秒出。"""
    root = Path(root)
    cache_path = Path(cache_dir) / CACHE_FILE if cache_dir else None
    cache = {}
    if cache_path and cache_path.exists():
        try:
            cache = json.loads(cache_path.read_text(encoding="utf-8"))
        except Exception:
            cache = {}
    key = str(root.resolve())
    new_cache = dict(cache)
    try:
        st = root.stat()
        ent = cache.get(key)
        if ent and ent.get("mtime") == st.st_mtime:
            return {k: list(v) for k, v in ent["index"].items()}
    except OSError:
        return {}
    index = {}
    for p in sorted(root.rglob("*")):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        rel = p.relative_to(root).as_posix()
        index.setdefault(norm_stem(p.name), []).append(rel)
    new_cache[key] = {"mtime": st.st_mtime, "index": index}
    if cache_path:
        try:
            cache_path.write_text(json.dumps(new_cache, ensure_ascii=False),
                                  encoding="utf-8")
        except Exception:
            pass
    return {k: list(v) for k, v in index.items()}


def invalidate_cache(cache_dir, root):
    """根目录内容可能已变化时调用：删除该根目录的缓存条目。"""
    cache_path = Path(cache_dir) / CACHE_FILE
    if not cache_path.exists():
        return
    try:
        cache = json.loads(cache_path.read_text(encoding="utf-8"))
    except Exception:
        return
    cache.pop(str(Path(root).resolve()), None)
    try:
        cache_path.write_text(json.dumps(cache, ensure_ascii=False), encoding="utf-8")
    except Exception:
        pass


def filter_index(index, exts):
    """按扩展名过滤资源索引。exts 为 None（全部）时原样返回。
    无扩展名的文件用 '' 表示。"""
    if exts is None:
        return index
    out = {}
    for k, rels in index.items():
        keep = [rel for rel in rels
                if PurePosixPath(rel).suffix.lower() in exts]
        if keep:
            out[k] = keep
    return out


def scope_index(index, prefixes=(), suffixes=()):
    """按前后缀锁定资源范围：只有文件名（词干，小写）符合任一限定前缀
    且符合任一限定后缀的资源才参与。前缀/后缀为空 = 该方向不限。"""
    if not prefixes and not suffixes:
        return index
    out = {}
    for k, rels in index.items():
        keep = []
        for rel in rels:
            stem = PurePosixPath(rel).stem.lower()
            if prefixes and not any(stem.startswith(p) for p in prefixes):
                continue
            if suffixes and not any(stem.endswith(s) for s in suffixes):
                continue
            keep.append(rel)
        if keep:
            out[k] = keep
    return out


def exts_of(index):
    """资源索引中出现的所有扩展名（小写，无扩展名记为 ''），排序返回。"""
    exts = {PurePosixPath(rel).suffix.lower()
            for rels in index.values() for rel in rels}
    return sorted(exts)


# ---------- 名称与匹配 ----------

def name_variants(name):
    """名称 → 匹配候选集：拼音全拼连写、首字母连写（多音字全读音组合）、
    原名小写（去空白）。纯拉丁名只有原名候选。返回去重列表（保持顺序）。"""
    s = str(name).strip()
    if not s:
        return []
    out = [re.sub(r"\s+", "", s.lower())]
    if pinyin and re.search(r"[一-鿿]", s):
        for style in (Style.NORMAL, Style.FIRST_LETTER):
            readings = pinyin(s, style=style, heteronym=True)
            # 全读音组合（多音字笛卡尔积），限制组合数防爆炸
            combos = [""]
            for rd in readings:
                combos = [c + r for c in combos for r in rd][:64]
            for v in combos:
                v = re.sub(r"\s+", "", v.lower())
                if v and v not in out:
                    out.append(v)
    return out


def parse_patterns(text):
    """解析忽略前后缀输入：逗号/空格/分号分隔，去空去重。"""
    parts = re.split(r"[,，;；\s]+", str(text))
    seen, out = set(), []
    for p in parts:
        p = p.strip().lower()
        if p and p not in seen:
            seen.add(p)
            out.append(p)
    return out


def tokenize(stem):
    """词干拆词元：按 _ - 空白 分隔，纯数字段丢弃，末位词元剥掉尾部数字。
    如 V1_zzcy_d_cy_fqcd1 -> [zzcy, d, cy, fqcd]；icon_hb_shuidi -> [icon, hb, shuidi]。"""
    toks = [t for t in re.split(r"[_\-\s]+", stem.lower()) if t]
    toks = [t for t in toks if not t.isdigit()]
    if toks:
        toks[-1] = re.sub(r"\d+$", "", toks[-1]) or toks[-1]
    return [t for t in toks if t]


def _tier_of(rel, variants):
    """判断单个资源对候选集的命中层级。
    返回 (层级, 打分)；层级 1=整名相等 2=连续词元段相等 3=词元子串；未命中返回 None。
    打分用于同级内排序（越小越好）：词元拼接后与候选的长度差。"""
    stem = Path(rel).stem.lower()
    norm = re.sub(r"[_\-\s]+", "", stem)
    for v in variants:
        if norm == v:
            return (1, 0)
    toks = tokenize(stem)
    for v in variants:
        joined = "".join(toks)
        if joined == v:
            return (2, 0)
        # 连续词元段相等
        for i in range(len(toks)):
            acc = ""
            for j in range(i, len(toks)):
                acc += toks[j]
                if acc == v:
                    return (2, abs(len(joined) - len(v)))
                if len(acc) > len(v):
                    break
    for v in variants:
        for t in toks:
            if (len(v) >= 3 and v in t) or (len(t) >= 4 and t in v):
                return (3, abs(len(t) - len(v)))
    return None


def match_resource(name, index):
    """按名称在资源索引中匹配。
    返回 (status, candidates)：
      ok       candidates=[唯一命中的相对路径]（仅整名相等/连续词元段相等）
      conflict candidates=[候选相对路径...]（并列最优、重名、或第3级子串疑似命中）
      missing  candidates=[]
    只在最高命中层级内取最优；并列不猜；第3级（子串包含，如 hycm ⊂ hycmjh）
    容易张冠李戴（红颜草莓精华 撞上 红颜草莓），一律交人工选定。
    资源范围可先由 scope_index 限定。"""
    variants = name_variants(name)
    if not variants:
        return ("missing", [])
    hits = {}  # rel -> (tier, score)
    for norm_name, rels in index.items():
        for rel in rels:
            r = _tier_of(rel, variants)
            if r is None:
                continue
            if rel not in hits or r < hits[rel]:
                hits[rel] = r
    if not hits:
        return ("missing", [])
    best_tier = min(t for t, _ in hits.values())
    best_score = min(s for t, s in hits.values() if t == best_tier)
    best = sorted(rel for rel, (t, s) in hits.items()
                  if t == best_tier and s == best_score)
    if len(best) == 1 and best_tier <= 2:
        return ("ok", best)
    return ("conflict", best)


def split_dup_groups(changes, plan, fmt, keep_ext, prefix):
    """把多名称命中同一资源的改动分组并分类。
    changes 元素为 (行号, 名称, ..., 资源)（资源在末位）。
    返回 (groups, preselect, outside_occ)：
      groups:      {资源: [名称...]} 需要人工决定的重复组
      preselect:   {资源: 名称} 资源当前已被组内某名称的行填入（默认保留该名称）
      outside_occ: {资源} 资源当前已被组外名称的行填入（交占用检测，不算重复待定）
    """
    by_rel = {}
    for ch in changes:
        by_rel.setdefault(ch[-1], set()).add(ch[1])
    name_by_row = {}
    for r, n in plan["same"]:
        name_by_row[r] = n
    for entry in plan["ok"] + plan["conflict"]:
        name_by_row.setdefault(entry[0], entry[1])
    for r, n in plan["missing"]:
        name_by_row.setdefault(r, n)
    old_values = plan.get("old_values", {})
    groups, preselect, outside_occ = {}, {}, set()
    for rel, names in by_rel.items():
        if len(names) < 2:
            continue
        val = norm(format_value(rel, fmt, keep_ext, prefix))
        owners = {name_by_row[r] for r, old in old_values.items()
                  if r in name_by_row and norm(old) == val}
        if owners - names:
            outside_occ.add(rel)  # 组外已有人占用
            continue
        groups[rel] = sorted(names)
        inside = owners & names
        if len(inside) == 1:
            preselect[rel] = next(iter(inside))
    return groups, preselect, outside_occ


def preview_changes(plan, resolutions, dup_decisions, fmt, keep_ext, prefix):
    """把匹配计划和人工决定合成改动清单（不写盘，供全部改动窗口展示）。
    返回 [(行号, 名称, 旧值, 新值, 资源相对路径)]，按行号排序。
    重复资源未决定的组不出现在清单里（反正不会写入）。"""
    old = plan.get("old_values", {})
    changes = [(r, n, old.get(r, ""), v, rel) for r, n, v, rel in plan["ok"]]
    for r, name, cands in plan["conflict"]:
        pick = resolutions.get(r)
        if pick:
            changes.append((r, name, old.get(r, ""),
                            format_value(pick, fmt, keep_ext, prefix), pick))
    for r, name in plan["missing"]:
        pick = resolutions.get(r)
        if pick:
            changes.append((r, name, old.get(r, ""),
                            format_value(pick, fmt, keep_ext, prefix), pick))
    groups, preselect, _ = split_dup_groups(changes, plan, fmt, keep_ext, prefix)
    out = []
    for ch in changes:
        rel = ch[-1]
        if rel in groups:
            keep = (dup_decisions or {}).get(rel) or preselect.get(rel)
            if not keep or ch[1] != keep:
                continue
        out.append(ch)
    return sorted(out)


def format_value(rel, fmt, keep_ext=False, prefix=""):
    """把资源相对路径格式化为写入值。
    文件名: 词干（默认不含扩展名）。
    相对路径: 相对路径/词干；prefix（路径起始点到资源库根目录的相对路径）
    非空时加在最前面，如 prefix=图标库、rel=chashu/x.json -> 图标库/chashu/x。"""
    p = PurePosixPath(rel)
    name = p.name if keep_ext else p.stem
    if fmt == FMT_NAME:
        return name
    parent = p.parent.as_posix()
    parts = []
    if prefix:
        parts.append(prefix.strip("/"))
    if parent != ".":
        parts.append(parent)
    parts.append(name)
    return "/".join(parts)


# ---------- 配置表读写 ----------

OOXML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OOXML_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
OOXML_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"

for _prefix, _ns in [
    ("", OOXML_NS), ("r", OOXML_REL_NS),
    ("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006"),
    ("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"),
    ("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"),
    ("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2"),
    ("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3"),
]:
    try:
        ET.register_namespace(_prefix, _ns)
    except ValueError:
        pass


def _sheet_xml_path(wb_xml, rels_xml, sheet):
    """工作表名 -> xl/worksheets/sheetN.xml 路径。"""
    wb = ET.fromstring(wb_xml)
    rid = None
    for sh in wb.iter(f"{{{OOXML_NS}}}sheet"):
        if sh.get("name") == sheet:
            rid = sh.get(f"{{{OOXML_REL_NS}}}id")
            break
    if rid is None:
        raise KeyError(f"工作表「{sheet}」不存在")
    rels = ET.fromstring(rels_xml)
    for rel in rels.iter(f"{{{OOXML_PKG_REL_NS}}}Relationship"):
        if rel.get("Id") == rid:
            target = rel.get("Target")
            return target.lstrip("/") if target.startswith("/") else "xl/" + target
    raise KeyError(f"工作表「{sheet}」的关系目标找不到")


def _si_text(si):
    """sharedStrings 一个 si 项的纯文本（富文本取所有 t 拼接）。"""
    return "".join(t.text or "" for t in si.iter(f"{{{OOXML_NS}}}t"))


def _col_idx_of(ref):
    """单元格引用（如 BC12）-> 列号（1-based）。"""
    letters = re.match(r"[A-Z]+", ref).group(0)
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n


def write_cells_surgical(path, sheet, changes, backup_dir, stamp):
    """外科式修改 xlsx：只改目标工作表里目标单元格的值，zip 内其余文件
    （其他工作表、公式缓存值、样式、主题等）字节级原样保留——整表重写会
    丢失公式缓存导致转表失败，这里不做整表重写。
    changes: [(行号, 列号, 新字符串值)]。返回备份文件名（未备份返回 None）。"""
    tp = Path(path)
    with zipfile.ZipFile(tp) as zin:
        infos = zin.infolist()
        data = {i.filename: zin.read(i) for i in infos}

    sheet_xml = _sheet_xml_path(data["xl/workbook.xml"],
                                data["xl/_rels/workbook.xml.rels"], sheet)

    # 共享字符串表：已有值复用索引，新值追加；没有 sharedStrings 则用内联字符串
    ss_xml = "xl/sharedStrings.xml"
    ss_root = None
    ss_index = {}
    if ss_xml in data:
        ss_root = ET.fromstring(data[ss_xml])
        for si in ss_root:
            txt = _si_text(si)
            if txt not in ss_index:
                ss_index[txt] = len(ss_index)
    n_new_strings = 0

    root = ET.fromstring(data[sheet_xml])
    sheet_data = root.find(f"{{{OOXML_NS}}}sheetData")
    rows = {int(r.get("r")): r for r in sheet_data.findall(f"{{{OOXML_NS}}}row")}

    for r, c, val in changes:
        val = str(val)
        ref = f"{get_column_letter(c)}{r}"
        row_el = rows.get(r)
        if row_el is None:  # 行不存在：按行号顺序插入
            row_el = ET.Element(f"{{{OOXML_NS}}}row", {"r": str(r)})
            pos = len(sheet_data)
            for i, other in enumerate(sheet_data):
                if int(other.get("r")) > r:
                    pos = i
                    break
            sheet_data.insert(pos, row_el)
            rows[r] = row_el
        cell = None
        cells = row_el.findall(f"{{{OOXML_NS}}}c")
        for ce in cells:
            if ce.get("r") == ref:
                cell = ce
                break
        if cell is None:  # 单元格不存在：按列序插入
            cell = ET.Element(f"{{{OOXML_NS}}}c", {"r": ref})
            pos = len(cells)
            for i, ce in enumerate(cells):
                if _col_idx_of(ce.get("r")) > c:
                    pos = i
                    break
            row_el.insert(pos, cell)
        for child in list(cell):  # 清掉旧的 v/is/f（公式行上游已跳过）
            cell.remove(child)
        if ss_root is not None:
            if val not in ss_index:
                si = ET.SubElement(ss_root, f"{{{OOXML_NS}}}si")
                te = ET.SubElement(si, f"{{{OOXML_NS}}}t")
                te.text = val
                if val != val.strip():
                    te.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                ss_index[val] = len(ss_index)
                n_new_strings += 1
            cell.set("t", "s")
            ET.SubElement(cell, f"{{{OOXML_NS}}}v").text = str(ss_index[val])
        else:
            cell.set("t", "inlineStr")
            is_el = ET.SubElement(cell, f"{{{OOXML_NS}}}is")
            te = ET.SubElement(is_el, f"{{{OOXML_NS}}}t")
            te.text = val
            if val != val.strip():
                te.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")

    data[sheet_xml] = ET.tostring(root, encoding="UTF-8", xml_declaration=True)
    if ss_root is not None and n_new_strings:
        ss_root.set("count", str(int(ss_root.get("count", len(ss_index))) + len(changes)))
        ss_root.set("uniqueCount", str(len(ss_index)))
        data[ss_xml] = ET.tostring(ss_root, encoding="UTF-8", xml_declaration=True)

    bak_name = None
    if backup_dir:
        bdir = Path(backup_dir)
        bdir.mkdir(parents=True, exist_ok=True)
        bak_name = f"{tp.stem}.{stamp}{tp.suffix}"
        shutil.copy2(tp, bdir / bak_name)  # 此刻磁盘上仍是原始文件
    tmp = tp.with_name(tp.name + ".tmp")
    with zipfile.ZipFile(tmp, "w") as zout:
        for info in infos:  # 保持原有条目顺序和压缩方式
            zout.writestr(info, data[info.filename])
    os.replace(tmp, tp)
    return bak_name

def norm(v):
    """把单元格值归一化成可比较的字符串（1 / 1.0 / '1' 视为相同）。"""
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return str(v).strip()


def is_formula_like(v):
    """判断值是否形如公式（以 = 开头的文本，写入后会变成公式，有注入风险）。"""
    return isinstance(v, str) and v.startswith("=")


def is_formula_cell(cell):
    """判断单元格是否是公式: 覆盖普通公式、数组公式和以 = 开头的文本。"""
    return cell.data_type == "f" or is_formula_like(cell.value)


def header_map_of(ws):
    """读取已打开工作表的第 1 行表头，返回 {列名: 列号(1-based)}。"""
    hmap = {}
    for c in ws[1]:
        h = norm(c.value)
        if h and h not in hmap:
            hmap[h] = c.column
    return hmap


def valid_headers(ws):
    hdr = list(next(ws.iter_rows(max_row=1, values_only=True), None) or [])
    while hdr and hdr[-1] is None:
        hdr.pop()
    if not hdr or not all(isinstance(h, str) and h.strip() for h in hdr):
        return None
    if len(set(hdr)) != len(hdr):
        return None
    return [h.strip() for h in hdr]


def load_table_sheets(path):
    """读取 xlsx 的可用子表（表头合规的），返回 [(子表名, [表头...])]。"""
    wb = openpyxl.load_workbook(path, read_only=True)
    sheets = []
    for ws in wb.worksheets:
        hdr = valid_headers(ws)
        if hdr:
            sheets.append((ws.title, hdr))
    wb.close()
    return sheets


def read_sheet(path, sheet):
    """打开目标表格，返回 (workbook, worksheet, {列名: 列号})。"""
    wb = openpyxl.load_workbook(path)
    if sheet not in wb.sheetnames:
        raise KeyError(f"工作表「{sheet}」不存在于 {Path(path).name}")
    ws = wb[sheet]
    return wb, ws, header_map_of(ws)


# ---------- 批量匹配 ----------

def plan_match(path, sheet, name_col, value_col, index, fmt,
               scope_prefixes=(), scope_suffixes=(), keep_ext=False, prefix=""):
    """扫描目标子表，为每行生成匹配计划（不写盘）。
    scope_prefixes/scope_suffixes 限定参与匹配的资源范围（文件名前后缀）。
    返回 dict:
      ok:         [(行号, 名称, 新值, 资源相对路径)]   自动匹配（新填入+覆盖）
      overwrite:  [(行号, 名称, 旧值, 新值)]           已有配置值将被覆盖的行
      same:       [(行号, 名称)]                        值已一致，无需改动
      conflict:   [(行号, 名称, [候选相对路径...])]
      missing:    [(行号, 名称)]
      skipped:    {empty: n, formula: n}                空名称/目标公式 统计
      scoped_out: [相对路径...]                         被限定前后缀排除的资源
      hit_rels:   {相对路径...}                         被命中（含冲突候选）的资源
      old_values: {行号: 旧值}                          每个处理行的当前值
    """
    scoped = scope_index(index, scope_prefixes, scope_suffixes)
    if scope_prefixes or scope_suffixes:
        all_rels = {rel for rels in index.values() for rel in rels}
        in_rels = {rel for rels in scoped.values() for rel in rels}
        scoped_out = sorted(all_rels - in_rels)
    else:
        scoped_out = []
    wb, ws, hdr = read_sheet(path, sheet)
    try:
        if name_col not in hdr or value_col not in hdr:
            raise KeyError(f"表头里找不到列「{name_col}」或「{value_col}」")
        nc, vc = hdr[name_col], hdr[value_col]
        plan = {"ok": [], "same": [], "conflict": [], "missing": [], "overwrite": [],
                "skipped": {"empty": 0, "formula": 0}, "old_values": {},
                "scoped_out": scoped_out, "hit_rels": set()}
        for r in range(2, ws.max_row + 1):
            name = norm(ws.cell(r, nc).value)
            if not name:
                plan["skipped"]["empty"] += 1
                continue
            cell = ws.cell(r, vc)
            if is_formula_cell(cell):
                plan["skipped"]["formula"] += 1
                continue
            plan["old_values"][r] = norm(cell.value)
            status, cands = match_resource(name, scoped)
            plan["hit_rels"].update(cands)
            if status == "ok":
                val = format_value(cands[0], fmt, keep_ext, prefix)
                if norm(cell.value) == norm(val):
                    plan["same"].append((r, name))
                else:
                    plan["ok"].append((r, name, val, cands[0]))
                    if norm(cell.value):
                        plan["overwrite"].append((r, name, norm(cell.value), val))
            elif status == "conflict":
                # 候选的写入值全部相同时（如同名文件分处不同目录）不算冲突
                vals = {format_value(c, fmt, keep_ext, prefix) for c in cands}
                if len(vals) == 1:
                    val = vals.pop()
                    if norm(cell.value) == norm(val):
                        plan["same"].append((r, name))
                    else:
                        plan["ok"].append((r, name, val, cands[0]))
                        if norm(cell.value):
                            plan["overwrite"].append((r, name, norm(cell.value), val))
                else:
                    plan["conflict"].append((r, name, cands))
            else:
                plan["missing"].append((r, name))
        return plan
    finally:
        wb.close()


def find_in_use(path, sheet, name_col, value_col, changes):
    """占用检测：按最终状态（当前值 + 本轮改动）计算，返回新值被
    不同名称的其他行占用（资源重复分配）的改动列表：
    [(行号, 名称, 新值, 资源相对路径, [占用方描述...])]。
    必须在所有人工决定（重复组/排除/指定）应用之后调用。
    changes 元素为 (行号, 名称, ..., 新值, 资源)（4 或 5 元组均可）。"""
    if not changes:
        return []
    wb, ws, hdr = read_sheet(path, sheet)
    try:
        nc, vc = hdr[name_col], hdr[value_col]
        final_vals, names = {}, {}
        for r in range(2, ws.max_row + 1):
            final_vals[r] = norm(ws.cell(r, vc).value)
            names[r] = norm(ws.cell(r, nc).value)
        for ch in changes:
            final_vals[ch[0]] = norm(ch[-2])
        out = []
        for ch in changes:
            r, name, val, rel = ch[0], ch[1], ch[-2], ch[-1]
            nv = norm(val)
            holders = [f"行{x}「{names.get(x)}」" for x, v in final_vals.items()
                       if x != r and v == nv and names.get(x) != name]
            if holders:
                out.append((r, name, val, rel, holders))
        return out
    finally:
        wb.close()


def run_match(path, sheet, name_col, value_col, index, fmt,
              scope_prefixes=(), scope_suffixes=(), keep_ext=False, prefix="",
              resolutions=None, dup_decisions=None, exclude_rows=None,
              include_rows=None, backup_dir=None, dry=True, log=print):
    """执行批量匹配写入（dry=True 只预览不写盘）。
    resolutions: {行号: 选定的资源相对路径}，冲突/未匹配行的人工指定结果。
    dup_decisions: {资源相对路径: 保留的名称}，"" 表示该资源的行都不写；
    未决定的重复组一律不写入。
    exclude_rows: {行号}，「全部改动」窗口里被排除的行。
    include_rows: {行号}，「全部改动」窗口里被人工恢复的占用行。
    返回改动单元格数。日志按统一模板输出。"""
    resolutions = resolutions or {}
    plan = plan_match(path, sheet, name_col, value_col, index, fmt,
                      scope_prefixes, scope_suffixes, keep_ext, prefix)
    _cfg, _res = [], []  # 配置行 / 执行结果行

    label = f"{Path(path).name} | {sheet}"
    _cfg.append(f"目标: {label} | 名称字段「{name_col}」→ 资源字段「{value_col}」")
    prefix_desc = f"，路径前缀 {prefix}" if (fmt == FMT_PATH and prefix) else ""
    n_scoped = sum(len(v) for v in scope_index(index, scope_prefixes, scope_suffixes).values())
    _cfg.append(f"写入格式: {FMT_NAMES[fmt]}{prefix_desc} | 资源库 {n_scoped} 个文件参与")
    if scope_prefixes or scope_suffixes:
        _cfg.append(f"限定前缀: {', '.join(scope_prefixes) or '无'} | 限定后缀: {', '.join(scope_suffixes) or '无'}")

    # 应用人工指定
    resolved_rows = []  # (行号, 名称, 新值, 资源相对路径)
    leftover_conflict, leftover_missing = [], []
    wb, ws, hdr = read_sheet(path, sheet)
    vc = hdr[value_col]
    try:
        for r, name, cands in plan["conflict"]:
            pick = resolutions.get(r)
            if pick:
                val = format_value(pick, fmt, keep_ext, prefix)
                if norm(ws.cell(r, vc).value) == norm(val):
                    plan["same"].append((r, name))
                else:
                    resolved_rows.append((r, name, val, pick))
            else:
                leftover_conflict.append((r, name, cands))
        for r, name in plan["missing"]:
            pick = resolutions.get(r)
            if pick:
                val = format_value(pick, fmt, keep_ext, prefix)
                if norm(ws.cell(r, vc).value) == norm(val):
                    plan["same"].append((r, name))
                else:
                    resolved_rows.append((r, name, val, pick))
            else:
                leftover_missing.append((r, name))
        changes = plan["ok"] + resolved_rows
        # 多名称命中同一资源（首字母碰撞风险）：需人工决定，未决定的全部不写入。
        # 资源已被组外名称的行填入的，交占用检测处理，不算重复待定；
        # 已被组内某名称的行填入的，默认保留该名称（可在「解决重复资源」改）。
        # 同一名称多行共享资源暂不拦截（无法自动判断），但会在日志里列出供核对。
        dup_pending = []
        kept_dup = []
        dup_decisions = dup_decisions or {}
        groups, preselect, outside_occ = split_dup_groups(
            changes, plan, fmt, keep_ext, prefix)
        for rel, names in groups.items():
            keep = dup_decisions.get(rel)  # None=未定; ""=都不写; 名称=只写该名称的行
            if keep is None:
                keep = preselect.get(rel)
            if keep is None:
                dup_pending.append((rel, names))
                changes = [ch for ch in changes if ch[3] != rel]
            else:
                before = len(changes)
                changes = [ch for ch in changes
                           if ch[3] != rel or (keep and ch[1] == keep)]
                kept_dup.append((rel, keep, before - len(changes)))
        # 「全部改动」窗口里被排除的行
        excluded_n = 0
        if exclude_rows:
            before = len(changes)
            changes = [ch for ch in changes if ch[0] not in exclude_rows]
            excluded_n = before - len(changes)
        # 新值已被不同名称的行占用：默认排除（include_rows 里人工恢复的除外）。
        # 按所有人工决定（重复组/排除/指定）之后的最终状态计算，避免误报。
        include_rows = include_rows or set()
        in_use_all = find_in_use(path, sheet, name_col, value_col, changes)
        in_use_map = {r: holders for r, _, _, _, holders in in_use_all}
        in_use_rows = {r for r, *_ in in_use_all}
        in_use_pending = []
        if in_use_rows:
            kept = []
            for ch in changes:
                if ch[0] in in_use_rows and ch[0] not in include_rows:
                    in_use_pending.append(ch)
                else:
                    kept.append(ch)
            changes = kept
        total = len(changes)
        # 覆盖清单按最终实际会写入的行重新统计（重复待定被剔除的行不算）
        remaining = {(r, name) for r, name, val, rel in changes}
        overwrite = [o for o in plan["overwrite"] if (o[0], o[1]) in remaining]

        L = []
        L.append(f"【改动】共 {total} 处"
                 + (f"（含人工指定 {len(resolved_rows)} 行）" if resolved_rows else ""))
        for r, name, val, rel in changes[:30]:
            L.append(f"  行{r} 「{name}」-> {val}（{rel}）")
        if len(changes) > 30:
            L.append(f"  …… 其余 {len(changes) - 30} 处省略，完整列表见「全部改动」窗口")
        if overwrite:
            L.append(f"  其中覆盖已填 {len(overwrite)} 行（旧 -> 新，前 10 条）:")
            for r, name, old, new in overwrite[:10]:
                L.append(f"    行{r} 「{name}」: {old} -> {new}")
            if len(overwrite) > 10:
                L.append(f"    …… 其余 {len(overwrite) - 10} 处省略")
        if plan["same"]:
            L.append(f"  值已一致无需改动: {len(plan['same'])} 行")

        pending = []
        if leftover_conflict:
            pending.append(f"  冲突未决 {len(leftover_conflict)} 行（点「解决冲突」处理）: "
                           + ", ".join(f"行{r}「{n}」" for r, n, _ in leftover_conflict[:5])
                           + (" ……" if len(leftover_conflict) > 5 else ""))
        if dup_pending:
            pending.append(f"  重复待定 {len(dup_pending)} 组（点「解决重复资源」处理）: "
                           + "; ".join(f"{rel} <- {', '.join(names)}"
                                       for rel, names in dup_pending[:5])
                           + (" ……" if len(dup_pending) > 5 else ""))
        if in_use_pending:
            pending.append(f"  占用排除 {len(in_use_pending)} 行"
                           "（资源已被其他名称占用，「全部改动」可恢复）:")
            for r, name, val, rel in in_use_pending[:5]:
                pending.append(f"    行{r} 「{name}」-> {val}"
                               f"（占用: {'、'.join(in_use_map.get(r, [])[:3])}）")
            if len(in_use_pending) > 5:
                pending.append(f"    …… 其余 {len(in_use_pending) - 5} 行省略")
        if leftover_missing:
            pending.append(f"  未匹配 {len(leftover_missing)} 行（点「手动指定」处理）: "
                           + ", ".join(f"行{r}「{n}」" for r, n in leftover_missing[:5])
                           + (" ……" if len(leftover_missing) > 5 else ""))
        if excluded_n:
            pending.append(f"  人工排除 {excluded_n} 行（「全部改动」中勾选）")
        if pending:
            L.append("【未写入/待处理】")
            L.extend(pending)

        sk = plan["skipped"]
        skips = []
        if sk["empty"]:
            skips.append(f"名称为空 {sk['empty']} 行")
        if sk["formula"]:
            skips.append(f"资源字段是公式 {sk['formula']} 行（工具不改动公式）")
        if skips:
            L.append("【跳过】" + "，".join(skips))
        for rel, keep, dropped in kept_dup:
            L.append(f"【重复资源】{rel} 按选定只写入「{keep or '无'}」的行（排除 {dropped} 行）")

        # 资源侧核查：限定范围外的 & 范围内未被命中的 & 同名称多行
        side = []
        if plan["scoped_out"]:
            side.append(f"  不符合限定前后缀 {len(plan['scoped_out'])} 个"
                        "（未参与匹配，检查命名）: "
                        + ", ".join(plan["scoped_out"][:5])
                        + (" ……" if len(plan["scoped_out"]) > 5 else ""))
        used = set(plan["hit_rels"]) | set(resolutions.values())
        scoped_rels = {rel for rels in
                       scope_index(index, scope_prefixes, scope_suffixes).values()
                       for rel in rels}
        unused = sorted(scoped_rels - used)
        if unused:
            side.append(f"  范围内未被命中 {len(unused)} 个（可能多余或名称对不上）: "
                        + ", ".join(unused[:5])
                        + (" ……" if len(unused) > 5 else ""))
        by_name_rel = {}
        for ch in changes:
            by_name_rel.setdefault((ch[1], ch[3]), []).append(ch[0])
        same_name_rows = [(n, rel, rows) for (n, rel), rows in by_name_rel.items()
                          if len(rows) > 1]
        if same_name_rows:
            side.append(f"  同名称多行共享资源 {len(same_name_rows)} 组"
                        "（道具ID可能不同，请核对）: "
                        + "; ".join(f"「{n}」行{rows[:5]}"
                                    for n, rel, rows in same_name_rows[:5])
                        + (" ……" if len(same_name_rows) > 5 else ""))
        if side:
            L.append("【资源侧核查】")
            L.extend(side)

        saved = 0
        result_line = ""
        if total and not dry:
            try:
                stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
                cell_changes = [(r, vc, val) for r, name, val, rel in changes]
                bak = write_cells_surgical(path, sheet, cell_changes, backup_dir, stamp)
                saved = total
                _res.append(f"[已保存] {Path(path).name}（共 {total} 处改动，外科式写入，已按原始状态备份）")
                result_line = (f"{label}: {total} 处改动，已保存"
                               + (f"（备份: {bak}）" if bak else "（未备份）"))
            except Exception as ex:
                _res.append(f"[错误] 保存失败: {ex}（原文件未被修改）")
                result_line = f"{label}: 保存失败（{ex}），原文件未被修改"
        elif total:
            result_line = f"{label}: 预览 {total} 处改动"
        else:
            result_line = f"{label}: 无需改动"
    finally:
        wb.close()

    out = [f"【批量匹配】{datetime.now():%Y-%m-%d %H:%M:%S}"]
    out.extend(_cfg)
    out.append("─" * 46)
    out.extend(L)
    out.append("─" * 46)
    out.append("执行结果: " + result_line)
    out.extend(f"  {l}" for l in _res)
    status = "预览(未写入)" if dry else ("成功" if saved == total else "失败")
    out.append(f"汇总: 计划改动 {total} | 实际保存 {saved} | "
               f"其中覆盖已填 {len(overwrite)} | "
               f"人工排除 {excluded_n} | 占用排除 {len(in_use_pending)} | "
               f"冲突未决 {len(leftover_conflict)} | 未匹配 {len(leftover_missing)} | "
               f"重复待定 {len(dup_pending)} | "
               f"范围外资源 {len(plan['scoped_out'])} | 未命中资源 {len(unused)} | 状态: {status}")
    log("\n".join(out))
    return total


# ---------- 存在性校验 ----------

def check_existence(path, sheet, value_cols, index, fmt, prefix="", log=print):
    """校验已填资源值在资源库中是否存在（只读不写表）。
    按文件名: 值按词干查索引。
    按相对路径: 值先按「路径前缀/值」拼回库内路径查（起始点为父目录时），
    再按值本身就是根目录相对路径查，最后按词干查兜底。
    返回 dict: missing/conflict/ok/empty/formula 统计与清单。"""
    wb, ws, hdr = read_sheet(path, sheet)
    try:
        existing_all = set(index.keys())
        rels_all = {rel for rels in index.values() for rel in rels}
        rel_noext = {}
        for rel in rels_all:
            rel_noext.setdefault(PurePosixPath(rel).with_suffix("").as_posix(),
                                 []).append(rel)

        def path_hit(v):
            vv = v.replace("\\", "/").rstrip("/")
            cands = [vv]
            if prefix:
                pre = prefix.strip("/")
                if vv.startswith(pre + "/"):
                    cands.insert(0, vv[len(pre) + 1:])  # 值含前缀，剥掉前缀查库
                else:
                    cands.append(pre + "/" + vv)        # 值不含前缀，补上前缀查
            hits = []
            for cand in cands:
                hits.extend(rel_noext.get(cand, []))
                if cand in rels_all:
                    hits.append(cand)
            return sorted(set(hits))

        missing, conflict, ok_n, empty_n, formula_n = [], [], 0, 0, 0
        for col in value_cols:
            if col not in hdr:
                log(f"[跳过] {col}: 表头里找不到该列")
                continue
            c = hdr[col]
            for r in range(2, ws.max_row + 1):
                cell = ws.cell(r, c)
                v = norm(cell.value)
                if not v:
                    empty_n += 1
                    continue
                if is_formula_cell(cell):
                    formula_n += 1
                    continue
                if fmt == FMT_NAME:
                    key = norm_stem(v)
                    if key in existing_all:
                        if len(index[key]) > 1:
                            conflict.append((r, col, v, index[key]))
                        else:
                            ok_n += 1
                    else:
                        missing.append((r, col, v))
                else:
                    hit = path_hit(v)
                    if len(hit) == 1:
                        ok_n += 1
                    elif len(hit) > 1:
                        conflict.append((r, col, v, hit))
                    else:
                        # 兼容只填了文件名的情况：按词干查
                        key = norm_stem(PurePosixPath(v.replace("\\", "/")).name)
                        if key in existing_all:
                            ok_n += 1
                        else:
                            missing.append((r, col, v))
    finally:
        wb.close()

    label = f"{Path(path).name} | {sheet}"
    prefix_desc = f"，路径前缀 {prefix}" if (fmt == FMT_PATH and prefix) else ""
    out = [f"【存在性校验】{datetime.now():%Y-%m-%d %H:%M:%S}"]
    out.append(f"目标: {label} | 校验方式: 按{FMT_NAMES[fmt]}{prefix_desc} | 校验字段: {', '.join(value_cols)}")
    out.append("─" * 46)
    out.append(f"通过: {ok_n} 个值")
    if missing:
        out.append(f"缺失（资源库找不到）: {len(missing)} 个")
        for r, col, v in missing[:100]:
            out.append(f"  行{r} 「{col}」: {v!r}")
        if len(missing) > 100:
            out.append(f"  …… 其余 {len(missing) - 100} 个省略")
    if conflict:
        out.append(f"冲突（一个值对应多个文件）: {len(conflict)} 个")
        for r, col, v, rels in conflict[:50]:
            out.append(f"  行{r} 「{col}」: {v!r} -> {' | '.join(rels[:5])}")
    if empty_n:
        out.append(f"空值: {empty_n} 个单元格（未填，未校验）")
    if formula_n:
        out.append(f"公式: {formula_n} 个单元格（未校验）")
    out.append("─" * 46)
    out.append(f"汇总: 通过 {ok_n} | 缺失 {len(missing)} | 冲突 {len(conflict)} | "
               f"空值 {empty_n} | 公式 {formula_n}")
    log("\n".join(out))
    return {"missing": missing, "conflict": conflict, "ok": ok_n,
            "empty": empty_n, "formula": formula_n}


# ---------- 界面 ----------

class ResolveDialog(tk.Toplevel):
    """冲突解决/手动指定窗口：先选资源，再选行，把资源指定给选中的行。
    self.resolutions = {行号: 资源相对路径}（含带入的已有指定）。"""

    def __init__(self, master, items, all_rels, resolutions, title="解决冲突 / 手动指定"):
        """items: [(行号, 名称, [候选相对路径...], 旧值)]，候选为空表示未匹配。
        resolutions: {行号: 资源相对路径} 已有指定。"""
        super().__init__(master)
        self.title(title)
        self.geometry("1060x560")
        self.resolutions = dict(resolutions)
        self._all = sorted(all_rels)
        self._items = items
        self._by_row = {r: (name, cands, old) for r, name, cands, old in items}

        ttk.Label(self, padding=(8, 8, 8, 0),
                  text="左侧搜索并选中资源（选中行后可只看该行的候选），右侧搜索并勾选要写入的行（可多选），点「← 指定给选中的行」。",
                  foreground="#888888").pack(anchor=tk.W)

        pane = ttk.Frame(self, padding=8)
        pane.pack(fill=tk.BOTH, expand=True)
        pane.columnconfigure(0, weight=1)
        pane.columnconfigure(1, weight=1)
        pane.rowconfigure(0, weight=1)

        # 左：资源
        f1 = ttk.LabelFrame(pane, text="1. 选资源", padding=4)
        f1.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        self.rel_filter = tk.StringVar()
        self.rel_filter.trace_add("write", lambda *a: self._refresh_rels())
        ttk.Entry(f1, textvariable=self.rel_filter).pack(fill=tk.X, pady=(0, 2))
        self.cand_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(f1, text="只看选中行的候选", variable=self.cand_var,
                        command=self._refresh_rels).pack(anchor=tk.W)
        self.rel_list = tk.Listbox(f1, exportselection=False)
        self.rel_list.pack(fill=tk.BOTH, expand=True)

        # 右：行
        f2 = ttk.LabelFrame(pane, text="2. 选行（可多选；显示 行号 | 名称 | 旧值）", padding=4)
        f2.grid(row=0, column=1, sticky="nsew")
        self.row_filter = tk.StringVar()
        self.row_filter.trace_add("write", lambda *a: self._refresh_rows())
        ttk.Entry(f2, textvariable=self.row_filter).pack(fill=tk.X, pady=(0, 2))
        self.row_list = tk.Listbox(f2, selectmode=tk.EXTENDED, exportselection=False)
        self.row_list.pack(fill=tk.BOTH, expand=True)
        self.row_list.bind("<<ListboxSelect>>", lambda e: self._refresh_rels())

        bar = ttk.Frame(self, padding=8)
        bar.pack(fill=tk.X)
        ttk.Button(bar, text="← 指定给选中的行", command=self._assign).pack(side=tk.LEFT)
        ttk.Button(bar, text="清除选中行的指定", command=self._clear).pack(
            side=tk.LEFT, padx=8)
        self.status = ttk.Label(bar, text="")
        self.status.pack(side=tk.LEFT, padx=8)
        ttk.Button(bar, text="完成", command=self.destroy).pack(side=tk.RIGHT)

        self._refresh_rows()
        self._refresh_rels()
        self.transient(master)

    def _selected_rows(self):
        return [self._shown_rows[i] for i in self.row_list.curselection()]

    def _row_label(self, r):
        name, cands, old = self._by_row[r]
        done = "✔ " if r in self.resolutions else ""
        state = f"冲突{len(cands)}候选" if cands else "未匹配"
        return f"{done}行{r} | {name} | {state} | 旧值 {old or '（空）'}"

    def _refresh_rows(self):
        kw = self.row_filter.get().strip().lower()
        self._shown_rows = []
        self.row_list.delete(0, tk.END)
        for r, name, cands, old in self._items:
            label = self._row_label(r)
            if not kw or kw in label.lower():
                self._shown_rows.append(r)
                self.row_list.insert(tk.END, label)

    def _rel_pool(self):
        """当前资源池：勾选「只看候选」且有选中行时，为选中行候选的并集；否则全部。"""
        if self.cand_var.get():
            rows = self._selected_rows()
            pool = sorted({c for r in rows for c in self._by_row[r][1]})
            if pool:
                return pool
        return self._all

    def _refresh_rels(self):
        kw = self.rel_filter.get().strip().lower()
        self.rel_list.delete(0, tk.END)
        for rel in self._rel_pool():
            if not kw or kw in rel.lower():
                self.rel_list.insert(tk.END, rel)

    def _assign(self):
        rows = self._selected_rows()
        sel = self.rel_list.curselection()
        if not rows or not sel:
            messagebox.showwarning("缺少选择", "请先选中资源和要写入的行", parent=self)
            return
        rel = self.rel_list.get(sel[0])
        for r in rows:
            self.resolutions[r] = rel
        self._refresh_rows()
        self._refresh_status()

    def _clear(self):
        for r in self._selected_rows():
            self.resolutions.pop(r, None)
        self._refresh_rows()
        self._refresh_status()

    def _refresh_status(self):
        done = sum(1 for r, _, _, _ in self._items if r in self.resolutions)
        self.status.config(text=f"已指定 {done} / {len(self._items)} 行")


class DupDialog(tk.Toplevel):
    """重复资源决定窗口：同一资源被多个不同名称命中时，选定只写入哪个名称（或都不写）。
    结果写入 self.decisions = {资源相对路径: 名称 或 ""}。"""

    def __init__(self, master, groups):
        """groups: [(资源相对路径, [名称...]), ...]"""
        super().__init__(master)
        self.title("解决重复资源")
        self.geometry("640x520")
        self.decisions = dict(getattr(master, "_dup_decisions", {}))
        self._vars = {}

        ttk.Label(self, padding=8,
                  text="同一资源被多个不同名称命中（疑似首字母碰撞），请选定每个资源只写入哪个名称的行；"
                       "选「都不写」则该资源相关的行全部跳过。未选择的组按「都不写」处理。",
                  foreground="#888888", wraplength=610, justify=tk.LEFT).pack(
            anchor=tk.W, fill=tk.X)

        # 底部按钮栏先 pack 到底部，避免被滚动区挤走
        bar = ttk.Frame(self, padding=8)
        bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status = ttk.Label(bar, text="")
        self.status.pack(side=tk.LEFT)
        ttk.Button(bar, text="完成", command=self._done).pack(side=tk.RIGHT)

        # 可滚动区域
        canvas = tk.Canvas(self, highlightthickness=0)
        sb = ttk.Scrollbar(self, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
        inner = ttk.Frame(canvas, padding=8)
        win = canvas.create_window((0, 0), window=inner, anchor=tk.NW)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-e.delta / 120), "units"))

        for rel, names in groups:
            g = ttk.LabelFrame(inner, text=rel, padding=4)
            g.pack(fill=tk.X, pady=3)
            var = tk.StringVar(value=self.decisions.get(rel) or "__skip__")
            self._vars[rel] = var
            for name in names:
                ttk.Radiobutton(g, text=f"只写入「{name}」的行", variable=var,
                                value=name).pack(anchor=tk.W)
            ttk.Radiobutton(g, text="都不写", variable=var,
                            value="__skip__").pack(anchor=tk.W)

        self._refresh_status()
        self.transient(master)

    def _done(self):
        for rel, var in self._vars.items():
            v = var.get()
            self.decisions[rel] = "" if v == "__skip__" else v
        self.destroy()

    def _refresh_status(self):
        n = sum(1 for v in self._vars.values() if v.get() != "__skip__")
        self.status.config(text=f"共 {len(self._vars)} 组重复资源，已选保留 {n} 组")


class ChangesDialog(tk.Toplevel):
    """全部改动窗口：列出所有待写入改动，可双击/按钮排除部分行再执行。
    self.excluded = {行号}，关闭后由调用方读取。"""

    def __init__(self, master, changes, excluded, auto_excluded=(), included=(),
                 occupy=None):
        """changes: [(行号, 名称, 旧值, 新值, 资源相对路径)]；excluded: {行号} 已排除的。
        auto_excluded: {行号} 新值被其他行占用而默认排除的；included: {行号} 已人工恢复的占用行。
        occupy: {行号: 占用方描述}。"""
        super().__init__(master)
        self.title("全部改动")
        self.geometry("1080x540")
        self.excluded = set(excluded)
        self.included = set(included)
        self._auto = set(auto_excluded)
        self._occupy = occupy or {}
        self._data = {r: (name, old, val, rel) for r, name, old, val, rel in changes}

        ttk.Label(self, padding=(8, 8, 8, 0),
                  text="双击（或选中后按空格）切换 写入/排除；「占用」行的资源已被其他名称的行使用，"
                       "默认排除，可手动恢复。",
                  foreground="#888888").pack(anchor=tk.W)

        cols = ("state", "row", "name", "old", "value", "rel", "note")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for cid, text, w in [("state", "写入", 56), ("row", "行号", 56),
                             ("name", "名称", 120), ("old", "旧值", 170),
                             ("value", "新值", 190), ("rel", "资源", 280),
                             ("note", "备注", 160)]:
            self.tree.heading(cid, text=text)
            self.tree.column(cid, width=w, anchor=tk.W)
        sb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)
        self.tree.bind("<Double-Button-1>", lambda e: self._toggle_selected())
        self.tree.bind("<space>", lambda e: self._toggle_selected())

        for r, name, old, val, rel in changes:
            self.tree.insert("", "end", iid=str(r), values=self._values(r))

        bar = ttk.Frame(self, padding=8)
        bar.pack(fill=tk.X)
        self.status = ttk.Label(bar, text="")
        self.status.pack(side=tk.LEFT)
        ttk.Button(bar, text="排除选中", command=lambda: self._set_selected(False)).pack(
            side=tk.LEFT, padx=8)
        ttk.Button(bar, text="恢复选中", command=lambda: self._set_selected(True)).pack(
            side=tk.LEFT)
        ttk.Button(bar, text="全部恢复", command=self._restore_all).pack(side=tk.LEFT, padx=8)
        ttk.Button(bar, text="完成", command=self.destroy).pack(side=tk.RIGHT)
        self._refresh_status()
        self.transient(master)

    def _checked(self, r):
        """行当前是否勾选写入：人工排除 > 占用默认排除（可人工恢复）。"""
        return r not in self.excluded and (r not in self._auto or r in self.included)

    def _values(self, r):
        name, old, val, rel = self._data[r]
        if self._checked(r):
            state = "✓"
        elif r in self._auto:
            state = "—占用"
        else:
            state = "—"
        return (state, r, name, old, val, rel, self._occupy.get(r, ""))

    def _set_row(self, r, write):
        """勾选/取消勾选一行（处理占用行的恢复逻辑）。"""
        if write:
            if r in self._auto:
                self.included.add(r)
            else:
                self.excluded.discard(r)
        else:
            self.included.discard(r)
            if r not in self._auto:
                self.excluded.add(r)

    def _toggle_selected(self):
        for iid in self.tree.selection():
            r = int(iid)
            self._set_row(r, not self._checked(r))
            self.tree.item(iid, values=self._values(r))
        self._refresh_status()

    def _set_selected(self, write):
        for iid in self.tree.selection():
            r = int(iid)
            self._set_row(r, write)
            self.tree.item(iid, values=self._values(r))
        self._refresh_status()

    def _restore_all(self):
        self.excluded.clear()
        self.included = set(self._auto)
        for iid in self.tree.get_children():
            self.tree.item(iid, values=self._values(int(iid)))
        self._refresh_status()

    def _refresh_status(self):
        n_write = sum(1 for r in self._data if self._checked(r))
        self.status.config(
            text=f"共 {len(self._data)} 处改动，写入 {n_write} 行，"
                 f"排除 {len(self._data) - n_write} 行（含占用默认排除 {len(self._auto - self.included)} 行）")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("美术资源配置工具")
        self.geometry("1060x720")
        self.app_dir = app_dir()
        self.res_root = get_config_dir(self.app_dir, "res_root") \
            or get_config_dir(self.app_dir, "icon_root")  # 兼容旧配置
        self.backup_dir = get_config_dir(self.app_dir, "backup_dir") \
            or (self.app_dir / BACKUP_DIR)
        self._index = {}        # 资源库索引 {规范化词干: [相对路径...]}
        self._resolutions = {}  # {(文件路径, 子表, 资源字段, 行号): 资源相对路径}（会话内保持）
        self._dup_decisions = {}  # {资源相对路径: 保留的名称 或 ""}（会话内保持）
        self._excluded = set()  # {(文件路径, 子表, 资源字段, 行号)}（会话内保持）
        self._included = set()  # 人工恢复的占用行，同上维度（会话内保持）
        self._start_dirs = []   # 路径起始点候选（资源库根目录及其各层父目录）
        # 两个页签各自独立的配置表选择状态
        for prefix in ("m", "c"):
            setattr(self, f"{prefix}_file", None)    # 已选配置表路径
            setattr(self, f"{prefix}_sheets", [])    # [(子表名, [表头...])]
        self._build()
        self._rescan(log=False)
        self._log("[启动] 美术资源配置工具已就绪"
                  + ("（未安装 pypinyin，中文名只按原名匹配）" if not pinyin else ""))

    # ---------- 界面搭建 ----------

    def _build(self):
        # 顶部路径设置栏
        bar = ttk.LabelFrame(self, text="目录设置", padding=6)
        bar.pack(fill=tk.X, padx=6, pady=(6, 0))
        self._dir_row(bar, 0, "资源库根目录:", "res_root",
                      lambda: self._pick_dir("res_root"))
        ttk.Button(bar, text="重新扫描", command=lambda: self._rescan(log=True)).grid(
            row=1, column=1, sticky=tk.W, pady=(4, 0))
        self.count_var = tk.StringVar(value="资源库未扫描")
        ttk.Label(bar, textvariable=self.count_var, foreground="#555555").grid(
            row=1, column=2, sticky=tk.W, padx=8)
        ttk.Button(bar, text="设置备份目录", command=self._set_backup_dir).grid(
            row=1, column=3, sticky=tk.E, padx=4)

        ttk.Label(bar, text="资源类型（多选，默认全部）:").grid(
            row=0, column=4, sticky=tk.W, padx=(16, 4))
        self.ext_list = tk.Listbox(bar, selectmode=tk.EXTENDED, height=3,
                                   width=14, exportselection=False)
        self.ext_list.grid(row=1, column=4, sticky=tk.NW, padx=(16, 4), pady=(4, 0))
        ttk.Button(bar, text="全选", command=self._select_all_exts).grid(
            row=1, column=5, sticky=tk.NW, pady=(4, 0))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        self._build_match_tab()
        self._build_check_tab()

        self.log_area = scrolledtext.ScrolledText(self, height=16, font=("Consolas", 9))
        self.log_area.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))

    def _dir_row(self, parent, row, label, attr, cmd):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky=tk.W)
        var = tk.StringVar(value=str(getattr(self, attr) or "（未设置）"))
        setattr(self, f"{attr}_var", var)
        ttk.Label(parent, textvariable=var, width=70,
                  foreground="#333333").grid(row=row, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Button(parent, text="选择…", command=cmd).grid(row=row, column=3, sticky=tk.E)

    def _table_row(self, parent, prefix):
        """配置表选择行：直接弹文件框选一个 xlsx。"""
        ttk.Label(parent, text="配置表:").grid(row=0, column=0, sticky=tk.W)
        var = tk.StringVar(value="（未选择）")
        setattr(self, f"{prefix}_path_var", var)
        ttk.Label(parent, textvariable=var, width=50,
                  foreground="#333333").grid(row=0, column=1, columnspan=3, sticky=tk.W, padx=4)
        ttk.Button(parent, text="选择配置表…",
                   command=lambda: self._pick_table(prefix)).grid(
            row=0, column=4, columnspan=2, sticky=tk.W)

    def _base_row(self, parent, row, prefix):
        """路径起始点下拉（资源库根目录及其父目录，仅 相对路径+文件名 时生效）。"""
        ttk.Label(parent, text="路径起始点:").grid(row=row, column=0, sticky=tk.W)
        var = tk.StringVar(value=ROOT_LABEL)
        box = ttk.Combobox(parent, textvariable=var, state="readonly", width=34)
        box.grid(row=row, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(parent, text="（写相对路径时从哪一级目录开始算，可选资源库的父目录）",
                  foreground="#888888").grid(row=row, column=3, columnspan=3, sticky=tk.W)
        setattr(self, f"{prefix}_base_var", var)
        setattr(self, f"{prefix}_base_box", box)

    def _selected_prefix(self, prefix):
        """路径起始点 → 写入值的路径前缀（起始点到资源库根目录的相对路径）。"""
        pos = getattr(self, f"{prefix}_base_box").current()
        if pos <= 0 or not self.res_root:  # 0=资源库根目录，无前缀
            return ""
        start = self._start_dirs[pos]
        try:
            return self.res_root.relative_to(start).as_posix()
        except ValueError:
            return ""

    def _refresh_base_boxes(self):
        """路径起始点候选：资源库根目录 + 各层父目录（到盘符为止）。"""
        self._start_dirs = []
        if self.res_root:
            d = self.res_root.resolve()
            self._start_dirs.append(d)
            self._start_dirs.extend(d.parents)
        vals = [ROOT_LABEL] + [str(d) for d in self._start_dirs[1:]]
        for prefix in ("m", "c"):
            box = getattr(self, f"{prefix}_base_box", None)
            var = getattr(self, f"{prefix}_base_var", None)
            if box is None:
                continue
            old = var.get()
            box["values"] = vals
            var.set(old if old in vals else ROOT_LABEL)

    def _build_match_tab(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  批量匹配  ")

        sel = ttk.LabelFrame(tab, text="① 选择配置表与字段", padding=6)
        sel.pack(fill=tk.X, padx=8, pady=(6, 0))

        self._table_row(sel, "m")

        ttk.Label(sel, text="子表:").grid(row=1, column=0, sticky=tk.W)
        self.m_sheet_var = tk.StringVar()
        self.m_sheet_box = ttk.Combobox(sel, textvariable=self.m_sheet_var,
                                        state="readonly", width=24)
        self.m_sheet_box.grid(row=1, column=1, sticky=tk.W, padx=4)
        self.m_sheet_box.bind("<<ComboboxSelected>>", lambda e: self._m_on_sheet())

        ttk.Label(sel, text="匹配字段（名称）:").grid(row=1, column=2, sticky=tk.W)
        self.m_name_var = tk.StringVar()
        self.m_name_box = ttk.Combobox(sel, textvariable=self.m_name_var,
                                       state="readonly", width=18)
        self.m_name_box.grid(row=1, column=3, sticky=tk.W, padx=4)

        ttk.Label(sel, text="要填的字段（资源）:").grid(row=1, column=4, sticky=tk.W)
        self.m_value_var = tk.StringVar()
        self.m_value_box = ttk.Combobox(sel, textvariable=self.m_value_var,
                                        state="readonly", width=18)
        self.m_value_box.grid(row=1, column=5, sticky=tk.W, padx=4)

        opt = ttk.LabelFrame(tab, text="② 匹配设置", padding=6)
        opt.pack(fill=tk.X, padx=8, pady=4)

        ttk.Label(opt, text="写入格式:").grid(row=0, column=0, sticky=tk.W)
        self.m_fmt_var = tk.StringVar(value=FMT_NAME)
        ttk.Radiobutton(opt, text="文件名", variable=self.m_fmt_var,
                        value=FMT_NAME).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(opt, text="相对路径+文件名", variable=self.m_fmt_var,
                        value=FMT_PATH).grid(row=0, column=2, sticky=tk.W)
        self.m_keep_ext = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="保留扩展名", variable=self.m_keep_ext).grid(
            row=0, column=3, sticky=tk.W, padx=8)

        self._base_row(opt, 1, "m")

        cfg = load_config(self.app_dir)
        ttk.Label(opt, text="限定前缀:").grid(row=2, column=0, sticky=tk.W)
        self.m_prefix_var = tk.StringVar(value=cfg.get("scope_prefix", ""))
        ttk.Entry(opt, textvariable=self.m_prefix_var, width=34).grid(
            row=2, column=1, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(opt, text="限定后缀:").grid(row=2, column=3, sticky=tk.W)
        self.m_suffix_var = tk.StringVar(value=cfg.get("scope_suffix", ""))
        ttk.Entry(opt, textvariable=self.m_suffix_var, width=24).grid(
            row=2, column=4, columnspan=2, sticky=tk.W, padx=4)
        ttk.Label(opt, text="（多个用逗号分隔；只有文件名符合限定前后缀的资源才参与匹配，如 V1_icon_zw_）",
                  foreground="#888888").grid(row=3, column=0, columnspan=6, sticky=tk.W)

        act = ttk.Frame(tab, padding=6)
        act.pack(fill=tk.X)
        ttk.Button(act, text="预览", command=lambda: self._run_match(dry=True)).pack(side=tk.LEFT)
        ttk.Button(act, text="解决冲突", command=lambda: self._open_resolve("conflict")).pack(
            side=tk.LEFT, padx=(0, 8))
        ttk.Button(act, text="手动指定", command=lambda: self._open_resolve("missing")).pack(
            side=tk.LEFT)
        ttk.Button(act, text="解决重复资源", command=self._open_dups).pack(side=tk.LEFT, padx=8)
        ttk.Button(act, text="全部改动…", command=self._open_changes).pack(
            side=tk.LEFT, padx=8)
        ttk.Button(act, text="开始写入", command=lambda: self._run_match(dry=False)).pack(
            side=tk.RIGHT)

    def _build_check_tab(self):
        tab = ttk.Frame(self.nb)
        self.nb.add(tab, text="  存在性校验  ")

        sel = ttk.LabelFrame(tab, text="① 选择配置表与资源字段（字段可多选）", padding=6)
        sel.pack(fill=tk.BOTH, expand=True, padx=8, pady=(6, 0))

        self._table_row(sel, "c")

        ttk.Label(sel, text="子表:").grid(row=1, column=0, sticky=tk.W)
        self.c_sheet_var = tk.StringVar()
        self.c_sheet_box = ttk.Combobox(sel, textvariable=self.c_sheet_var,
                                        state="readonly", width=24)
        self.c_sheet_box.grid(row=1, column=1, sticky=tk.W, padx=4)
        self.c_sheet_box.bind("<<ComboboxSelected>>", lambda e: self._c_on_sheet())

        ttk.Label(sel, text="资源字段:").grid(row=1, column=2, sticky=tk.NW)
        self.c_field_list = tk.Listbox(sel, selectmode=tk.EXTENDED, height=6,
                                       exportselection=False)
        self.c_field_list.grid(row=1, column=3, sticky=tk.NSEW, padx=4)

        opt = ttk.Frame(tab, padding=6)
        opt.pack(fill=tk.X)
        ttk.Label(opt, text="② 校验方式:").grid(row=0, column=0, sticky=tk.W)
        self.c_fmt_var = tk.StringVar(value=FMT_NAME)
        ttk.Radiobutton(opt, text="按文件名", variable=self.c_fmt_var,
                        value=FMT_NAME).grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(opt, text="按相对路径", variable=self.c_fmt_var,
                        value=FMT_PATH).grid(row=0, column=2, sticky=tk.W)
        self._base_row(opt, 1, "c")
        ttk.Button(opt, text="开始校验（只读，不写表）",
                   command=self._run_check).grid(row=0, column=3, sticky=tk.E, padx=8)

    # ---------- 数据装载 ----------

    def _pick_dir(self, attr):
        d = filedialog.askdirectory(parent=self, title="选择目录",
                                    initialdir=str(getattr(self, attr) or self.app_dir))
        if not d:
            return
        setattr(self, attr, Path(d))
        getattr(self, f"{attr}_var").set(d)
        save_config_dir(self.app_dir, attr, d)
        if attr == "res_root":
            self._rescan(log=True)

    def _set_backup_dir(self):
        d = filedialog.askdirectory(parent=self, initialdir=str(self.backup_dir),
                                    title="选择备份目录")
        if not d:
            return
        self.backup_dir = Path(d)
        save_config_dir(self.app_dir, "backup_dir", d)
        self._log(f"[备份目录] 已设置为: {self.backup_dir}")

    def _rescan(self, log=True):
        self._index = scan_resources(self.res_root, cache_dir=self.app_dir) \
            if self.res_root else {}
        self._refresh_base_boxes()
        # 扩展名列表：默认全选（= 不过滤）
        self.ext_list.delete(0, tk.END)
        for e in exts_of(self._index):
            self.ext_list.insert(tk.END, e or "（无扩展名）")
        self._select_all_exts()
        n = sum(len(v) for v in self._index.values())
        self.count_var.set(f"资源库 {n} 个文件")
        if log:
            self._log(f"[扫描] 资源库 {n} 个文件")

    def _select_all_exts(self):
        self.ext_list.select_set(0, tk.END)

    def _cur_index(self):
        """按勾选的资源类型（扩展名）过滤后的资源索引。全选/未选 = 不过滤。"""
        total = self.ext_list.size()
        sel = self.ext_list.curselection()
        if not sel or len(sel) == total:
            return self._index
        exts = set()
        for i in sel:
            e = self.ext_list.get(i)
            exts.add("" if e == "（无扩展名）" else e)
        return filter_index(self._index, exts)

    # ---------- 配置表选择 ----------

    def _pick_table(self, prefix):
        """直接弹文件框选一个配置表 xlsx，加载其子表。"""
        init = get_config_dir(self.app_dir, "table_dir") or self.app_dir
        f = filedialog.askopenfilename(
            parent=self, title="选择配置表", initialdir=str(init),
            filetypes=[("Excel 表格", "*.xlsx")])
        if not f:
            return
        p = Path(f)
        try:
            sheets = load_table_sheets(p)
        except Exception as ex:
            messagebox.showerror("读取失败", f"{p.name}: {ex}")
            return
        if not sheets:
            messagebox.showwarning("没有可用子表", f"{p.name} 里没有表头合规的子表")
            return
        save_config_dir(self.app_dir, "table_dir", p.parent)  # 记住目录，下次直接打开这里
        setattr(self, f"{prefix}_file", p)
        setattr(self, f"{prefix}_sheets", sheets)
        getattr(self, f"{prefix}_path_var").set(str(p))
        sheet_box = getattr(self, f"{prefix}_sheet_box")
        sheet_box["values"] = [s for s, _ in sheets]
        getattr(self, f"{prefix}_sheet_var").set("")
        if prefix == "m":
            self.m_name_box["values"] = []
            self.m_value_box["values"] = []
        else:
            self.c_field_list.delete(0, tk.END)
        self._log(f"[配置表] 已选择 {p}（{len(sheets)} 个可用子表）")

    def _cur_sheet_hdr(self, prefix):
        """当前选中的 (子表名, 表头)；未选返回 None。"""
        sheets = getattr(self, f"{prefix}_sheets")
        pos = getattr(self, f"{prefix}_sheet_box").current()
        return sheets[pos] if 0 <= pos < len(sheets) else None

    # ---------- 页签1 事件 ----------

    def _m_on_sheet(self):
        cur = self._cur_sheet_hdr("m")
        if not cur:
            return
        self.m_name_box["values"] = cur[1]
        self.m_value_box["values"] = cur[1]

    def _m_selection(self):
        """校验并返回 (文件路径, 子表, 名称字段, 资源字段, 写入格式, 前缀, 后缀, 路径前缀)，不完整返回 None。"""
        if not self.m_file:
            messagebox.showwarning("缺少选择", "请先选择配置表")
            return None
        cur = self._cur_sheet_hdr("m")
        if not cur:
            messagebox.showwarning("缺少选择", "请先选择子表")
            return None
        name_col, value_col = self.m_name_var.get(), self.m_value_var.get()
        if not name_col or not value_col:
            messagebox.showwarning("缺少选择", "请选择匹配字段和要填的字段")
            return None
        if name_col == value_col:
            messagebox.showwarning("选择有误", "匹配字段和要填的字段不能相同")
            return None
        if not self.res_root:
            messagebox.showwarning("缺少设置", "请先在顶部设置资源库根目录")
            return None
        fmt = self.m_fmt_var.get()
        scope_prefixes = parse_patterns(self.m_prefix_var.get())
        scope_suffixes = parse_patterns(self.m_suffix_var.get())
        path_prefix = self._selected_prefix("m")
        save_config(self.app_dir,
                    scope_prefix=self.m_prefix_var.get(),
                    scope_suffix=self.m_suffix_var.get())
        return (self.m_file, cur[0], name_col, value_col,
                fmt, scope_prefixes, scope_suffixes, path_prefix)

    def _run_match(self, dry):
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, prefixes, suffixes, path_prefix = sel
        exclude_rows = self._excluded_rows(path, sheet, value_col)
        include_rows = self._included_rows(path, sheet, value_col)
        res_rows = self._res_rows(path, sheet, value_col)
        if not dry:
            prefix_desc = f"\n路径起始点: {self.m_base_var.get()}" if fmt == FMT_PATH else ""
            if not messagebox.askyesno(
                "确认写入",
                f"文件: {path.name}\n子表: {sheet}\n"
                f"匹配字段「{name_col}」→ 要填的字段「{value_col}」\n"
                f"写入格式: {FMT_NAMES[fmt]}{prefix_desc}\n"
                f"已手动指定 {len(res_rows)} 行，"
                f"已决定重复资源 {len(self._dup_decisions)} 组，"
                f"已排除 {len(exclude_rows)} 行\n\n"
                "只会改写数据区，表头不动。确定执行？"):
                return
        self._log("=" * 60)
        try:
            run_match(path, sheet, name_col, value_col, self._cur_index(), fmt,
                      prefixes, suffixes, self.m_keep_ext.get(), path_prefix,
                      resolutions=res_rows, dup_decisions=self._dup_decisions,
                      exclude_rows=exclude_rows, include_rows=include_rows,
                      backup_dir=self.backup_dir, dry=dry, log=self._log)
        except Exception as ex:
            self._log(f"[错误] {ex}")

    def _open_resolve(self, mode):
        """mode='conflict' 只列冲突行；mode='missing' 只列未匹配行。"""
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, prefixes, suffixes, path_prefix = sel
        cur_index = self._cur_index()
        try:
            plan = plan_match(path, sheet, name_col, value_col, cur_index, fmt,
                              prefixes, suffixes, self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("读取失败", str(ex))
            return
        if mode == "conflict":
            items = [(r, name, cands, plan["old_values"].get(r, ""))
                     for r, name, cands in plan["conflict"]]
            title, empty_msg = "解决冲突", "当前没有冲突的行"
        else:
            items = [(r, name, [], plan["old_values"].get(r, ""))
                     for r, name in plan["missing"]]
            title, empty_msg = "手动指定", "当前没有未匹配的行"
        if not items:
            messagebox.showinfo("无需指定", empty_msg)
            return
        all_rels = sorted(rel for rels in
                          scope_index(cur_index, prefixes, suffixes).values()
                          for rel in rels)
        res_rows = self._res_rows(path, sheet, value_col)
        dlg = ResolveDialog(self, sorted(items), all_rels, res_rows, title=title)
        self.wait_window(dlg)
        self._save_res_rows(path, sheet, value_col, dlg.resolutions)
        self._log(f"[{title}] 当前共 {len(dlg.resolutions)} 行已指定资源")
        self._run_match(dry=True)  # 关闭后自动刷新预览

    def _dup_groups(self, plan, res_rows, fmt, keep_ext, prefix):
        """从匹配计划算出待人工决定的重复资源组: {资源相对路径: [不同名称...]}。
        已被组外名称的行填入的资源交给占用检测，不出现在这里；
        已被组内某名称的行填入的，默认保留该名称（写入 _dup_decisions）。
        同一名称多行共享资源暂不拦截。"""
        changes = [(r, n, None, rel) for r, n, v, rel in plan["ok"]]
        for r, name, cands in plan["conflict"]:
            pick = res_rows.get(r)
            if pick:
                changes.append((r, name, None, pick))
        for r, name in plan["missing"]:
            pick = res_rows.get(r)
            if pick:
                changes.append((r, name, None, pick))
        groups, preselect, _ = split_dup_groups(changes, plan, fmt, keep_ext, prefix)
        for rel, name in preselect.items():
            self._dup_decisions.setdefault(rel, name)
        return groups

    def _open_dups(self):
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, scope_prefixes, scope_suffixes, path_prefix = sel
        try:
            plan = plan_match(path, sheet, name_col, value_col, self._cur_index(), fmt,
                              scope_prefixes, scope_suffixes,
                              self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("读取失败", str(ex))
            return
        groups = self._dup_groups(plan, self._res_rows(path, sheet, value_col),
                                  fmt, self.m_keep_ext.get(), path_prefix)
        if not groups:
            messagebox.showinfo("无重复资源", "当前没有同一资源被多个不同名称命中的情况")
            return
        dlg = DupDialog(self, sorted(groups.items()))
        self.wait_window(dlg)
        self._dup_decisions.update(dlg.decisions)
        self._log(f"[重复资源] 当前共 {len(self._dup_decisions)} 组已决定")
        self._run_match(dry=True)  # 关闭后自动刷新预览

    def _excl_key(self, path, sheet, value_col):
        return (str(path), sheet, value_col)

    def _excluded_rows(self, path, sheet, value_col):
        key = self._excl_key(path, sheet, value_col)
        return {r for f, s, c, r in self._excluded if (f, s, c) == key}

    def _included_rows(self, path, sheet, value_col):
        key = self._excl_key(path, sheet, value_col)
        return {r for f, s, c, r in self._included if (f, s, c) == key}

    def _res_rows(self, path, sheet, value_col):
        """当前 文件+子表+字段 的人工指定：{行号: 资源相对路径}。"""
        key = self._excl_key(path, sheet, value_col)
        return {r: rel for (f, s, c, r), rel in self._resolutions.items()
                if (f, s, c) == key}

    def _save_res_rows(self, path, sheet, value_col, res_rows):
        """把对话框的行级指定写回全局 _resolutions（覆盖该上下文的旧值）。"""
        key = self._excl_key(path, sheet, value_col)
        self._resolutions = {k: v for k, v in self._resolutions.items()
                             if k[:3] != key}
        self._resolutions.update({key + (r,): rel for r, rel in res_rows.items()})

    def _open_changes(self):
        sel = self._m_selection()
        if not sel:
            return
        path, sheet, name_col, value_col, fmt, prefixes, suffixes, path_prefix = sel
        try:
            plan = plan_match(path, sheet, name_col, value_col, self._cur_index(), fmt,
                              prefixes, suffixes, self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("读取失败", str(ex))
            return
        try:
            changes = preview_changes(plan, self._res_rows(path, sheet, value_col),
                                      self._dup_decisions,
                                      fmt, self.m_keep_ext.get(), path_prefix)
        except Exception as ex:
            messagebox.showerror("计算改动失败", str(ex))
            return
        if not changes:
            messagebox.showinfo("没有改动", "当前没有待写入的改动")
            return
        key = self._excl_key(path, sheet, value_col)
        excl = self._excluded_rows(path, sheet, value_col)
        try:
            in_use = find_in_use(path, sheet, name_col, value_col, changes)
        except Exception as ex:
            messagebox.showerror("占用检测失败", str(ex))
            return
        auto = {r for r, *_ in in_use}
        occupy = {r: "占用: " + "、".join(holders)
                  for r, _, _, _, holders in in_use}
        incl = self._included_rows(path, sheet, value_col)
        dlg = ChangesDialog(self, changes, excl, auto, incl, occupy)
        self.wait_window(dlg)
        self._excluded = {k for k in self._excluded if k[:3] != key}
        self._excluded |= {(key[0], key[1], key[2], r) for r in dlg.excluded}
        self._included = {k for k in self._included if k[:3] != key}
        self._included |= {(key[0], key[1], key[2], r) for r in dlg.included}
        n_write = sum(1 for r, *_ in changes if dlg._checked(r))
        self._log(f"[全部改动] 共 {len(changes)} 处改动，写入 {n_write} 行，"
                  f"排除 {len(changes) - n_write} 行")
        self._run_match(dry=True)  # 关闭后自动刷新预览

    # ---------- 页签2 事件 ----------

    def _c_on_sheet(self):
        cur = self._cur_sheet_hdr("c")
        if not cur:
            return
        self.c_field_list.delete(0, tk.END)
        for h in cur[1]:
            self.c_field_list.insert(tk.END, h)

    def _run_check(self):
        if not self.c_file:
            messagebox.showwarning("缺少选择", "请先选择配置表")
            return
        cur = self._cur_sheet_hdr("c")
        if not cur:
            messagebox.showwarning("缺少选择", "请先选择子表")
            return
        cols = [self.c_field_list.get(i) for i in self.c_field_list.curselection()]
        if not cols:
            messagebox.showwarning("缺少选择", "请勾选要校验的资源字段")
            return
        if not self.res_root:
            messagebox.showwarning("缺少设置", "请先在顶部设置资源库根目录")
            return
        self._log("=" * 60)
        try:
            check_existence(self.c_file, cur[0], cols,
                            self._cur_index(), self.c_fmt_var.get(),
                            prefix=self._selected_prefix("c"), log=self._log)
        except Exception as ex:
            self._log(f"[错误] {ex}")

    def _log(self, msg):
        self.log_area.insert(tk.END, msg + "\n")
        self.log_area.see(tk.END)
        self.update_idletasks()


if __name__ == "__main__":
    app = App()
    app.mainloop()

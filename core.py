# -*- coding: utf-8 -*-
"""美术资源配置工具 - 核心逻辑（无界面依赖）。

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
"""
import json
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath
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
BACKUP_DIR = "资源写入备份"

FMT_NAME = "name"  # 写入/校验：文件名（不含扩展名）
FMT_PATH = "path"  # 写入/校验：相对路径+文件名（不含扩展名）
FMT_NAMES = {FMT_NAME: "文件名", FMT_PATH: "相对路径+文件名"}


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
    每次都真实遍历：目录 mtime 不会因果嵌套子目录内的增删改而变化，
    依赖缓存会拿到过期索引，而本扫描只走文件名、本身很快。
    cache_dir 参数保留仅为兼容旧调用，不再使用。"""
    root = Path(root)
    try:
        root.stat()
    except OSError:
        return {}
    index = {}
    for p in sorted(root.rglob("*")):
        if not p.is_file() or p.name.startswith("~$"):
            continue
        rel = p.relative_to(root).as_posix()
        index.setdefault(norm_stem(p.name), []).append(rel)
    return index


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
    """解析限定前后缀输入：逗号/空格/分号分隔，去空去重。"""
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


def build_matcher(index):
    """为资源索引预构建匹配器（批量匹配时只建一次，避免每个名称重复拆词）。
    返回 (t1, t2, files, jlen)：
      t1:    {整名规范化: [相对路径...]}           —— 第 1 级整名相等，字典直查
      t2:    {连续词元段: [相对路径...]}           —— 第 2 级词元段相等，字典直查
      files: [(相对路径, [词元...])]               —— 第 3 级子串判定用
      jlen:  {相对路径: 词元拼接长度}              —— 第 2 级打分用
    """
    t1, t2, files, jlen = {}, {}, [], {}
    for rels in index.values():
        for rel in rels:
            stem = Path(rel).stem.lower()
            nrm = re.sub(r"[_\-\s]+", "", stem)
            t1.setdefault(nrm, []).append(rel)
            toks = tokenize(stem)
            files.append((rel, toks))
            joined = "".join(toks)
            jlen[rel] = len(joined)
            for i in range(len(toks)):
                acc = ""
                for j in range(i, len(toks)):
                    acc += toks[j]
                    t2.setdefault(acc, []).append(rel)
    return t1, t2, files, jlen


def matcher_match(matcher, variants):
    """在预构建的匹配器上匹配候选集。返回 (status, candidates)，规则：
      ok       唯一命中（仅第 1/2 级）
      conflict 并列最优、重名、或第 3 级子串疑似命中（交人工选定）
      missing  无命中
    只在最高命中层级内取最优；并列不猜。"""
    if not variants:
        return ("missing", [])
    t1, t2, files, jlen = matcher
    hits = {}  # rel -> (tier, score)
    for v in variants:
        for rel in t1.get(v, ()):
            hits[rel] = (1, 0)
    if not hits:
        for v in variants:
            for rel in t2.get(v, ()):
                s = (2, abs(jlen[rel] - len(v)))
                if rel not in hits or s < hits[rel]:
                    hits[rel] = s
    if not hits:
        for rel, toks in files:
            for v in variants:
                hit3 = None
                for t in toks:
                    if (len(v) >= 3 and v in t) or (len(t) >= 4 and t in v):
                        hit3 = (3, abs(len(t) - len(v)))
                        break
                if hit3:
                    hits[rel] = hit3
                    break
    if not hits:
        return ("missing", [])
    best_tier = min(t for t, _ in hits.values())
    best_score = min(s for t, s in hits.values() if t == best_tier)
    best = sorted(rel for rel, (t, s) in hits.items()
                  if t == best_tier and s == best_score)
    if len(best) == 1 and best_tier <= 2:
        return ("ok", best)
    return ("conflict", best)


def match_resource(name, index):
    """按名称在资源索引中匹配（单次调用；批量场景用 build_matcher + matcher_match）。
    返回 (status, candidates)：ok / conflict / missing。
    第3级（子串包含，如 hycm ⊂ hycmjh）容易张冠李戴（红颜草莓精华 撞上
    红颜草莓），一律交人工选定。资源范围可先由 scope_index 限定。"""
    return matcher_match(build_matcher(index), name_variants(name))


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
    """直接改 xlsx 内目标工作表的 XML：只动目标单元格的值，zip 内其余文件
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
    n_si = 0           # si 项实际数量（可能有重复文本，不等于去重后的文本数）
    count_delta = 0    # 共享字符串引用总次数的变化（仅 非字符串->字符串 的单元格 +1）
    if ss_xml in data:
        ss_root = ET.fromstring(data[ss_xml])
        for si in ss_root:
            txt = _si_text(si)
            if txt not in ss_index:
                ss_index[txt] = n_si
            n_si += 1
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
            if cell.get("t") != "s":  # 空值/数字/内联字符串 → 共享字符串：引用数 +1
                count_delta += 1      # 旧值已是共享字符串的，替换引用，总数不变
            if val not in ss_index:
                si = ET.SubElement(ss_root, f"{{{OOXML_NS}}}si")
                te = ET.SubElement(si, f"{{{OOXML_NS}}}t")
                te.text = val
                if val != val.strip():
                    te.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                ss_index[val] = n_si + n_new_strings  # 追加在末尾，索引从实际 si 数起
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
    if ss_root is not None and (n_new_strings or count_delta):
        if ss_root.get("count") is not None:
            ss_root.set("count", str(int(ss_root.get("count")) + count_delta))
        ss_root.set("uniqueCount", str(n_si + n_new_strings))  # si 项实际总数
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
    matcher = build_matcher(scoped)
    if scope_prefixes or scope_suffixes:
        all_rels = {rel for rels in index.values() for rel in rels}
        in_rels = {rel for rels in scoped.values() for rel in rels}
        scoped_out = sorted(all_rels - in_rels)
    else:
        scoped_out = []
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise KeyError(f"工作表「{sheet}」不存在于 {Path(path).name}")
        ws = wb[sheet]
        hdr_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None) or ()
        hdr = {}
        for i, h in enumerate(hdr_row, 1):
            h = norm(h)
            if h and h not in hdr:
                hdr[h] = i
        if name_col not in hdr or value_col not in hdr:
            raise KeyError(f"表头里找不到列「{name_col}」或「{value_col}」")
        nc, vc = hdr[name_col], hdr[value_col]
        plan = {"ok": [], "same": [], "conflict": [], "missing": [], "overwrite": [],
                "skipped": {"empty": 0, "formula": 0}, "old_values": {},
                "scoped_out": scoped_out, "hit_rels": set(), "vc": vc}
        for row in ws.iter_rows(min_row=2):
            if not row:
                continue
            r = row[0].row
            name = norm(row[nc - 1].value if len(row) >= nc else None)
            if not name:
                plan["skipped"]["empty"] += 1
                continue
            cell = row[vc - 1] if len(row) >= vc else None
            if cell is not None and is_formula_cell(cell):
                plan["skipped"]["formula"] += 1
                continue
            old = norm(cell.value) if cell is not None else ""
            plan["old_values"][r] = old
            status, cands = matcher_match(matcher, name_variants(name))
            plan["hit_rels"].update(cands)
            if status == "ok":
                val = format_value(cands[0], fmt, keep_ext, prefix)
                if old == norm(val):
                    plan["same"].append((r, name))
                else:
                    plan["ok"].append((r, name, val, cands[0]))
                    if old:
                        plan["overwrite"].append((r, name, old, val))
            elif status == "conflict":
                # 候选的写入值全部相同时（如同名文件分处不同目录）不算冲突
                vals = {format_value(c, fmt, keep_ext, prefix) for c in cands}
                if len(vals) == 1:
                    val = vals.pop()
                    if old == norm(val):
                        plan["same"].append((r, name))
                    else:
                        plan["ok"].append((r, name, val, cands[0]))
                        if old:
                            plan["overwrite"].append((r, name, old, val))
                elif old in {norm(v) for v in vals}:
                    # 现值已是候选之一，无需处理
                    plan["same"].append((r, name))
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
    wb = openpyxl.load_workbook(path, read_only=True)
    try:
        if sheet not in wb.sheetnames:
            raise KeyError(f"工作表「{sheet}」不存在于 {Path(path).name}")
        ws = wb[sheet]
        hdr_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None) or ()
        hdr = {}
        for i, h in enumerate(hdr_row, 1):
            h = norm(h)
            if h and h not in hdr:
                hdr[h] = i
        if name_col not in hdr or value_col not in hdr:
            raise KeyError(f"表头里找不到列「{name_col}」或「{value_col}」")
        nc, vc = hdr[name_col], hdr[value_col]
        final_vals, names = {}, {}
        for row in ws.iter_rows(min_row=2):
            if not row:
                continue
            r = row[0].row
            final_vals[r] = norm(row[vc - 1].value if len(row) >= vc else None)
            names[r] = norm(row[nc - 1].value if len(row) >= nc else None)
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

    # 应用人工指定（旧值取自 plan，不重复打开工作簿）
    resolved_rows = []  # (行号, 名称, 新值, 资源相对路径)
    leftover_conflict, leftover_missing = [], []
    vc = plan["vc"]
    old_of = plan["old_values"].get
    try:
        for r, name, cands in plan["conflict"]:
            pick = resolutions.get(r)
            if pick:
                val = format_value(pick, fmt, keep_ext, prefix)
                if norm(old_of(r, "")) == norm(val):
                    plan["same"].append((r, name))
                else:
                    resolved_rows.append((r, name, val, pick))
            else:
                leftover_conflict.append((r, name, cands))
        for r, name in plan["missing"]:
            pick = resolutions.get(r)
            if pick:
                val = format_value(pick, fmt, keep_ext, prefix)
                if norm(old_of(r, "")) == norm(val):
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
        L.append(f"【改动】共 {total} 处（其中覆盖已填 {len(overwrite)} 行）"
                 + (f"，含人工指定 {len(resolved_rows)} 行" if resolved_rows else ""))
        old_values = plan.get("old_values", {})
        for r, name, val, rel in changes[:200]:
            L.append(f"  行{r} 「{name}」: {old_values.get(r) or '（空）'} -> {val}（{rel}）")
        if len(changes) > 200:
            L.append(f"  …… 其余 {len(changes) - 200} 处省略，完整列表见「全部改动」窗口")
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
                _res.append(f"[已保存] {Path(path).name}（共 {total} 处改动，已按原始状态备份）")
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
        pass  # 没有需要关闭的文件句柄（数据来自 plan）

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

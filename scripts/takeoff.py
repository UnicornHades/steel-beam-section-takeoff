# -*- coding: utf-8 -*-
"""钢结构梁编号-截面对照表提取：从梁翻样图 DWG 批量提取梁截面并与构件清单核对。

四阶段流水线：prepare（复制DWG到ASCII目录）→ convert（DWG转DXF）→
extract（提取BH截面文字）→ report（生成对照表Excel）。

关键约束（不要改动，否则会回到失败状态）：
  1. DWG→DXF 必须用 Python subprocess 调用 accoreconsole.exe，
     绝不能通过 git-bash/MSYS 调用 —— /i、/s 参数会被路径转换破坏，命令不执行。
  2. 中文路径下的 DWG 转换不稳定，prepare 阶段必须先复制成 ASCII 序号文件名。
  3. 文字提取必须调用 dwg-offline-extract 技能的 extract_text.py，
     ezdxf 直接读 modelspace 读不到文字（文字在块引用 INSERT 里）。
  4. accoreconsole 一次只能开一个 DWG：用 /i 传 DWG，脚本里只写 DXFOUT + QUIT。
"""
import argparse
import glob
import json
import os
import re
import shutil
import subprocess
import sys
import time

# ---------- 默认配置（可用命令行覆盖） ----------
DEFAULT_DIR_GLOB = "*梁布置图及翻样图*"
DEFAULT_FLOOR_RE = r"筒内地上(.+?)层梁"
DEFAULT_CODE_RE = r"(I-[^\s]+)"          # 从 DWG 文件名提取构件编号
DEFAULT_MANIFEST_GLOB = "*附图构件清单.xlsx"
DWG_SCRIPT = '_.FILEDIA 0\n_.DXFOUT "{dxf}" 16\n_.QUIT\n'


def log(*args):
    print(*args, flush=True)


# ---------- 工具函数 ----------
def cn_to_int(s):
    """中文数字转整数（支持 一~九十九），失败返回 None。"""
    if s is None:
        return None
    s = str(s).strip()
    if s.isdigit():
        return int(s)
    digit = {"零": 0, "一": 1, "二": 2, "三": 3, "四": 4,
             "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
    try:
        if s == "十":
            return 10
        if "十" in s:
            a, b = s.split("十", 1)
            tens = digit.get(a, 1) if a else 1
            ones = digit.get(b, 0) if b else 0
            return tens * 10 + ones
        if len(s) == 1 and s in digit:
            return digit[s]
        # 逐位相加（如 "十九" 已处理；"二十三" 已处理）
        total = 0
        for ch in s:
            if ch in digit:
                total = total * 10 + digit[ch]
        return total if total else None
    except Exception:
        return None


def find_accoreconsole(explicit=None):
    """定位 accoreconsole.exe。优先用显式参数，再扫描常见 Autodesk 安装路径。"""
    if explicit and os.path.exists(explicit):
        return explicit
    candidates = []
    for drive in ("C:", "D:", "E:"):
        base = drive + os.sep + "Program Files"
        for ver in range(2026, 2014, -1):
            candidates.append(os.path.join(base, "Autodesk",
                              "AutoCAD %d" % ver, "accoreconsole.exe"))
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


def find_extractor(explicit=None):
    """定位 dwg-offline-extract 技能的 extract_text.py。"""
    if explicit and os.path.exists(explicit):
        return explicit
    home = os.path.expanduser("~")
    p = os.path.join(home, ".workbuddy", "skills", "dwg-offline-extract",
                     "scripts", "extract_text.py")
    return p if os.path.exists(p) else None


def load_mapping(work):
    with open(os.path.join(work, "mapping.json"), encoding="utf-8") as f:
        return json.load(f)


# ---------- 阶段 1：prepare ----------
def stage_prepare(args):
    os.makedirs(os.path.join(args.work, "dwg"), exist_ok=True)
    os.makedirs(os.path.join(args.work, "dxf"), exist_ok=True)
    os.makedirs(os.path.join(args.work, "txt"), exist_ok=True)

    dirs = sorted(glob.glob(os.path.join(args.base, args.dir_glob)))
    if not dirs:
        raise SystemExit("未找到楼层文件夹，检查 --base 与 --dir-glob。base=%s" % args.base)

    floor_re = re.compile(args.floor_re)
    code_re = re.compile(args.code_re)
    records = []
    seq = 0
    for d in dirs:
        dz = os.path.basename(d)
        m = floor_re.search(dz)
        floor = m.group(1) if m else dz
        for f in sorted(glob.glob(os.path.join(d, "CAD", "*.dwg"))):
            fn = os.path.basename(f)
            m2 = code_re.search(fn)
            code = m2.group(1) if m2 else ""
            seq += 1
            newname = "%04d.dwg" % seq
            dst = os.path.join(args.work, "dwg", newname)
            shutil.copy2(f, dst)
            records.append({
                "seq": seq, "file": newname, "src": f, "floor": floor,
                "floor_num": cn_to_int(floor),
                "code": code, "srcname": fn,
                "is_plan": ("布置图" in fn and "翻样图" not in fn),
            })

    with open(os.path.join(args.work, "mapping.json"), "w", encoding="utf-8") as f:
        json.dump(records, f, ensure_ascii=False, indent=2)

    log("[prepare] 复制 %d 个 DWG 到 %s（布置图 %d）" %
        (len(records), os.path.join(args.work, "dwg"),
         sum(1 for r in records if r["is_plan"])))
    return records


# ---------- 阶段 2：convert ----------
def stage_convert(args):
    exe = find_accoreconsole(args.accoreconsole)
    if not exe:
        raise SystemExit("未找到 accoreconsole.exe，请用 --accoreconsole 指定完整路径。")
    work = args.work
    dwgdir = os.path.join(work, "dwg")
    dxfdir = os.path.join(work, "dxf")
    os.makedirs(dxfdir, exist_ok=True)
    scr = os.path.join(work, "one.scr")
    logfile = os.path.join(work, "conv.log")

    dwgs = sorted(glob.glob(os.path.join(dwgdir, "*.dwg")))
    if not dwgs:
        raise SystemExit("没有待转换 DWG，请先跑 prepare 阶段。")

    ok = fail = 0
    t0 = time.time()
    with open(logfile, "w", encoding="utf-8") as lf:
        for i, dwg in enumerate(dwgs, 1):
            base = os.path.splitext(os.path.basename(dwg))[0]
            dxf = os.path.join(dxfdir, base + ".dxf")
            if os.path.exists(dxf) and os.path.getsize(dxf) > 1000:
                ok += 1
                continue
            dxf_fwd = dxf.replace("\\", "/")
            with open(scr, "w", encoding="ascii") as sf:
                sf.write(DWG_SCRIPT.format(dxf=dxf_fwd))
            try:
                # 关键：Python subprocess 直调，绕开 git-bash 参数转换
                subprocess.run([exe, "/i", dwg, "/s", scr],
                               stdout=subprocess.DEVNULL,
                               stderr=subprocess.DEVNULL,
                               timeout=args.timeout)
            except subprocess.TimeoutExpired:
                pass
            if os.path.exists(dxf) and os.path.getsize(dxf) > 1000:
                ok += 1
                status = "OK"
            else:
                fail += 1
                status = "FAIL"
            lf.write("%d/%d %s %s\n" % (i, len(dwgs), base, status))
            lf.flush()
            if i % 10 == 0 or i == len(dwgs):
                log("[convert] %d/%d (%s) 已用 %.0fs" %
                    (i, len(dwgs), status, time.time() - t0))
    log("[convert] 完成 OK=%d FAIL=%d 用时 %.0fs" % (ok, fail, time.time() - t0))
    return fail


# ---------- 阶段 3：extract ----------
def stage_extract(args):
    extractor = find_extractor(args.extractor)
    if not extractor:
        raise SystemExit("未找到 extract_text.py，请安装 dwg-offline-extract 技能或用 --extractor 指定。")
    work = args.work
    dxfdir = os.path.join(work, "dxf")
    txtdir = os.path.join(work, "txt")
    os.makedirs(txtdir, exist_ok=True)
    records = load_mapping(work)

    env = dict(os.environ)
    env["PYTHONUTF8"] = "1"
    results = []
    for r in records:
        if r["is_plan"]:
            continue
        dxf = os.path.join(dxfdir, r["file"].replace(".dwg", ".dxf"))
        rec = dict(r)
        rec.update({"bh": "", "bh_candidates": [], "piece_no": "", "text_count": 0})
        if os.path.exists(dxf):
            out = os.path.join(txtdir, r["file"].replace(".dwg", ".txt"))
            subprocess.run([sys.executable, extractor, dxf, out],
                           stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
                           env=env)
            if os.path.exists(out):
                texts = []
                for ln in open(out, encoding="utf-8", errors="ignore").read().splitlines():
                    if ln.startswith("#") or not ln.strip():
                        continue
                    parts = ln.split("\t")
                    if len(parts) >= 6:
                        texts.append(parts[-1].strip())
                cands = []
                for t in texts:
                    tc = t.replace(" ", "")
                    if re.match(args.bh_re, tc, re.I):
                        cands.append(tc)
                rec["bh_candidates"] = sorted(set(cands))
                rec["bh"] = cands[0] if cands else ""
                for t in texts:
                    if re.match(r"^I-\d", t.strip(), re.I):
                        rec["piece_no"] = t.strip()
                        break
                rec["text_count"] = len(texts)
        results.append(rec)

    with open(os.path.join(work, "sections.json"), "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=2)

    with_bh = sum(1 for r in results if r["bh"])
    log("[extract] 梁 %d 根，提取到截面 %d 根，无截面 %d 根" %
        (len(results), with_bh, len(results) - with_bh))
    for r in results:
        if not r["bh"]:
            log("  无截面: %s %s (文字数 %d)" % (r["floor"], r["code"], r["text_count"]))
    return results


# ---------- 阶段 4：report ----------
def stage_report(args):
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    work = args.work
    secs = json.load(open(os.path.join(work, "sections.json"), encoding="utf-8"))
    sec_map = {r["code"]: r for r in secs}
    excluded = set(x.strip() for x in args.excluded.split(",") if x.strip())

    dirs = glob.glob(os.path.join(args.base, args.dir_glob))
    dirs.sort(key=lambda p: (cn_to_int(
        (re.search(args.floor_re, os.path.basename(p)) or [None, "999"])[1]) or 999))

    rows_out = []
    for d in dirs:
        m = re.search(args.floor_re, os.path.basename(d))
        floor = m.group(1) if m else ""
        floor_num = cn_to_int(floor) or 0
        for xl in glob.glob(os.path.join(d, args.manifest_glob)):
            wb = openpyxl.load_workbook(xl, read_only=True, data_only=True)
            ws = wb.active
            for r in list(ws.iter_rows(values_only=True))[1:]:
                if len(r) <= max(args.qty_col, args.weight_col) or not r[args.code_col]:
                    continue
                code = str(r[args.code_col]).strip()
                if "合计" in code or code == "":
                    continue
                src = sec_map.get(code, {})
                bh = src.get("bh", "")
                if code in excluded:
                    note = "截面非标准H型钢（非BH），已确认排除"
                    bh = "—"
                elif not bh:
                    note = "图中未找到BH主材行，需人工复核"
                elif len(src.get("bh_candidates", [])) > 1:
                    note = "多规格:" + "/".join(src["bh_candidates"])
                else:
                    note = ""
                rows_out.append({
                    "floor": floor_num, "code": code, "bh": bh,
                    "qty": r[args.qty_col], "weight": r[args.weight_col],
                    "drawing": src.get("srcname", ""), "note": note,
                })

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "梁编号-截面对照表"
    headers = ["楼层", "构件编号(梁编号)", "梁截面(BH规格)", "数量", "单重(kg)", "来源图纸", "备注"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(1, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="305496")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in rows_out:
        ws.append([row["floor"], row["code"], row["bh"], row["qty"],
                   row["weight"], row["drawing"], row["note"]])
        ri = ws.max_row
        for c in range(1, len(headers) + 1):
            ws.cell(ri, c).border = border
            ws.cell(ri, c).alignment = Alignment(vertical="center", wrap_text=(c == 7))
        if not row["bh"] or "复核" in row["note"]:
            for c in range(1, len(headers) + 1):
                ws.cell(ri, c).fill = PatternFill("solid", fgColor="FFF2CC")
    for i, w in enumerate([8, 22, 24, 8, 12, 55, 34], 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w
    ws.freeze_panes = "A2"

    ws2 = wb.create_sheet("未提取到截面")
    ws2.append(["楼层", "构件编号", "来源图纸", "说明"])
    for c in range(1, 5):
        cell = ws2.cell(1, c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="C00000")
    for row in rows_out:
        if "复核" in row["note"]:
            ws2.append([row["floor"], row["code"], row["drawing"], row["note"]])
    for i, w in enumerate([8, 22, 55, 44], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    out_dir = os.path.dirname(os.path.abspath(args.out))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    wb.save(args.out)
    log("[report] 已生成 %s：共 %d 根，有截面 %d，待复核 %d" %
        (args.out, len(rows_out), sum(1 for r in rows_out if r["bh"] and "复核" not in r["note"]),
         sum(1 for r in rows_out if "复核" in r["note"])))
    return args.out


def build_parser():
    p = argparse.ArgumentParser(description="钢结构梁编号-截面对照表提取")
    p.add_argument("--stage", choices=["all", "prepare", "convert", "extract", "report"],
                   default="all", help="执行阶段，默认 all 全流程")
    p.add_argument("--base", help="图纸根目录（内含各层终审单文件夹）")
    p.add_argument("--work", required=True, help="工作目录，建议纯ASCII路径（存 dwg/dxf/txt/中间json）")
    p.add_argument("--out", help="输出对照表 xlsx 路径（report 阶段必需）")
    p.add_argument("--accoreconsole", help="accoreconsole.exe 完整路径（默认自动探测）")
    p.add_argument("--extractor", help="extract_text.py 完整路径（默认自动探测 dwg-offline-extract 技能）")
    p.add_argument("--timeout", type=int, default=90, help="单文件转换超时秒数，默认90")
    p.add_argument("--dir-glob", default=DEFAULT_DIR_GLOB, help="楼层文件夹匹配模式")
    p.add_argument("--floor-re", default=DEFAULT_FLOOR_RE, help="从文件夹名提取楼层的正则")
    p.add_argument("--code-re", default=DEFAULT_CODE_RE, help="从DWG文件名提取构件编号的正则")
    p.add_argument("--manifest-glob", default=DEFAULT_MANIFEST_GLOB, help="构件清单xlsx匹配模式")
    p.add_argument("--code-col", type=int, default=3, help="清单中构件编号列(0-based)，默认3")
    p.add_argument("--qty-col", type=int, default=4, help="清单中数量列(0-based)，默认4")
    p.add_argument("--weight-col", type=int, default=5, help="清单中单重列(0-based)，默认5")
    p.add_argument("--bh-re", default=r"^BH\d+[\*xX×]\d+", help="截面规格正则，默认匹配 BH数字*数字")
    p.add_argument("--excluded", default="", help="逗号分隔的已确认排除构件编号（非标准H型钢）")
    return p


def main():
    args = build_parser().parse_args()
    stages = ["prepare", "convert", "extract", "report"] if args.stage == "all" else [args.stage]
    if "prepare" in stages and not args.base:
        raise SystemExit("prepare 阶段需要 --base")
    if "report" in stages and not (args.base and args.out):
        raise SystemExit("report 阶段需要 --base 和 --out")
    for s in stages:
        {"prepare": stage_prepare, "convert": stage_convert,
         "extract": stage_extract, "report": stage_report}[s](args)
    log("[done] 全部阶段完成")


if __name__ == "__main__":
    main()

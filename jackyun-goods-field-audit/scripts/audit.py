# -*- coding: utf-8 -*-
"""吉客云货品档案 · 字段信息抽查引擎 v1.1

按流星封版口径实现规则 1-8：只筛疑点，不自动判错。
v1.1（2026-10-22）：规则1 属性码放宽为结构校验（不限 XL/SL，JL 等亦可）；
                    规则2 计价商品银重禁空白/文字/区间（如 18-20），工费禁空白。

用法:
    python audit.py --src "10-15-2货品导出.xlsx"
    python audit.py --src "货品导出.xlsx" --out "疑点清单.xlsx" [--sheet sheetTitle]

默认输出: 与被核查表同目录、同名的 `<被核查表名>_抽查疑点清单_<YYYYMMDD>.xlsx`
输出含 7 个 sheet: 疑点清单 / 汇总统计 / 条码形态待识别 / 属性码未收录 /
                 天然石词表边界 / 手镯无重量信息 / 规则说明
"""
import argparse, collections, datetime, os, re, sys

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ============================ 字段定位 ============================
# 优先按表头名匹配，匹配不到时回退到默认列序（0-based）
FIELD_ALIASES = {
    "条码":     ["条码", "货品条码"],
    "规格":     ["规格", "规格名称"],
    "货品编号": ["货品编号", "货品编号(规)"],
    "货品名称": ["货品名称", "货品名"],
    "成分":     ["成分", "计价方式"],
    "银重":     ["银重(规)", "银重"],
    "工费":     ["工费价格(规)", "工费价格", "工费"],
    "字印":     ["字印要求(规)", "字印要求"],
    "分类":     ["分类", "分类名称"],
    "抽检":     ["抽检比例(规)", "抽检比例", "抽检"],
}
FIELD_DEFAULT_IDX = {"条码": 0, "规格": 1, "货品编号": 2, "货品名称": 3, "成分": 4,
                     "银重": 5, "工费": 6, "字印": 8, "分类": 13, "抽检": 15}


def resolve_columns(header):
    """表头 → 字段名到列索引的映射"""
    norm = [str(h or "").strip() for h in header]
    idx = {}
    for field, aliases in FIELD_ALIASES.items():
        idx[field] = FIELD_DEFAULT_IDX[field]
        for a in aliases:
            if a in norm:
                idx[field] = norm.index(a)
                break
    return idx


# ============================ 业务常量 ============================
# ★ 属性码不限定枚举（2026-10-22 放宽）：只要货品编号满足
#   属性2位 + 产品代码5位 + 出条码年月4位 + 品牌代号1位 = 12 位，即视为合法。
#   ATTR_NAME 只用于「编号属性 ↔ 分类」的中文比对；未收录的码不做该比对，
#   仅登记到「属性码未收录」sheet，供人工补表（不算疑点）。
ATTR_NAME = {"XL": "项链", "SL": "手链", "ES": "耳饰", "JZ": "戒指", "SZ": "手镯",
             "DZ": "吊坠", "BC": "包材", "YL": "银礼", "GC": "金礼", "GG": "金礼",
             "JL": "脚链"}
ATTRS = set(ATTR_NAME)

# 编号结构：属性2 + 产品代码5 + 出条码年月4 + 品牌代号1
CODE_PAT = re.compile(r"^([A-Z]{2})([A-Z0-9]{5})(\d{4})([A-Z])$")

# 品牌：规格【品牌】↔ 货品编号末位代号 ↔ 条码品牌标记
BRAND_CODE = {"中央": "K", "唯一": "A", "老鳳祥": "F", "老凤祥": "F", "中国珠宝": "R",
              "珍尚银": "Y", "IU": "B", "TY": "T", "萃华": "H", "周大金": "Z"}
CODE_BRAND = {v: k for k, v in BRAND_CODE.items()}
MARK2BRAND = {"ZY": "中央", "FX": "老鳳祥", "SY": "珍尚银", "IU": "IU",
              "CH": "萃华", "DJ": "周大金", "ZZ": "中国珠宝", "TY": "TY"}
# ★唯一 与 TY 为相等信息：编号末位代号 A / T 互通，条码标记亦互通
BRAND_CODE_ALIAS = {"唯一": {"A", "T"}, "TY": {"A", "T"}}

NATURE_SUF = ("WC", "CK")      # WC 外采商品 / CK 工厂测款
EXTRA_SUF = ("SH", "ZZ", "TY")

# 品牌组 → 必须含「配国检证书」；非品牌组 → 不得出现「配国检证书」，按材质判「配证书」
BRAND_GROUP_CERT = {"中央", "老鳳祥", "老凤祥", "中国珠宝", "珍尚银", "萃华"}
NONBRAND_GROUP = {"唯一", "IU", "TY", "周大金"}   # ★周大金不需要配国检证书

SKIP_RULE_ATTRS = {"BC"}       # 包材整行免检

STONES = ["和田玉", "玛瑙", "玉髓", "青金石", "芙蓉石", "蜜蜡", "木变石", "翡翠", "水晶",
          "发晶", "欧泊", "海蓝宝", "石榴石", "长石", "葡萄石", "尖晶石", "碧玺", "孔雀石"]

# 银含量等级归一（9999 ≡ 999.9）
YIN_PAT = [(r"足银\s*999\.9", "9999"), (r"足银\s*9999", "9999"), (r"足银\s*999", "999"),
           (r"足银", "足银"), (r"(?:s|S|ag|AG)\s*925", "925"), (r"925\s*银", "925"),
           (r"银\s*925", "925"), (r"925", "925")]


# ============================ 工具函数 ============================
def s(v):
    if v is None:
        return ""
    return str(v).replace("\n", " ").replace("\r", " ").strip()


def yin_level(txt):
    for p, lv in YIN_PAT:
        if re.search(p, txt, re.I):
            return lv
    return ""


def fmt_num(x):
    return str(int(x)) if abs(x - round(x)) < 1e-9 else str(round(x, 4))


# ---------- 规则2 计价商品的数值字段判定 ----------
NUM_RE = re.compile(r"^\d+(?:\.\d+)?$")
RANGE_RE = re.compile(r"^\d+(?:\.\d+)?\s*[-~—－–至到]\s*\d+(?:\.\d+)?$")


def check_amount(val):
    """计价商品的『银重(规)』『工费价格(规)』判定。
    ★ 只接受单一数字；空白、区间（18-20）、文字备注一律为疑点。
    返回 (是否合规, 问题描述)"""
    if not val:
        return False, "不得空白，必须填写数字"
    if NUM_RE.match(val):
        return True, ""
    if RANGE_RE.match(val):
        return False, f"不得填写区间（{val}），必须填单一数字"
    return False, f"不得出现文字备注（{val}），必须为纯数字"


def spec_weight(sp):
    """规格里的重量 → (期望银重, 说明)。区间取中间值，唯一值取该值；无重量信息返回 (None, "")"""
    m = re.search(r"(\d+(?:\.\d+)?)\s*[-~—～至]\s*(\d+(?:\.\d+)?)\s*(?:g|G|克)", sp)
    if m:
        a, b = float(m.group(1)), float(m.group(2))
        if a > b:
            a, b = b, a
        return (a + b) / 2, f"规格区间 {fmt_num(a)}-{fmt_num(b)}g → 银重应为中间值 {fmt_num((a+b)/2)}g（区间外或无值均不合格）"
    m2 = re.search(r"(\d+(?:\.\d+)?)\s*(?:g|G|克)", sp)
    if m2:
        v = float(m2.group(1))
        return v, f"规格唯一重量 {fmt_num(v)}g → 银重应写 {fmt_num(v)}g"
    return None, ""


def brand_of(sp):
    m = re.search(r"【(.+?)】", sp)
    return m.group(1) if m else "(无标签)"


def prodname_of(sp):
    m = re.search(r"【.+?】(.+?)[（{(]", sp)
    return m.group(1).strip() if m else ""


def parse_code(code):
    """按结构解析 12 位货品编号：属性2 + 产品代码5 + 年月4 + 品牌代号1。
    ★ 不限定属性码枚举（XL/SL/JZ/SZ/DZ/JL/GG… 均可），结构成立即合法。
    返回 (属性码, 品牌代号)，结构不符返回 (None, None)。"""
    m = CODE_PAT.match(code or "")
    if not m:
        return None, None
    return m.group(1), m.group(4)


def parse_barcode(bc):
    """多模式解析条码 → dict(mark, attr, nature, core)；无法识别返回 None"""
    if not bc:
        return None
    b = bc
    nature = ""
    for n in NATURE_SUF:
        if b.endswith(n):
            nature, b = n, b[:-len(n)]
            break
    mark = ""
    for x in ("ZY", "FX", "SY", "IU", "CH", "DJ"):
        if b.startswith(x):
            mark, b = x, b[len(x):]
            break
    for x in EXTRA_SUF:
        if b.endswith(x):
            if not mark:
                mark = x
            b = b[:-len(x)]
            break
    b2 = re.sub(r"[Bb][Cc]\d+$", "", b)
    m = re.fullmatch(r"\d+([A-Z]{2})\d+", b2)
    # ★ 属性码不限枚举：两侧被数字包围的 2 位大写字母即视为属性码（JL 等新码同样适用）
    attr = m.group(1) if m else ""
    if not attr and not mark:
        return None
    return {"mark": mark, "attr": attr, "nature": nature, "core": b}


# ============================ 主流程 ============================
def run(src, out=None, sheet=None):
    wb = openpyxl.load_workbook(src, data_only=True)
    ws = wb[sheet] if sheet else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit("[ERROR] 表是空的")
    col = resolve_columns(rows[0])
    data = rows[1:]
    print(f"[i] sheet={ws.title} 数据行={len(data)}")
    print(f"[i] 字段列定位: " + ", ".join(f"{k}={chr(65+v)}" for k, v in sorted(col.items(), key=lambda x: x[1])))

    issues, unparsed, skipped, boundary, noweight = [], [], [], [], []
    unknown_attr = []          # 编号属性码未收录中文名 → 只登记，不算疑点
    stats = collections.Counter()

    for i, r in enumerate(data):
        rowno = i + 2

        def g(f):
            k = col[f]
            return s(r[k]) if k < len(r) else ""

        bc, sp, code, name = g("条码"), g("规格"), g("货品编号"), g("货品名称")
        cheng, yin, gong = g("成分"), g("银重"), g("工费")
        ziyin, cls, tj = g("字印"), g("分类"), g("抽检")

        brand = brand_of(sp)
        code_attr, code_brandletter = parse_code(code)
        pbc = parse_barcode(bc)
        if pbc is None:
            unparsed.append((rowno, bc, code, brand, name))

        def add(rule, rname, field, val, exp):
            issues.append([rowno, bc, code, brand, rule, rname, field, val, exp, ""])
            stats[f"{rule} {rname}"] += 1

        # 包材整行免检
        if code[:2] in SKIP_RULE_ATTRS or cls == "包材":
            skipped.append((rowno, bc, code, brand, name))
            continue

        # ---------- 规则1 四列属性与品牌一致 ----------
        if not code or len(code) != 12 or code_attr is None:
            add("规则1", "属性/品牌四列一致", "货品编号", code,
                "应为 12 位：属性2位 + 产品代码5位 + 出条码年月4位 + 品牌代号1位（属性码不限 XL/SL，JL 等亦可）")
        else:
            attr_expect = code_attr
            if attr_expect not in ATTR_NAME:
                # 属性码结构合法但未收录中文名 → 不做分类比对，只登记待补
                unknown_attr.append((rowno, bc, code, attr_expect, cls, brand))
            elif cls and cls != ATTR_NAME[attr_expect]:
                add("规则1", "属性/品牌四列一致", "编号属性↔分类",
                    f"编号前2位 {attr_expect} / 分类 {cls}",
                    f"编号前2位 {attr_expect}({ATTR_NAME[attr_expect]}) 应与分类一致")
            if pbc is not None and pbc["attr"] and pbc["attr"] != attr_expect:
                add("规则1", "属性/品牌四列一致", "条码属性↔编号属性",
                    f"条码含 {pbc['attr']} / 编号前2位 {attr_expect}", "两者应为同一产品属性")
            if brand in BRAND_CODE:
                allowed = BRAND_CODE_ALIAS.get(brand, {BRAND_CODE[brand]})
                if code_brandletter not in allowed:
                    exp_txt = BRAND_CODE[brand] if len(allowed) == 1 else " / ".join(sorted(allowed))
                    add("规则1", "属性/品牌四列一致", "编号末位↔规格品牌",
                        f"末位 {code_brandletter} / 规格【{brand}】", f"末位代号应为 {exp_txt}（{brand}）")
            else:
                add("规则1", "属性/品牌四列一致", "规格品牌", brand,
                    "品牌不在已知代号表（中央K/唯一A/老凤祥F/中国珠宝R/珍尚银Y/IU-B/TY-T/萃华H/周大金Z），待人工")
            if pbc is not None and pbc["mark"]:
                mk_brand = MARK2BRAND.get(pbc["mark"], "")
                same = (mk_brand == brand) or ({mk_brand, brand} <= {"唯一", "TY"})
                if brand in BRAND_CODE and mk_brand and not same:
                    add("规则1", "属性/品牌四列一致", "条码标记↔规格品牌",
                        f"条码标记 {pbc['mark']}={mk_brand} / 规格【{brand}】", "两者应为同一品牌（唯一 与 TY 视为等同）")

        # ---------- 规则2 成分 ↔ 银重/工费 ----------
        # ★ 计价：银重(规) 必须有单一数字（文字备注、区间 18-20、空白均为疑点）
        #          工费价格(规) 不得空白（区间/文字同判）
        # ★ 一口价：整条豁免
        if cheng not in ("计价", "一口价"):
            add("规则2", "成分与价格", "成分", cheng or "(空)", "只能为『计价』或『一口价』两种选择")
        elif cheng == "计价":
            ok, msg = check_amount(yin)
            if not ok:
                add("规则2", "成分与价格", "银重(规)", yin or "(空)", "计价商品银重" + msg)
            ok, msg = check_amount(gong)
            if not ok:
                add("规则2", "成分与价格", "工费价格(规)", gong or "(空)", "计价商品工费" + msg)

        # ---------- 规则3 货品名称 ↔ 证书措辞 ----------
        lv = yin_level(sp)
        has_stone = any(x in sp for x in STONES)
        spec_gj = "配国检证书" in sp
        spec_pz = ("配证书" in sp) and not spec_gj
        name_gj = "配国检证书" in name
        name_pz = ("配证书" in name) and not name_gj
        if brand in BRAND_GROUP_CERT:
            if not name_gj:
                add("规则3", "货品名称证书措辞", "货品名称", name, f"品牌【{brand}】必须含『配国检证书』")
        elif brand in NONBRAND_GROUP:
            if spec_gj or name_gj:
                add("规则3", "货品名称证书措辞", "货品名称", name,
                    f"非品牌组【{brand}】不应出现『配国检证书』（只有品牌组写国检证书）")
            elif spec_pz:
                if not name_pz:
                    add("规则3", "货品名称证书措辞", "货品名称", name, "规格已写『配证书』，货品名称应体现『配证书』")
            else:
                need = has_stone or (lv in ("999", "9999", "足银"))
                reason = f"材质 {lv}" + ("+天然石" if has_stone else "")
                if need and not name_pz:
                    add("规则3", "货品名称证书措辞", "货品名称", name, f"应含『配证书』（依据：{reason}）")
                if (not need) and name_pz:
                    add("规则3", "货品名称证书措辞", "货品名称", name, f"不应写『配证书』（依据：{reason}，无天然石）")
        else:
            if brand != "(无标签)":
                add("规则3", "货品名称证书措辞", "规格品牌", brand, "品牌未纳入证书规则，证书措辞待人工确认")
        # 天然石词表边界提示
        if not has_stone and re.search(r"(玉|晶|珀|翠|砂|珍珠|贝壳|猫眼|托帕|松石)", sp) and "钻石" not in sp:
            boundary.append((rowno, bc, brand, sp[:70]))

        # ---------- 规则4 银含量三列统一 ----------
        lv_sp, lv_nm, lv_zy = yin_level(sp), yin_level(name), yin_level(ziyin)
        lvs = [lv_sp, lv_nm, lv_zy]
        if all(v == "" for v in lvs):
            pass                                  # 三列均未写 = 本就不含银 → 免检
        elif len(set(lvs)) > 1:
            add("规则4", "银含量三列统一", "规格/货品名称/字印要求",
                f"规格={lv_sp or '未写'}｜货品名称={lv_nm or '未写'}｜字印={lv_zy or '未写'}",
                "三列银含量等级应一致（写法可不同，等级须同）")

        # ---------- 规则5 产品名一致 ----------
        pn = prodname_of(sp)
        if pn:
            if pn not in name:
                add("规则5", "产品名一致", "货品名称", name, f"应包含规格产品名『{pn}』")
        else:
            add("规则5", "产品名一致", "规格", sp[:40], "规格未按『【品牌】产品名（…）』格式书写，无法提取产品名")

        # ---------- 规则6 字印要求前缀 ----------
        y = ziyin
        if "WC" in bc or "CK" in bc:
            pass                                  # 外采/工厂测款不查字印
        elif not y:
            if lv:                                # 有银含量 = 金属材质，字印不该为空
                add("规则6", "字印要求前缀", "字印要求(规)", "(空)", f"品牌【{brand}】字印要求为空")
        else:
            ok, exp = True, ""
            if brand in ("老鳳祥", "老凤祥"):
                ok, exp = (y.startswith("老鳳祥") or y.startswith("老凤祥")), "老鳳祥"
            elif brand == "中国珠宝":
                ok, exp = y.startswith("中珠"), "中珠"
            elif brand == "珍尚银":
                ok, exp = y.startswith("中国黄金"), "中国黄金"
            elif brand == "周大金":
                ok, exp = y.startswith("周大金"), "周大金"
            elif brand in ("唯一", "IU", "TY"):
                ok, exp = any(y.startswith(x) for x in ("FY", "MY", "ZZ", "zz", "ZG", "YG", "足银")), "工厂代号(FY/MY/ZZ/足银 等)"
            elif brand == "萃华":
                ok, exp = y.startswith("CH"), "CH"
            elif brand == "中央":
                ok = bool(re.fullmatch(r"(足银\s*999\.9|足银\s*9999|足银\s*999|足银|s925|S925|925银|925)[\s\w]*", y, re.I)) \
                     and not re.search(r"(LOGO|中珠|老鳳祥|FY|MY)", y)
                exp = "只有银含量描述"
            else:
                ok, exp = False, f"品牌【{brand}】的字印前缀规则未定义，待人工"
            if not ok:
                add("规则6", "字印要求前缀", "字印要求(规)", y, f"应以『{exp}』开头")

        # ---------- 规则7 抽检比例不得为空 ----------
        if not tj:
            add("规则7", "抽检比例非空", "抽检比例(规)", "(空)", "抽检比例列必须有内容，不得为空")

        # ---------- 规则8 手镯银重 = 规格重量中间值 / 唯一值 ----------
        if code_attr == "SZ" or cls == "手镯":
            exp_w, why = spec_weight(sp)
            if exp_w is None:
                noweight.append((rowno, bc, code, brand, sp[:60], yin))
            else:
                try:
                    actual = float(yin)
                except (TypeError, ValueError):
                    actual = None
                if actual is None or abs(actual - exp_w) > 0.005:
                    add("规则8", "手镯银重取中间值", "银重(规)", yin or "(空)", why)

    write_report(out or default_out(src), issues, stats, data, unparsed, skipped,
                 boundary, noweight, unknown_attr)
    return issues, stats


# ============================ 输出 ============================
def default_out(src):
    stem, _ = os.path.splitext(os.path.basename(src))
    return os.path.join(os.path.dirname(os.path.abspath(src)),
                        f"{stem}_抽查疑点清单_{datetime.date.today():%Y%m%d}.xlsx")


def write_report(OUT, issues, stats, data, unparsed, skipped, boundary, noweight, unknown_attr):
    thin = Side(style="thin", color="D0D0D0")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    HFILL = PatternFill("solid", fgColor="1F3864")
    HFONT = Font(color="FFFFFF", bold=True, size=11)

    def header(ws, ncol):
        ws.freeze_panes = "A2"
        for c in range(1, ncol + 1):
            cell = ws.cell(1, c)
            cell.fill, cell.font, cell.border = HFILL, HFONT, bd
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.row_dimensions[1].height = 26

    out = openpyxl.Workbook()
    # Sheet1 疑点清单
    w1 = out.active
    w1.title = "疑点清单"
    H = ["行号", "条码", "货品编号", "规格品牌", "命中规则", "规则名称",
         "问题字段", "实际值", "期望/说明", "人工结论"]
    w1.append(H)
    for row in issues:
        w1.append(row)
    for i, width in enumerate([6, 18, 15, 10, 7, 24, 16, 42, 50, 14], start=1):
        w1.column_dimensions[chr(64 + i)].width = width
    header(w1, len(H))
    for r in range(2, w1.max_row + 1):
        for c in range(1, len(H) + 1):
            cell = w1.cell(r, c)
            cell.border = bd
            cell.alignment = Alignment(vertical="top", wrap_text=(c in (8, 9)))
    w1.auto_filter.ref = w1.dimensions

    # Sheet2 汇总统计
    w2 = out.create_sheet("汇总统计")
    w2.append(["规则", "命中条数"])
    for k, v in sorted(stats.items()):
        w2.append([k, v])
    w2.append(["—— 合计（行·次）", ""])
    w2.cell(w2.max_row, 2).value = len(issues)
    w2.append(["涉及行数（去重）", len(set(x[0] for x in issues))])
    w2.append(["总数据行", len(data)])
    w2.append(["无任何疑点的行数", len(data) - len(set(x[0] for x in issues))])
    w2.append(["包材整行免检（属性码 BC）", len(skipped)])
    w2.append(["手镯无重量信息（规则8 跳过）", len(noweight)])
    w2.append(["条码形态无法识别", len(unparsed)])
    w2.append(["属性码未收录中文名（不参与分类比对）", len(unknown_attr)])
    w2.column_dimensions["A"].width = 34
    w2.column_dimensions["B"].width = 14
    header(w2, 2)

    # Sheet3 条码形态待识别
    w3 = out.create_sheet("条码形态待识别")
    w3.append(["行号", "条码", "货品编号", "规格品牌", "货品名称"])
    for x in unparsed:
        w3.append(list(x))
    for i, wd in enumerate([6, 18, 15, 10, 40], start=1):
        w3.column_dimensions[chr(64 + i)].width = wd
    header(w3, 5)

    # Sheet4 属性码未收录（编号结构合法，但属性码没有对应的中文分类名）
    w3b = out.create_sheet("属性码未收录", 3)
    w3b.append(["行号", "条码", "货品编号", "属性码", "分类", "规格品牌"])
    for x in unknown_attr:
        w3b.append(list(x))
    for i, wd in enumerate([6, 18, 15, 9, 12, 10], start=1):
        w3b.column_dimensions[chr(64 + i)].width = wd
    header(w3b, 6)
    note_r = w3b.max_row + 2
    w3b.cell(note_r, 1).value = (
        "说明：以下属性码在编号结构上合法（属性2位+代码5位+年月4位+品牌1位），"
        "但未收录中文分类名，因此未做「编号属性↔分类」比对，不计入疑点。"
        "如需纳入比对，请在 audit.py 的 ATTR_NAME 中补充「属性码: 分类中文名」。")
    w3b.cell(note_r, 1).font = Font(italic=True, color="808080", size=9)

    # Sheet5 天然石词表边界
    w4 = out.create_sheet("天然石词表边界")
    w4.append(["行号", "条码", "规格品牌", "规格（疑似含天然石但不在清单）"])
    for x in boundary:
        w4.append(list(x))
    for i, wd in enumerate([6, 18, 10, 70], start=1):
        w4.column_dimensions[chr(64 + i)].width = wd
    header(w4, 4)

    # Sheet5 手镯无重量信息
    w5 = out.create_sheet("手镯无重量信息")
    w5.append(["行号", "条码", "货品编号", "规格品牌", "规格", "银重(规)"])
    for x in noweight:
        w5.append(list(x))
    for i, wd in enumerate([6, 18, 15, 10, 62, 12], start=1):
        w5.column_dimensions[chr(64 + i)].width = wd
    header(w5, 6)

    # Sheet6 规则说明
    by_rule = collections.Counter()
    for k, v in stats.items():
        by_rule[k.split()[0]] += v
    w6 = out.create_sheet("规则说明")
    w6.append(["规则", "名称", "判定口径（命中即入疑点清单，不自动判错）", "命中条数"])
    for x in [
        ("规则1", "属性/品牌四列一致",
         "货品编号结构：属性2位 + 产品代码5位 + 出条码年月4位 + 品牌代号1位 = 12 位；"
         "属性码不限枚举（XL 项链/SL 手链/JZ 戒指/SZ 手镯/ES 耳饰/DZ 吊坠/JL 脚链/GG金礼…均可），"
         "只要结构成立即合法；未收录中文名的属性码只登记不判错。"
         "条码属性 = 编号前2位 = 分类对应属性；条码品牌标记 ↔ 编号末位代号 ↔ 规格【品牌】。数字段一律不校。"
         "★唯一 与 TY 视为同品牌信息，末位代号 A / T 互通、条码标记亦互通。", by_rule.get("规则1", 0)),
        ("规则2", "成分与价格",
         "成分只有『计价』『一口价』两种。计价商品：银重(规) 必须是单一数字 —— "
         "空白、文字备注、区间（如 18-20）均为疑点；工费价格(规) 不得空白（区间/文字同判）。"
         "一口价商品整条豁免。", by_rule.get("规则2", 0)),
        ("规则3", "货品名称证书措辞",
         "品牌组（中央/老凤祥/珍尚银/中国珠宝/萃华）须含『配国检证书』；"
         "非品牌组（唯一/IU/TY/周大金）不得出现『配国检证书』；规格已写『配证书』则货品名称须体现；"
         "否则足银999系或含天然石须含『配证书』，925且无天然石不得写。合成立方氧化锆不算天然石。", by_rule.get("规则3", 0)),
        ("规则4", "银含量三列统一",
         "规格 / 货品名称 / 字印要求 三列银含量等级须一致（s925=S925=925银；9999=999.9）；"
         "三列均未写 = 本就不含银，免检。", by_rule.get("规则4", 0)),
        ("规则5", "产品名一致",
         "货品名称 = 规格产品名 + 银含量 + 证书后缀（规格【品牌】后、首个括号前的产品名须出现在货品名称中）。", by_rule.get("规则5", 0)),
        ("规则6", "字印要求前缀",
         "老凤祥→老鳳祥｜中国珠宝→中珠｜珍尚银→中国黄金｜周大金→周大金｜唯一/IU/TY→工厂代号(FY/MY/ZZ/足银 等)｜"
         "萃华→CH｜中央→只有银含量。条码带 WC/CK（外采/测款）整项免检；无银含量（非金属材质）字印可为空。", by_rule.get("规则6", 0)),
        ("规则7", "抽检比例非空", "抽检比例(规) 必须有内容，为空即入疑点清单。", by_rule.get("规则7", 0)),
        ("规则8", "手镯银重取中间值",
         "手镯（属性码 SZ / 分类『手镯』）：规格写重量区间（如 11-13g）时银重须为中间值（12）；"
         "规格写唯一值（如 12g）时须填该值；规格无重量信息（玉髓等）跳过。", by_rule.get("规则8", 0)),
    ]:
        w6.append(list(x))
    w6.append([])
    w6.append(["附：品牌 — 编号末位代号 — 条码标记"])
    for x in ["中央 K（条码前缀 ZY）", "唯一 A（条码无标记）", "老凤祥 F（条码前缀 FX）",
              "中国珠宝 R（条码后缀 ZZ）", "珍尚银 Y（条码前缀 SY）", "IU B（条码前缀 IU）",
              "TY T（条码后缀 TY）", "萃华 H（条码前缀 CH）", "周大金 Z（条码前缀 DJ）"]:
        w6.append([x])
    w6.append([])
    w6.append(["附：属性码 — 分类"])
    w6.append(["XL 项链｜SL 手链｜ES 耳饰｜JZ 戒指｜SZ 手镯｜DZ 吊坠｜BC 包材｜YL 银礼｜金礼（条码 GC / 编号 GG）"])
    w6.append([])
    w6.append(["附：商品性质后缀（条码）"])
    w6.append(["WC 外采商品｜CK 工厂提供测款商品"])
    w6.append([])
    w6.append(["附：天然石（需配证书；合成立方氧化锆不算）"])
    w6.append(["和田玉、玛瑙、玉髓、青金石、芙蓉石、蜜蜡、木变石、翡翠、水晶、发晶、欧泊、海蓝宝、石榴石、长石、葡萄石、尖晶石、碧玺、孔雀石"])
    w6.column_dimensions["A"].width = 10
    w6.column_dimensions["B"].width = 22
    w6.column_dimensions["C"].width = 95
    w6.column_dimensions["D"].width = 10
    header(w6, 4)
    w6.append([])
    w6.append(["人工口径修订记录"])
    for t, d in [
        ("口径1", "规格里已写『配证书』时，货品名称体现『配证书』即为通过"),
        ("口径2", "三列均未写银含量 = 本就不含银，规则4 免检"),
        ("口径3", "非金属材质（无银含量）无法打字印，字印为空不算问题"),
        ("口径4", "条码带 WC / CK（外采、工厂测款）的商品，不检查字印要求"),
        ("口径5", "包材产品（属性码 BC）整行跳过核对"),
        ("口径6", "萃华归入品牌组，须含『配国检证书』"),
        ("口径7", "周大金品牌字印以『周大金』开头为正确"),
        ("口径8", "抽检比例列不得为空（规则7）"),
        ("口径9", "手镯银重须为规格重量区间的中间值，或规格唯一值（规则8）"),
        ("口径10", "规格材质完整性 / 规格工艺完整性 / 禁用商贸俗称材质名 → 需看图识别，不纳入自动核查"),
        ("口径11", "TY 与 唯一 为相等信息：编号末位代号 T / A 互通，条码标记亦互通"),
        ("口径12", "周大金不需要配国检证书 → 归入非品牌组口径（按材质判断是否写『配证书』）"),
    ]:
        w6.append([t, d])

    out.save(OUT)
    print(f"\n[OK] 疑点清单 → {OUT}")
    print(f"     疑点 {len(issues)} 条 / 涉及 {len(set(x[0] for x in issues))} 行 / 总 {len(data)} 行")
    print(f"     条码形态待识别 {len(unparsed)} ｜ 包材免检 {len(skipped)} ｜ 手镯无重量 {len(noweight)}")
    print("     分规则：" + "、".join(f"{k.split()[0]}={v}" for k, v in sorted(stats.items()) if k.startswith("规则"))
          if stats else "     分规则：无")
    return OUT


def main():
    ap = argparse.ArgumentParser(description="吉客云货品档案 · 字段信息抽查引擎 v1.0")
    ap.add_argument("--src", required=True, help="被核查的货品导出表 (.xlsx)")
    ap.add_argument("--out", default=None, help="输出的疑点清单 (.xlsx)，默认同目录自动命名")
    ap.add_argument("--sheet", default=None, help="工作表名，默认取第一个 sheet")
    a = ap.parse_args()
    if not os.path.isfile(a.src):
        sys.exit(f"[ERROR] 文件不存在: {a.src}")
    run(a.src, a.out, a.sheet)


if __name__ == "__main__":
    main()

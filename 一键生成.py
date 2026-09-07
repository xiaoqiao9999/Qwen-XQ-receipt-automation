"""
单据类一键自动化生成脚本 v4
步骤1: OCR识别入库单扫描件 → 生成采购台账(宋体)
步骤2: 解包验收单模板 → 生成验收单docx → 转PDF
步骤7.5: 自动从发票PDF提取信息 → 生成发票记录.json (pdfplumber)
步骤8~14: 读取发票记录 → 匹配数据库 → 生成付款单xlsx → 转PDF(自动验证)
v4: 每个步骤前检查对应文件/数据，不存在则跳过
    OCR数据涉及数量、单价、金额时必须交叉验证(数量×单价=金额，各项之和=合计)
    发票信息从PDF文本/表格提取，严禁编造或自动补全
"""
import os, shutil, zipfile, subprocess, re, json, glob, sys, time
from lxml import etree
import xlrd, xlutils.copy as xcopy
from xlwt import XFStyle, Font
import paths, naming, db
sys.path.insert(0, os.path.join(paths.BASE, "80_工具"))
import payment_ledger

# ============================================================
# 路径配置（统一走 paths.py，不散写绝对路径）
# ============================================================
BASE = paths.BASE
TEMPLATE_XLS = paths.TEMPLATE_XLS
OUTPUT_XLS = paths.ledger_path(2026)
TEMPLATE_DOCX = paths.TEMPLATE_DOCX
TEMPLATE_PAY_XLSX = paths.TEMPLATE_PAY_XLSX
DATABASE_XLSX = paths.DATABASE_XLSX
INVOICE_DIR = paths.INVOICE_IN_DIR                       # 发票输入 10_输入\发票
INVOICE_DONE_DIR = paths.INVOICE_OUT_DIR                 # 发票归档 20_输出\发票
PURCHASE_DIR = paths.PURCHASE_IN_DIR                     # 申购单输入 10_输入\申购单
APPROVAL_DIR = paths.APPROVAL_IN_DIR                     # 事项审批输入 10_输入\事项审批
SCAN_DIR = paths.SCAN_DIR                                # 入库单扫描件 10_输入
SCAN_DONE_DIR = os.path.join(paths.ARCHIVE_DIR, "入库单扫描件")  # 扫描件归档 40_归档
OUTPUT_DOCX_DIR = paths.ACCEPT_DIR                       # 验收单输出(Word) 20_输出\验收单
OUTPUT_PDF_ACCEPT_DIR = paths.PRINT_DIR                  # 验收单PDF直接放 30_待打印
OUTPUT_PAY_XLSX_DIR = paths.PAY_DIR                      # 付款单源文件(xlsx) 20_输出\付款单
OUTPUT_PDF_PAY_DIR = paths.PRINT_DIR                     # 付款单PDF直接放 30_待打印
WORK_DIR = paths.WORK_DIR
UNPACKED_DOCX = os.path.join(WORK_DIR, "unpacked_tpl")
UNPACKED_PAY = os.path.join(WORK_DIR, "unpacked_payment_tpl")
OCR_SCRIPT = paths.OCR_SCRIPT

for d in [OUTPUT_DOCX_DIR, OUTPUT_PAY_XLSX_DIR, OUTPUT_PDF_ACCEPT_DIR, OUTPUT_PDF_PAY_DIR]:
    os.makedirs(d, exist_ok=True)

# ============================================================
# 初始化全局变量
# ============================================================
scan_results = []
all_records = []
all_receipts = []
rename_map = {}
docx_files = []
pay_count = 0
usage_conflicts = []   # 申购单与事项审批用途不一致项，流程末尾汇总提示

# ============================================================
# 步骤0：自动分拣输入文件
# 规则：丢进 10_输入 根目录（非子目录）的文件按以下顺序判定归位：
#   1) 文件名关键词（入库单/发票/申购单/审批）
#   2) 图片扩展名 → 入库单扫描件；pdf 读首页文本关键词
#   3) 判定不了 → 10_输入\待确认（WARNING 提示人工处理，绝不丢弃）
# 同名冲突：内容相同视为重复副本删除源文件；同名不同内容自动加序号。
# 直接放在各子目录里的文件保持原样，不受本步骤影响。
# ============================================================
print("=" * 60)
print("步骤0：自动分拣输入文件（10_输入 根目录散放文件归位）")
print("=" * 60)

INPUT_PENDING_DIR = os.path.join(paths.INPUT_DIR, "待确认")
_IMG_EXTS = {".jpg", ".jpeg", ".png"}

def _name_category(fname):
    if "入库单" in fname or "扫描件" in fname:
        return paths.SCAN_DIR
    if "发票" in fname:
        return INVOICE_DIR
    if "申购单" in fname or "申购" in fname:
        return PURCHASE_DIR
    if "事项审批" in fname or "审批" in fname:
        return APPROVAL_DIR
    return None

def _pdf_category(path):
    """读 PDF 首页文本判定类型；失败/无特征返回 None。"""
    try:
        import pdfplumber
        with pdfplumber.open(path) as pdf:
            text = (pdf.pages[0].extract_text() or "")[:1500]
    except Exception as e:
        print("  WARNING: %s 内容识别失败(%s)，转入待确认" % (os.path.basename(path), e))
        return None
    if not text.strip():
        return None
    if ("价税合计" in text or "（小写）" in text or "(小写)" in text
            or "发票" in text or "销售方" in text):
        return INVOICE_DIR
    if "申购" in text:
        return PURCHASE_DIR
    if "审批" in text:
        return APPROVAL_DIR
    if "入库单" in text:
        print("  提示: %s 是 PDF 版入库单，主流程仅支持图片扫描件，转入待确认" % os.path.basename(path))
        return INPUT_PENDING_DIR
    return None

def _same_content(a, b):
    try:
        import hashlib
        ha = hashlib.md5(open(a, "rb").read()).hexdigest()
        hb = hashlib.md5(open(b, "rb").read()).hexdigest()
        return ha == hb
    except Exception:
        return False

def _unique_dest(dst):
    base, ext = os.path.splitext(dst)
    n = 2
    while os.path.exists(dst):
        dst = "%s(%d)%s" % (base, n, ext)
        n += 1
    return dst

loose_files = []
if os.path.isdir(paths.INPUT_DIR):
    for fn in sorted(os.listdir(paths.INPUT_DIR)):
        fp = os.path.join(paths.INPUT_DIR, fn)
        if os.path.isfile(fp) and not fn.startswith("~$"):
            loose_files.append((fn, fp))

if not loose_files:
    print("跳过: 10_输入 根目录没有待分拣文件")
else:
    moved = pending = 0
    for fn, fp in loose_files:
        ext = os.path.splitext(fn)[1].lower()
        cat = _name_category(fn)
        if cat == paths.SCAN_DIR and ext == ".pdf":
            cat = None  # 图片扫描件流程不处理 PDF，按内容重新判定
        if cat is None and ext in _IMG_EXTS:
            cat = paths.SCAN_DIR
        elif cat is None and ext == ".pdf":
            cat = _pdf_category(fp)
        if cat is None:
            cat = INPUT_PENDING_DIR
            pending += 1
            print("  [待确认] %s —— 无法判定类型，请人工归类" % fn)
        os.makedirs(cat, exist_ok=True)
        dst = os.path.join(cat, fn)
        if os.path.exists(dst):
            if _same_content(fp, dst):
                os.remove(fp)
                print("  去重: %s 与目标目录已有文件内容相同，删除散放副本" % fn)
                moved += 1
                continue
            dst = _unique_dest(dst)
            print("  重名避让: %s → %s" % (fn, os.path.basename(dst)))
        shutil.move(fp, dst)
        rel = os.path.relpath(cat, paths.INPUT_DIR)
        print("  归位: %s → %s\\%s" % (fn, rel, os.path.basename(dst)))
        moved += 1
    print("分拣完成: 归位 %d 个（其中待确认 %d 个）" % (moved, pending))

# ============================================================
# 步骤1：OCR识别入库单扫描件
# ============================================================
print("=" * 60)
print("步骤1：OCR识别入库单扫描件")
print("=" * 60)

def ocr_single(img_path, index=0):
    """使用NPU OCR识别单张图片（CPU模式），失败返回空。
    注意：ppocr.exe不支持中文路径和中文文件名，需要先复制到纯英文临时目录并重命名。"""
    img_path = os.path.abspath(img_path)
    tmp_dir = os.path.join(WORK_DIR, "ocr_tmp")
    os.makedirs(tmp_dir, exist_ok=True)
    tmp_img = os.path.join(tmp_dir, "img_%04d.jpg" % index)
    shutil.copy2(img_path, tmp_img)
    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-File", OCR_SCRIPT, tmp_img, "-Device", "cpu"],
            capture_output=True, timeout=300
        )
        try:
            stdout_text = result.stdout.decode('utf-8')
        except UnicodeDecodeError:
            stdout_text = result.stdout.decode('gbk')
        if result.returncode != 0:
            err_lines = [l for l in stdout_text.splitlines() if '[ERROR]' in l]
            print("    OCR错误: %s" % "; ".join(err_lines) if err_lines else "exit code %d" % result.returncode)
            return []
        lines = []
        in_result = False
        for line in stdout_text.splitlines():
            s = line.strip()
            if s.startswith('---') and len(s) >= 40:
                in_result = True
                continue
            if s.startswith('---') and in_result:
                break
            if in_result and s:
                lines.append(s)
        return lines
    except Exception as e:
        print("    OCR异常: %s" % str(e))
        return []
    finally:
        if os.path.exists(tmp_img):
            os.remove(tmp_img)

def ocr_directory(dir_path):
    """OCR识别目录下所有jpg图片"""
    jpg_files = sorted(glob.glob(os.path.join(dir_path, "*.jpg")))
    if not jpg_files:
        jpg_files = sorted(glob.glob(os.path.join(dir_path, "*.jpeg")))
        if not jpg_files:
            jpg_files = sorted(glob.glob(os.path.join(dir_path, "*.png")))
    all_results = []
    for idx, fp in enumerate(jpg_files):
        print("  OCR (%d/%d): %s" % (idx + 1, len(jpg_files), os.path.basename(fp)))
        lines = ocr_single(fp, index=idx)
        all_results.append({"file": os.path.basename(fp), "lines": lines})
    return all_results

def extract_receipt_id_from_ocr(lines):
    """从OCR行中提取入库单号"""
    for line in lines:
        m = re.search(r'(CGSH\d{12,})', line)
        if m:
            return m.group(1)
    return ""

# ============================================================
# OCR数据自动提取（单号匹配失败时用于自动补全JSON）
# ============================================================
UNITS_SET = {"个", "套", "台", "米", "根", "卷", "只", "把", "片", "包", "桶", "盒", "瓶", "件", "条", "支", "捆", "组", "双", "箱", "块", "张", "付"}
NOISE_WORDS = {"限公司仓库", "州川油平刑i小分", "州川型", "用刑", "荆州牛", "消小务有", "荆州酒中刑小分", "伞病", "合", "计：", "采购员：", "仓库保管员："}

def _is_code(s):
    return bool(re.match(r'^[A-Za-z0-9]{8,}$', s))

def _is_float(s):
    try:
        float(s)
        return True
    except:
        return False

def _plausible_unit(s):
    # 数量前字段作为单位的宽容判定：未在 UNITS_SET 时，短汉字也视为单位
    s = (s or "").strip()
    if not s or len(s) > 4:
        return False
    return bool(re.fullmatch(r'[\u4e00-\u9fa5]{1,4}', s))

def _spec_garbled(s):
    # 规格型号识别不清（乱码）判定：含字母数字或常见规格符号则可信；
    # 否则为全汉字且偏长 → 视为识别不清，置空跳过
    s = (s or "").strip()
    if not s:
        return False
    if re.search(r'[A-Za-z0-9\-+×/.·*#%()（）★Ω]', s):
        return False
    return len(s) > 4

def extract_receipt_date(lines):
    for line in lines:
        m = re.search(r'(\d{4})年(\d{1,2})月(\d{1,2})', line)
        if m:
            return "%s-%02d-%02d" % (m.group(1), int(m.group(2)), int(m.group(3)))
        m = re.search(r'(\d{4})-(\d{2})-(\d{2})', line)
        if m:
            return m.group(0)
    return ""

def extract_receipt_supplier(lines):
    for i, line in enumerate(lines):
        if "供应商名称" in line:
            for j in range(i + 1, min(i + 6, len(lines))):
                candidate = lines[j].strip()
                if candidate and candidate not in ("制单人", "备注", "物料编号", "物料名称") and len(candidate) > 4:
                    if "经营部" in candidate or "公司" in candidate or "厂" in candidate or "店" in candidate or "个体" in candidate:
                        return candidate
    return ""

def extract_receipt_total(lines):
    for i, line in enumerate(lines):
        if "合计" in line or "计：" in line:
            for j in range(i, min(i + 6, len(lines))):
                m = re.search(r'(\d+\.\d{2})', lines[j])
                if m:
                    return float(m.group(1))
    return 0

def parse_items_from_ocr(lines):
    items = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        if not _is_code(line):
            i += 1
            continue
        code = line
        collected = []
        j = i + 1
        while j < len(lines):
            nxt = lines[j].strip()
            if nxt in NOISE_WORDS or nxt.startswith("限公司") or nxt.startswith("州川") or nxt.startswith("荆州"):
                j += 1
                continue
            if _is_code(nxt) or "合计" in nxt:
                break
            collected.append(nxt)
            j += 1

        qty = None
        qty_idx = -1
        for idx, v in enumerate(collected):
            if re.match(r'^\d+\.\d{2}$', v):
                qty = float(v)
                qty_idx = idx
                break
        if qty is None:
            for idx, v in enumerate(collected):
                if _is_float(v):
                    rest_nums = [x for x in collected[idx + 1:] if _is_float(x)]
                    if len(rest_nums) >= 2:
                        qty = float(v)
                        qty_idx = idx
                        break
        if qty is None or qty_idx < 0:
            i += 1
            continue

        unit = ""
        if qty_idx > 0:
            if collected[qty_idx - 1] in UNITS_SET:
                unit = collected[qty_idx - 1]
            else:
                for k in range(qty_idx - 1, -1, -1):
                    if collected[k] in UNITS_SET:
                        unit = collected[k]
                        break
        if not unit and qty_idx > 0 and _plausible_unit(collected[qty_idx - 1]):
            unit = collected[qty_idx - 1]

        spec_end = qty_idx - 1 if qty_idx > 0 and collected[qty_idx - 1] == unit else qty_idx
        name = collected[0] if collected else ""
        spec = collected[1] if spec_end > 1 else ""
        if _spec_garbled(spec):
            spec = ""

        rem = collected[qty_idx + 1:]
        prices = []
        amounts = []
        for v in rem:
            if re.match(r'^\d+\.\d{4}$', v):
                prices.append(float(v))
            elif re.match(r'^\d+(\.\d{2})?$', v) and not re.match(r'^\d+\.\d{4}$', v):
                amounts.append(float(v))

        price_no_tax = prices[0] if len(prices) >= 1 else None
        price_tax = prices[1] if len(prices) >= 2 else None
        amount_no_tax = amounts[0] if len(amounts) >= 1 else None
        amount_tax = amounts[1] if len(amounts) >= 2 else None

        if amount_tax is None and price_tax is not None:
            amount_tax = round(price_tax * qty, 2)
        if amount_no_tax is None and price_no_tax is not None:
            amount_no_tax = round(price_no_tax * qty, 2)

        if name and unit and qty is not None:
            items.append({
                "code": code, "name": name, "spec": spec, "unit": unit,
                "qty": qty,
                "price": price_tax if price_tax is not None else 0,
                "amount": amount_tax if amount_tax is not None else 0
            })
            i = j
            continue
        i += 1
    return items

def cross_validate_items(items, total_ocr=0, context=""):
    """交叉验证：数量×单价=金额，各项金额之和=OCR合计。
    不匹配时打印WARNING提示人工确认，严禁自动补全或编造数据。"""
    issues = []
    for i, it in enumerate(items):
        qty = it.get("qty", 0) or 0
        price = it.get("price", 0) or 0
        amount = it.get("amount", 0) or 0
        if qty > 0 and price > 0:
            expected = round(qty * price, 2)
            if abs(expected - amount) > 0.5:
                issues.append("    第%d项 %s: 数量%s × 单价%.2f = %.2f ≠ 金额%.2f" % (
                    i + 1, it.get("name", "?"), qty, price, expected, amount))
    if total_ocr > 0:
        items_total = round(sum(it.get("amount", 0) or 0 for it in items), 2)
        if abs(items_total - total_ocr) > 1:
            issues.append("    物料合计 %.2f ≠ 入库单合计 %.2f（可能存在遗漏或多识别）" % (items_total, total_ocr))
    if issues:
        prefix = "  " + context + " " if context else "  "
        print(prefix + "交叉验证 WARNING:")
        for iss in issues:
            print(iss)
    return len(issues) == 0

def _load_receipt_db():
    """从DB加载全部入库单为 {receipt_id: {...含items}} 映射。"""
    return {rid: db.get_receipt(rid) for rid in db.receipt_ids()}


def _usage_from_name(filename):
    """从“（用途）金额”形文件名提取用途，无则空。"""
    m = re.search(r"[（(](.+?)\s*金额", filename or "")
    return m.group(1).strip().rstrip("）)") if m else ""


def _usage_by_amount(folder, amount):
    """在 folder 目录内按“（用途）金额”+金额匹配，返回用途或空。"""
    for pf in glob.glob(os.path.join(folder, "*.pdf")):
        pm = re.search(r"[（(](.+?)\s*金额(\d+)", os.path.basename(pf))
        if pm and float(pm.group(2)) == int(round(amount)):
            return pm.group(1).strip().rstrip("）)")
    return ""


def resolve_invoice_usage(filename, amount, existing="", context=""):
    """统一解析发票用途：已有值 → 发票文件名 → 申购单/事项审批按金额匹配（二者须一致）。
    返回归一化用途；无法确定或申购单与事项审批不一致时返回空串（不编造）。"""
    usage = (existing or "").strip()
    if usage:
        return usage
    if filename:
        usage = _usage_from_name(filename)
        if usage:
            if context:
                print("  %s: 用途从文件名提取: %s" % (context, usage))
            return usage
    usage_p = _usage_by_amount(PURCHASE_DIR, amount)
    usage_a = _usage_by_amount(APPROVAL_DIR, amount)
    if usage_p and usage_a:
        if usage_p == usage_a:
            if context:
                print("  %s: 申购单与事项审批用途一致: %s" % (context, usage_p))
            return usage_p
        else:
            print("  ##### 用途不一致 #######程序提示##########################")
            print("  涉及发票: %s" % (filename or "（未知）"))
            print("  申购单用途: %s" % usage_p)
            print("  事项审批用途: %s" % usage_a)
            print("  两者不一致，已放弃自动采用，请人工核对后补填用途！")
            print("  ############################################################")
            usage_conflicts.append({"filename": filename, "amount": amount,
                                    "purchase": usage_p, "approval": usage_a})
            return ""
    if usage_p:
        if context:
            print("  %s: 用途从申购单提取: %s" % (context, usage_p))
        return usage_p
    if usage_a:
        if context:
            print("  %s: 用途从事项审批提取: %s" % (context, usage_a))
        return usage_a
    return ""


def audit_and_fix_ledger(output_xls, month_label, receipts_info):
    """台账自检：OCR提取的单据数据总金额要与台账内金额保持一致。
    逐单号检查行数是否匹配DB receipts项数，缺失则自动补入。
    无法修复时返回False并打印WARNING告知用户。"""
    if not receipts_info:
        return True
    if not os.path.exists(output_xls):
        return True

    # 从DB读取
    receipt_db = _load_receipt_db()

    # 读取台账
    rb = xlrd.open_workbook(output_xls, formatting_info=True)
    sheet_names = rb.sheet_names()
    if month_label not in sheet_names:
        print("  台账自检: 未找到工作表 '%s'，跳过自检" % month_label)
        return True
    sheet_idx = sheet_names.index(month_label)
    sheet = rb.sheet_by_index(sheet_idx)

    # 统计台账中每个入库单号的行数与金额
    ledger_counts = {}
    ledger_amounts = {}
    for row_idx in range(1, sheet.nrows):
        rid = str(sheet.cell_value(row_idx, 8)).strip() if sheet.cell_value(row_idx, 8) else ""
        if not rid or "合计" in rid:
            continue
        ledger_counts[rid] = ledger_counts.get(rid, 0) + 1
        try:
            amt = float(sheet.cell_value(row_idx, 6))
        except:
            amt = 0
        ledger_amounts[rid] = ledger_amounts.get(rid, 0) + amt

    issues = []
    missing_records = []

    for rid_info in receipts_info:
        rid = rid_info["id"]
        if rid not in receipt_db:
            continue
        json_items = receipt_db[rid]["items"]
        json_count = len(json_items)
        ledger_count = ledger_counts.get(rid, 0)
        json_total = sum(it.get("amount", 0) for it in json_items)
        ledger_total = ledger_amounts.get(rid, 0)

        if json_count != ledger_count or abs(json_total - ledger_total) > 1:
            issues.append("  %s: JSON=%d项/%.0f元, 台账=%d行/%.0f元" % (
                rid, json_count, json_total, ledger_count, ledger_total))
            # 准备补入缺失记录
            if ledger_count == 0:
                # 整单缺失
                for it in json_items:
                    missing_records.append({
                        "code": it.get("code", ""), "name": it["name"], "spec": it.get("spec", ""),
                        "unit": it["unit"], "qty": it["qty"], "price": it.get("price", 0),
                        "amount": it.get("amount", 0), "date": receipt_db[rid]["date"],
                        "receipt": rid, "supplier": receipt_db[rid]["supplier"], "reason": ""
                    })
            elif ledger_count < json_count:
                # 部分缺失：对比名称找出缺失项
                ledger_names = []
                for row_idx in range(1, sheet.nrows):
                    cell_rid = str(sheet.cell_value(row_idx, 8)).strip() if sheet.cell_value(row_idx, 8) else ""
                    if cell_rid == rid:
                        ledger_names.append(str(sheet.cell_value(row_idx, 1)).strip())
                for it in json_items:
                    if it["name"] not in ledger_names:
                        missing_records.append({
                            "code": it.get("code", ""), "name": it["name"], "spec": it.get("spec", ""),
                            "unit": it["unit"], "qty": it["qty"], "price": it.get("price", 0),
                            "amount": it.get("amount", 0), "date": receipt_db[rid]["date"],
                            "receipt": rid, "supplier": receipt_db[rid]["supplier"], "reason": ""
                        })

    if not issues:
        print("  台账自检: 全部%d个单号一致" % len(receipts_info))
        return True

    print("  台账自检 WARNING: 发现%d个单号不一致" % len(issues))
    for iss in issues:
        print(iss)

    if not missing_records:
        print("  台账自检: 无法自动修复（台账行数≥JSON项数但金额不匹配），请人工检查")
        return False

    # 自动修复：补入缺失记录
    print("  台账自检: 尝试自动修复，补入%d条缺失记录" % len(missing_records))
    wb = xcopy.copy(rb)
    ws = wb.get_sheet(sheet_idx)

    style_song = XFStyle()
    style_song.font = Font()
    style_song.font.name = 'SimSun'
    style_song.font.height = 200
    style_song_bold = XFStyle()
    style_song_bold.font = Font()
    style_song_bold.font.name = 'SimSun'
    style_song_bold.font.bold = True
    style_song_bold.font.height = 200

    # 找到数据末尾行和合计行
    last_data_row = 0
    for row_idx in range(1, sheet.nrows):
        rid = str(sheet.cell_value(row_idx, 8)).strip() if sheet.cell_value(row_idx, 8) else ""
        if rid and "合计" not in rid:
            last_data_row = row_idx

    total_row_idx = last_data_row + 1
    old_total = 0
    if total_row_idx < sheet.nrows:
        try:
            old_total = float(sheet.cell_value(total_row_idx, 6))
        except:
            old_total = 0

    # 清除旧合计行
    for col_idx in range(11):
        ws.write(total_row_idx, col_idx, "")

    # 在旧合计行位置写入缺失记录
    start_row = total_row_idx
    for i, rec in enumerate(missing_records):
        vals = [rec["code"], rec["name"], rec["spec"], rec["unit"],
                rec["qty"], rec["price"], rec["amount"], rec["date"],
                rec["receipt"], rec["supplier"], rec["reason"]]
        for col, val in enumerate(vals):
            ws.write(start_row + i, col, val, style_song)

    # 写入新合计行
    new_total = old_total + sum(r["amount"] for r in missing_records)
    ws.write(start_row + len(missing_records), 6, new_total, style_song_bold)
    wb.save(output_xls)
    print("  台账已修复: 补入%d条, 合计 %.2f → %.2f" % (len(missing_records), old_total, new_total))

    # 二次验证
    rb2 = xlrd.open_workbook(output_xls, formatting_info=True)
    sheet2 = rb2.sheet_by_index(sheet_idx)
    ledger_counts2 = {}
    ledger_amounts2 = {}
    for row_idx in range(1, sheet2.nrows):
        rid = str(sheet2.cell_value(row_idx, 8)).strip() if sheet2.cell_value(row_idx, 8) else ""
        if rid and "合计" not in rid:
            ledger_counts2[rid] = ledger_counts2.get(rid, 0) + 1
            try:
                amt = float(sheet2.cell_value(row_idx, 6))
            except:
                amt = 0
            ledger_amounts2[rid] = ledger_amounts2.get(rid, 0) + amt

    still_issues = []
    for rid_info in receipts_info:
        rid = rid_info["id"]
        if rid not in receipt_db:
            continue
        json_count = len(receipt_db[rid]["items"])
        ledger_count = ledger_counts2.get(rid, 0)
        json_total = sum(it.get("amount", 0) for it in receipt_db[rid]["items"])
        ledger_total = ledger_amounts2.get(rid, 0)
        if json_count != ledger_count or abs(json_total - ledger_total) > 1:
            still_issues.append("  %s: JSON=%d项/%.0f元, 台账=%d行/%.0f元" % (
                rid, json_count, json_total, ledger_count, ledger_total))

    if still_issues:
        print("  台账自检: 修复后仍有问题，请人工检查:")
        for iss in still_issues:
            print(iss)
        return False
    else:
        print("  台账自检: 修复后全部一致")
        return True

def auto_add_receipt_to_db(receipt_id, lines):
    """从OCR行自动提取数据并写入DB，返回是否成功。"""
    date = extract_receipt_date(lines)
    supplier = extract_receipt_supplier(lines)
    items = parse_items_from_ocr(lines)
    total_ocr = extract_receipt_total(lines)

    if not date:
        print("    自动提取: 未能识别日期，跳过")
        return False
    if not supplier:
        print("    自动提取: 未能识别供应商，跳过")
        return False
    if not items:
        print("    自动提取: 未能识别物料明细，跳过")
        return False

    # supplier_short: 去掉省市区前缀
    supplier_short = supplier
    for prefix in ["荆州市", "荆州区", "沙市区", "荆州经济技术开发区"]:
        if supplier_short.startswith(prefix):
            supplier_short = supplier_short[len(prefix):]

    # 交叉验证：数量×单价=金额，各项金额之和=OCR合计
    cross_validate_items(items, total_ocr, "自动提取")

    items_total = sum(it.get("amount", 0) for it in items)
    total = int(total_ocr) if total_ocr > 0 else int(items_total)

    db.add_receipt(receipt_id, {
        "date": date,
        "supplier": supplier,
        "supplier_short": supplier_short,
        "items": items,
        "total": total,
    })
    print("    自动提取: 已添加 %s → %s (%d项, %d元)" % (receipt_id, supplier_short, len(items), total))
    return True

# 检查入库单扫描件目录是否有图片
jpg_files = sorted(glob.glob(os.path.join(SCAN_DIR, "*.jpg")))
if not jpg_files:
    jpg_files = sorted(glob.glob(os.path.join(SCAN_DIR, "*.jpeg")))
    if not jpg_files:
        jpg_files = sorted(glob.glob(os.path.join(SCAN_DIR, "*.png")))
has_scan_files = len(jpg_files) > 0

if has_scan_files:
    scan_results = ocr_directory(SCAN_DIR)
    print("识别了 %d 张入库单扫描件" % len(scan_results))

    ocr_success = any(len(sr["lines"]) > 0 for sr in scan_results)
    if not ocr_success:
        print("WARNING: OCR未能从任何入库单中提取有效数据，后续步骤将跳过")
        scan_results = []
else:
    print("跳过: 入库单扫描件目录无图片文件(jpg/jpeg/png)")

# ============================================================
# 步骤2：匹配入库单数据（DB）
# ============================================================
print("\n" + "=" * 60)
print("步骤2：匹配入库单数据（DB）")
print("=" * 60)

has_json_db = len(db.receipt_ids()) > 0

if has_scan_files and scan_results and has_json_db:
    receipt_db = _load_receipt_db()
    print("数据映射表: %d 条入库单记录" % len(receipt_db))

    unmatched_count = 0

    def _do_match(sr, rid, data):
        supplier_short = data["supplier_short"]
        total = data["total"]
        rename_map[sr["file"]] = naming.scan_name(rid, data["date"], supplier_short)
        for it in data["items"]:
            all_records.append({
                "code": it.get("code", ""), "name": it["name"], "spec": it.get("spec", ""),
                "unit": it["unit"], "qty": it["qty"], "price": it.get("price", 0),
                "amount": it.get("amount", 0), "date": data["date"],
                "receipt": rid, "supplier": data["supplier"], "reason": ""
            })
        items = [{"name": it["name"], "spec": it.get("spec", ""), "unit": it["unit"], "qty": it["qty"]} for it in data["items"]]
        all_receipts.append({
            "id": rid, "date": data["date"], "supplier_short": supplier_short,
            "items": items, "total": total
        })
        print("  匹配: %s → %s (%d项, %d元)" % (rid, supplier_short, len(items), total))
        # 交叉验证：数量×单价=金额，各项金额之和=合计
        cross_validate_items(data.get("items", []), data.get("total", 0), rid)

    last_rid = None          # 上一张已处理入库单号（续页沿用）
    continuations = {}       # rid -> {"items":[...], "total":float}  分页续页待合并

    for sr in scan_results:
        rid = extract_receipt_id_from_ocr(sr["lines"])
        if rid and rid in receipt_db:
            last_rid = rid
            _do_match(sr, rid, receipt_db[rid])
        elif rid:
            print("  WARNING: 单号 %s 未在数据映射表中找到，尝试自动提取..." % rid)
            if auto_add_receipt_to_db(rid, sr["lines"]):
                receipt_db = _load_receipt_db()
                if rid in receipt_db:
                    last_rid = rid
                    _do_match(sr, rid, receipt_db[rid])
                    continue
            unmatched_count += 1
        else:
            # 无单号 → 续页（清单超一页的第二页），沿用上一张单号并合并物料
            if last_rid and last_rid in receipt_db:
                cont_items = parse_items_from_ocr(sr["lines"])
                if cont_items:
                    if last_rid not in continuations:
                        continuations[last_rid] = {"items": [], "total": 0}
                    continuations[last_rid]["items"].extend(cont_items)
                    cont_total = extract_receipt_total(sr["lines"])
                    if cont_total > 0 and continuations[last_rid]["total"] == 0:
                        continuations[last_rid]["total"] = cont_total
                    print("  续页: %s 无单号，沿用上一张 %s，识别 %d 项物料" % (sr["file"], last_rid, len(cont_items)))
                else:
                    print("  WARNING: 续页 %s 未能解析出物料，未合并" % sr["file"])
                    unmatched_count += 1
            else:
                print("  WARNING: 未能从 %s 中识别入库单号" % sr["file"])
                unmatched_count += 1

    # 合并续页物料到对应入库单（DB + 台账 + 验收单），并按最后一页合计校验
    for rid, cont in continuations.items():
        data = receipt_db[rid]
        merged_items = data["items"] + cont["items"]
        merged_total = cont["total"] if cont["total"] > 0 else data["total"]
        db.add_receipt(rid, {
            "date": data["date"], "supplier": data["supplier"],
            "supplier_short": data["supplier_short"],
            "items": merged_items, "total": merged_total,
        })
        receipt_db[rid] = db.get_receipt(rid)
        for a in all_receipts:
            if a["id"] == rid:
                a["items"] = [{"name": it["name"], "spec": it.get("spec", ""), "unit": it["unit"], "qty": it["qty"]} for it in merged_items]
                a["total"] = merged_total
                break
        for it in cont["items"]:
            all_records.append({
                "code": it.get("code", ""), "name": it["name"], "spec": it.get("spec", ""),
                "unit": it["unit"], "qty": it["qty"], "price": it.get("price", 0),
                "amount": it.get("amount", 0), "date": data["date"],
                "receipt": rid, "supplier": data["supplier"], "reason": ""
            })
        sum_merged = round(sum(it.get("amount", 0) or 0 for it in merged_items), 2)
        if merged_total > 0 and abs(sum_merged - merged_total) > 1:
            print("  合并金额校验 WARNING: %s 合并物料合计 %.2f ≠ 单据总金额 %.2f（%d项）" % (rid, sum_merged, merged_total, len(merged_items)))
        else:
            print("  合并金额校验 OK: %s 物料合计 %.2f = 总金额 %.2f（%d项）" % (rid, sum_merged, merged_total, len(merged_items)))
        cross_validate_items(merged_items, merged_total, rid)

    if unmatched_count > 0:
        print("WARNING: %d 张扫描件未匹配且自动提取失败，请手动补充到: %s (用 manage_db.py add)" % (unmatched_count, db.DB_PATH))

    print("匹配完成: %d条记录, %d份验收单" % (len(all_records), len(all_receipts)))
elif not has_scan_files or not scan_results:
    print("跳过: 无入库单扫描件或OCR未识别到数据")
elif not has_json_db:
    print("跳过: 数据库中无入库单记录 (用 manage_db.py add 补录)")

# ============================================================
# 步骤3：扫描件归档
# ============================================================
print("\n" + "=" * 60)
print("步骤3：扫描件归档")
print("=" * 60)

has_rename = len(rename_map) > 0
if has_rename:
    os.makedirs(SCAN_DONE_DIR, exist_ok=True)
    for orig_name, new_name in rename_map.items():
        src = os.path.join(SCAN_DIR, orig_name)
        dst = os.path.join(SCAN_DONE_DIR, new_name)
        if os.path.exists(src):
            shutil.move(src, dst)
            print("  移动: %s → %s" % (orig_name, new_name))
    print("扫描件已移动到 40_归档\\入库单扫描件\\")
else:
    print("跳过: 无匹配成功的扫描件需要归档")

# ============================================================
# 步骤4：生成采购物资台帐
# ============================================================
print("\n" + "=" * 60)
print("步骤4：生成采购物资台帐")
print("=" * 60)

has_records = len(all_records) > 0
has_tpl_xls = os.path.exists(TEMPLATE_XLS)

if has_records and has_tpl_xls:
    # 优先打开已有台账（保留之前数据），不存在则用模板
    source_xls = OUTPUT_XLS if os.path.exists(OUTPUT_XLS) else TEMPLATE_XLS
    if source_xls == OUTPUT_XLS:
        print("打开已有台账: %s" % OUTPUT_XLS)
    else:
        print("从模板创建新台账: %s" % OUTPUT_XLS)
    rb = xlrd.open_workbook(source_xls, formatting_info=True)
    wb = xcopy.copy(rb)
    sheet_names = rb.sheet_names()

    def _month_label_of(date_str):
        dm = re.search(r'(\d{4})-(\d{2})', date_str or "")
        return ("%d月" % int(dm.group(2))) if dm else None

    # 修复：原先按第一条记录的月份把整批记录写进同一页，导致跨月单据错位。
    # 现按每条记录自身日期分组，逐月写入对应工作表。
    month_groups = {}
    for rec in all_records:
        ml = _month_label_of(rec.get("date", ""))
        if ml is None:
            ml = "6月"
            print("WARNING: 无法从 %s 的日期提取月份，默认使用6月" % rec.get("receipt", "?"))
        month_groups.setdefault(ml, []).append(rec)

    style_song = XFStyle()
    style_song.font = Font()
    style_song.font.name = 'SimSun'
    style_song.font.height = 200
    style_song_bold = XFStyle()
    style_song_bold.font = Font()
    style_song_bold.font.name = 'SimSun'
    style_song_bold.font.bold = True
    style_song_bold.font.height = 200

    written_months = []
    for month_label in sorted(month_groups, key=lambda m: int(m.rstrip("月"))):
        if month_label not in sheet_names:
            print("ERROR: 模板中不存在工作表 '%s'，跳过该月 %d 条" % (
                month_label, len(month_groups[month_label])))
            continue
        sheet_idx = sheet_names.index(month_label)
        ws = wb.get_sheet(sheet_idx)
        existing_sheet = rb.sheet_by_index(sheet_idx)
        print("采购台账写入工作表: %s" % month_label)

        # 收集该月已有入库单号，找到数据末尾行（I列=索引8有值的最后一行）
        existing_receipts = set()
        last_data_row = 0
        for row_idx in range(1, existing_sheet.nrows):
            receipt_val = existing_sheet.cell_value(row_idx, 8)
            if receipt_val:
                receipt_str = str(receipt_val).strip()
                if receipt_str and "合计" not in receipt_str:
                    existing_receipts.add(receipt_str)
                    last_data_row = row_idx

        # 过滤掉该月已存在的入库单号（避免重复追加）
        group = month_groups[month_label]
        new_records = [r for r in group if str(r["receipt"]).strip() not in existing_receipts]
        skipped_count = len(group) - len(new_records)

        if skipped_count > 0:
            print("  %s: 跳过%d条已存在的记录" % (month_label, skipped_count))

        if not new_records:
            print("  %s: 所有入库单记录已存在，无需追加" % month_label)
            continue

        # 读取旧合计金额（旧合计行在 last_data_row + 1 的 G列=索引6）
        existing_total = 0
        total_row_idx = last_data_row + 1
        if total_row_idx < existing_sheet.nrows:
            old_total_val = existing_sheet.cell_value(total_row_idx, 6)
            if isinstance(old_total_val, (int, float)):
                existing_total = old_total_val
            else:
                try:
                    existing_total = float(old_total_val)
                except (ValueError, TypeError):
                    existing_total = 0

        # 从 last_data_row + 1 开始写入新记录（自然覆盖旧合计行）
        start_row = last_data_row + 1
        for i, rec in enumerate(new_records):
            vals = [rec["code"], rec["name"], rec["spec"], rec["unit"],
                    rec["qty"], rec["price"], rec["amount"], rec["date"],
                    rec["receipt"], rec["supplier"], rec["reason"]]
            for col, val in enumerate(vals):
                ws.write(start_row + i, col, val, style_song)

        # 写入新合计行（旧合计 + 新记录合计）
        new_total = existing_total + sum(r["amount"] for r in new_records)
        ws.write(start_row + len(new_records), 6, new_total, style_song_bold)
        written_months.append(month_label)
        print("采购台账 %s: 追加%d条 (已有%d条, 合计%.2f元)" % (
            month_label, len(new_records), len(existing_receipts), new_total))

    if written_months:
        wb.save(OUTPUT_XLS)
        print("台账已保存 → %s" % OUTPUT_XLS)

        # 步骤4.5：台账自检（逐月核对DB金额与台账一致，不一致尝试修复）
        print("\n  步骤4.5：台账自检")
        for month_label in written_months:
            month_receipts = [a for a in all_receipts
                              if (_month_label_of(a.get("date", "")) or "6月") == month_label]
            audit_and_fix_ledger(OUTPUT_XLS, month_label, month_receipts)
else:
    if not has_records:
        print("跳过: 无入库单记录")
    if not has_tpl_xls:
        print("跳过: 采购台账模板不存在: %s" % TEMPLATE_XLS)

# ============================================================
# 步骤5~6：解包验收单模板 + 生成验收单docx
# ============================================================
print("\n" + "=" * 60)
print("步骤5~6：生成验收单docx")
print("=" * 60)

has_tpl_docx = os.path.exists(TEMPLATE_DOCX)
has_receipts = len(all_receipts) > 0

if has_tpl_docx and has_receipts:
    NS_W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    XML_NS = 'http://www.w3.org/XML/1998/namespace'

    def qn_w(tag):
        if ':' in tag:
            prefix, local = tag.split(':', 1)
            nsmap = {'w': NS_W, 'r': NS_R}
            return '{%s}%s' % (nsmap.get(prefix, ''), local)
        return '{%s}%s' % (NS_W, tag)

    def set_cell_text(tc, text):
        p = tc.find(qn_w('w:p'))
        if p is None: return
        runs = p.findall(qn_w('w:r'))
        t_elems = [r.find(qn_w('w:t')) for r in runs]
        t_elems = [t for t in t_elems if t is not None]
        if t_elems:
            t_elems[0].text = text
            t_elems[0].set('{%s}space' % XML_NS, 'preserve')
            r_elem = t_elems[0].getparent()
            rpr = r_elem.find(qn_w('w:rPr'))
            if rpr is None: rpr = etree.SubElement(r_elem, qn_w('w:rPr'))
            rf = rpr.find(qn_w('w:rFonts'))
            if rf is None: rf = etree.SubElement(rpr, qn_w('w:rFonts'))
            rf.set(qn_w('w:ascii'), '宋体')
            rf.set(qn_w('w:hAnsi'), '宋体')
            rf.set(qn_w('w:eastAsia'), '宋体')
            for extra_t in t_elems[1:]:
                extra_t.text = ''
        else:
            r = p.find(qn_w('w:r'))
            if r is None: r = etree.SubElement(p, qn_w('w:r'))
            t = etree.SubElement(r, qn_w('w:t'))
            t.text = text
            t.set('{%s}space' % XML_NS, 'preserve')
            rpr = r.find(qn_w('w:rPr'))
            if rpr is None: rpr = etree.SubElement(r, qn_w('w:rPr'))
            rf = rpr.find(qn_w('w:rFonts'))
            if rf is None: rf = etree.SubElement(rpr, qn_w('w:rFonts'))
            rf.set(qn_w('w:ascii'), '宋体')
            rf.set(qn_w('w:hAnsi'), '宋体')
            rf.set(qn_w('w:eastAsia'), '宋体')

    def fill_acceptance(receipt, output_path):
        work = output_path + '_w'
        if os.path.exists(work): shutil.rmtree(work)
        shutil.copytree(UNPACKED_DOCX, work)
        doc_xml = os.path.join(work, 'word', 'document.xml')
        tree = etree.parse(doc_xml)
        root = tree.getroot()
        tbl = root.find('.//' + qn_w('w:tbl'))
        rows = tbl.findall(qn_w('w:tr'))
        set_cell_text(rows[2].findall(qn_w('w:tc'))[3], receipt['date'])
        items = receipt['items']
        for i in range(17):
            tcs = rows[5 + i].findall(qn_w('w:tc'))
            if i < len(items):
                it = items[i]
                set_cell_text(tcs[0], str(i + 1))
                set_cell_text(tcs[1], it['name'])
                set_cell_text(tcs[2], it.get('spec', '') or '')
                set_cell_text(tcs[3], it['unit'])
                q = it['qty']
                set_cell_text(tcs[4], str(int(q)) if q == int(q) else str(q))
            else:
                for tc in tcs[:5]:
                    set_cell_text(tc, '')
        tree.write(doc_xml, xml_declaration=True, encoding='UTF-8', standalone=True)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for dp, dn, fn in os.walk(work):
                for f in fn:
                    fp = os.path.join(dp, f)
                    zf.write(fp, os.path.relpath(fp, work))
        shutil.rmtree(work)

    if os.path.exists(UNPACKED_DOCX): shutil.rmtree(UNPACKED_DOCX)
    with zipfile.ZipFile(TEMPLATE_DOCX, 'r') as zf:
        zf.extractall(UNPACKED_DOCX)
    print("验收单模板解包完成")

    docx_files = []
    for r in all_receipts:
        fname = naming.acceptance_name(r['id'], r['date'], r['supplier_short'], r['total'])
        out = os.path.join(OUTPUT_DOCX_DIR, fname)
        fill_acceptance(r, out)
        docx_files.append(out)
    print("生成 %d 个验收单docx" % len(docx_files))
else:
    if not has_tpl_docx:
        print("跳过: 验收单模板不存在: %s" % TEMPLATE_DOCX)
    if not has_receipts:
        print("跳过: 无验收单数据")

# ============================================================
# 步骤7：验收单转PDF
# ============================================================
print("\n" + "=" * 60)
print("步骤7：验收单转PDF")
print("=" * 60)

def _strip_revisions(docx_path):
    """清理文档中会让 Word 无头转换挂死的标记。

    验收单模板带 documentProtection（限制编辑保护）：Word 打开受保护文档时会弹出
    密码/确认对话框，在 COM 无头调用下永久挂住。同时移除可能存在的修订记录与批注
    引用，避免 Word 默认不导出被删除行导致续页表格少印内容。"""
    changed = False
    tmp = docx_path + ".rev.tmp"
    with zipfile.ZipFile(docx_path) as zin:
        names = zin.namelist()
        items = {n: zin.read(n) for n in names}

    def clean(xml_bytes):
        text = xml_bytes.decode("utf-8")
        text = re.sub(r'<w:del\b[^>]*>.*?</w:del>', '', text, flags=re.S)
        text = re.sub(r'<w:del\b[^>]*/>', '', text)
        text = re.sub(r'<w:moveFrom\b[^>]*>.*?</w:moveFrom>', '', text, flags=re.S)
        text = re.sub(r'<w:moveTo\b[^>]*>(.*?)</w:moveTo>', r'\1', text, flags=re.S)
        text = re.sub(r'<w:ins\b[^>]*>(.*?)</w:ins>', r'\1', text, flags=re.S)
        text = re.sub(r'<w:ins\b[^>]*/>', '', text)
        text = re.sub(r'<w:commentRangeStart\b[^>]*/>', '', text)
        text = re.sub(r'<w:commentRangeEnd\b[^>]*/>', '', text)
        text = re.sub(r'<w:r[^>]*>\s*<w:commentReference\b[^>]*/>.*?</w:r>', '', text, flags=re.S)
        return text

    for n in [x for x in names if x.startswith("word/") and x.endswith(".xml")]:
        raw_new = clean(items[n])
        if n == "word/settings.xml":
            raw_new = re.sub(r'<w:documentProtection\b[^>]*/>', '', raw_new)
            raw_new = re.sub(r'<w:documentProtection\b[^>]*>.*?</w:documentProtection>', '', raw_new, flags=re.S)
        data = raw_new.encode("utf-8")
        if data != items[n]:
            items[n] = data
            changed = True
    if "word/comments.xml" in items:
        del items["word/comments.xml"]
        changed = True
    if not changed:
        return

    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for n, data in items.items():
            zout.writestr(n, data)
    os.replace(tmp, docx_path)


def _docx_to_pdf(docx_path, out_dir):
    """docx 转 PDF：优先本机 Word 组件，回退 soffice。返回是否成功。"""
    pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
    pdf_path = os.path.join(out_dir, pdf_name)
    try:
        import win32com.client
        _strip_revisions(docx_path)
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        try:
            doc = word.Documents.Open(os.path.abspath(docx_path), ReadOnly=True,
                                      AddToRecentFiles=False, Visible=False)
            try:
                doc.AcceptAllRevisions()
            except Exception:
                pass
            doc.ExportAsFixedFormat(pdf_path, 17)  # 17 = wdExportDocumentPDF
            doc.Close(False)
        finally:
            word.Quit()
        return os.path.exists(pdf_path)
    except Exception as e:
        print("  Word 组件转换失败，尝试 soffice: %s" % e)
        subprocess.run(['soffice', '--headless', '--convert-to', 'pdf', '--outdir', out_dir, docx_path],
                       capture_output=True)
        return os.path.exists(pdf_path)


if docx_files:
    # 沙箱/受限环境下 Word 组件无法启动，可用 SKIP_PDF=1 先出 docx，
    # PDF 由 80_工具\to_pdf.py 单独转换（需在正常环境运行）。
    _skip_pdf = os.environ.get("SKIP_PDF") == "1"
    for docx_path in docx_files:
        pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + '.pdf'
        if _skip_pdf:
            print("  PDF: %s  [已跳过，稍后用 80_工具\\to_pdf.py 转换]" % pdf_name)
            continue
        ok = _docx_to_pdf(docx_path, OUTPUT_PDF_ACCEPT_DIR)
        print("  PDF: %s%s" % (pdf_name, "" if ok else "  [WARNING] 生成失败"))
else:
    print("跳过: 无验收单docx文件")

# ============================================================
# 步骤7.5：自动从发票PDF提取信息，写入DB(invoices表)
# ============================================================
print("\n" + "=" * 60)
print("步骤7.5：自动提取发票信息 → 写入DB(invoices表)")
print("=" * 60)

# 仅处理尚未入库的新发票（filename 不在 invoices 表）
invoice_pdfs_in_dir = [fn for fn in os.listdir(INVOICE_DIR) if fn.lower().endswith(".pdf")] if os.path.isdir(INVOICE_DIR) else []
fresh_pdfs = []
for fn in sorted(invoice_pdfs_in_dir):
    if db.get_invoice(fn) is None:
        fresh_pdfs.append(fn)
    else:
        print("  跳过: %s 已入库" % fn)

if fresh_pdfs:
    try:
        import pdfplumber
        invoice_records = []
        for inv_fn in sorted(fresh_pdfs):
            inv_path = os.path.join(INVOICE_DIR, inv_fn)
            try:
                with pdfplumber.open(inv_path) as pdf:
                    page = pdf.pages[0]
                    text = page.extract_text() or ""
                    tables = page.extract_tables()

                    # 提取价税合计
                    m_amt = re.search(r'（小写）[^\d]*(\d+[\.,]\d+)', text)
                    amount = round(float(m_amt.group(1).replace(',', '')), 2) if m_amt else None

                    # 提取销售方名称 - 优先从table cell（C3='销售方信息' → C4=名称）
                    seller = None
                    for table in tables:
                        for ri, row in enumerate(table):
                            for ci, cell in enumerate(row):
                                if cell and '销售方' in cell.replace('\n', ''):
                                    if ci + 1 < len(row) and row[ci + 1]:
                                        m_s = re.search(r'名称[：:]\s*(.+)', row[ci + 1])
                                        if m_s:
                                            seller = m_s.group(1).strip().split('\n')[0].strip()
                                    break
                            if seller:
                                break
                        if seller:
                            break

                    # 回退：从文本行中查找
                    if not seller:
                        for line in text.split('\n'):
                            if '售' in line and '方' in line and '名称' in line:
                                m_s = re.search(r'名称[：:]\s*(.+)', line)
                                if m_s:
                                    seller = m_s.group(1).strip()
                                    break

                    if seller and amount:
                        # 用途：优先发票文件名“(用途)金额”，回退申购单按金额匹配
                        usage = resolve_invoice_usage(inv_fn, amount, "", context="发票提取")
                        db.add_invoice(inv_fn, seller, amount, usage)
                        invoice_records.append({"filename": inv_fn, "amount": amount, "seller": seller, "usage": usage})
                        print("  提取成功并入库: %s → %s / %.2f / 用途:%s" % (inv_fn, seller, amount, usage or "(空)"))
                    else:
                        print("  WARNING: %s 提取不完整(seller=%s, amount=%s)，跳过" % (inv_fn, seller, amount))
            except Exception as e:
                print("  WARNING: %s 读取失败: %s" % (inv_fn, str(e)))

        if invoice_records:
            print("发票信息已入库: %d 张" % len(invoice_records))
        else:
            print("跳过: 未能从任何发票PDF中提取有效信息")
    except ImportError:
        print("跳过: 缺少pdfplumber库，无法提取发票信息 (pip install pdfplumber)")
else:
    if invoice_pdfs_in_dir:
        print("跳过: 所有发票均已在DB中，无需重复提取")
    else:
        print("跳过: 发票目录无PDF文件")

# 存量补写用途：已入库但 usage 为空、且发票PDF仍在输入目录的，尝试补写并回写DB
for fn in sorted(invoice_pdfs_in_dir):
    iv = db.get_invoice(fn)
    if iv and not (iv.get("usage") or "").strip():
        new_usage = resolve_invoice_usage(fn, iv.get("amount") or 0, "", context="存量补写")
        if new_usage:
            db.update_invoice_usage(fn, new_usage)
            print("  存量补写用途: %s -> %s" % (fn, new_usage))

# ============================================================
# 步骤8~14：生成付款单（支持多发票批量处理）
# ============================================================
print("\n" + "=" * 60)
print("步骤8~14：生成付款单")
print("=" * 60)

# 从DB读取未处理发票（processed=0）
invoices = [iv for iv in db.all_invoices() if not iv.get("processed")]
has_invoice_json = len(invoices) > 0
has_tpl_pay = os.path.exists(TEMPLATE_PAY_XLSX)
has_db = os.path.exists(DATABASE_XLSX)

if has_invoice_json and has_tpl_pay and has_db:
    print("发票记录(待处理): %d 张" % len(invoices))

    import pandas as pd
    df_db = pd.read_excel(DATABASE_XLSX, sheet_name="Sheet1", header=None)
    pay_count = 0
    invoice_pdfs = [fn for fn in os.listdir(INVOICE_DIR) if fn.lower().endswith(".pdf")]

    for inv_idx, invoice in enumerate(invoices):
        print("\n--- 发票 %d/%d: %s ---" % (inv_idx + 1, len(invoices), invoice.get("filename", "")))
        amount = invoice.get("amount") or 0
        seller = invoice.get("seller", "")

        # 确定对应的PDF文件
        invoice_filename = invoice.get("filename", "")
        target_pdf = None
        for pf in invoice_pdfs:
            if pf == invoice_filename:
                target_pdf = pf
                break
        if target_pdf is None and invoice_pdfs:
            target_pdf = invoice_pdfs[0]

        # 提取用途：统一解析（已有值/文件名/申购单），并确保回写DB
        usage = resolve_invoice_usage(invoice_filename, amount, invoice.get("usage"), context="付款单")
        old_usage = (invoice.get("usage") or "").strip()
        if usage and usage != old_usage:
            db.update_invoice_usage(invoice_filename, usage)
            print("  用途已回写DB: %s" % usage)
        if not usage:
            print("  WARNING: 未找到用途，付款单用途将为空")

        m2 = re.search(r"([^省市县区]+(?:经营部|公司|厂|店))", seller)
        supplier_short_pay = m2.group(1) if m2 else seller
        print("发票: %s, 金额: %.2f" % (supplier_short_pay, amount))

        # 移动发票PDF
        os.makedirs(INVOICE_DONE_DIR, exist_ok=True)
        if target_pdf and os.path.exists(os.path.join(INVOICE_DIR, target_pdf)):
            src = os.path.join(INVOICE_DIR, target_pdf)
            new_name = naming.invoice_name(supplier_short_pay, amount)
            dst = os.path.join(INVOICE_DONE_DIR, new_name)
            shutil.move(src, dst)
            db.mark_invoice_processed(target_pdf)
            print("  移动: %s → %s" % (target_pdf, new_name))
            if target_pdf in invoice_pdfs:
                invoice_pdfs.remove(target_pdf)

        # 匹配付款单数据库
        db_row = None
        keywords = [k for k in seller.replace("荆州","").replace("沙市区","").replace("荆州区","").replace("市","").split(" ") if len(k) >= 2]
        for idx, row in df_db.iterrows():
            if idx < 2: continue
            cell_val = str(row.iloc[2]) if pd.notna(row.iloc[2]) else ""
            for kw in keywords:
                if len(kw) >= 2 and kw in cell_val:
                    db_row = row
                    break
            if db_row is not None: break

        if db_row is not None:
            supplier_full = str(db_row.iloc[2])
            account = str(db_row.iloc[3])
            bank = str(db_row.iloc[4])
            phone = str(db_row.iloc[5]).strip() if pd.notna(db_row.iloc[5]) and str(db_row.iloc[5]).strip() else ""
            location = str(db_row.iloc[7]).strip() if pd.notna(db_row.iloc[7]) else ""
            print("  匹配: %s" % supplier_full)
        else:
            supplier_full, account, bank, phone, location = seller, "", "", "", ""
            print("  WARNING: 未匹配到供应商")

        # 每张发票重新解包模板
        if os.path.exists(UNPACKED_PAY): shutil.rmtree(UNPACKED_PAY)
        with zipfile.ZipFile(TEMPLATE_PAY_XLSX, 'r') as zf:
            zf.extractall(UNPACKED_PAY)

        NS_SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        ss_path = os.path.join(UNPACKED_PAY, "xl", "sharedStrings.xml")
        tree = etree.parse(ss_path)
        root = tree.getroot()
        ssts = root.findall("{%s}si" % NS_SS)

        cn_digits = ["零","壹","贰","叁","肆","伍","陆","柒","捌","玖"]
        amount_int = int(amount * 100)
        yuan, jiao, fen = amount_int // 100, (amount_int % 100) // 10, amount_int % 10
        digit_vals = [yuan//1000, (yuan%1000)//100, (yuan%100)//10, yuan%10, jiao, fen]
        for pos, run_idx in enumerate([3,5,7,9,11,13]):
            ssts[13].findall("{%s}r" % NS_SS)[run_idx].find("{%s}t" % NS_SS).text = cn_digits[digit_vals[pos]]
        ssts[13].findall("{%s}r" % NS_SS)[15].find("{%s}t" % NS_SS).text = "%.2f" % amount

        def add_ss(text):
            si = etree.SubElement(root, "{%s}si" % NS_SS)
            r = etree.SubElement(si, "{%s}r" % NS_SS)
            t = etree.SubElement(r, "{%s}t" % NS_SS)
            t.text = text
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            return len(root.findall("{%s}si" % NS_SS)) - 1

        idx_usage = add_ss(usage)
        idx_supplier = add_ss(supplier_full)
        idx_account = add_ss(account)
        idx_bank = add_ss(bank)
        idx_location = add_ss(location)
        idx_phone = add_ss(phone)

        tree.write(ss_path, xml_declaration=True, encoding="UTF-8", standalone=True)
        with open(ss_path, "r", encoding="utf-8") as f:
            content = f.read()
        total_ss = len(root.findall("{%s}si" % NS_SS))
        content = re.sub(r'count="\d+"', 'count="%d"' % total_ss, content)
        content = re.sub(r'uniqueCount="\d+"', 'uniqueCount="%d"' % total_ss, content)
        with open(ss_path, "w", encoding="utf-8") as f:
            f.write(content)

        sheet_path = os.path.join(UNPACKED_PAY, "xl", "worksheets", "sheet1.xml")
        tree = etree.parse(sheet_path)
        root = tree.getroot()
        for cell_ref, ss_idx in [("H4", idx_usage), ("G7", idx_supplier), ("G8", idx_account), ("J8", idx_bank), ("G9", idx_location), ("K9", idx_phone)]:
            c = root.find('.//{%s}c[@r="%s"]' % (NS_SS, cell_ref))
            if c is not None:
                c.set("t", "s")
                for ch in list(c): c.remove(ch)
                v = etree.SubElement(c, "{%s}v" % NS_SS)
                v.text = str(ss_idx)
        tree.write(sheet_path, xml_declaration=True, encoding="UTF-8", standalone=True)

        output_name = naming.invoice_payment_name(supplier_short_pay, int(amount))
        output_xlsx = os.path.join(OUTPUT_PAY_XLSX_DIR, output_name)
        with zipfile.ZipFile(output_xlsx, 'w', zipfile.ZIP_DEFLATED) as zf:
            for dp, dn, fn in os.walk(UNPACKED_PAY):
                for f in fn:
                    fp = os.path.join(dp, f)
                    zf.write(fp, os.path.relpath(fp, UNPACKED_PAY))
        print("  付款单xlsx: %s" % output_name)

        try:
            import win32com.client
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.DisplayAlerts = False
            wb = xl.Workbooks.Open(os.path.abspath(output_xlsx))
            pdf_path = os.path.join(OUTPUT_PDF_PAY_DIR, output_name.replace(".xlsx", ".pdf"))
            wb.ExportAsFixedFormat(0, pdf_path)
            wb.Close(False)
            xl.Quit()
            print("  付款单PDF: %s" % os.path.basename(pdf_path))
        except Exception:
            subprocess.run(["soffice","--headless","--convert-to","pdf","--outdir",OUTPUT_PDF_PAY_DIR,output_xlsx], capture_output=True)
            print("  付款单PDF(soffice): %s" % output_name.replace(".xlsx", ".pdf"))

        # 验证
        print("  --- 验证 ---")
        vdf = pd.read_excel(output_xlsx, header=None)
        for label, (r, c), exp in [("H4用途",(3,7),usage),("G7收款单位",(6,6),supplier_full),("G8账号",(7,6),account),("J8开户行",(7,9),bank),("G9地点",(8,6),location),("K9电话",(8,10),phone)]:
            try: act = str(vdf.iloc[r,c]).strip() if pd.notna(vdf.iloc[r,c]) else ""
            except IndexError: act = ""
            print("    %s: %s" % (label, "OK" if act == exp else "FAIL [%s]" % act))
        pay_count += 1

        # 阶段5：写入付款台账（含付款单字段 + 关联物料明细）
        ledger_pay = {
            "sign": os.path.splitext(output_name)[0],
            "date": time.strftime("%Y-%m-%d"),
            "seller": supplier_full,
            "amount": amount,
            "usage": usage,
            "account": account,
            "bank": bank,
        }
        ledger_path_i = paths.payment_ledger_path()
        materials = payment_ledger.match_materials(seller, amount, all_records)
        payment_ledger.append_payment(ledger_path_i, ledger_pay, materials)
        print("  付款台账: %s (%d行物料)" % (os.path.basename(ledger_path_i), len(materials)))

else:
    if not has_invoice_json:
        print("跳过: 无未处理的发票（invoices表均 processed=1 或为空）")
    if not has_tpl_pay:
        print("跳过: 付款单模板不存在: %s" % TEMPLATE_PAY_XLSX)
    if not has_db:
        print("跳过: 付款单数据库不存在: %s" % DATABASE_XLSX)

print("\n" + "=" * 60)
print("全部完成! 采购台账(%d条) + 验收单(%d份) + 付款单(%d份)" % (len(all_records), len(all_receipts), pay_count))
print("=" * 60)

if usage_conflicts:
    print("\n" + "#" * 60)
    print("# 以下 %d 张发票申购单与事项审批用途不一致，请逐张核对后补填用途：" % len(usage_conflicts))
    print("#" * 60)
    for ci, c in enumerate(usage_conflicts, 1):
        print("  %d) 发票: %s" % (ci, c["filename"] or "（未知）"))
        print("     金额: %s   | 申购单用途: %s   | 事项审批用途: %s"
              % (c["amount"], c["purchase"], c["approval"]))
    print("  ↳ 处理后请在该发票对应的付款单用途栏人工补填，或重新生成。")
    print("#" * 60)

# 机制可观测性：列出 DB 中用途仍为空的发票（含存量历史发票），便于人工发现并补填
# 已确认忽略项：dzfp_...荆州浦华荆清水务有限公司_7620.pdf（金额7620）用途为空，无需提示
_IGNORE_EMPTY_USAGE = ("7620.0", "荆州浦华荆清水务")
_empty_usage = [
    iv for iv in db.all_invoices()
    if not (iv.get("usage") or "").strip()
    and not (str(iv.get("amount")) == _IGNORE_EMPTY_USAGE[0] and _IGNORE_EMPTY_USAGE[1] in (iv.get("filename") or ""))
]
if _empty_usage:
    print("\n" + "#" * 60)
    print("# 以下 %d 张发票用途仍为空（DB 未记录）：请人工确认后补填" % len(_empty_usage))
    print("# 补填方式：付款单用途栏人工填写，或使用 manage_db 更新该发票 usage")
    for _iv in _empty_usage:
        print("  - %s (金额=%s，processed=%s)" % (_iv["filename"], _iv.get("amount"), _iv.get("processed")))
    print("#" * 60)

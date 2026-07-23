from __future__ import annotations

import json
import re
import shutil
import sys
import threading
import time
import traceback
import unicodedata
import zipfile
from collections import Counter
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable
from xml.etree import ElementTree

import numpy as np
import pypdfium2 as pdfium
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Cm
from rapidocr_onnxruntime import RapidOCR


IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".webp", ".tif", ".tiff"}
OUTPUT_NAME = "整理结果"
MAX_IMAGE_HEIGHT_CM = 23.0
OCR = RapidOCR()


class RunLogger:
    """Emit immediate console progress and retain a redacted run log."""

    def __init__(self, output: Path) -> None:
        self.started_at = time.monotonic()
        self.log_path = output / "运行日志.txt"
        self.current = ""
        self._lock = threading.Lock()
        self._stop = threading.Event()
        self._heartbeat = threading.Thread(target=self._emit_heartbeat, daemon=True)
        self.log_path.write_text("", encoding="utf-8-sig")

    def start(self) -> None:
        self._heartbeat.start()

    def info(self, message: str) -> None:
        elapsed = time.monotonic() - self.started_at
        line = f"[{elapsed:6.1f}s] {message}"
        with self._lock:
            print(line, flush=True)
            with self.log_path.open("a", encoding="utf-8") as log_file:
                log_file.write(line + "\n")

    def set_current(self, message: str) -> None:
        with self._lock:
            self.current = message
        self.info(message)

    def clear_current(self) -> None:
        with self._lock:
            self.current = ""

    def stop(self) -> None:
        self._stop.set()
        self._heartbeat.join(timeout=1)

    def _emit_heartbeat(self) -> None:
        while not self._stop.wait(10):
            with self._lock:
                current = self.current
            if current:
                self.info(f"仍在处理：{current}，请勿关闭窗口。")


@dataclass
class Evidence:
    path: Path
    kind: str
    text: str
    lines: list[str]
    amount: Decimal | None = None
    created_date: str | None = None
    invoice_date: str | None = None
    order_ids: set[str] = field(default_factory=set)
    product_name: str | None = None


def normalize_text(text: str) -> str:
    return text.replace("：", ":").replace("，", ",").replace("￥", "¥")


def decimal_value(value: str) -> Decimal | None:
    try:
        return Decimal(value.replace(",", "")).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def ocr_image(image: Image.Image) -> tuple[str, list[str]]:
    result, _ = OCR(np.asarray(image))
    if not result:
        return "", []
    lines = [str(item[1]).strip() for item in result if len(item) >= 2 and str(item[1]).strip()]
    return normalize_text("\n".join(lines)), lines


def extract_orders(text: str) -> set[str]:
    candidates = set(re.findall(r"(?<!\d)\d{16,32}(?!\d)", text.replace(" ", "")))
    return {item for item in candidates if not re.fullmatch(r"20\d{12,18}", item)}


def extract_date(text: str) -> str | None:
    compact = re.sub(r"\s+", " ", text)
    patterns = [
        r"创建时间\s*[:：]?\s*(20\d{2})[年./-](\d{1,2})[月./-](\d{1,2})",
        r"订单信息\s*(20\d{2})[年./-](\d{1,2})[月./-](\d{1,2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, compact)
        if match:
            return f"{int(match.group(1)):04d}.{int(match.group(2)):02d}.{int(match.group(3)):02d}"
    return None


def extract_invoice_date(text: str) -> str | None:
    compact = re.sub(r"\s+", " ", text)
    match = re.search(r"(20\d{2})[年./-](\d{1,2})[月./-](\d{1,2})(?:日)?", compact)
    if not match:
        return None
    return f"{int(match.group(1)):04d}.{int(match.group(2)):02d}.{int(match.group(3)):02d}"


def amounts_near_label(text: str, labels: Iterable[str]) -> list[Decimal]:
    found: list[Decimal] = []
    compact = re.sub(r"\s+", " ", text)
    for label in labels:
        for match in re.finditer(label + r".{0,45}?[-¥]?\s*(\d{1,8}(?:\.\d{1,2})?)", compact, re.I):
            value = decimal_value(match.group(1))
            if value is not None and value > 0:
                found.append(value)
    return found


def extract_payment_amount(text: str, lines: list[str]) -> Decimal | None:
    labeled: list[Decimal] = []
    strict_money = re.compile(r"(?:[¥￥]\s*(\d{1,8}(?:\.\d{1,2})?)|(?<![\d.-])(\d{1,8}\.\d{2})(?!\d))")
    for index, line in enumerate(lines):
        if "实付款" not in line:
            continue
        for nearby in lines[index : index + 4]:
            for match in strict_money.finditer(nearby):
                raw = match.group(1) or match.group(2)
                value = decimal_value(raw)
                if value is not None and value > 0:
                    labeled.append(value)
    if not labeled:
        compact = re.sub(r"\s+", " ", text)
        match = re.search(r"实付款.{0,100}?(?:[¥￥]\s*(\d{1,8}(?:\.\d{1,2})?)|(\d{1,8}\.\d{2}))", compact)
        if match:
            value = decimal_value(match.group(1) or match.group(2))
            if value is not None and value > 0:
                labeled.append(value)
    standalone: list[Decimal] = []
    for line in lines:
        match = re.fullmatch(r"\s*-?\s*[¥￥]?\s*(\d{1,7}\.\d{2})\s*", line)
        if match:
            value = decimal_value(match.group(1))
            if value and value > 0:
                standalone.append(value)
    if labeled:
        return max(labeled)
    return max(standalone) if standalone else None


def extract_invoice_amount(text: str, lines: list[str]) -> Decimal | None:
    money = []
    for line in lines:
        for raw in re.findall(r"[¥￥]\s*(\d{1,8}(?:\.\d{1,2})?)", line):
            value = decimal_value(raw)
            if value and value > 0:
                money.append(value)
    if money:
        return max(money)
    compact = re.sub(r"\s+", " ", text)
    match = re.search(r"(?:小写|价税合计).{0,80}?[¥￥]\s*(\d{1,8}(?:\.\d{1,2})?)", compact)
    return decimal_value(match.group(1)) if match else None


def classify_screenshot(text: str) -> str:
    payment_score = sum(token in text for token in ("账单详情", "当前状态", "支付成功", "收单机构", "支付方式", "付款方式", "交易详情"))
    product_score = sum(token in text for token in ("交易成功", "订单信息", "实付款", "成交时间", "发货时间"))
    if payment_score >= 2 and payment_score >= product_score:
        return "payment"
    if product_score >= 2:
        return "product"
    return "unknown"


def abbreviate_invoice_project_name(value: str, max_length: int = 20) -> str | None:
    """Convert an invoice 项目名称 into a filename-friendly fallback name."""
    name = unicodedata.normalize("NFKC", value)
    name = re.sub(r"^\*[^*]{1,30}\*", "", name).strip()
    name = re.split(r"(?:颜色分类|规格型号|商品编码)\s*[:：]?", name, maxsplit=1)[0]
    name = re.sub(r"\s+", "", name)
    # OCR can merge an invoice item's unit, quantity, unit price, amount, and tax
    # columns directly after its name (for example: 钳子把34.8514851485).
    name = re.sub(
        r"(?<=[\u4e00-\u9fff])(?:只|个|件|套|把|米|台|支|盒|包|张|瓶|千克|公斤|组)\d+(?:\.\d+)?(?:\d+(?:\.\d+)?)*.*$",
        "",
        name,
    )
    name = re.sub(r"(?:[¥￥]|\b(?:税率|单价|数量|金额)\b).*?$", "", name, flags=re.I)
    # Invoice item names are often followed by a long size/specification list.
    name = re.sub(
        r"(?<=[\u4e00-\u9fff])\d+(?:\.\d+)?(?:[/×xX*]\d+(?:\.\d+)?)+.*$",
        "",
        name,
    )
    name = re.sub(r"[\\/:*?\"<>|]+", "", name).strip("._-，,；;、")
    if not name:
        return None
    if len(name) > max_length:
        name = name[:max_length].rstrip("._-，,；;、")
    return name or None


def extract_invoice_project_name(text: str, lines: list[str]) -> str | None:
    """Read 项目名称 from the invoice item row; screenshot text must not name files."""
    normalized_lines = [unicodedata.normalize("NFKC", line).strip() for line in lines]
    for index, line in enumerate(normalized_lines):
        match = re.match(r"^\*[^*]{1,30}\*(.+)$", line)
        if not match:
            continue

        parts = [match.group(1).strip()]
        for continuation in normalized_lines[index + 1 : index + 8]:
            if not continuation:
                continue
            if re.match(r"^\*[^*]{1,30}\*", continuation):
                break
            if any(token in continuation for token in ("颜色分类", "合计", "价税合计", "备注", "开票人")):
                break
            # Unit/quantity/unit-price/amount/tax columns mark the end of the item description.
            if re.search(r"(?:^|\s)(?:只|个|件|套|米|台|支|盒|包)\s+\d", continuation) and re.search(
                r"\d+(?:\.\d+)?\s+\d+(?:\.\d+)?", continuation
            ):
                break
            parts.append(continuation)

        project_name = abbreviate_invoice_project_name("".join(parts))
        if project_name:
            return project_name

    # Fallback for OCR that merges the item row into one line.
    match = re.search(r"\*[^*\n]{1,30}\*([^\n]{2,120})", unicodedata.normalize("NFKC", text))
    return abbreviate_invoice_project_name(match.group(1)) if match else None


def extract_model_tokens(text: str) -> list[str]:
    normalized = unicodedata.normalize("NFKC", text).upper()
    patterns = (
        r"(?<![A-Z0-9])([A-Z]\d+(?:[A-Z]\d+)+)(?![A-Z0-9])",
        r"(?<![A-Z0-9])([A-Z]{1,8}\d+(?:\.\d+)?[A-Z]?)(?![A-Z0-9])",
    )
    found: list[str] = []
    for pattern in patterns:
        found.extend(re.findall(pattern, normalized))
    return sorted(
        set(
            item
            for item in found
            if not item.startswith("20")
            and len(item) <= 8
            and not re.fullmatch(r"[A-Z]?\d+", item)
        ),
        key=lambda item: (-len(item), item),
    )


def abbreviate_transaction_product_name(value: str, max_length: int = 16) -> str | None:
    """Keep the identifying portion of a product-title line from a transaction page."""
    name = unicodedata.normalize("NFKC", value)
    name = re.sub(r"[¥￥].*$", "", name)
    name = re.split(r"[;；]", name, maxsplit=1)[0]
    name = re.split(r"(?:颜色分类|规格|型号|款式|套餐|服务保障|退货|商品总价|实付款)", name, maxsplit=1)[0]
    name = re.sub(r"(?:\s|，|,|;|；)*(?:x|×)\s*\d+.*$", "", name, flags=re.I)
    name = re.sub(r"长\d+(?:\.\d+)?\s*(?:cm|毫米|mm).*$", "", name, flags=re.I)
    # Descriptive titles often append a size/model sequence after a complete
    # Chinese product category, e.g. 水口剪钳170塑料钳子 -> 水口剪钳.
    name = re.sub(r"(?<=[\u4e00-\u9fff])\d{2,}.*$", "", name)
    name = re.sub(r"(?<=[\u4e00-\u9fff])\d+(?:\.\d+)*[.。:：]*$", "", name)
    name = re.sub(r"\s+", "", name)
    name = re.sub(r"[\\/:*?\"<>|]+", "", name).strip("._-，,；;、")
    if not name or len(re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", name)) < 2:
        return None
    return name[:max_length].rstrip("._-，,；;、") or None


def is_transaction_nonproduct_line(value: str) -> bool:
    return any(
        token in value
        for token in (
            "交易成功", "实付款", "订单信息", "订单编号", "创建时间", "成交时间", "发货时间", "付款时间",
            "商品总价", "收货地址", "进店逛逛", "申请退款", "加入购物车", "服务保障", "微信交易号",
            "支付方式", "查看更多", "催发货", "交易快照", "订单状态", "待发货", "天猫", "旗舰店",
            "账单详情", "全部账单", "安全电流", "采购权益", "进店逛逛", "厂家直销", "品质保证",
            "多色可选", "闲鱼转卖", "申请售后", "申请退款", "假一赔", "极速退款", "好评率",
        )
    )


def screenshot_name_candidate(text: str, lines: list[str]) -> str | None:
    """Use the displayed product title, rather than invoice wording, for names."""
    ranked: list[tuple[int, int, str]] = []
    product_words = ("端子线", "并联线", "转接线", "公转母", "转母", "插头", "垫片", "垫圈", "钳", "线束")
    for line in lines:
        original = unicodedata.normalize("NFKC", line).strip()
        if not original or is_transaction_nonproduct_line(original):
            continue
        candidate = abbreviate_transaction_product_name(original)
        if not candidate:
            continue
        # A displayed price or a meaningful product/model word makes this a
        # title line; nearby UI labels and standalone quantities do not qualify.
        if (
            "¥" in original
            or "￥" in original
            or extract_model_tokens(candidate)
            or len(re.findall(r"[\u4e00-\u9fff]", candidate)) >= 3
        ):
            category_score = sum(token in candidate for token in product_words)
            model_score = len(extract_model_tokens(candidate))
            chinese_count = len(re.findall(r"[\u4e00-\u9fff]", candidate))
            # Prefer real product/category wording over bare models or variants.
            price_score = 15 if "¥" in original or "￥" in original else 0
            score = category_score * 100 + model_score * 20 + min(chinese_count, 16) + price_score
            ranked.append((score, -len(ranked), candidate))
    return max(ranked)[2] if ranked else None


def select_product_name(invoice: Evidence, product: Evidence | None, payment: Evidence | None) -> tuple[str | None, str]:
    """Select a filename title with transaction evidence taking precedence."""
    if product and product.product_name:
        return normalize_product_alias(product.product_name), "商品交易截图"
    if payment and payment.product_name:
        return payment.product_name, "支付交易截图（弱兜底）"
    if invoice.product_name:
        return invoice.product_name, "发票项目名称（兜底）"
    return None, ""


def normalize_product_alias(name: str) -> str:
    """Apply concise, user-facing names for well-known electrical components."""
    if "MR30" in name.upper() and any(token in name for token in ("公转母", "公头母头", "三芯")):
        return "MR30三相线"
    return name


def inspect_screenshot(path: Path) -> Evidence:
    with Image.open(path) as source:
        image = source.convert("RGB")
        text, lines = ocr_image(image)
    kind = classify_screenshot(text)
    amount = extract_payment_amount(text, lines)
    return Evidence(
        path=path,
        kind=kind,
        text=text,
        lines=lines,
        amount=amount,
        created_date=extract_date(text),
        order_ids=extract_orders(text),
        product_name=screenshot_name_candidate(text, lines),
    )


def inspect_invoice(path: Path, cache_dir: Path) -> Evidence:
    document = pdfium.PdfDocument(str(path))
    page = document[0]
    text_page = page.get_textpage()
    text = normalize_text(text_page.get_text_bounded() or "")
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if len(re.sub(r"\s+", "", text)) < 30 or extract_invoice_amount(text, lines) is None:
        rendered = cache_dir / f"{path.stem}.png"
        image = page.render(scale=2.2).to_pil().convert("RGB")
        image.save(rendered)
        text, lines = ocr_image(image)
    text_page.close()
    page.close()
    document.close()
    return Evidence(
        path=path,
        kind="invoice",
        text=text,
        lines=lines,
        amount=extract_invoice_amount(text, lines),
        invoice_date=extract_invoice_date(text),
        order_ids=extract_orders(text),
        product_name=extract_invoice_project_name(text, lines),
    )


def shared_order(a: Evidence, b: Evidence) -> bool:
    if not a.order_ids or not b.order_ids:
        return False
    return bool(a.order_ids & b.order_ids)


def unique_match(source: Evidence, candidates: list[Evidence]) -> Evidence | None:
    same_amount = [item for item in candidates if item.amount == source.amount and item.amount is not None]
    if len(same_amount) == 1:
        return same_amount[0]
    by_order = [item for item in same_amount if shared_order(source, item)]
    return by_order[0] if len(by_order) == 1 else None


def safe_name(value: str) -> str:
    value = re.sub(r"[\\/:*?\"<>|]", " ", value)
    value = re.sub(r"\s+", " ", value).strip(" .")
    return value[:80]


def add_image_page(document: Document, path: Path, first: bool) -> None:
    paragraph = document.add_paragraph()
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Cm(0)
    paragraph.paragraph_format.space_after = Cm(0)
    with Image.open(path) as image:
        width, height = image.size
    max_width = 19.0
    max_height = MAX_IMAGE_HEIGHT_CM
    scale = min(max_width / width, max_height / height)
    paragraph.add_run().add_picture(str(path), width=Cm(width * scale), height=Cm(height * scale))


def write_word(path: Path, images: list[Path]) -> None:
    if not images:
        raise ValueError("Word 至少需要一张交易截图")
    document = Document()
    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = section.bottom_margin = Cm(1)
    section.left_margin = section.right_margin = Cm(1)
    for index, image in enumerate(images):
        add_image_page(document, image, index == 0)
        if index < len(images) - 1:
            document.paragraphs[-1].add_run().add_break(WD_BREAK.PAGE)
    document.save(path)


def validate_word(path: Path, expected_images: int) -> list[str]:
    namespaces = {
        "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    }
    with zipfile.ZipFile(path) as archive:
        media = [name for name in archive.namelist() if name.startswith("word/media/")]
        root = ElementTree.fromstring(archive.read("word/document.xml"))
    errors = []
    if len(media) != expected_images:
        errors.append(f"Word 图片数量为 {len(media)}，应为 {expected_images}")
    if root.findall(".//w:t", namespaces):
        errors.append("Word 中存在额外文字")
    expected_breaks = expected_images - 1
    actual_breaks = len(root.findall(".//w:br[@w:type='page']", namespaces))
    if actual_breaks != expected_breaks:
        errors.append(f"Word 分页符数量为 {actual_breaks}，应为 {expected_breaks}")
    heights = [int(node.attrib["cy"]) for node in root.findall(".//wp:extent", namespaces)]
    if len(heights) != expected_images or any(value > 8_280_000 for value in heights):
        errors.append("Word 图片高度超过 23 cm")
    return errors


def amount_name(amount: Decimal) -> str:
    return f"{amount:.2f}"


def make_output_root(root: Path) -> Path:
    candidate = root / OUTPUT_NAME
    if not candidate.exists():
        candidate.mkdir()
        return candidate
    index = 2
    while (root / f"{OUTPUT_NAME}_{index}").exists():
        index += 1
    candidate = root / f"{OUTPUT_NAME}_{index}"
    candidate.mkdir()
    return candidate


def redacted_orders(values: set[str]) -> str:
    return ",".join("***" + item[-4:] for item in sorted(values)) or "-"


def main(root: Path) -> int:
    output = make_output_root(root)
    cache = output / "识别缓存"
    cache.mkdir()
    logger = RunLogger(output)
    logger.start()
    logs: list[dict[str, str]] = []
    issues: list[str] = []

    pdf_paths = sorted(path for path in root.glob("*.pdf") if path.is_file())
    image_paths = sorted(path for path in root.iterdir() if path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS)
    logger.info(f"开始整理：发现 {len(pdf_paths)} 张发票、{len(image_paths)} 张截图。")
    logger.info(f"结果将写入：{output.name}")
    if not pdf_paths:
        issues.append("文件夹根目录中必须存在 PDF 发票。")
        logger.info("提示：根目录中未发现发票 PDF。")

    screenshots = []
    for index, path in enumerate(image_paths, start=1):
        logger.set_current(f"正在识别截图 {index}/{len(image_paths)}：{path.name}")
        evidence = inspect_screenshot(path)
        screenshots.append(evidence)
        logger.info(f"截图识别完成：类型={evidence.kind}，金额={evidence.amount or '未识别'}。")
        logs.append({
            "file": path.name,
            "type": evidence.kind,
            "amount": str(evidence.amount or ""),
            "date": evidence.created_date or "",
            "orders": redacted_orders(evidence.order_ids),
            "product": evidence.product_name or "",
        })

    invoices = []
    for index, path in enumerate(pdf_paths, start=1):
        logger.set_current(f"正在读取发票 {index}/{len(pdf_paths)}：{path.name}")
        evidence = inspect_invoice(path, cache)
        invoices.append(evidence)
        logger.info(f"发票识别完成：金额={evidence.amount or '未识别'}，项目={evidence.product_name or '未识别'}。")
        logs.append({
            "file": path.name,
            "type": "invoice",
            "amount": str(evidence.amount or ""),
            "date": evidence.invoice_date or "",
            "orders": redacted_orders(evidence.order_ids),
            "product": evidence.product_name or "",
        })

    products = [item for item in screenshots if item.kind == "product"]
    payments = [item for item in screenshots if item.kind == "payment"]
    unknown = [item for item in screenshots if item.kind == "unknown"]
    for item in unknown:
        issues.append(f"无法判断截图类型：{item.path.name}")

    product_by_amount = Counter(item.amount for item in products if item.amount is not None)
    payment_by_amount = Counter(item.amount for item in payments if item.amount is not None)
    used_products: set[Path] = set()
    used_payments: set[Path] = set()
    used_invoices: set[Path] = set()
    completed = 0
    for index, invoice in enumerate(invoices, start=1):
        logger.set_current(f"正在配对并生成资料包 {index}/{len(invoices)}：{invoice.path.name}")
        missing = []
        if invoice.amount is None:
            missing.append("发票金额")
        product = unique_match(invoice, [item for item in products if item.path not in used_products])
        payment = unique_match(invoice, [item for item in payments if item.path not in used_payments])
        if invoice.amount is not None and product_by_amount[invoice.amount] > 1 and not invoice.order_ids:
            product = None
        if invoice.amount is not None and payment_by_amount[invoice.amount] > 1 and not invoice.order_ids:
            payment = None
        images = [item.path for item in (product, payment) if item is not None]
        if not images:
            missing.append("至少一张商品或支付交易截图")
        product_name, name_source = select_product_name(invoice, product, payment)
        if not product_name:
            missing.append("发票项目名称")
        created_date = (
            (product.created_date if product else None)
            or (payment.created_date if payment else None)
            or invoice.invoice_date
        )
        if not created_date:
            missing.append("商品创建时间、支付创建时间或发票开票日期")
        if missing:
            issues.append(f"{invoice.path.name}：缺少或无法确认 {'、'.join(missing)}")
            logger.info(f"待确认：{'、'.join(missing)}。")
            continue

        name = safe_name(f"{amount_name(invoice.amount)} {product_name} {created_date}")
        package = output / name
        package.mkdir(exist_ok=False)
        pdf_target = package / f"{name}.pdf"
        word_target = package / f"{name}.docx"
        shutil.copy2(invoice.path, pdf_target)
        write_word(word_target, images)
        errors = validate_word(word_target, len(images))
        if errors:
            issues.append(f"{name}：" + "；".join(errors))
            shutil.rmtree(package)
            logger.info(f"待确认：Word 校验失败：{'；'.join(errors)}。")
            continue
        if product:
            used_products.add(product.path)
        if payment:
            used_payments.add(payment.path)
        used_invoices.add(invoice.path)
        completed += 1
        logs.append({
            "file": invoice.path.name,
            "type": "naming",
            "amount": str(invoice.amount or ""),
            "date": created_date,
            "orders": redacted_orders(invoice.order_ids),
            "product": product_name,
            "name_source": name_source,
            "output": name,
        })
        logger.info(f"完成 {completed}/{len(invoices)}：{name}（名称来自{name_source}）。")

    for item in products:
        if item.path not in used_products:
            issues.append(f"未配对商品截图：{item.path.name}")
    for item in payments:
        if item.path not in used_payments:
            issues.append(f"未配对支付截图：{item.path.name}")

    (output / "识别日志.json").write_text(json.dumps(logs, ensure_ascii=False, indent=2), encoding="utf-8")
    if issues:
        (output / "待确认.txt").write_text("\n".join(issues) + "\n", encoding="utf-8-sig")
    if cache.exists():
        shutil.rmtree(cache)
    logger.clear_current()
    logger.info(f"整理完成：共 {completed} 组，待确认 {len(issues)} 项。")
    logger.info(f"输出目录：{output}")
    logger.info(f"运行日志：{logger.log_path}")
    logger.stop()
    return 0 if completed > 0 and not issues else 2


if __name__ == "__main__":
    try:
        target = Path(sys.argv[1]).resolve() if len(sys.argv) > 1 else Path.cwd()
        raise SystemExit(main(target))
    except Exception:
        error_file = Path(sys.argv[1] if len(sys.argv) > 1 else ".") / "发票整理错误日志.txt"
        error_file.write_text(traceback.format_exc(), encoding="utf-8-sig")
        print(f"运行失败，错误详情：{error_file}")
        raise SystemExit(99)

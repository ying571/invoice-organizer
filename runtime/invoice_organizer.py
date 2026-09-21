from __future__ import annotations

import json
import re
import secrets
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
UNMATCHED_NAME = "未配对材料"
DUPLICATES_NAME = "重复文件"
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
    transfer_product_name: str | None = None


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
        r"创建时间\s*[:：]?\s*(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})",
        r"订单信息\s*(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})",
        r"转账时间\s*[:：]?\s*(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})",
        r"收款时间\s*[:：]?\s*(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})",
        r"支付时间\s*[:：]?\s*(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})",
    ]
    for pattern in patterns:
        match = re.search(pattern, compact)
        if match:
            return f"{int(match.group(1)):04d}.{int(match.group(2)):02d}.{int(match.group(3)):02d}"
    return None


def extract_invoice_date(text: str) -> str | None:
    compact = re.sub(r"\s+", " ", text)
    match = re.search(r"(20\d{2})\s*[年./-]\s*(\d{1,2})\s*[月./-]\s*(\d{1,2})(?:\s*日)?", compact)
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
    strict_money = re.compile(r"(?:[¥￥]\s*(\d{1,8}(?:\.\d{1,2})?)|(?<![\d.-])(\d{1,8}\.\d{1,2})(?!\d))")
    for index, line in enumerate(lines):
        if "实付款" not in line:
            continue
        # The price-breakdown lines immediately after 实付款 often contain a
        # discount (for example 共减¥6.2).  They are not the paid amount.
        for nearby in lines[index : index + 1]:
            for match in strict_money.finditer(nearby):
                raw = match.group(1) or match.group(2)
                value = decimal_value(raw)
                if value is not None and value > 0:
                    labeled.append(value)
    if not labeled:
        compact = re.sub(r"\s+", " ", text)
        match = re.search(r"实付款.{0,100}?(?:[¥￥]\s*(\d{1,8}(?:\.\d{1,2})?)|(\d{1,8}\.\d{1,2}))", compact)
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
    transfer_score = sum(token in text for token in ("已收款", "转账时间", "收款时间", "转账给", "转账成功"))
    if transfer_score >= 2 and transfer_score >= payment_score and transfer_score >= product_score:
        return "transfer"
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
        r"(?:pcs|只|个|件|套|把|米|台|支|盒|包|张|瓶|份|千克|公斤|组)\s*\d+(?:\.\d+)?(?:\d+(?:\.\d+)?)*.*$",
        "",
        name,
        flags=re.I,
    )
    name = re.sub(r"(?:[¥￥]|\b(?:税率|单价|数量|金额)\b).*?$", "", name, flags=re.I)
    # Invoice item names are often followed by a long size/specification list.
    name = re.sub(
        r"(?<=[\u4e00-\u9fff])\d+(?:\.\d+)?(?:[/×xX*]\d+(?:\.\d+)?)+.*$",
        "",
        name,
    )
    # Some e-invoice text layers concatenate the unit/quantity columns to a
    # product name without spaces, e.g. 数据线条24.09... -> 数据线.
    name = re.sub(r"(?:条|个|件|套|把|米|台|支|盒|包|份)\s*\d+(?:\.\d+)?(?:\d+(?:\.\d+)?)*.*$", "", name)
    name = re.sub(r"(?:家用|工业级|商用|办公).*$", "", name)
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


def extract_transfer_invoice_project_name(lines: list[str]) -> str | None:
    """Return the complete invoice item name after its final asterisk for transfers."""
    for line in (unicodedata.normalize("NFKC", item).strip() for item in lines):
        if line.count("*") < 2:
            continue
        name = line.rsplit("*", maxsplit=1)[-1].strip()
        # PDF text layers frequently append the unit, quantity, unit price,
        # amount, and tax columns after the project-name cell. They are not
        # part of the 项目名称, so remove only that table-column suffix.
        name = re.split(
            r"\s+(?:只|个|件|套|把|米|台|支|盒|包|张|瓶|千克|公斤|组)\s*\d",
            name,
            maxsplit=1,
        )[0].strip()
        name = re.sub(
            r"(?<=[\u4e00-\u9fff])(?:只|个|件|套|把|米|台|支|盒|包|张|瓶|千克|公斤|组)\d+(?:\.\d+)?(?:\d+(?:\.\d+)?)*.*$",
            "",
            name,
        )
        # A valid invoice item description must contain actual name text, rather
        # than only the value columns that may follow it in OCR output.
        if name and re.search(r"[\u4e00-\u9fffA-Za-z0-9]", name):
            return name
    return None


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
    name = re.sub(r"[（(]\s*\d+\s*(?:个|件|套|只|包|张).*$", "", name)
    name = re.sub(r"(?:\s|，|,|;|；)*(?:x|×)\s*\d+.*$", "", name, flags=re.I)
    name = re.sub(r"长\d+(?:\.\d+)?\s*(?:cm|毫米|mm).*$", "", name, flags=re.I)
    name = re.sub(r"(?:家用|工业级|商用|办公).*$", "", name)
    # Descriptive titles often append a size/model sequence after a complete
    # Chinese product category, e.g. 水口剪钳170塑料钳子 -> 水口剪钳.
    # Keep electrical ratings such as 12V/24V; strip only unqualified numeric
    # suffixes that are usually dimensions, quantities, or OCR noise.
    name = re.sub(r"(?<=[\u4e00-\u9fff])\d{2,}(?!\s*(?:V|A|P|MM|CM|W)).*\Z", "", name, flags=re.I)
    name = re.sub(r"(?<=[\u4e00-\u9fff])\d+(?:\.\d+)*[.。:：]*$", "", name)
    # '*' is illegal in Windows filenames; retain its dimensional meaning.
    name = re.sub(r"(?<=\d)\*(?=\d)", "x", name)
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
            # These are order-page labels, discounts, or seller metadata.  OCR
            # can make them look more "model-like" than the product title.
            "商家编码", "实付价", "实付款", "淘金币抵", "店铺优惠抵", "官方立减", "平台优惠",
            "红包抵", "优惠抵", "共减", "商品总价", "订单保障", "赠品",
            "上门取件", "运费", "延长收货", "确认收货", "已发货", "自动确认", "大促价保", "申请价保",
            "价格明细", "两条特惠装", "退货宝",
            "华南农业大学", "查看发票", "收货信息", "积分", "店铺优惠", "微信支付", "支付宝支付",
        )
    )


def has_shared_product_term(title: str, specification: str) -> bool:
    """Require a product-word overlap before a grey specification can replace a title."""
    title_terms = {
        run[start : end]
        for run in re.findall(r"[\u4e00-\u9fff]{2,}", title)
        for start in range(len(run) - 1)
        for end in range(start + 2, min(len(run), start + 4) + 1)
    }
    return any(term in specification for term in title_terms)


def preferred_specification_line(lines: list[str], title_index: int, title: str) -> str | None:
    """Keep only interface-defining pin counts from an adjacent spec line."""
    # P means pin count only for wires/connectors.  It is not retained for
    # LEDs, labels, tools, or unrelated product variants.
    wire_terms = ("线", "端子")
    if not any(term in title for term in wire_terms):
        return None
    # OCR may insert thumbnail text between the black title and its grey
    # specification, so examine the following short block rather than only the
    # immediately adjacent line.
    for raw_line in lines[title_index + 1 : title_index + 10]:
        original = unicodedata.normalize("NFKC", raw_line).strip()
        if not original or is_transaction_nonproduct_line(original):
            continue
        # A colour may be a meaningful specification.  Gifts and promotions
        # are not a product identity and must never replace the title.
        if any(token in original for token in ("送", "赠", "套餐", "优惠")):
            continue
        # For wires, the pin count changes the connector/interface.  Colour,
        # LED package, quantity, pitch, or generic material descriptions do
        # not normally belong in the concise filename.
        pin = re.search(r"(?<![A-Z0-9])(\d{1,2})\s*P(?![A-Z0-9])", original, re.I)
        if pin and not re.search(r"\d{1,2}\s*P", title, re.I):
            return f"{title} {pin.group(1)}P"
    return None


def screenshot_name_candidate(text: str, lines: list[str]) -> str | None:
    """Use the displayed product title, rather than invoice wording, for names."""
    # Payment OCR can lose Chinese glyphs while preserving the package-size
    # sequence and resistor values.  This pattern is distinctive enough to
    # recover the concise product family for common SMD resistor orders.
    package_hits = re.findall(r"(?:0402|0603|0805|1206|1210|2010|2512)", text)
    if len(set(package_hits)) >= 2 and re.search(r"(?:\bR\b|\d+R|\d+K|\d+M|欧)", text, re.I):
        return "贴片电阻"
    ranked: list[tuple[int, int, str, int]] = []
    product_words = (
        "端子线", "并联线", "转接线", "公转母", "转母", "插头", "垫片", "垫圈", "钳", "线束",
        "传感器", "排母", "排针", "分析仪", "卡纸", "电机", "电池", "数据线", "保护套", "剪刀",
        "灯带", "灯条", "电阻", "电阻器", "电阻网络",
    )
    for index, line in enumerate(lines):
        original = unicodedata.normalize("NFKC", line).strip()
        if not original or is_transaction_nonproduct_line(original):
            continue
        candidate = abbreviate_transaction_product_name(original)
        if not candidate:
            continue
        # A colour, length or charging specification without a product noun is
        # a variant line, not a reliable name (e.g. 1m白色老式安卓4A快充).
        if re.match(r"^\d+(?:\.\d+)?(?:cm|mm|m)?(?:白色|黑色|红色|蓝色|老式|新款|快充|规格)", candidate, re.I):
            continue
        # A displayed price or a meaningful product/model word makes this a
        # title line; nearby UI labels and standalone quantities do not qualify.
        if (
            "¥" in original
            or "￥" in original
            or extract_model_tokens(candidate)
            or any(token in candidate for token in product_words)
        ):
            category_score = sum(token in candidate for token in product_words)
            model_score = len(extract_model_tokens(candidate))
            chinese_count = len(re.findall(r"[\u4e00-\u9fff]", candidate))
            # Prefer real product/category wording over bare models or variants.
            # The actual title nearly always has its item price beside it; give
            # that stronger weight than an isolated model or specification.
            price_score = 60 if "¥" in original or "￥" in original else 0
            score = category_score * 100 + model_score * 20 + min(chinese_count, 16) + price_score
            ranked.append((score, -index, candidate, index))
    if not ranked:
        return None
    _, _, title, title_index = max(ranked)
    return preferred_specification_line(lines, title_index, title) or title


def select_product_name(
    invoice: Evidence,
    product: Evidence | None,
    payment: Evidence | None,
    transfer: Evidence | None,
) -> tuple[str | None, str]:
    """Select a filename title with transaction evidence taking precedence."""
    if product and product.product_name:
        return normalize_product_alias(product.product_name), "商品交易截图"
    if payment and payment.product_name:
        return normalize_product_alias(payment.product_name), "支付交易截图（弱兜底）"
    if transfer and invoice.transfer_product_name:
        return invoice.transfer_product_name, "发票项目名称（转账凭证）"
    if invoice.product_name:
        return invoice.product_name, "发票项目名称（兜底）"
    return None, ""


def normalize_product_alias(name: str) -> str:
    """Return a concise product entity, retaining only discriminating specifications."""
    normalized = unicodedata.normalize("NFKC", name)
    upper = normalized.upper()
    if "MR30" in upper and any(token in normalized for token in ("公转母", "公头母头", "三芯")):
        return "MR30三相线"
    if "TYPEC" in upper and any(token in upper for token in ("XH", "PH", "转接", "转接线")):
        return "Type-C转接线"
    # LED package, colour and quantity are selectable variants.  The concise
    # material name intentionally remains at the product-family level.
    if "贴片LED" in upper:
        return "贴片LED"
    if any(token in normalized for token in ("贴片电阻", "电阻器", "电阻网络", "电阻")):
        return "贴片电阻" if "贴片" in normalized else "电阻"
    if "灯带" in normalized or "灯条" in normalized:
        return "灯带"
    # Marketing copy (office/home use, sharp, anti-stick, wear-resistant) does
    # not identify the purchased good.
    if "剪刀" in normalized:
        return "剪刀"
    # Retain the actual label product, not paper stock or adhesive attributes.
    if "小标签贴" in normalized or "标签贴" in normalized:
        return "小标签贴"
    # Product compatibility prose after “适用于” is neither the model nor the
    # product identity.
    if "逻辑分析仪" in normalized:
        return "USB逻辑分析仪" if "USB" in upper else "逻辑分析仪"
    # KF-style OCR strings can fuse a model, pitch and pin count.  They are
    # error-prone and redundant when the connector construction is available.
    if "接线端子" in normalized:
        pin = re.search(r"(?<![A-Z0-9])(\d{1,2})\s*P(?![A-Z0-9])", upper)
        pin_suffix = f" {pin.group(1)}P" if pin else ""
        if "螺钉式" in normalized and "PCB" in upper:
            return "螺钉式PCB接线端子" + pin_suffix
        return "接线端子" + pin_suffix
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
        transfer_product_name=extract_transfer_invoice_project_name(lines),
    )


def shared_order(a: Evidence, b: Evidence) -> bool:
    if not a.order_ids or not b.order_ids:
        return False
    return bool(a.order_ids & b.order_ids)


def product_name_match_score(source: Evidence, candidate: Evidence) -> int:
    """Return a conservative title-overlap score for same-amount disambiguation."""
    if not source.product_name or not candidate.product_name:
        return 0
    source_name = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", source.product_name).casefold()
    candidate_name = re.sub(r"[^\u4e00-\u9fffA-Za-z0-9]", "", candidate.product_name).casefold()
    if len(source_name) < 2 or len(candidate_name) < 2:
        return 0
    if source_name in candidate_name or candidate_name in source_name:
        return min(len(source_name), len(candidate_name)) * 10
    # A two-character Chinese product term (e.g. 剪刀) is sufficient evidence;
    # one shared generic character is not.
    shared = set(source_name) & set(candidate_name)
    return len(shared) * 5 if len(shared) >= 2 else 0


def evidence_identity(item: Evidence) -> tuple[object, ...]:
    """Fields used to decide whether two OCR results are indistinguishable."""
    return (
        item.kind,
        item.amount,
        item.created_date or item.invoice_date or "",
        tuple(sorted(item.order_ids)),
        re.sub(r"\s+", "", item.product_name or "").casefold(),
        re.sub(r"\s+", "", item.transfer_product_name or "").casefold(),
    )


def identical_group(items: list[Evidence]) -> list[Evidence] | None:
    """Return a group whose known OCR fields agree; blank fields are missing data."""
    if len(items) < 2:
        return None
    fields = (
        lambda item: item.kind,
        lambda item: item.amount,
        lambda item: item.created_date or item.invoice_date or "",
        lambda item: tuple(sorted(item.order_ids)),
        lambda item: re.sub(r"\s+", "", item.product_name or "").casefold(),
        lambda item: re.sub(r"\s+", "", item.transfer_product_name or "").casefold(),
    )
    for get_value in fields:
        known_values = {get_value(item) for item in items if get_value(item) not in (None, "", ())}
        if len(known_values) > 1:
            return None
    return items


def randomly_select_identical(items: list[Evidence], label: str) -> tuple[Evidence, str]:
    """Choose one indistinguishable duplicate while leaving the rest unmatched."""
    selected = secrets.choice(items)
    paths = "；".join(str(item.path.resolve()) for item in sorted(items, key=lambda item: item.path.name.casefold()))
    detail = (
        f"存在 {len(items)} 个完全相同的{label}；随机选取 {selected.path.resolve()} 配对，"
        f"其余文件保留在原位置且不复制到未配对材料（重复文件地址：{paths}）"
    )
    return selected, detail


def evidence_label(kind: str) -> str:
    return {
        "product": "商品订单截图",
        "payment": "付款截图",
        "transfer": "转账截图",
        "invoice": "发票",
    }.get(kind, kind)


def unique_match(source: Evidence, candidates: list[Evidence]) -> tuple[Evidence | None, str, set[Path]]:
    same_amount = [item for item in candidates if item.amount == source.amount and item.amount is not None]
    if len(same_amount) == 1:
        return same_amount[0], "金额唯一", set()
    by_order = [item for item in same_amount if shared_order(source, item)]
    if len(by_order) == 1:
        return by_order[0], "金额和订单号", set()
    scored = [(product_name_match_score(source, item), item) for item in same_amount]
    scored = [(score, item) for score, item in scored if score > 0]
    if not scored:
        duplicates = identical_group(same_amount)
        if duplicates:
            selected, detail = randomly_select_identical(duplicates, evidence_label(duplicates[0].kind))
            return selected, detail, {item.path for item in duplicates if item.path != selected.path}
        return None, "同金额候选无订单号或商品名称交集", set()
    best_score = max(score for score, _ in scored)
    best = [item for score, item in scored if score == best_score]
    if len(best) == 1:
        return best[0], "金额和商品名称（待用户复核）", set()
    duplicates = identical_group(best)
    if duplicates:
        selected, detail = randomly_select_identical(duplicates, evidence_label(duplicates[0].kind))
        return selected, detail, {item.path for item in duplicates if item.path != selected.path}
    return None, "同金额候选的商品名称仍无法唯一确认", set()


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


def copy_review_materials(output: Path, folder_name: str, paths: Iterable[Path], logger: RunLogger) -> int:
    """Copy source files into one review folder without changing originals."""
    source_paths = sorted(set(paths), key=lambda path: path.name.casefold())
    if not source_paths:
        return 0

    review_dir = output / folder_name
    review_dir.mkdir(exist_ok=True)
    copied = 0
    for source in source_paths:
        target = review_dir / source.name
        if target.exists():
            # Source files normally share one input directory, but keep every
            # file if a caller ever supplies duplicate basenames.
            index = 2
            while True:
                target = unmatched_dir / f"{source.stem} ({index}){source.suffix}"
                if not target.exists():
                    break
                index += 1
        shutil.copy2(source, target)
        copied += 1
    logger.info(f"已复制 {copied} 个文件到：{folder_name}")
    return copied


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
    transfers = [item for item in screenshots if item.kind == "transfer"]
    unknown = [item for item in screenshots if item.kind == "unknown"]
    for item in unknown:
        issues.append(f"无法判断截图类型：{item.path.name}")

    product_by_amount = Counter(item.amount for item in products if item.amount is not None)
    payment_by_amount = Counter(item.amount for item in payments if item.amount is not None)
    transfer_by_amount = Counter(item.amount for item in transfers if item.amount is not None)
    used_products: set[Path] = set()
    used_payments: set[Path] = set()
    used_transfers: set[Path] = set()
    used_invoices: set[Path] = set()
    intentionally_unmatched: set[Path] = set()
    duplicate_notes: list[str] = []
    duplicate_paths: set[Path] = set()
    invoices_to_pair: list[Evidence] = []
    invoice_groups: dict[tuple[object, ...], list[Evidence]] = {}
    for invoice in invoices:
        invoice_groups.setdefault(evidence_identity(invoice), []).append(invoice)
    for group in invoice_groups.values():
        if len(group) == 1:
            invoices_to_pair.append(group[0])
            continue
        selected, detail = randomly_select_identical(group, "发票")
        intentionally_unmatched.update(item.path for item in group if item.path != selected.path)
        duplicate_paths.update(item.path for item in group)
        invoices_to_pair.append(selected)
        duplicate_notes.append(detail)
    completed = 0
    for index, invoice in enumerate(invoices_to_pair, start=1):
        logger.set_current(f"正在配对并生成资料包 {index}/{len(invoices_to_pair)}：{invoice.path.name}")
        missing = []
        if invoice.amount is None:
            missing.append("发票金额")
        product, product_match, duplicate_products = unique_match(invoice, [item for item in products if item.path not in used_products])
        payment, payment_match, duplicate_payments = unique_match(invoice, [item for item in payments if item.path not in used_payments])
        transfer, transfer_match, duplicate_transfers = unique_match(invoice, [item for item in transfers if item.path not in used_transfers])
        if invoice.amount is not None and product_by_amount[invoice.amount] > 1 and not invoice.order_ids and not duplicate_products:
            product = None
            product_match = "发票无订单号，多个同金额商品截图"
            duplicate_products = set()
        if invoice.amount is not None and payment_by_amount[invoice.amount] > 1 and not invoice.order_ids and not duplicate_payments:
            payment = None
            payment_match = "发票无订单号，多个同金额支付截图"
            duplicate_payments = set()
        if invoice.amount is not None and transfer_by_amount[invoice.amount] > 1 and not invoice.order_ids and not duplicate_transfers:
            transfer = None
            transfer_match = "发票无订单号，多个同金额转账截图"
            duplicate_transfers = set()
        images = [item.path for item in (product, payment) if item is not None]
        if not images and transfer:
            images = [transfer.path]
        if not images:
            missing.append("至少一张商品、支付或转账截图")
        product_name, name_source = select_product_name(invoice, product, payment, transfer if not product and not payment else None)
        if not product_name:
            missing.append("发票项目名称")
        created_date = (
            invoice.invoice_date
            or
            (product.created_date if product else None)
            or (payment.created_date if payment else None)
            or (transfer.created_date if transfer else None)
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
            intentionally_unmatched.update(duplicate_products)
            if duplicate_products:
                duplicate_paths.update(duplicate_products | {product.path})
                duplicate_notes.append(product_match)
        if payment:
            used_payments.add(payment.path)
            intentionally_unmatched.update(duplicate_payments)
            if duplicate_payments:
                duplicate_paths.update(duplicate_payments | {payment.path})
                duplicate_notes.append(payment_match)
        if transfer and not product and not payment:
            used_transfers.add(transfer.path)
            intentionally_unmatched.update(duplicate_transfers)
            if duplicate_transfers:
                duplicate_paths.update(duplicate_transfers | {transfer.path})
                duplicate_notes.append(transfer_match)
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
            "product_match": product_match,
            "payment_match": payment_match,
            "transfer_match": transfer_match,
            "matched_product_file": product.path.name if product else "",
            "matched_payment_file": payment.path.name if payment else "",
            "matched_transfer_file": transfer.path.name if transfer else "",
            "evidence_mode": (
                "商品+支付截图" if product and payment else
                "仅商品订单截图" if product else
                "仅支付截图" if payment else
                "仅转账截图" if transfer else
                ""
            ),
            "output": name,
        })
        evidence_mode = (
            "商品+支付截图" if product and payment else
            "仅商品订单截图" if product else
            "仅支付截图" if payment else
            "仅转账截图" if transfer else
            ""
        )
        logger.info(
            f"完成 {completed}/{len(invoices_to_pair)}：{name}（名称来自{name_source}；"
            f"商品匹配={product_match}；支付匹配={payment_match}；证据={evidence_mode}）。"
        )

    for item in products:
        if item.path not in used_products and item.path not in intentionally_unmatched:
            issues.append(f"未配对商品截图：{item.path.name}")
    for item in payments:
        if item.path not in used_payments and item.path not in intentionally_unmatched:
            issues.append(f"未配对支付截图：{item.path.name}")
    for item in transfers:
        if item.path not in used_transfers and item.path not in intentionally_unmatched:
            issues.append(f"未配对转账截图：{item.path.name}")

    unmatched_paths = [item.path for item in invoices if item.path not in used_invoices and item.path not in intentionally_unmatched]
    unmatched_paths.extend(
        item.path for item in screenshots
        if item.path not in used_products
        and item.path not in used_payments
        and item.path not in used_transfers
        and item.path not in intentionally_unmatched
    )
    copy_review_materials(output, UNMATCHED_NAME, unmatched_paths, logger)
    duplicate_copied = copy_review_materials(output, DUPLICATES_NAME, duplicate_paths, logger)

    for detail in duplicate_notes:
        logs.append({
            "type": "random_duplicate_selection",
            "detail": detail,
        })
    logs.append({
        "type": "duplicate_summary",
        "groups": str(len(duplicate_notes)),
        "files": str(len(duplicate_paths)),
        "copied": str(duplicate_copied),
        "folder": DUPLICATES_NAME,
    })

    (output / "识别日志.json").write_text(json.dumps(logs, ensure_ascii=False, indent=2), encoding="utf-8")
    if issues:
        (output / "待确认.txt").write_text("\n".join(issues) + "\n", encoding="utf-8-sig")
    if cache.exists():
        shutil.rmtree(cache)
    logger.clear_current()
    logger.info(f"整理完成：共 {completed} 组，待确认 {len(issues)} 项。")
    logger.info(f"输出目录：{output}")
    logger.info(f"运行日志：{logger.log_path}")
    logger.info(f"重复文件汇总：共 {len(duplicate_notes)} 组、{len(duplicate_paths)} 个文件，已复制 {duplicate_copied} 个文件到：{DUPLICATES_NAME}")
    for detail in duplicate_notes:
        logger.info(f"重复材料处理：{detail}")
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

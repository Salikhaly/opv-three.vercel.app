# -*- coding: utf-8 -*-
"""
parser_core.py — обёртка над main.py для веб-сервиса.
Принимает байты PDF, возвращает (превью-словарь, байты Excel).
Ничего не хранит: все временные файлы удаляются сразу.
"""

import os
import tempfile
from pathlib import Path
from typing import Optional, Dict, Any, List

import main as parser

ASSETS = Path(__file__).parent / "assets"
TEMPLATE = ASSETS / "template.xlsx"


def _contract_to_dict(c) -> Dict[str, Any]:
    return {
        "creditor": c.creditor,
        "financing_type": c.financing_type,
        "contract_amount": c.contract_amount,
        "periodic_payment": c.periodic_payment,
        "outstanding": c.outstanding,
        "current_overdue_days": c.current_overdue_days,
        "current_overdue_amount": c.current_overdue_amount,
        "max_overdue_days": c.max_overdue_days,
        "max_overdue_amount": c.max_overdue_amount,
        "end_date": c.end_date,
        "status": c.status,
        "cessionary": c.cessionary,
    }


def _write_temp(data: bytes, suffix: str) -> str:
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(data)
    return path


def parse_report(credit_pdf: Optional[bytes] = None,
                 pension_pdf: Optional[bytes] = None,
                 enable_ocr: bool = False,
                 pension_orgs: Optional[str] = None,
                 credit_indices: Optional[str] = None,
                 include_excel: bool = True,
                 print_layout: bool = False) -> Dict[str, Any]:
    """
    Возвращает: {
      ok, source, iin, fio, pkr, counts, monthly_load,
      active[], completed_recent[], completed_old[],
      excel_b64  # готовый файл в base64
    }
    """
    import base64

    if not credit_pdf and not pension_pdf:
        raise ValueError("Загрузите кредитный отчёт или пенсионную выписку.")
    tmp_credit = _write_temp(credit_pdf, ".pdf") if credit_pdf else None
    tmp_pension = _write_temp(pension_pdf, ".pdf") if pension_pdf else None
    tmp_out = _write_temp(b"", ".xlsx")
    os.unlink(tmp_out)  # нужен только путь; write создаст заново

    try:
        review_warnings = []
        cd = {}
        if tmp_credit:
            try:
                cd = parser.extract_any_credit_report(
                    pdf_path=Path(tmp_credit), enable_ocr=enable_ocr, ocr_lang="rus",
                    poppler_path=None, tesseract_cmd=None, tessdata_dir=None,
                    plumber_layout=False,
                )
            except Exception as exc:
                raise ValueError("Кредитный отчёт: PDF повреждён, защищён паролем или не содержит читаемого текста. Проверьте файл.") from exc
            parser.validate_credit_data(cd, strict=False)
    
            if not cd.get("iin") and not cd.get("pkr") and not cd.get("active"):
                return {"ok": False, "error": "Кредитный отчёт: не удалось прочитать данные. "
                        "Проверьте, что это отчёт ПКБ (1cb.kz) или ГКБ/МКБ (mkb.kz) и что PDF текстовый."}
        pension = None
        srzp = 0
        warnings = []
        pension_organizations = []
        if tmp_pension:
            try:
                pension = parser.parse_pension_pdf_tables(Path(tmp_pension))
                parser.validate_pension_data(pension, strict=False)
            except Exception:
                return {"ok": False, "error": "Не удалось прочитать пенсионную выписку. Проверьте файл PDF."}
            if not pension.get("rows"):
                return {"ok": False, "error": "В пенсионной выписке не найдены платежи."}
            from collections import Counter
            identities = Counter((x.date, x.period, x.sender_bin or x.sender, x.amount, x.status) for x in pension["rows"])
            for key, count in identities.items():
                if count > 1:
                    review_warnings.append(f"Пенсионная выписка: {count} одинаковых поступления от {key[0]}, период {key[1] or 'не указан'}, сумма {key[3]}. Сравните с оригиналом; суммы учтены как в файле.")
            for x in pension["rows"]:
                if x.amount is None or x.amount < 0:
                    review_warnings.append(f"Пенсионная выписка: проверьте сумму поступления от {x.date}.")
            pension_organizations = parser.list_pension_orgs(pension["rows"])
            pension["rows"] = [] if pension_orgs == "none" else parser.select_pension_orgs(pension["rows"], pension_orgs)
            srzp = parser.calc_total_srzp(pension["rows"])
            selected_keys = set(parser.group_pension_by_org(pension["rows"]))
            for key, org_rows in parser.group_pension_by_org(pension["rows"]).items():
                if len(org_rows) > 12:
                    warnings.append(f"Организация {key}: в Анкете показаны первые 12 платежей; СРЗП рассчитана по всем {len(org_rows)}.")
            for org in pension_organizations:
                org["selected"] = org["key"] in selected_keys

        active = cd.get("active") or []
        if credit_indices == "none":
            active = []
        elif credit_indices:
            try:
                indices = {int(x.strip()) for x in credit_indices.split(",")}
            except ValueError:
                raise ValueError("Некорректный выбор кредитов.")
            if any(i < 0 or i >= len(active) for i in indices):
                raise ValueError("Выбранный кредит не найден в отчёте.")
            active = [c for i, c in enumerate(active) if i in indices]
        recent = cd.get("completed_recent") or []
        old = cd.get("completed_old") or []
        revoked = cd.get("revoked") or []

        excel_bytes = None
        if include_excel:
            if print_layout:
                from print_export import print_model, write_printable, restore_original_questionnaire, restore_original_credits
                model = print_model(cd, pension, active, recent, old, revoked, srzp, bool(credit_pdf), bool(pension_pdf))
                write_printable(tmp_out, model)
                restore_original_questionnaire(tmp_out,TEMPLATE,pension,model["fio"],srzp)
                restore_original_credits(tmp_out,TEMPLATE,cd,active,recent,old,revoked,model["fio"])
            else:
                parser.write_output_legacy_excel(
                    template_path=TEMPLATE, output_path=Path(tmp_out),
                    pkr=cd.get("pkr"), active=active, recent=recent, old=old,
                    revoked=revoked, blank_zero=True, only_loans_active=False,
                    srzp=srzp,
                )
                if pension:
                    parser.fill_anketa_from_pension(Path(tmp_out), pension, sort_rows=False)
        

            if print_layout:
                from openpyxl import load_workbook
                workbook = load_workbook(tmp_out)
                for sheet in workbook:
                    sheet.oddHeader.left.text = "BRE Calc · BEREKELY EREKE"
                    sheet.oddHeader.left.size = 8
                    sheet.oddHeader.left.color = "7623B6"
                workbook.save(tmp_out)
            with open(tmp_out, "rb") as f:
                excel_bytes = f.read()
        monthly = sum((x.periodic_payment or 0) for x in active)

        for label, count, limit in (("Действующие договоры", len(active), 18),
                                    ("Завершённые и отозванные", len(recent) + len(revoked), 10),
                                    ("Договоры старше 5 лет", len(old), 10)):
            if count > limit:
                warnings.append(f"{label}: {count}; в листе Кредиты показаны первые {limit}.")
        return {
            "ok": True,
            "warnings": warnings,
            "review_warnings": review_warnings,
            "print_html": __import__("print_preview").workbook_html(excel_bytes, pension.get("fio") if pension else cd.get("fio") or "Анкета клиента") if excel_bytes is not None and print_layout else None,
            "pension_organizations": pension_organizations,
            "has_credit": bool(credit_pdf),
            "has_pension": bool(pension_pdf),
            "source": cd.get("source") or "pension",
            "iin": cd.get("iin") or (pension.get("iin") if pension else None),
            "fio": (pension.get("fio") if pension else None) or cd.get("fio"),
            "pkr": cd.get("pkr"),
            "generated_at": cd.get("generated_at"),
            "srzp": srzp,
            "monthly_load": monthly,
            "counts": {
                "active": len(active),
                "completed_recent": len(recent),
                "completed_old": len(old),
            },
            "active": [_contract_to_dict(c) for c in active],
            "completed_recent": [_contract_to_dict(c) for c in recent],
            "completed_old": [_contract_to_dict(c) for c in old],
            "excel_filename": __import__('print_export').client_filename((pension.get('fio') if pension else None) or cd.get('fio')),
            "excel_b64": base64.b64encode(excel_bytes).decode("ascii") if excel_bytes is not None else None,
        }
    finally:
        for p in (tmp_credit, tmp_pension, tmp_out):
            if p and os.path.exists(p):
                try:
                    os.unlink(p)
                except OSError:
                    pass

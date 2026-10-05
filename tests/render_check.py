"""Render sample documents through the installed Office and report audits.

Not a pytest test: it needs Microsoft Word/PowerPoint (or LibreOffice) and
writes PNGs to tests/out/ for a person to look at. Run before a release:

    python tests/render_check.py
"""
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "tests" / "out"
for skill in ("arabic-docx", "arabic-pptx"):
    sys.path.insert(0, str(ROOT / skill / "scripts"))

import arabic_docx as D   # noqa: E402
import arabic_pptx as P   # noqa: E402

MIXED = ("نص عربي مع PostgreSQL و FastAPI ولغات C# و C++ و .NET "
         "والهاتف +963 912 345 ونسبة 50% و DECIMAL(5,2).")


def docx_sample():
    doc = D.ArabicDocx()
    doc.title("تقرير تجريبي")
    doc.toc()
    doc.heading("المقدمة", page_break=True)
    doc.para(MIXED)
    doc.para("فقرة عريضة مع Nuxt 3.", bold=True)
    doc.heading("القوائم", level=2)
    doc.bullets(["نقطة أولى طويلة بما يكفي لتلتف إلى سطر ثانٍ داخل الصفحة حتى "
                 "نرى أين يبدأ السطر الثاني من القائمة", "نقطة ثانية مع Docker"])
    doc.numbered(["خطوة أولى", "خطوة ثانية"])
    doc.heading("الجداول", page_break=True)
    doc.table(["البند", "التفصيل", "ملاحظات"],
              [["قاعدة البيانات", "PostgreSQL 16", "أساسية"],
               ["الواجهة", "Nuxt 3", "DECIMAL(5,2)"]],
              [4, 5, 5], caption="جدول بعرض محدد", header_fill="D9E2F3")
    doc.para("فقرة بعد الجدول.")
    doc.page_numbers(skip_first=True, start=0)
    return doc.save(OUT / "sample_doc.docx")


def pptx_sample():
    deck = P.ArabicPptx(base_size=20, transition="fade")
    deck.title_slide("عنوان العرض", "عرض تجريبي لمكتبة arabic-pptx")
    deck.bullets_slide("المحاور", [MIXED, "أرقام الإصدارات مثل Nuxt 3", "نقطة ثالثة"],
                       reveal=True)
    deck.steps_slide("المراحل", [("جمع البيانات", "المصادر"), ("التحليل", "القواعد"),
                                 ("التنبيه", "الطبيب"), ("المتابعة", "التقارير")],
                     reveal=True)
    deck.table_slide("التقنيات", ["الطبقة", "التقنية"],
                     [["الواجهة", "Nuxt 3"], ["الخلفية", "FastAPI"]], [10, 12])
    return deck.save(OUT / "sample_deck.pptx")


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    d, p = docx_sample(), pptx_sample()
    print("docx audit:", D.audit_docx(d) or "clean")
    print("pptx audit:", P.audit_pptx(p) or "clean")
    print("docx pages:", D.preview(d, OUT))
    print("pptx slides:", P.preview(p, OUT))

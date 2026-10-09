# -*- coding: utf-8 -*-
"""
AlArab Order AI - النسخة الويب (متوافقة مع الآيفون، الأندرويد، وجميع المتصفحات)
- التحقق التلقائي من أرقام الهواتف ومفتاح العراق (+964 / 0).
- كشف الطلبات المكررة وتنبيه المستخدم.
- إمكانية حذف أو تعديل الطلبات مباشرة.
- التصدير المباشر لقالب الشركة بصيغة Excel.
"""

import io
import json
import re
import time
from pathlib import Path
import pandas as pd
from PIL import Image
import requests
import streamlit as st
from openpyxl import Workbook

# إعدادات صفحة الويب لتكون متجاوبة ونظيفة للهواتف
st.set_page_config(
    page_title="العراب - نظام التوصيل الذكي والمتقدم",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

# قائمة النماذج المتاحة وتضمين فلاش 3
PREFERRED_MODELS = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.7-flash",
    "gemini-3.8-flash",
    "gemini-2.5-flash",
    "gemini-2.0-flash",
    "gemini-1.5-flash",
    "gemini-1.5-pro",
]

DEFAULT_MODEL = "gemini-3.5-flash"

FIELDS = [
    "اسم الزبون", "رقم الهاتف الاساسي", "رقم الهاتف الثانوي", "المحافظة",
    "المنطقة", "نوع البضاعه", "عدد القطع", "السعر مع التوصيل", "حجم الطلب",
    "الملاحظات", "نوع الطلب"
]

SYSTEM_PROMPT = r"""
أنت نظام استخراج طلبات لمتجر عراقي اسمه "العراب".
حلّل صورة/لقطة شاشة طلب الزبون بالكامل، وليس جزءاً واحداً فقط.
المطلوب إخراج JSON فقط، بدون Markdown وبدون شرح.

القواعد:
1) اسم الزبون: خذه من اسم الحساب الظاهر أعلى المحادثة إذا كان واضحاً. إذا لم يوجد اكتب "غير متوفر".
2) رقم الهاتف الاساسي: استخرج رقم الهاتف الذي يظهر للزبون. إذا كتبه بدون مفتاح ابدأ بـ 07، وإذا كتبه مع مفتاح العراق (+964 أو 964) حوله إلى الصيغة المحلية المبدوءة بـ 07 ليكون متناسقاً.
3) رقم الهاتف الثانوي: إذا وُجد رقم آخر واضح اكتبه، وإلا اتركه فارغاً.
4) المحافظة: استنتجها بدقة من العنوان (مثل: بغداد، البصرة، ذي قار، بابل، نينوى، واسط، كربلاء، الانبار، ديالى، اربيل، كركوك، السليمانية، دهوك، صلاح الدين، القادسية، النجف، المثنى، ميسان).
5) المنطقة: استخرج اسم المنطقة/القضاء/الناحية بدقة (مثل: الأعظمية، الكاظمية، المنصور، الكرادة...).
6) أقرب نقطة دالة أو ملاحظات: ادمج أي تفاصيل إضافية أو نقاط دالة أو ملاحظات في حقل الملاحظات.
7) نوع البضاعه: اكتب اسم المنتج المطلوب كما يفهم من المحادثة.
8) عدد القطع: عدد القطع المطلوب فعلياً (رقم فقط). إذا لم يذكر فاعتبره 1.
9) السعر مع التوصيل: استخرج السعر النهائي (شامل التوصيل) كمبلغ صافي بدون فواصل وبدون كلمة دينار (مثل 60000 أو 25000). إذا كتب الزبون 60 أو 60 الف فالمقصود 60000.
10) حجم الطلب: اجعله دائماً "عادي".
11) نوع الطلب: "طلب جديد" أو "استبدال".
12) أخرج JSON بهذا الشكل بالضبط:
{
 "اسم الزبون":"",
 "رقم الهاتف الاساسي":"",
 "رقم الهاتف الثانوي":"",
 "المحافظة":"",
 "المنطقة":"",
 "نوع البضاعه":"",
 "عدد القطع":"",
 "السعر مع التوصيل":"",
 "حجم الطلب":"عادي",
 "الملاحظات":"",
 "نوع الطلب":"طلب جديد",
 "_ثقة":"عالية"
}
"""

def img_to_b64(img_file):
    img = Image.open(img_file).convert("RGB")
    max_side = 1600
    if max(img.size) > max_side:
        scale = max_side / max(img.size)
        img = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=85)
    import base64
    return base64.b64encode(buf.getvalue()).decode("ascii")

def normalize_num(s):
    if not s:
        return ""
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    return str(s).translate(trans)

def clean_phone_number(ph):
    if not ph:
        return "", "رقم مفقود"
    ph = normalize_num("".join(c for c in str(ph) if c.isdigit() or c == "+"))
    
    if ph.startswith("+964"):
        ph = "0" + ph[4:]
    elif ph.startswith("964") and len(ph) >= 12:
        ph = "0" + ph[3:]
    elif ph.startswith("7") and len(ph) == 10:
        ph = "0" + ph
        
    if ph.startswith("07") and len(ph) == 11:
        return ph, "سليم"
    elif len(ph) > 5:
        return ph, "⚠ رقم غير معتاد"
    else:
        return ph, "⚠ رقم قصير جداً"

def clean_price_format(val):
    if not val:
        return ""
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    s = str(val).translate(trans)
    digits_only = "".join(c for c in s if c.isdigit())
    if not digits_only:
        return s
    num = int(digits_only)
    if 0 < num < 100:
        num = num * 1000
    return str(num)

def clean_json_text(txt):
    txt = txt.strip()
    # تم تصحيح هذا السطر بالكامل لمنع أي أخطاء برمجية
    txt = re.sub(r"^

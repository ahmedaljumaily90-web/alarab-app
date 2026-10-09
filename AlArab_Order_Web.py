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

DEFAULT_MODEL = "gemini-2.5-flash"  # نموذج سريع ومستقر للويب

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
    txt = re.sub(r"^```(?:json)?\s*", "", txt, flags=re.I)
    txt = re.sub(r"\s*```$", "", txt)
    a, b = txt.find("{"), txt.rfind("}")
    if a >= 0 and b > a:
        txt = txt[a:b+1]
    return txt

def parse_response(text):
    data = json.loads(clean_json_text(text))
    out = {f: str(data.get(f, "") or "") for f in FIELDS}
    out["اسم الزبون"] = out["اسم الزبون"] or "غير متوفر"
    out["حجم الطلب"] = "عادي"
    out["نوع الطلب"] = out["نوع الطلب"] or "طلب جديد"
    out["_ثقة"] = str(data.get("_ثقة", "") or "")
    
    raw_phone = data.get("رقم الهاتف الاساسي", "")
    cleaned_ph, phone_status = clean_phone_number(raw_phone)
    out["رقم الهاتف الاساسي"] = cleaned_ph
    out["_phone_status"] = phone_status
    
    if data.get("رقم الهاتف الثانوي", ""):
        out["رقم الهاتف الثانوي"], _ = clean_phone_number(data.get("رقم الهاتف الثانوي", ""))
        
    out["رقم الهاتف الثانوي"] = normalize_num(out["رقم الهاتف الثانوي"])
    out["عدد القطع"] = normalize_num(out["عدد القطع"])
    out["السعر مع التوصيل"] = clean_price_format(data.get("السعر مع التوصيل", ""))
    return out

def call_gemini(api_key, model, img_file):
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    b64 = img_to_b64(img_file)
    payload = {
        "contents": [{
            "parts": [
                {"text": SYSTEM_PROMPT},
                {"inline_data": {"mime_type": "image/jpeg", "data": b64}}
            ]
        }],
        "generationConfig": {
            "temperature": 0,
            "responseMimeType": "application/json"
        }
    }
    headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
    r = requests.post(endpoint, headers=headers, json=payload, timeout=60)
    if r.status_code != 200:
        raise RuntimeError(f"Gemini API Error {r.status_code}: {r.text[:300]}...")
    data = r.json()
    try:
        text = data["candidates"][0]["content"]["parts"][0]["text"]
    except Exception:
        raise RuntimeError("فشل قراءة الرد من نموذج الذكاء الاصطناعي.")
    return parse_response(text)

def create_excel(df_data):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(FIELDS)
    for row in df_data:
        ws.append([
            row.get("اسم الزبون", ""),
            row.get("رقم الهاتف الاساسي", ""),
            row.get("رقم الهاتف الثانوي", ""),
            row.get("المحافظة", ""),
            row.get("المنطقة", ""),
            row.get("نوع البضاعه", ""),
            row.get("عدد القطع", ""),
            row.get("السعر مع التوصيل", ""),
            "عادي",
            row.get("الملاحظات", ""),
            row.get("نوع الطلب", "طلب جديد")
        ])
    ws.freeze_panes = "A2"
    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    return out_buf

# تصميم واجهة الويب
st.markdown("<h1 style='text-align: center; color: #d4af37;'>📦 العراب - نظام التوصيل الذكي والمتقدم</h1>", unsafe_allow_html=True)
st.markdown("<p style='text-align: center; color: gray;'>استخراج الطلبات من الصور بدقة عالية للآيفون والأندرويد</p>", unsafe_allow_html=True)

# الشريط الجانبي للإعدادات مع تضمين المفتاح
with st.sidebar:
    st.header("⚙️ الإعدادات")
    DEFAULT_API_KEY = "AQ.Ab8RN6Kp3WxLw5nZAw2zd2wPb3Fn0Spq3A5OVUr1X_wNI6pRwg"
    api_key_input = st.text_input("مفتاح Gemini API", type="password", value=DEFAULT_API_KEY)
    model_choice = st.selectbox("اختر نموذج الذكاء الاصطناعي", ["gemini-2.5-flash", "gemini-1.5-flash", "gemini-1.5-pro"], index=0)
    st.markdown("---")
    st.info("💡 **تعليمات الاستخدام:**\n1. ارفع صور المحادثات أو لقطات الشاشة.\n2. اضغط على زر بدء الاستخراج.\n3. قم بالتعديل أو الحذف إن وجد، ثم حمّل ملف الأكسل.")

if "processed_rows" not in st.session_state:
    st.session_state.processed_rows = []

# رفع الصور (تم إصلاح السطر بالكامل هنا)
uploaded_files = st.file_uploader("📂 اختر أو ارفع صور الطلبات (يدعم صور متعددة)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)

col1, col2 = st.columns(2)
with col1:
    start_btn = st.button("🤖 ابدأ استخراج الطلبات", type="primary", use_container_width=True)
with col2:
    clear_btn = st.button("🗑 مسح النتائج الحالية", use_container_width=True)

if clear_btn:
    st.session_state.processed_rows = []
    st.rerun()

if start_btn:
    if not api_key_input.strip():
        st.error("⚠ يرجى إدخال مفتاح Gemini API في الشريط الجانبي أولاً.")
    elif not uploaded_files:
        st.warning("⚠ يرجى رفع صورة واحدة على الأقل.")
    else:
        st.session_state.processed_rows = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        total = len(uploaded_files)
        for i, file_obj in enumerate(uploaded_files):
            status_text.text(f"جارٍ معالجة الصورة {i+1} من {total}: {file_obj.name} ...")
            try:
                row = call_gemini(api_key_input.strip(), model_choice, file_obj)
                row["_filename"] = file_obj.name
                st.session_state.processed_rows.append(row)
            except Exception as e:
                err_row = {f: "" for f in FIELDS}
                err_row["اسم الزبون"] = "خطأ في القراءة"
                err_row["الملاحظات"] = str(e)[:150]
                err_row["_filename"] = file_obj.name
                err_row["_phone_status"] = "خطأ"
                st.session_state.processed_rows.append(err_row)
            
            progress_bar.progress((i + 1) / total)
            
        status_text.text("✅ اكتملت عملية استخراج جميع الطلبات بنجاح!")
        st.success("تم الانتهاء من المعالجة!")

# عرض جدول البيانات والتحكم بها إذا كانت موجودة
if st.session_state.processed_rows:
    st.markdown("---")
    st.subheader("📋 جدول الطلبات المستخرجة (قابل للتعديل)")
    
    df_display = pd.DataFrame(st.session_state.processed_rows)
    phone_counts = df_display["رقم الهاتف الاساسي"].value_counts().to_dict()
    
    status_list = []
    for idx, r in df_display.iterrows():
        ph = str(r.get("رقم الهاتف الاساسي", "")).strip()
        p_status = r.get("_phone_status", "سليم")
        if ph and phone_counts.get(ph, 0) > 1:
            status_list.append("⚠ مكرر بنفس الرقم")
        elif p_status != "سليم":
            status_list.append(p_status)
        else:
            status_list.append("✅ سليم")
            
    df_display["الحالة"] = status_list
    
    columns_to_show = FIELDS + ["الحالة", "_filename"]
    existing_cols = [c for c in columns_to_show if c in df_display.columns]
    
    edited_df = st.data_editor(
        df_display[existing_cols],
        num_rows="dynamic",
        use_container_width=True,
        key="order_editor"
    )
    
    updated_rows = []
    for index, row in edited_df.iterrows():
        row_dict = {f: str(row.get(f, "")) for f in FIELDS}
        row_dict["_filename"] = row.get("_filename", "")
        cleaned_ph, p_status = clean_phone_number(row_dict["رقم الهاتف الاساسي"])
        row_dict["رقم الهاتف الاساسي"] = cleaned_ph
        row_dict["_phone_status"] = p_status
        updated_rows.append(row_dict)
    
    st.session_state.processed_rows = updated_rows

    st.markdown("---")
    excel_buffer = create_excel(st.session_state.processed_rows)
    st.download_button(
        label="📊 تحميل ملف الأكسل بصيغة الشركة (Excel)",
        data=excel_buffer,
        file_name=f"طلبات_العراب_{time.strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )

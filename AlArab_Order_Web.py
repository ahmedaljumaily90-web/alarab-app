# -*- coding: utf-8 -*-
"""
Al-Arab Order AI (نسخة الويب والهاتف - Streamlit)
- التحقق التلقائي من أرقام الهواتف ومفتاح العراق (+964 / 0).
- كشف الطلبات المكررة وتنبيه المستخدم وتمييزها بصرياً.
- زر حذف الطلب المباشر وتصدير Excel.
- تثبيت ودعم مفتاح AQ الجديد تلقائياً لمنع خطأ 401.
"""

import streamlit as st
import json, re, io, base64
from pathlib import Path
from PIL import Image
import requests
from openpyxl import Workbook
import pandas as pd

APP_TITLE = "العراب - نظام التوصيل الذكي والمتقدم (الويب)"
# تثبيت مفتاحك الجديد بشكل دائم وصحيح في النظام
DEFAULT_API_KEY = "AQ.Ab8RN6I2N21DuvR1msayHwB3Z2VtBVth2rcAvWCypm4w4cpSqw"
PREFERRED_MODELS = [
    "gemini-3.8-flash", "gemini-3.7-flash", "gemini-3.6-flash",
    "gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-2.5-flash"
]

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

def img_to_b64_from_uploaded(uploaded_file):
    img = Image.open(uploaded_file).convert("RGB")
    max_side = 2200
    if max(img.size) > max_side:
        scale = max_side / max(img.size)
        img = img.resize((int(img.width*scale), int(img.height*scale)), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=88)
    return base64.b64encode(buf.getvalue()).decode("ascii")

def normalize_num(s):
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
        return ph, "⚠ رقم غير معتاد (تحقق)"
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

def clean_model_name(model):
    model = (model or "").strip()
    model = model.replace("models/", "").replace(" ", "")
    return model

def test_gemini_connection(api_key, model):
    model = clean_model_name(model)
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    payload = {"contents": [{"parts": [{"text": "Hello"}]}]}
    
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    r = requests.post(endpoint, headers=headers, json=payload, timeout=30)
    if r.status_code != 200:
        headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
        r = requests.post(endpoint, headers=headers, json=payload, timeout=30)
        
    if r.status_code == 200:
        return True, "تم الاتصال بنجاح بـ Gemini API والموديل يعمل بشكل ممتاز!"
    else:
        return False, f"فشل الاتصال (كود الخطأ {r.status_code}): {r.text[:200]}"

def call_gemini(api_key, model, uploaded_file):
    model = clean_model_name(model)
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    b64 = img_to_b64_from_uploaded(uploaded_file)
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
    
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    r = requests.post(endpoint, headers=headers, json=payload, timeout=120)
    if r.status_code != 200:
        headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
        r = requests.post(endpoint, headers=headers, json=payload, timeout=120)
        
    if r.status_code != 200:
        raise RuntimeError(f"Gemini API {r.status_code}: {r.text[:1000]}")
        
    data = r.json()
    try:
        text = data["candidates"][0]["content"]["parts"]["text"] if "text" in data["candidates"][0]["content"]["parts"] else data["candidates"][0]["content"]["parts"][0]["text"]
    except Exception:
        try:
            text = data["candidates"][0]["content"]["parts"][0]["text"]
        except Exception:
            raise RuntimeError("لم يصل نص JSON من Gemini: " + json.dumps(data, ensure_ascii=False)[:1000])
    return parse_response(text)

def create_excel_file(rows):
    output = io.BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(FIELDS)
    
    for row in rows:
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
    wb.save(output)
    output.seek(0)
    return output

# --- واجهة Streamlit للويب ---
st.set_page_config(page_title=APP_TITLE, layout="wide")

st.markdown("<h2 style='text-align: center; color: #ffd400;'>العراب - نظام التوصيل الذكي والمتقدم (الهاتف والويب)</h2>", unsafe_allow_html=True)

if "api_key" not in st.session_state or not st.session_state.api_key:
    st.session_state.api_key = DEFAULT_API_KEY

with st.sidebar:
    st.header("⚙ الإعدادات والذكاء الاصطناعي")
    
    user_api_key = st.text_input("مفتاح Gemini API", value=st.session_state.api_key, type="password")
    if user_api_key:
        st.session_state.api_key = user_api_key
        
    selected_model = st.selectbox("اختر النموذج", PREFERRED_MODELS, index=4)
    
    if st.button("🔌 اختبار الاتصال بالذكاء الاصطناعي"):
        if not st.session_state.api_key:
            st.error("الرجاء إدخال المفتاح أولاً.")
        else:
            with st.spinner("جاري اختبار الاتصال..."):
                success, msg = test_gemini_connection(st.session_state.api_key, selected_model)
                if success:
                    st.success(msg)
                else:
                    st.error(msg)
                    
    st.info("تم تثبيت المفتاح الجديد وتفعيل كشف التكرار وحذف الطلبات تلقائياً.")

uploaded_files = st.file_uploader("📂 اختر صور الطلبات (يمكن اختيار عدة صور)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)

if "extracted_rows" not in st.session_state:
    st.session_state.extracted_rows = []

col_btn1, col_btn2 = st.columns([3, 1])
with col_btn1:
    start_clicked = st.button("🤖 ابدأ استخراج الطلبات", type="primary", use_container_width=True)
with col_btn2:
    clear_clicked = st.button("🗑 حذف الكل", use_container_width=True)

if clear_clicked:
    st.session_state.extracted_rows = []
    st.rerun()

if start_clicked:
    if not st.session_state.api_key:
        st.error("الرجاء إدخال مفتاح Gemini API في الشريط الجانبي.")
    elif not uploaded_files:
        st.warning("الرجاء رفع صورة واحدة على الأقل.")
    else:
        rows = []
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        for idx, file in enumerate(uploaded_files):
            status_text.text(f"جارٍ معالجة الصورة {idx+1} من {len(uploaded_files)}: {file.name}")
            try:
                res = call_gemini(st.session_state.api_key, selected_model, file)
                res["_filename"] = file.name
                rows.append(res)
            except Exception as e:
                err_row = {f: "" for f in FIELDS}
                err_row["اسم الزبون"] = "خطأ في القراءة"
                err_row["الملاحظات"] = str(e)[:200]
                err_row["_filename"] = file.name
                err_row["_phone_status"] = "خطأ"
                rows.append(err_row)
            progress_bar.progress((idx + 1) / len(uploaded_files))
            
        st.session_state.extracted_rows = rows
        status_text.text("اكتمل استخراج جميع الطلبات بنجاح!")
        st.success("تمت المعالجة بنجاح!")

if st.session_state.extracted_rows:
    st.subheader("📊 جدول الطلبات المستخرجة (مع كشف التكرار)")
    
    phone_counts = {}
    for r in st.session_state.extracted_rows:
        ph = (r.get("رقم الهاتف الاساسي", "")).strip()
        if ph:
            phone_counts[ph] = phone_counts.get(ph, 0) + 1

    indices_to_delete = []
    
    for idx, r in enumerate(list(st.session_state.extracted_rows)):
        ph = (r.get("رقم الهاتف الاساسي", "")).strip()
        status = r.get("_phone_status", "سليم")
        is_duplicate = False
        
        if ph and phone_counts.get(ph, 0) > 1:
            status = "⚠ طلب مكرر بنفس الرقم!"
            is_duplicate = True
            
        with st.container(border=True):
            cols = st.columns([2, 2, 2, 2, 1])
            with cols[0]:
                st.markdown(f"**الزبون:** {r.get('اسم الزبون', '')}")
                st.markdown(f"**الهاتف:** `{r.get('رقم الهاتف الاساسي', '')}`")
            with cols[1]:
                st.markdown(f"**المحافظة:** {r.get('المحافظة', '')}")
                st.markdown(f"**المنطقة:** {r.get('المنطقة', '')}")
            with cols[2]:
                st.markdown(f"**المنتج:** {r.get('نوع البضاعه', '')}")
                st.markdown(f"**السعر:** {r.get('السعر مع التوصيل', '')}")
            with cols[3]:
                if is_duplicate:
                    st.markdown(f"<span style='color: red; font-weight: bold;'>{status}</span>", unsafe_allow_html=True)
                else:
                    st.markdown(f"<span style='color: green;'>{status}</span>", unsafe_allow_html=True)
                st.markdown(f"*(ملف: {r.get('_filename', '')})*")
            with cols[4]:
                if st.button("🗑 حذف", key=f"del_{idx}"):
                    indices_to_delete.append(idx)

    if indices_to_delete:
        for i in sorted(indices_to_delete, reverse=True):
            del st.session_state.extracted_rows[i]
        st.rerun()

    st.divider()
    
    excel_io = create_excel_file(st.session_state.extracted_rows)
    st.download_button(
        label="📥 تحميل ملف الأكسل بقالب الشركة (Excel)",
        data=excel_io,
        file_name="طلبات_العراب_قالب_الشركة.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

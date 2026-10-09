# -*- coding: utf-8 -*-
"""
Al-Arab Order Web App - النسخة المحصنة النهائية بدون أخطاء
"""
import streamlit as st
import pandas as pd
import json, os, re, io, base64
from pathlib import Path
import requests
from openpyxl import Workbook

st.set_page_config(
    page_title="العراب - نظام التوصيل الذكي",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <div style='background-color: #ffd400; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1);'>
        <h1 style='color: #000; margin:0; font-size: 24px;'>العراب - نظام التوصيل الذكي (نسخة الهاتف)</h1>
        <p style='color: #333; margin:5px 0 0 0; font-size: 14px;'>حماية قصوى ضد تكرار نظام شركة التوصيل، استخراج ذكي، وتعديل فوري</p>
    </div>
    <br>
""", unsafe_allow_html=True)

default_api_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        default_api_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if not default_api_key:
    default_api_key = os.environ.get("GEMINI_API_KEY", "")

with st.sidebar:
    st.header("⚙ إعدادات النظام")
    api_key_input = st.text_input("مفتاح Gemini API:", type="password", value=default_api_key)
    model_choice = st.selectbox("اختر نموذج الذكاء الاصطناعي:", ["gemini-3.5-flash-lite", "gemini-2.5-flash", "gemini-3.7-flash"], index=0)
    if default_api_key:
        st.success("✔ تم تحميل المفتاح المحفوظ تلقائياً.")
    else:
        st.info("💡 أدخل المفتاح مرة واحدة في إعدادات Secrets على Streamlit ليبقى محفوظاً.")

FIELDS = [
    "اسم الزبون", "رقم الهاتف الاساسي", "رقم الهاتف الثانوي", "المحافظة",
    "المنطقة", "نوع البضاعه", "عدد القطع", "السعر مع التوصيل", "حجم الطلب",
    "الملاحظات", "نوع الطلب", "_phone_status", "_filename"
]

SYSTEM_PROMPT = r"""
أنت نظام استخراج طلبات لمتجر عراقي اسمه "العراب".
حلّل صورة/لقطة شاشة طلب الزبون بالكامل، سواء كانت من فيسبوك أو واتساب.
المطلوب إخراج JSON فقط، بدون Markdown وبدون شرح.

القواعد:
1) اسم الزبون: خذه من اسم الحساب الظاهر أعلى المحادثة أو من بيانات واتساب. إذا كانت محادثة واتساب ولا يوجد اسم صريح واكتفى برقم هاتف أو جزء منه، اكتب "واتساب: [الرقم أو الأجزاء الظاهرة مثل +964...]" بدلاً من غير متوفر.
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

def normalize_num(s):
    trans = str.maketrans("٠١ي٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    return str(s).translate(trans)

def clean_phone_number(ph):
    if not ph:
        return "", "⚠ رقم مفقود"
    ph_digits = "".join(c for c in str(ph) if c.isdigit())
    if len(ph_digits) >= 10:
        last_10 = ph_digits[-10:]
        ph = "0" + last_10
    
    if ph.startswith("07") and len(ph) == 11: 
        return ph, "سليم"
    elif len(ph) > 5: 
        return ph, "⚠ رقم غير معتاد"
    else: 
        return ph, "⚠ رقم قصير"

def clean_price_format(val):
    if not val: return ""
    trans = str.maketrans("٠١ي٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    s = str(val).translate(trans)
    digits_only = "".join(c for c in s if c.isdigit())
    if not digits_only: return s
    num = int(digits_only)
    if 0 < num < 100: num = num * 1000
    return str(num)

def call_gemini_api(api_key, model, img_bytes, mime_type, filename):
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    b64 = base64.b64encode(img_bytes).decode("ascii")
    
    payload = {
        "contents": [{
            "parts": [
                {"text": SYSTEM_PROMPT},
                {"inline_data": {"mime_type": mime_type, "data": b64}}
            ]
        }],
        "generationConfig": {"temperature": 0, "responseMimeType": "application/json"}
    }
    headers = {"x-goog-api-key": api_key, "Content-Type": "application/json"}
    r = requests.post(endpoint, headers=headers, json=payload, timeout=120)
    if r.status_code != 200:
        raise RuntimeError(f"API Error {r.status_code}: {r.text[:300]}")
    
    res_json = r.json()
    text = res_json["candidates"][0]["content"]["parts"][0]["text"]
    
    text = text.strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.I)
    text = re.sub(r"\s*```$", "", text)
    data = json.loads(text)
    
    out = {f: str(data.get(f, "") or "") for f in FIELDS if f not in ["_filename", "_phone_status"]}
    
    raw_name = str(data.get("اسم الزبون", "")).strip()
    raw_ph = data.get("رقم الهاتف الاساسي", "")
    cleaned_ph, p_status = clean_phone_number(raw_ph)
    
    if not raw_name or raw_name.lower() in ["غير متوفر", "unknown", "none", ""]:
        if cleaned_ph and len(cleaned_ph) >= 6:
            out["اسم الزبون"] = f"واتساب: +964 {cleaned_ph[-6:]}"
        else:
            out["اسم الزبون"] = "واتساب: +964..."
    else:
        out["اسم الزبون"] = raw_name

    out["حجم الطلب"] = "عادي"
    out["نوع الطلب"] = out["نوع الطلب"] or "طلب جديد"
    out["رقم الهاتف الاساسي"] = cleaned_ph
    out["_phone_status"] = p_status
    out["رقم الهاتف الثانوي"] = normalize_num(data.get("رقم الهاتف الثانوي", ""))
    out["عدد القطع"] = normalize_num(data.get("عدد القطع", ""))
    out["السعر مع التوصيل"] = clean_price_format(data.get("السعر مع التوصيل", ""))
    out["_filename"] = filename
    return out

if "image_cache" not in st.session_state:
    st.session_state["image_cache"] = {}

uploaded_files = st.file_uploader("📂 اختر أو التقط صور الطلبات (دفعة واحدة)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True, key="bulk_upload")

if uploaded_files:
    if st.button("🤖 ابدأ استخراج الطلبات للصور المرفوعة", type="primary"):
        if not api_key_input:
            st.error("الرجاء إدخال مفتاح Gemini API في القائمة الجانبية.")
        else:
            extracted_rows = st.session_state.get("extracted_rows", [])
            existing_files = {r.get("_filename") for r in extracted_rows}
            
            progress_bar = st.progress(0)
            status_text = st.empty()

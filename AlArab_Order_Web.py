# -*- coding: utf-8 -*-
"""
Al-Arab Order Web App (نسخة الهاتف واللابتوب عبر المتصفح)
تشغيل البرنامج عبر الويب باستخدام Streamlit
"""
import streamlit as st
import pandas as pd
import json, os, re, io, base64
from pathlib import Path
import requests
from openpyxl import Workbook

st.set_page_config(page_title="العراب - استخراج الطلبات", page_icon="📦", layout="wide")

st.markdown("""
    <div style='background-color: #ffd400; padding: 15px; border-radius: 10px; text-align: center;'>
        <h1 style='color: #000; margin:0;'>العراب - نظام التوصيل الذكي (نسخة الهاتف)</h1>
        <p style='color: #333; margin:5px 0 0 0;'>استخراج بيانات الزبائن من صور المحادثات وتصديرها بقالب شركة التوصيل</p>
    </div>
    <br>
""", unsafe_allow_html=True)

# إعدادات مفتاح API والنموذج
with st.sidebar:
    st.header("⚙ إعدادات الذكاء الاصطناعي")
    api_key_input = st.text_input("أدخل مفتاح Gemini API:", type="password", value=os.environ.get("GEMINI_API_KEY", ""))
    model_choice = st.selectbox("اختر النموذج:", ["gemini-3.5-flash-lite", "gemini-2.5-flash", "gemini-3.7-flash"], index=0)
    st.info("💡 المفتاح الخاص بك آمن ولا يتم مشاركته.")

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

def normalize_num(s):
    trans = str.maketrans("٠١ي٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    return str(s).translate(trans)

def clean_phone_number(ph):
    if not ph:
        return "", "رقم مفقود"
    ph = normalize_num("".join(c for c in str(ph) if c.isdigit() or c == "+"))
    if ph.startswith("+964"): ph = "0" + ph[4:]
    elif ph.startswith("964") and len(ph) >= 12: ph = "0" + ph[3:]
    elif ph.startswith("7") and len(ph) == 10: ph = "0" + ph
    
    if ph.startswith("07") and len(ph) == 11: return ph, "سليم"
    elif len(ph) > 5: return ph, "⚠ رقم غير معتاد"
    else: return ph, "⚠ رقم قصير"

def clean_price_format(val):
    if not val: return ""
    trans = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
    s = str(val).translate(trans)
    digits_only = "".join(c for c in s if c.isdigit())
    if not digits_only: return s
    num = int(digits_only)
    if 0 < num < 100: num = num * 1000
    return str(num)

def call_gemini_web(api_key, model, uploaded_file):
    endpoint = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
    bytes_data = uploaded_file.getvalue()
    b64 = base64.b64encode(bytes_data).decode("ascii")
    
    payload = {
        "contents": [{
            "parts": [
                {"text": SYSTEM_PROMPT},
                {"inline_data": {"mime_type": uploaded_file.type, "data": b64}}
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
    
    out = {f: str(data.get(f, "") or "") for f in FIELDS}
    out["اسم الزبون"] = out["اسم الزبون"] or "غير متوفر"
    out["حجم الطلب"] = "عادي"
    out["نوع الطلب"] = out["نوع الطلب"] or "طلب جديد"
    
    raw_ph = data.get("رقم الهاتف الاساسي", "")
    cleaned_ph, p_status = clean_phone_number(raw_ph)
    out["رقم الهاتف الاساسي"] = cleaned_ph
    out["_phone_status"] = p_status
    out["رقم الهاتف الثانوي"] = normalize_num(data.get("رقم الهاتف الثانوي", ""))
    out["عدد القطع"] = normalize_num(data.get("عدد القطع", ""))
    out["السعر مع التوصيل"] = clean_price_format(data.get("السعر مع التوصيل", ""))
    return out

# رفع الصور من الموبايل أو اللابتوب
uploaded_files = st.file_uploader("📂 اختر أو التقط صور الطلبات (يمكنك اختيار صور متعددة)", type=["png", "jpg", "jpeg", "webp"], accept_multiple_files=True)

if uploaded_files:
    st.success(f"تمت إضافة {len(uploaded_files)} صورة بنجاح.")
    
    if st.button("🤖 ابدأ استخراج الطلبات بالذكاء الاصطناعي", type="primary"):
        if not api_key_input:
            st.error("الرجاء إدخال مفتاح Gemini API في القائمة الجانبية أولاً.")
        else:
            extracted_rows = []
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for i, file in enumerate(uploaded_files):
                status_text.text(f"جارٍ معالجة الصورة ({i+1}/{len(uploaded_files)}): {file.name}")
                try:
                    row = call_gemini_web(api_key_input, model_choice, file)
                    row["_filename"] = file.name
                    extracted_rows.append(row)
                except Exception as e:
                    err_row = {f: "" for f in FIELDS}
                    err_row["اسم الزبون"] = "خطأ في القراءة"
                    err_row["الملاحظات"] = str(e)
                    err_row["_phone_status"] = "خطأ"
                    err_row["_filename"] = file.name
                    extracted_rows.append(err_row)
                progress_bar.progress((i + 1) / len(uploaded_files))
                
            status_text.text("اكتملت المعالجة بنجاح!")
            st.session_state["extracted_rows"] = extracted_rows

if "extracted_rows" in st.session_state and st.session_state["extracted_rows"]:
    st.subheader("📋 جدول الطلبات المستخرجة:")
    df_view = pd.DataFrame(st.session_state["extracted_rows"])
    
    # عرض الجدول للمستخدم
    st.dataframe(df_view, use_container_width=True)
    
    # زر تصدير أكسل بقالب الشركة
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(FIELDS)
    
    for r in st.session_state["extracted_rows"]:
        ws.append([
            r.get("اسم الزبون", ""),
            r.get("رقم الهاتف الاساسي", ""),
            r.get("رقم الهاتف الثانوي", ""),
            r.get("المحافظة", ""),
            r.get("المنطقة", ""),
            r.get("نوع البضاعه", ""),
            r.get("عدد القطع", ""),
            r.get("السعر مع التوصيل", ""),
            "عادي",
            r.get("الملاحظات", ""),
            r.get("نوع الطلب", "طلب جديد")
        ])
    
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    st.download_button(
        label="📊 تحميل ملف الأكسل جاهز لشركة التوصيل",
        data=output,
        file_name="الطلبات_جاهزة_للتوصيل.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

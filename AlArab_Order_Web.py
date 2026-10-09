# -*- coding: utf-8 -*-
"""
Al-Arab Order Web App - النسخة المحصنة ضد تكرار شركة التوصيل
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
            
            new_added_count = 0
            for i, file in enumerate(uploaded_files):
                if file:
                    img_bytes = file.getvalue()
                    st.session_state["image_cache"][file.name] = {
                        "bytes": img_bytes,
                        "type": file.type
                    }
                    
                    if file.name not in existing_files:
                        status_text.text(f"جارٍ معالجة الصورة ({i+1}/{len(uploaded_files)}): {file.name}")
                        try:
                            row = call_gemini_api(api_key_input, model_choice, img_bytes, file.type, file.name)
                            extracted_rows.append(row)
                            new_added_count += 1
                        except Exception as e:
                            err_row = {f: "" for f in FIELDS}
                            err_row["اسم الزبون"] = "خطأ في القراءة"
                            err_row["الملاحظات"] = str(e)
                            err_row["_filename"] = file.name
                            err_row["_phone_status"] = "خطأ في القراءة"
                            extracted_rows.append(err_row)
                progress_bar.progress((i + 1) / len(uploaded_files))
                
            status_text.text(f"اكتملت المعالجة! تمت إضافة {new_added_count} طلب جديد بنجاح.")
            st.session_state["extracted_rows"] = extracted_rows

if "extracted_rows" in st.session_state and st.session_state["extracted_rows"]:
    st.subheader("📋 جدول الطلبات (كشف التكرار والتعديل المباشر):")
    
    df_check = pd.DataFrame(st.session_state["extracted_rows"])
    if not df_check.empty:
        phone_counts = df_check["رقم الهاتف الاساسي"].value_counts() if "رقم الهاتف الاساسي" in df_check.columns else {}
        name_counts = df_check["اسم الزبون"].value_counts() if "اسم الزبون" in df_check.columns else {}
        
        for idx, row in df_check.iterrows():
            ph = str(row.get("رقم الهاتف الاساسي", "")).strip()
            name = str(row.get("اسم الزبون", "")).strip()
            current_status = str(row.get("_phone_status", "سليم"))
            
            is_duplicate = False
            if ph and len(ph) >= 10 and phone_counts.get(ph, 0) > 1:
                is_duplicate = True
            if name and "واتساب" not in name and "غير متوفر" not in name and name_counts.get(name, 0) > 1:
                is_duplicate = True
                
            if is_duplicate and "خطأ" not in current_status and "مفقود" not in current_status:
                df_check.loc[idx, "_phone_status"] = "⚠ مكرر (رقم أو اسم)"
            elif not is_duplicate and "مكرر" in current_status:
                df_check.loc[idx, "_phone_status"] = "سليم"

    failed_rows = [r.get("_filename"] for r in st.session_state["extracted_rows"] if "خطأ" in str(r.get("_phone_status", "")) or r.get("اسم الزبون"] == "خطأ في القراءة"]
    if failed_rows:
        selected_retry_file = st.selectbox("⚠ يوجد ملفات فيها خطأ في القراءة. اختر الملف لإعادة معالجته:", failed_rows, key="retry_select")
        if st.button("🔄 إعادة قراءة وتحديث هذه الصورة من الذاكرة"):
            if not api_key_input:
                st.error("أدخل مفتاح Gemini API أولاً.")
            else:
                cache_item = st.session_state["image_cache"].get(selected_retry_file)
                if cache_item:
                    try:
                        with st.spinner("جارٍ إعادة تحليل الصورة..."):
                            new_row = call_gemini_api(api_key_input, model_choice, cache_item["bytes"], cache_item["type"], selected_retry_file)
                            st.session_state["extracted_rows"] = [new_row if r.get("_filename"] == selected_retry_file else r for r in st.session_state["extracted_rows"]]
                            st.success(f"تمت إعادة قراءة الصورة ({selected_retry_file}) بنجاح!")
                            st.rerun()
                    except Exception as err:
                        st.error(f"فشلت إعادة المعالجة: {err}")
                else:
                    st.error("الصورة غير موجودة في الذاكرة المؤقتة.")

    col1, col2 = st.columns([2, 3])
    with col1:
        if st.button("🗑 حذف كافة الصفوف المكررة بالكامل"):
            if not df_check.empty and "رقم الهاتف الاساسي" in df_check.columns:
                df_check = df_check.drop_duplicates(subset=["رقم الهاتف الاساسي"], keep="first")
                st.session_state["extracted_rows"] = df_check.to_dict(orient="records")
                st.success("تم حذف جميع السطور المكررة بالكامل بنجاح!")
                st.rerun()

    st.info("💡 الجدول يوضح حالة الأرقام والتكرارات. يمكنك النقر على أي خلية لتعديلها مباشرة.")
    
    edited_df = st.data_editor(df_check, num_rows="dynamic", use_container_width=True)
    st.session_state["extracted_rows"] = edited_df.to_dict(orient="records")
    
    EXPORT_FIELDS = [
        "اسم الزبون", "رقم الهاتف الاساسي", "رقم الهاتف الثانوي", "المحافظة",
        "المنطقة", "نوع البضاعه", "عدد القطع", "السعر مع التوصيل", "حجم الطلب",
        "الملاحظات", "نوع الطلب"
    ]
    
    # 🛡️ حماية صارمة قبل إنشاء ملف الإكسل لمنع تكرار النظام للأسماء والهواتف بشكل قطعي
    df_export = pd.DataFrame(st.session_state["extracted_rows"])
    if not df_export.empty:
        # إزالة أي تكرار مطابق تماماً برقم الهاتف أو الاسم قبل التصدير
        df_export = df_export.drop_duplicates(subset=["رقم الهاتف الاساسي"], keep="first")
        df_export = df_export.drop_duplicates(subset=["اسم الزبون"], keep="first")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(EXPORT_FIELDS)
    
    for _, r in df_export.iterrows():
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
        label="📊 تحميل ملف الأكسل المحصن (جاهز وآمن لشركة التوصيل)",
        data=output,
        file_name="الطلبات_محصنة_بدون_تكرار.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )

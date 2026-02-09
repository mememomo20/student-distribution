from flask import Flask, render_template, request, send_from_directory
import os
import pandas as pd
from datetime import datetime
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER


# =========================
# تنظيف النصوص
# =========================
def clean_text(x):
    if pd.isna(x):
        return ""
    s = str(x).replace("\u00a0", " ").strip()
    s = " ".join(s.split())
    if s.lower() == "nan":
        return ""
    return s


def clean_columns(df):
    df = df.copy()
    df.columns = [clean_text(c) for c in df.columns]
    return df


# =========================
# ✅ اكتشاف أعمدة الرغبات ديناميكياً
# =========================
_AR_ORDINAL = {
    "الأول": 1, "الاول": 1,
    "الثاني": 2, "الثانى": 2,
    "الثالث": 3,
    "الرابع": 4,
    "الخامس": 5,
    "السادس": 6,
    "السابع": 7,
    "الثامن": 8,
    "التاسع": 9,
    "العاشر": 10,
    "الحادي عشر": 11, "الحادى عشر": 11,
    "الثاني عشر": 12, "الثانى عشر": 12,
    "الثالث عشر": 13,
    "الرابع عشر": 14,
    "الخامس عشر": 15,
    "السادس عشر": 16,
    "السابع عشر": 17,
    "الثامن عشر": 18,
    "التاسع عشر": 19,
    "العشرون": 20,
}

def _wish_index_from_col(col: str):
    """
    يرجّع رقم الرغبة إذا العمود يمثل رغبة/اختيار.
    يدعم:
    - الرغبة 1 / رغبة1 / رغبة 12
    - الاختيار 1
    - الاختيار الأول / الثاني / ... (لحد 20 تقريباً)
    """
    c = clean_text(col)

    # الرغبة + رقم
    m = re.search(r"(?:^|\s)(?:الرغبة|رغبة)\s*([0-9]+)\s*$", c)
    if m:
        return int(m.group(1))

    # الاختيار + رقم
    m = re.search(r"(?:^|\s)(?:الاختيار|اختيار)\s*([0-9]+)\s*$", c)
    if m:
        return int(m.group(1))

    # الاختيار + كلمة ترتيب (الأول..)
    m = re.search(r"(?:^|\s)(?:الاختيار|اختيار)\s*(.+)\s*$", c)
    if m:
        key = clean_text(m.group(1))
        key = key.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
        # نرجع نطابق بالقاموس مع نسختين
        for k, v in _AR_ORDINAL.items():
            kk = k.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
            if key == kk:
                return v

    return None


def get_wish_columns(df):
    """
    يرجّع قائمة أعمدة الرغبات الموجودة فعلاً بالملف مرتبة (الرغبة 1..N).
    """
    pairs = []
    for c in df.columns:
        idx = _wish_index_from_col(c)
        if idx is not None:
            pairs.append((idx, c))

    pairs.sort(key=lambda x: x[0])
    return [c for _, c in pairs]


def normalize_columns(df):
    df = clean_columns(df)

    mapping = {
        "ت": "تسلسل الطالب (ID)",
        "اسم الطالب": "اسم الطالب",
        "المعدل": "المعدل (من 0 إلى 100)",
        "قناة القبول": "قناة القبول (عام/ذوي الشهداء/موازي)",
        "الملاحظات": "ملاحظات",
        "ملاحضات": "ملاحظات",
        "ملاحظة": "ملاحظات",
        "ملاحظه": "ملاحظات",
    }

    # ✅ نعيد تسمية الاختيار الأول..الخامس إلى الرغبة 1..5 (مثل ما عندج)
    mapping.update({
        "الاختيار الأول": "الرغبة 1",
        "الاختيار الثاني": "الرغبة 2",
        "الاختيار الثالث": "الرغبة 3",
        "الاختيار الرابع": "الرغبة 4",
        "الاختيار الخامس": "الرغبة 5",
    })

    df = df.rename(columns=mapping)

    # ✅ إذا الملف ما بيه ولا رغبة، نخلي افتراضي 1..5 حتى لا ينهار النظام
    wish_cols = get_wish_columns(df)
    if len(wish_cols) == 0:
        for i in range(1, 6):
            col = f"الرغبة {i}"
            if col not in df.columns:
                df[col] = ""

    if "هل الطالب من أبناء الأساتذة؟ (نعم/لا)" not in df.columns:
        df["هل الطالب من أبناء الأساتذة؟ (نعم/لا)"] = ""

    if "ملاحظات" not in df.columns:
        df["ملاحظات"] = ""

    return df


def clean_values(df):
    df = df.copy()
    for c in df.columns:
        if df[c].dtype == "object":
            df[c] = df[c].apply(clean_text)
    return df


def normalize_channel_values(df):
    df = df.copy()
    col = "قناة القبول (عام/ذوي الشهداء/موازي)"
    if col not in df.columns:
        return df

    def fix(v):
        s = clean_text(v)
        if s == "المركزي":
            return "عام"
        if s == "ذوي الشهداء":
            return "ذوي الشهداء"
        if s == "الموازي":
            return "موازي"

        if "شه" in s:
            return "ذوي الشهداء"
        if "موا" in s:
            return "موازي"
        if "مرك" in s or "عام" in s:
            return "عام"

        return "غير معروف"

    df[col] = df[col].apply(fix)
    return df


# =========================
# أبناء الأساتذة من "ملاحظات"
# =========================
def derive_teacher_flag_from_notes(df):
    df = df.copy()
    col_teacher = "هل الطالب من أبناء الأساتذة؟ (نعم/لا)"
    col_notes = "ملاحظات"

    def is_teacher_from_notes(note: str) -> bool:
        s = clean_text(note)
        s = s.replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
        keys = ["ابناء الاساتذة", "ابناء الاساتذه", "ابناء الاستاذه", "ابناء الأساتذة"]
        return any(k.replace("أ", "ا") in s for k in keys) or ("ابناء" in s and "اسات" in s)

    def norm_yes_no(v):
        s = clean_text(v)
        if s in ["نعم", "yes", "Yes", "Y", "y", "1", "true", "True"]:
            return "نعم"
        if s in ["لا", "no", "No", "N", "n", "0", "false", "False"]:
            return "لا"
        return ""

    df[col_teacher] = df[col_teacher].apply(norm_yes_no)

    def fill_row(row):
        v = row.get(col_teacher, "")
        if v in ["نعم", "لا"]:
            return v
        return "نعم" if is_teacher_from_notes(row.get(col_notes, "")) else "لا"

    df[col_teacher] = df.apply(fill_row, axis=1)
    return df


# =========================
# التوزيع (ديناميك رغبات)
# =========================
def extract_departments_from_file(df, wish_cols):
    depts = set()
    for w in wish_cols:
        for v in df[w].tolist():
            v = clean_text(v)
            if v:
                depts.add(v)
    return sorted(depts)


def calc_equal_capacity(num_students, num_departments):
    if num_departments <= 0:
        return 0, 0
    base = num_students // num_departments
    rem = num_students % num_departments
    return base, rem


def get_wishes(row, wish_cols):
    out = []
    for w in wish_cols:
        v = clean_text(row.get(w, ""))
        if v:
            out.append(v)
    return out


def build_dept_stats(df_alloc, channel_name):
    col_avg = "المعدل (من 0 إلى 100)"
    col_dept = "القسم المقبول فيه الطالب"

    accepted = df_alloc[df_alloc[col_dept] != "غير مقبول"].copy()
    if accepted.empty:
        return [], {}

    g = accepted.groupby(col_dept)[col_avg].agg(["count", "min"]).reset_index()
    g.columns = ["القسم", "عدد المقبولين", "المعدل الأدنى"]
    rows = g.to_dict(orient="records")

    for r in rows:
        r["القناة"] = channel_name

    mins = {r["القسم"]: float(r["المعدل الأدنى"]) for r in rows}
    return rows, mins


def allocate_normal_students(df_ch, capacities, wish_cols):
    remaining = capacities.copy()
    accepted_dept = []
    notes = []

    for _, row in df_ch.iterrows():
        chosen = "غير مقبول"
        why = ""
        for dept in get_wishes(row, wish_cols):
            if dept in remaining and remaining[dept] > 0:
                chosen = dept
                remaining[dept] -= 1
                why = "قبول ضمن الطاقة"
                break
        accepted_dept.append(chosen)
        notes.append(why)

    out = df_ch.copy()
    out["القسم المقبول فيه الطالب"] = accepted_dept
    out["سبب/ملاحظة"] = notes
    return out


def allocate_teachers_children(df_teachers, dept_min, wish_cols, margin=5):
    col_avg = "المعدل (من 0 إلى 100)"
    accepted_dept = []
    notes = []

    for _, row in df_teachers.iterrows():
        chosen = "غير مقبول"
        avg = float(row.get(col_avg, 0))

        for dept in get_wishes(row, wish_cols):
            if dept not in dept_min:
                continue
            if avg >= (dept_min[dept] - margin):
                chosen = dept
                break

        accepted_dept.append(chosen)
        notes.append(f"أبناء الأساتذة: قبول فوق الطاقة بشرط (>= الحد الأدنى - {margin})")

    out = df_teachers.copy()
    out["القسم المقبول فيه الطالب"] = accepted_dept
    out["سبب/ملاحظة"] = notes
    return out


def read_manual_caps_from_form(departments):
    caps = {}
    for i in range(len(departments)):
        dept_name = clean_text(request.form.get(f"cap_name__{i}", ""))
        seats_text = request.form.get(f"cap__{i}", "").strip()
        seats = int(seats_text) if seats_text.isdigit() else 0
        if dept_name:
            caps[dept_name] = seats
    return caps


def compute_stats(df):
    col_channel = "قناة القبول (عام/ذوي الشهداء/موازي)"
    return {
        "عدد_الكل": len(df),
        "عام": int((df[col_channel] == "عام").sum()),
        "ذوي_الشهداء": int((df[col_channel] == "ذوي الشهداء").sum()),
        "موازي": int((df[col_channel] == "موازي").sum()),
    }


@app.route("/", methods=["GET", "POST"])
def index():
    رسالة = ""
    احصائيات = None
    detected_departments = []
    filename = ""

    capacity_mode = "equal"
    last_caps = {}
    teacher_margin = 5

    if request.method == "POST":
        ملف = request.files.get("excel_file")
        if not ملف or ملف.filename == "":
            رسالة = "⚠️ اختاري ملف Excel أولاً."
        else:
            try:
                filename = ملف.filename
                path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
                ملف.save(path)

                df = pd.read_excel(path)
                df = normalize_columns(df)
                df = clean_values(df)
                df = normalize_channel_values(df)
                df = derive_teacher_flag_from_notes(df)

                col_channel = "قناة القبول (عام/ذوي الشهداء/موازي)"
                col_avg = "المعدل (من 0 إلى 100)"

                if col_channel not in df.columns or col_avg not in df.columns:
                    رسالة = "❌ الأعمدة الأساسية غير موجودة (قناة القبول + المعدل)."
                    filename = ""
                else:
                    df = df[df[col_channel].isin(["عام", "ذوي الشهداء", "موازي"])].copy()
                    احصائيات = compute_stats(df)

                    wish_cols = get_wish_columns(df)
                    if len(wish_cols) == 0:
                        رسالة = "❌ ماكو أعمدة رغبات بالملف (مثل: الرغبة 1 أو الاختيار الأول)."
                        filename = ""
                    else:
                        detected_departments = extract_departments_from_file(df, wish_cols)
                        if len(detected_departments) == 0:
                            رسالة = "❌ الرغبات موجودة بس الأقسام فارغة."
                            filename = ""
                        else:
                            رسالة = f"✅ تم رفع الملف. تم اكتشاف {len(wish_cols)} رغبة.  ."

            except Exception as e:
                رسالة = f"❌ صار خطأ: {e}"
                filename = ""

    return render_template(
        "index.html",
        رسالة=رسالة,
        احصائيات=احصائيات,
        detected_departments=detected_departments,
        filename=filename,
        teacher_margin=teacher_margin,
        capacity_mode=capacity_mode,
        last_caps=last_caps,
        can_rerun=False,
        active_tab="home",
    )


@app.route("/distribute", methods=["POST"])
def distribute():
    try:
        filename = request.form.get("filename", "")
        if not filename:
            return "❌ ماكو ملف مرفوع.", 400

        path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        if not os.path.exists(path):
            return "❌ الملف غير موجود داخل uploads.", 400

        df = pd.read_excel(path)
        df = normalize_columns(df)
        df = clean_values(df)
        df = normalize_channel_values(df)
        df = derive_teacher_flag_from_notes(df)

        col_avg = "المعدل (من 0 إلى 100)"
        col_channel = "قناة القبول (عام/ذوي الشهداء/موازي)"
        col_teacher = "هل الطالب من أبناء الأساتذة؟ (نعم/لا)"

        if col_avg not in df.columns or col_channel not in df.columns:
            return "❌ الأعمدة الأساسية غير موجودة (المعدل + قناة القبول).", 400

        df = df[df[col_channel].isin(["عام", "ذوي الشهداء", "موازي"])].copy()
        df[col_avg] = pd.to_numeric(df[col_avg], errors="coerce").fillna(0)

        wish_cols = get_wish_columns(df)
        if len(wish_cols) == 0:
            return "❌ ماكو أعمدة رغبات بالملف.", 400

        departments = extract_departments_from_file(df, wish_cols)
        departments = list(dict.fromkeys(departments))
        if len(departments) == 0:
            return "❌ ماكو أقسام داخل الرغبات.", 400

        margin_text = request.form.get("teacher_margin", "5").strip()
        teacher_margin = int(margin_text) if margin_text.isdigit() else 5

        capacity_mode = request.form.get("capacity_mode", "equal")
        last_caps = {}

        if capacity_mode == "manual":
            manual_caps = read_manual_caps_from_form(departments)
            last_caps = manual_caps.copy()
            if sum(manual_caps.values()) == 0:
                return "❌ الطاقة اليدوية كلها صفر! دخلي أرقام للطاقة.", 400
        else:
            manual_caps = {}

        results = []
        all_dept_stats_rows = []

        for channel_name in ["عام", "ذوي الشهداء", "موازي"]:
            df_ch = df[df[col_channel] == channel_name].copy()
            if df_ch.empty:
                continue

            df_ch = df_ch.sort_values(by=col_avg, ascending=False)

            if capacity_mode == "equal":
                base, rem = calc_equal_capacity(len(df_ch), len(departments))
                capacities = {d: base + (1 if i < rem else 0) for i, d in enumerate(departments)}
                if sum(capacities.values()) == 0:
                    return "❌ التوزيع المتساوي طلع طاقة صفر (عدد الأقسام أكبر من عدد الطلبة).", 400
            else:
                capacities = manual_caps.copy()

            teachers = df_ch[df_ch[col_teacher] == "نعم"].copy()
            normals = df_ch[df_ch[col_teacher] != "نعم"].copy()

            normals_alloc = allocate_normal_students(normals, capacities, wish_cols)
            _, dept_min_normals = build_dept_stats(normals_alloc, channel_name)

            teachers_alloc = allocate_teachers_children(teachers, dept_min_normals, wish_cols, margin=teacher_margin)

            alloc_channel = pd.concat([normals_alloc, teachers_alloc], ignore_index=True)
            alloc_channel["القناة"] = channel_name
            results.append(alloc_channel)

            rows_final, _ = build_dept_stats(alloc_channel, channel_name)
            all_dept_stats_rows.extend(rows_final)

        if not results:
            return "❌ ماكو نتائج.", 400

        final_df = pd.concat(results, ignore_index=True)

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_name = f"نتائج_{os.path.splitext(filename)[0]}_{stamp}.xlsx"
        out_path = os.path.join(app.config["UPLOAD_FOLDER"], out_name)
        final_df.to_excel(out_path, index=False)

        results_rows = final_df.to_dict(orient="records")
        احصائيات = compute_stats(df)
        رسالة = f"✅ تم التوزيع بنجاح. تم اكتشاف {len(wish_cols)} رغبة. هامش أبناء الأساتذة = {teacher_margin}."

        return render_template(
            "index.html",
            رسالة=رسالة,
            احصائيات=احصائيات,
            detected_departments=departments,
            filename=filename,
            download_file=out_name,
            results_rows=results_rows,
            dept_stats_rows=all_dept_stats_rows,
            teacher_margin=teacher_margin,
            capacity_mode=capacity_mode,
            last_caps=last_caps,
            can_rerun=True,
            active_tab="home",
        )

    except Exception as e:
        return f"❌ صار خطأ أثناء التوزيع: {e}", 500


@app.route("/download/<path:fname>")
def download(fname):
    return send_from_directory(app.config["UPLOAD_FOLDER"], fname, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)

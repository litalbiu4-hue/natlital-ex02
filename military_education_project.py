# ==========================================================
# Military Education Classification System
# Version 1.0
# ==========================================================

import pandas as pd
import math
from collections import defaultdict

# ==========================================================
# הגדרות מערכת
# ==========================================================

INPUT_FILE = "Input_400_Candidates.xlsx"
OUTPUT_FILE = "Output_Military_Education.xlsx"

MIN_CLASS_SIZE = 20
MAX_CLASS_SIZE = 35

# ==========================================================
# כיתות קיימות
# ==========================================================

CLASSES = [
    {
        "class_name": "מכינת תון צפון",
        "track": "מכינת תון",
        "region": "צפון"
    },
    {
        "class_name": "מכינת תון מרכז",
        "track": "מכינת תון",
        "region": "מרכז"
    },
    {
        "class_name": "טכנאי תון צפון",
        "track": "טכנאי תון",
        "region": "צפון"
    },
    {
        "class_name": "טכנאי תון מרכז",
        "track": "טכנאי תון",
        "region": "מרכז"
    },
    {
        "class_name": "הנדסאי תון מרכז 1",
        "track": "הנדסאי תון",
        "region": "מרכז"
    },
    {
        "class_name": "הנדסאי תון מרכז 2",
        "track": "הנדסאי תון",
        "region": "מרכז"
    },
    {
        "class_name": "תואר ראשון צפון",
        "track": "תואר ראשון",
        "region": "צפון"
    },
    {
        "class_name": "תואר ראשון מרכז",
        "track": "תואר ראשון",
        "region": "מרכז"
    },
    {
        "class_name": "תואר שני מרכז",
        "track": "תואר שני",
        "region": "מרכז"
    },
    {
        "class_name": "תואר שני דרום",
        "track": "תואר שני",
        "region": "דרום"
    }
]

# ==========================================================
# פונקציות עזר
# ==========================================================

def get_missing_subjects(row):

    missing = []

    if (
        row["ציון אנגלית"] < 60
        or row["יחל אנגלית"] < 3
    ):
        missing.append("אנגלית")

    if (
        row["ציון מתמטיקה"] < 60
        or row["יחל מתמטיקה"] < 3
    ):
        missing.append("מתמטיקה")

    if (
        row["ציון עברית"] < 60
        or row["יחל עברית"] < 2
    ):
        missing.append("עברית")

    return missing


def document_status(row):

    docs = [
        "צילום תז",
        "אישור 12 שנות לימוד",
        "תעודת בגרות",
        "תעודת תואר",
        "גיליון ציונים אקדמי"
    ]

    missing = []

    for doc in docs:

        if doc in row:

            if str(row[doc]).strip() == "לא":
                missing.append(doc)

    if len(missing) == 0:

        return (
            "תקין",
            ""
        )

    return (
        "נדרש בהשלמת מסמכים",
        ", ".join(missing)
    )# ==========================================================
# קבילות אקדמית
# ==========================================================

def technician_eligibility(row):

    age = row["גיל"]
    years = row["שנות לימוד"]

    missing = get_missing_subjects(row)

    if age >= 30:

        if years >= 12:

            return (
                True,
                "קביל לטכנאי תון",
                ""
            )

        return (
            False,
            "אינו קביל אקדמית",
            "פחות מ-12 שנות לימוד"
        )

    # מתחת לגיל 30

    if len(missing) == 0:

        return (
            True,
            "קביל לטכנאי תון",
            ""
        )

    if len(missing) == 1:

        return (
            True,
            "קביל אקדמית על תנאי לטכנאי תון",
            missing[0]
        )

    return (
        False,
        "קביל למכינת תון",
        ", ".join(missing)
    )


# ==========================================================
# הנדסאי
# ==========================================================

def practical_engineer_eligibility(row):

    age = row["גיל"]
    years = row["שנות לימוד"]

    missing = get_missing_subjects(row)

    if age >= 35:

        if years >= 12:

            return (
                True,
                "קביל להנדסאי תון",
                ""
            )

        return (
            False,
            "אינו קביל אקדמית",
            "פחות מ-12 שנות לימוד"
        )

    # מתחת לגיל 35

    if len(missing) == 0:

        return (
            True,
            "קביל להנדסאי תון",
            ""
        )

    if len(missing) == 1:

        return (
            True,
            "קביל אקדמית על תנאי להנדסאי תון",
            missing[0]
        )

    return (
        False,
        "קביל למכינת תון",
        ", ".join(missing)
    )


# ==========================================================
# מכינת תו"ן
# ==========================================================

def preparatory_program_eligibility(row):

    age = row["גיל"]
    years = row["שנות לימוד"]

    if age >= 30 and years >= 12:

        return (
            False,
            "לא נדרש למכינה",
            ""
        )

    return (
        True,
        "קביל למכינת תון",
        ""
    )


# ==========================================================
# תואר ראשון
# ==========================================================

def bachelor_eligibility(row):

    age = row["גיל"]

    years = row["שנות לימוד"]

    bagrut_avg = row["ממוצע בגרות"]

    if age < 30:

        if bagrut_avg >= 85:

            return (
                True,
                "קביל לתואר ראשון",
                "נדרש במכינה אקדמית כחלק מהמסלול"
            )

        return (
            False,
            "אינו קביל לתואר ראשון",
            "ממוצע בגרות נמוך מ-85"
        )

    # גיל 30 ומעלה

    if years >= 12:

        return (
            True,
            "קביל לתואר ראשון",
            "נדרש במכינה אקדמית כחלק מהמסלול"
        )

    return (
        False,
        "אינו קביל לתואר ראשון",
        "פחות מ-12 שנות לימוד"
    )


# ==========================================================
# תואר שני
# ==========================================================

def master_eligibility(row):

    degree_field = str(row["תחום תואר"]).strip()

    degree_avg = row["ממוצע תואר ראשון"]

    if degree_field == "מדעי":

        if degree_avg >= 75:

            return (
                True,
                "קביל לתואר שני",
                ""
            )

    else:

        if degree_avg >= 78:

            return (
                True,
                "קביל לתואר שני",
                ""
            )

    return (
        False,
        "אינו קביל לתואר שני",
        ""
    )


# ==========================================================
# מנוע קבילות ראשי
# ==========================================================

def evaluate_track(row, track):

    if track == "מכינת תון":
        return preparatory_program_eligibility(row)

    if track == "טכנאי תון":
        return technician_eligibility(row)

    if track == "הנדסאי תון":
        return practical_engineer_eligibility(row)

    if track == "תואר ראשון":
        return bachelor_eligibility(row)

    if track == "תואר שני":
        return master_eligibility(row)

    return (
        False,
        "מסלול לא מוכר",
        ""
    )# ==========================================================
# המלצה להנדסאי
# ==========================================================

def engineer_recommendation(row):

    preferred = str(row["מסלול מועדף"]).strip()

    if preferred != "טכנאי תון":

        return ""

    eligible, status, note = practical_engineer_eligibility(row)

    if eligible and "על תנאי" not in status:

        return "מומלץ לבחון שיבוץ להנדסאי תון"

    return ""


# ==========================================================
# קביעת מסלול סופי
# ==========================================================

def determine_final_track(row):

    preferred = str(row["מסלול מועדף"]).strip()

    alternative = str(row["מסלול חלופי"]).strip()

    preferred_ok, preferred_status, preferred_note = evaluate_track(
        row,
        preferred
    )

    if preferred_ok:

        return {
            "קבילות אקדמית": preferred_status,
            "שיבוץ מוצע": preferred,
            "הערות": "שובץ למסלול המועדף",
            "השלמות בגרות": preferred_note
        }

    alternative_ok, alternative_status, alternative_note = evaluate_track(
        row,
        alternative
    )

    if alternative_ok:

        return {
            "קבילות אקדמית": alternative_status,
            "שיבוץ מוצע": alternative,
            "הערות": "לא קביל למסלול המועדף, שובץ למסלול החלופי",
            "השלמות בגרות": alternative_note
        }

    prep_ok, prep_status, prep_note = preparatory_program_eligibility(row)

    if prep_ok:

        return {
            "קבילות אקדמית": prep_status,
            "שיבוץ מוצע": "מכינת תון",
            "הערות": "הופנה למכינת תון",
            "השלמות בגרות": prep_note
        }

    return {
        "קבילות אקדמית": "אינו קביל אקדמית",
        "שיבוץ מוצע": "ללא שיבוץ",
        "הערות": "לא נמצא מסלול מתאים",
        "השלמות בגרות": ""
    }


# ==========================================================
# טעינת קובץ
# ==========================================================

print("טוען קובץ מועמדים...")

df = pd.read_excel(INPUT_FILE)

print(f"נקראו {len(df)} מועמדים")


# ==========================================================
# בדיקת מסמכים
# ==========================================================

doc_status_list = []
missing_docs_list = []

for _, row in df.iterrows():

    status, missing_docs = document_status(row)

    doc_status_list.append(status)

    missing_docs_list.append(missing_docs)

df["סטטוס מסמכים"] = doc_status_list

df["מסמכים חסרים"] = missing_docs_list


# ==========================================================
# קביעת קבילות אקדמית
# ==========================================================

academic_status = []

assigned_tracks = []

notes = []

missing_subjects_column = []

recommendations = []

for _, row in df.iterrows():

    result = determine_final_track(row)

    academic_status.append(
        result["קבילות אקדמית"]
    )

    assigned_tracks.append(
        result["שיבוץ מוצע"]
    )

    notes.append(
        result["הערות"]
    )

    missing_subjects_column.append(
        result["השלמות בגרות"]
    )

    recommendations.append(
        engineer_recommendation(row)
    )

df["קבילות אקדמית"] = academic_status

df["שיבוץ מוצע"] = assigned_tracks

df["הערות"] = notes

df["השלמות בגרות"] = missing_subjects_column

df["המלצת מערכת"] = recommendations


# ==========================================================
# יצירת מבנה כיתות
# ==========================================================

class_capacity = {}

for classroom in CLASSES:

    class_capacity[
        classroom["class_name"]
    ] = []

waiting_lists = defaultdict(list)

print("הסתיים שלב הקבילות והשיבוץ הראשוני")
# ==========================================================
# פונקציה למציאת כיתה מתאימה
# ==========================================================

def find_class(track, region):

    possible_classes = []

    for classroom in CLASSES:

        if classroom["track"] == track:

            possible_classes.append(classroom)

    # קודם כל מחפשים באזור המגורים

    for classroom in possible_classes:

        if classroom["region"] == region:

            return classroom["class_name"]

    # אם לא נמצא - מחפשים במרכז

    for classroom in possible_classes:

        if classroom["region"] == "מרכז":

            return classroom["class_name"]

    return None


# ==========================================================
# שיבוץ בפועל לכיתות
# ==========================================================

assigned_classes = []

waiting_status = []

for idx, row in df.iterrows():

    track = row["שיבוץ מוצע"]

    region = row["אזור"]

    approved_center_transfer = str(
        row.get("מאשר מעבר למרכז", "לא")
    ).strip()

    class_name = find_class(track, region)

    if class_name is None:

        assigned_classes.append("ללא כיתה")

        waiting_status.append("לא שובץ")

        continue

    # בדיקת קיבולת

    if len(class_capacity[class_name]) < MAX_CLASS_SIZE:

        class_capacity[class_name].append(
            row["תעודת זהות"]
        )

        assigned_classes.append(class_name)

        waiting_status.append("שובץ")

        continue

    # כיתה מלאה

    moved_to_center = False

    if region in ["צפון", "דרום"]:

        if approved_center_transfer == "כן":

            center_class = find_class(
                track,
                "מרכז"
            )

            if center_class:

                if len(
                    class_capacity[center_class]
                ) < MAX_CLASS_SIZE:

                    class_capacity[
                        center_class
                    ].append(
                        row["תעודת זהות"]
                    )

                    assigned_classes.append(
                        center_class
                    )

                    waiting_status.append(
                        "שובץ במרכז"
                    )

                    moved_to_center = True

    if moved_to_center:

        continue

    waiting_lists[class_name].append(
        row["תעודת זהות"]
    )

    assigned_classes.append(
        f"רשימת המתנה - {class_name}"
    )

    waiting_status.append(
        "רשימת המתנה"
    )

df["כיתה מוצעת"] = assigned_classes

df["סטטוס שיבוץ"] = waiting_status


# ==========================================================
# תפוסות כיתות
# ==========================================================

class_summary = []

for class_name, students in class_capacity.items():

    class_summary.append(
        {
            "כיתה": class_name,
            "מספר משובצים": len(students),
            "מינימום": MIN_CLASS_SIZE,
            "מקסימום": MAX_CLASS_SIZE
        }
    )

class_summary_df = pd.DataFrame(
    class_summary
)


# ==========================================================
# רשימות המתנה
# ==========================================================

waiting_summary = []

for class_name, students in waiting_lists.items():

    waiting_summary.append(
        {
            "כיתה": class_name,
            "מספר ממתינים": len(students)
        }
    )

waiting_df = pd.DataFrame(
    waiting_summary
)


# ==========================================================
# המלצות פתיחת כיתות
# ==========================================================

recommendations = []

for _, row in class_summary_df.iterrows():

    count = row["מספר משובצים"]

    if count < MIN_CLASS_SIZE:

        recommendation = (
            "לא מומלץ לפתוח כיתה"
        )

    elif count <= MAX_CLASS_SIZE:

        recommendation = (
            "מומלץ לפתוח כיתה"
        )

    else:

        recommendation = (
            "מומלץ לבחון פתיחת כיתה נוספת"
        )

    recommendations.append(
        recommendation
    )

class_summary_df[
    "המלצת מערכת"
] = recommendations


# ==========================================================
# המלצה לכיתה נוספת לפי רשימת המתנה
# ==========================================================

extra_class_recommendations = []

for _, row in waiting_df.iterrows():

    waiting_count = row["מספר ממתינים"]

    if waiting_count >= 20:

        extra_class_recommendations.append(
            "מומלץ לבחון פתיחת כיתה נוספת"
        )

    else:

        extra_class_recommendations.append(
            "אין הצדקה לכיתה נוספת"
        )

if len(waiting_df) > 0:

    waiting_df[
        "המלצה"
    ] = extra_class_recommendations


# ==========================================================
# דוחות
# ==========================================================

eligibility_report = (
    df.groupby(
        "קבילות אקדמית"
    )
    .size()
    .reset_index(name="כמות")
)

conditional_report = df[
    df["קבילות אקדמית"]
    .str.contains(
        "על תנאי",
        na=False
    )
]

missing_docs_report = df[
    df["סטטוס מסמכים"]
    != "תקין"
]

gender_report = (
    df.groupby(
        ["מגדר", "קבילות אקדמית"]
    )
    .size()
    .reset_index(name="כמות")
)

region_report = (
    df.groupby(
        ["אזור", "קבילות אקדמית"]
    )
    .size()
    .reset_index(name="כמות")
)

rank_report = (
    df.groupby(
        "דרגה"
    )
    .size()
    .reset_index(name="כמות")
)

assignment_report = df[
    [
        "תעודת זהות",
        "שם מלא",
        "שיבוץ מוצע",
        "כיתה מוצעת",
        "סטטוס שיבוץ"
    ]
]


# ==========================================================
# דוח מנהלים
# ==========================================================

manager_report = pd.DataFrame(
    [
        {
            "מדד": "סהכ מועמדים",
            "ערך": len(df)
        },
        {
            "מדד": "קבילים",
            "ערך": len(
                df[
                    df["קבילות אקדמית"]
                    .str.contains(
                        "קביל",
                        na=False
                    )
                ]
            )
        },
        {
            "מדד": "קבילים על תנאי",
            "ערך": len(
                conditional_report
            )
        },
        {
            "מדד": "חסרי מסמכים",
            "ערך": len(
                missing_docs_report
            )
        },
        {
            "מדד": "ברשימות המתנה",
            "ערך": len(
                df[
                    df["סטטוס שיבוץ"]
                    == "רשימת המתנה"
                ]
            )
        }
    ]
)


# ==========================================================
# שמירת קובץ פלט
# ==========================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    df.to_excel(
        writer,
        sheet_name="מועמדים",
        index=False
    )

    eligibility_report.to_excel(
        writer,
        sheet_name="קבילויות",
        index=False
    )

    conditional_report.to_excel(
        writer,
        sheet_name="קבילויות על תנאי",
        index=False
    )

    missing_docs_report.to_excel(
        writer,
        sheet_name="חוסרי מסמכים",
        index=False
    )

    assignment_report.to_excel(
        writer,
        sheet_name="שיבוצים",
        index=False
    )

    class_summary_df.to_excel(
        writer,
        sheet_name="תפוסת כיתות",
        index=False
    )

    waiting_df.to_excel(
        writer,
        sheet_name="רשימות המתנה",
        index=False
    )

    gender_report.to_excel(
        writer,
        sheet_name="פילוח מגדרי",
        index=False
    )

    region_report.to_excel(
        writer,
        sheet_name="פילוח אזורי",
        index=False
    )

    rank_report.to_excel(
        writer,
        sheet_name="פילוח דרגות",
        index=False
    )

    class_summary_df.to_excel(
        writer,
        sheet_name="המלצות פתיחה",
        index=False
    )

    manager_report.to_excel(
        writer,
        sheet_name="דוח מנהלים",
        index=False
    )

print("=" * 50)
print("המערכת הסתיימה בהצלחה")
print(f"קובץ פלט נוצר: {OUTPUT_FILE}")
print("=" * 50)
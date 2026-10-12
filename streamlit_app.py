"""ALTERNATIVE (simpler) web version: Streamlit Community Cloud.

No Supabase / Vercel needed. Deploy free at share.streamlit.io from your
GitHub repo, then use from any system: upload menu.xlsx in the browser,
pick month/year/holidays, download the PDF. Menu edits happen by editing
the xlsx and re-uploading.

Deploy steps are in README_DEPLOY.md (Option B).
"""

import calendar
import datetime

import streamlit as st

import menu_app  # desktop core: build_pdf, monday_weeks, sheet_for_week, ...

st.set_page_config(page_title="Monthly Menu PDF Generator", layout="wide")
st.title("Monthly Menu PDF Generator")

up = st.file_uploader("1. Upload menu.xlsx", type=["xlsx"])
now = datetime.date.today()
month = st.selectbox("2. Month", list(calendar.month_name)[1:],
                     index=now.month - 1)
year = st.selectbox("3. Year", list(range(now.year - 2, now.year + 10)),
                    index=2)
hol_text = st.text_input("4. Holidays (dates)", placeholder="e.g. 12, 13, 19, 20")
second_sat = st.checkbox("Second Saturday is a holiday")

if up and st.button("Generate PDF"):
    m = list(calendar.month_name).index(month)
    days, bad = menu_app.parse_holiday_days(hol_text, year, m)
    dates = {datetime.date(year, m, d) for d in days}
    if second_sat:
        ss = menu_app.second_saturday(year, m)
        if ss:
            dates.add(ss)
    if bad:
        st.warning("Ignored invalid entries: " + ", ".join(bad))

    # The core reads menu.xlsx / format.pdf from disk, so stage the upload.
    import os
    up.seek(0)
    with open(menu_app.MENU_XLSX, "wb") as f:
        f.write(up.read())
    if not os.path.exists(menu_app.FORMAT_PDF):
        st.error("format.pdf not found next to the app.")
    else:
        out = "%s %d Menu.pdf" % (month, year)
        menu_app.build_pdf(year, m, dates, out)
        with open(out, "rb") as f:
            st.download_button("Download " + out, f.read(), file_name=out)

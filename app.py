# -*- coding: utf-8 -*-
"""
საწყობის ნაშთების აღრიცხვა — ყოველდღიური 1C ფაილის დაწერა სააღრიცხვო ფაილში.
გაშვება:  streamlit run app.py
"""
from datetime import datetime

import pandas as pd
import streamlit as st

import engine

st.set_page_config(page_title="საწყობის აღრიცხვა", page_icon="📦", layout="wide")
st.title("📦 საწყობის ნაშთების აღრიცხვა")
st.caption("ატვირთეთ ორი ფაილი: ყოველდღიური 1C რეპორტი (ოპერაციის ასოებით) და სააღრიცხვო ფაილი.")

with st.expander("ℹ️ ოპერაციის ასოები (ივსება ყოველდღიურ ფაილში, სვეტის სათაურის ზემოთ)"):
    st.markdown(
        "| ასო | ოპერაცია | ეფექტი ნაშთზე |\n|---|---|---|\n"
        "| **მ** | მიღება | + |\n| **დ** | დაბრუნება | + |\n"
        "| **გ** | გაცემა | − |\n| **ჩ** | ჩამოწერა | − |"
    )

c1, c2 = st.columns(2)
daily_up = c1.file_uploader("1️⃣ ყოველდღიური ფაილი (1C, ასოებით)", type=["xls", "xlsx"])
stock_up = c2.file_uploader("2️⃣ სააღრიცხვო ფაილი (ნაშთები + ჟურნალი)", type=["xlsx"])

if not (daily_up and stock_up):
    st.info("გასაგრძელებლად ატვირთეთ ორივე ფაილი.")
    st.stop()

# ---------------------------------------------------------------- parse both
parsed = engine.parse_daily_file(daily_up.getvalue(), daily_up.name)
bal, led, stock_problems = engine.load_stock_file(stock_up.getvalue())

if stock_problems:
    st.error("სააღრიცხვო ფაილის პრობლემები:")
    for p in stock_problems:
        st.write("• " + p)
    st.stop()

if parsed["errors"]:
    st.error("ყოველდღიური ფაილის შეცდომები — გაასწორეთ და ატვირთეთ თავიდან:")
    for e in parsed["errors"]:
        st.write("• " + e)
    if parsed["docs"]:
        st.dataframe(pd.DataFrame(parsed["docs"])[["doc_type", "doc_no", "date", "key"]]
                     .rename(columns={"doc_type": "ოპერაციის ტიპი", "doc_no": "დოკუმენტი",
                                      "date": "თარიღი", "key": "ასო"}),
                     use_container_width=True, hide_index=True)
    st.stop()

tx = parsed["transactions"]
if tx.empty:
    st.warning("ფაილში დასაწერი რაოდენობები ვერ მოიძებნა.")
    st.stop()

for w in parsed["warnings"]:
    st.warning(w)

# ---------------------------------------------------------------- duplicates
dups = engine.find_duplicates(led, tx)
if dups:
    st.error("ეს დოკუმენტები უკვე დაწერილია ჟურნალში (ფაილი მეორედ იტვირთება?):")
    st.dataframe(pd.DataFrame(dups, columns=["დოკუმენტი", "თარიღი"]),
                 use_container_width=True, hide_index=True)
    st.stop()

# ---------------------------------------------------------------- preview
st.subheader("გადასამოწმებელი ინფორმაცია")

doc_df = pd.DataFrame(parsed["docs"])
doc_df["ოპერაცია"] = doc_df["key"].map(lambda k: f"{k} — {engine.KEY_MAP[k][0]}")
doc_df["ეფექტი"] = doc_df["key"].map(lambda k: "+" if engine.KEY_MAP[k][1] > 0 else "−")
st.markdown("**დოკუმენტები და მონიშნული ოპერაციები**")
st.dataframe(doc_df[["date", "doc_no", "doc_type", "ოპერაცია", "ეფექტი"]]
             .rename(columns={"date": "თარიღი", "doc_no": "დოკუმენტი", "doc_type": "1C ტიპი"}),
             use_container_width=True, hide_index=True)

summ = (tx.groupby("ოპერაცია")
        .agg(დოკუმენტები=("დოკუმენტი", "nunique"), პოზიციები=("კოდი", "count"),
             რაოდენობა=("რაოდენობა", "sum"), წმინდა_ცვლილება=("ცვლილება (+/-)", "sum"))
        .reset_index())
c3, c4, c5 = st.columns(3)
c3.metric("ტრანზაქციები", f"{len(tx):,}")
c4.metric("სულ მოძრაობა (ცალი)", f"{tx['რაოდენობა'].sum():,.0f}")
c5.metric("წმინდა ცვლილება ნაშთზე", f"{tx['ცვლილება (+/-)'].sum():+,.0f}")
st.markdown("**ჯამები ოპერაციების მიხედვით** — შეადარეთ მოლოდინს დაწერამდე")
st.dataframe(summ, use_container_width=True, hide_index=True)

known = set(bal["კოდი"].dropna().astype(int))
new_skus = tx.loc[~tx["კოდი"].isin(known), ["კოდი", "ნომენკლატურა"]].drop_duplicates("კოდი")
if not new_skus.empty:
    st.warning(f"{len(new_skus)} კოდი არ არის ნაშთების ფურცელზე — დაემატება ავტომატურად, საწყისი ნაშთით 0:")
    st.dataframe(new_skus, use_container_width=True, hide_index=True)

with st.expander("ყველა ტრანზაქცია დეტალურად"):
    st.dataframe(tx, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------- post
st.divider()
if st.button("✅ დაწერე ჟურნალში და ჩამოტვირთე განახლებული ფაილი", type="primary"):
    data, info = engine.post_and_build(bal, led, tx, daily_up.name)
    st.success(f"დაიწერა {info['posted_rows']} ჩანაწერი.")

    if not info["negative"].empty:
        st.error(f"⚠️ {len(info['negative'])} SKU-ს ნაშთი უარყოფითია — გადაამოწმეთ საწყისი ნაშთი ან ოპერაციები:")
        st.dataframe(info["negative"][["კოდი", "ნომენკლატურა", "საწყისი ნაშთი", "მიმდინარე ნაშთი"]],
                     use_container_width=True, hide_index=True)

    with st.expander("განახლებული ნაშთები (მოძრაობა დღეს)"):
        moved = info["balances"][info["balances"]["კოდი"].isin(tx["კოდი"])]
        st.dataframe(moved[["კოდი", "ნომენკლატურა", "საწყისი ნაშთი", "მიმდინარე ნაშთი"]],
                     use_container_width=True, hide_index=True)

    fname = f"სააღრიცხვო_ფაილი_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
    st.download_button("⬇️ განახლებული სააღრიცხვო ფაილი", data=data, file_name=fname,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.info("ჩამოტვირთეთ ფაილი და ჩაანაცვლეთ ძველი ვერსია — შემდეგ დღეს ეს ფაილი ატვირთეთ.")

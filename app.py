# -*- coding: utf-8 -*-
"""
საწყობის ნაშთების აღრიცხვა — ყოველდღიური 1C ფაილის დაწერა სააღრიცხვო ფაილში
და ნაშთების შედარება 1C-ის ნაშთების რეპორტთან. დამთხვევა არტიკულით.
გაშვება:  streamlit run app.py
"""
from datetime import datetime

import pandas as pd
import streamlit as st

import engine

st.set_page_config(page_title="საწყობის აღრიცხვა", page_icon="📦", layout="wide")
st.title("📦 საწყობის ნაშთების აღრიცხვა")
st.caption("ატვირთეთ ყოველდღიური 1C რეპორტი (ოპერაციის ასოებით) და სააღრიცხვო ფაილი. "
           "შედარებისთვის დამატებით ატვირთეთ 1C ნაშთების რეპორტი (უწყისი — საქონელი საწყობში).")

with st.expander("ℹ️ ოპერაციის ასოები (ივსება ყოველდღიურ ფაილში, სვეტის სათაურის ზემოთ)"):
    st.markdown(
        "| ასო | ოპერაცია | ეფექტი ნაშთზე |\n|---|---|---|\n"
        "| **მ** | მიღება | + |\n| **დ** | დაბრუნება | + |\n"
        "| **გ** | გაცემა | − |\n| **ჩ** | ჩამოწერა | − |"
    )

c1, c2, c3 = st.columns(3)
daily_up = c1.file_uploader("1️⃣ თეოს ფაილი", type=["xls", "xlsx"])
stock_up = c2.file_uploader("2️⃣ ბაგრათის ფაილი (ნაშთები + ჟურნალი)", type=["xlsx"])
export_up = c3.file_uploader("3️⃣ მაუხრანის ფაილი", type=["xls", "xlsx"])

if not stock_up:
    st.info("გასაგრძელებლად ატვირთეთ სააღრიცხვო ფაილი (და ყოველდღიური ფაილი, თუ დღის დაწერა გინდათ).")
    st.stop()

bal, led, stock_problems = engine.load_stock_file(stock_up.getvalue())
if stock_problems:
    st.error("სააღრიცხვო ფაილის პრობლემები:")
    for p in stock_problems:
        st.write("• " + p)
    st.stop()

# balances after this session's posting (if any) live in session_state
posted_key = "posted_balances"

# ================================================================ posting
if daily_up:
    parsed = engine.parse_daily_file(daily_up.getvalue(), daily_up.name)

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

    dups = engine.find_duplicates(led, tx)
    if dups:
        st.error("ეს დოკუმენტები უკვე დაწერილია ჟურნალში (ფაილი მეორედ იტვირთება?):")
        st.dataframe(pd.DataFrame(dups, columns=["დოკუმენტი", "თარიღი"]),
                     use_container_width=True, hide_index=True)
        st.stop()

    st.subheader("გადასამოწმებელი ინფორმაცია")

    doc_df = pd.DataFrame(parsed["docs"])
    doc_df["ოპერაცია"] = doc_df["key"].map(lambda k: f"{k} — {engine.KEY_MAP[k][0]}")
    doc_df["ეფექტი"] = doc_df["key"].map(lambda k: "+" if engine.KEY_MAP[k][1] > 0 else "−")
    st.markdown("**დოკუმენტები და მონიშნული ოპერაციები**")
    st.dataframe(doc_df[["date", "doc_no", "doc_type", "ოპერაცია", "ეფექტი"]]
                 .rename(columns={"date": "თარიღი", "doc_no": "დოკუმენტი", "doc_type": "1C ტიპი"}),
                 use_container_width=True, hide_index=True)

    summ = (tx.groupby("ოპერაცია")
            .agg(დოკუმენტები=("დოკუმენტი", "nunique"), პოზიციები=("არტიკული", "count"),
                 რაოდენობა=("რაოდენობა", "sum"), წმინდა_ცვლილება=("ცვლილება (+/-)", "sum"))
            .reset_index())
    m1, m2, m3 = st.columns(3)
    m1.metric("ტრანზაქციები", f"{len(tx):,}")
    m2.metric("სულ მოძრაობა (ცალი)", f"{tx['რაოდენობა'].sum():,.0f}")
    m3.metric("წმინდა ცვლილება ნაშთზე", f"{tx['ცვლილება (+/-)'].sum():+,.0f}")
    st.markdown("**ჯამები ოპერაციების მიხედვით** — შეადარეთ მოლოდინს დაწერამდე")
    st.dataframe(summ, use_container_width=True, hide_index=True)

    known = set(bal["არტიკული"])
    new_skus = tx.loc[~tx["არტიკული"].isin(known), ["არტიკული", "ნომენკლატურა"]].drop_duplicates("არტიკული")
    if not new_skus.empty:
        st.warning(f"{len(new_skus)} არტიკული არ არის ნაშთების ფურცელზე — დაემატება ავტომატურად, საწყისი ნაშთით 0:")
        st.dataframe(new_skus, use_container_width=True, hide_index=True)

    with st.expander("ყველა ტრანზაქცია დეტალურად"):
        st.dataframe(tx, use_container_width=True, hide_index=True)

    st.divider()
    if st.button("✅ დაწერე ჟურნალში და ჩამოტვირთე განახლებული ფაილი", type="primary"):
        data, info = engine.post_and_build(bal, led, tx, daily_up.name)
        st.session_state[posted_key] = info["balances"]
        st.success(f"დაიწერა {info['posted_rows']} ჩანაწერი.")

        if not info["negative"].empty:
            st.error(f"⚠️ {len(info['negative'])} SKU-ს ნაშთი უარყოფითია — გადაამოწმეთ საწყისი ნაშთი ან ოპერაციები:")
            st.dataframe(info["negative"][["არტიკული", "ნომენკლატურა", "საწყისი ნაშთი", "მიმდინარე ნაშთი"]],
                         use_container_width=True, hide_index=True)

        with st.expander("განახლებული ნაშთები (მოძრაობა დღეს)"):
            moved = info["balances"][info["balances"]["არტიკული"].isin(tx["არტიკული"])]
            st.dataframe(moved[["არტიკული", "ნომენკლატურა", "საწყისი ნაშთი", "მიმდინარე ნაშთი"]],
                         use_container_width=True, hide_index=True)

        fname = f"სააღრიცხვო_ფაილი_{datetime.now():%Y-%m-%d_%H%M}.xlsx"
        st.download_button("⬇️ განახლებული სააღრიცხვო ფაილი", data=data, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.info("ჩამოტვირთეთ ფაილი და ჩაანაცვლეთ ძველი ვერსია — შემდეგ დღეს ეს ფაილი ატვირთეთ.")

# ================================================================ 1C compare
if export_up:
    st.divider()
    st.subheader("📊 ნაშთების შედარება 1C-სთან")

    export, exp_msgs = engine.parse_1c_stock_export(export_up.getvalue(), export_up.name)
    if export is None:
        for m in exp_msgs:
            st.error(m)
        st.stop()
    for m in exp_msgs:
        st.warning(m)

    # compare against post-posting balances if a posting happened this session,
    # otherwise against the uploaded accounting file as-is
    if posted_key in st.session_state:
        balances = st.session_state[posted_key]
        st.caption("შედარება ხდება ახლახან დაწერილი დღის ჩათვლით.")
    else:
        balances = engine.compute_balances(bal, led)
        st.caption("შედარება ხდება ატვირთული სააღრიცხვო ფაილის მდგომარეობასთან (დღე არ დაწერილა ამ სესიაში).")

    cmp, export_only = engine.compare_with_1c(balances, export)
    diffs = cmp[cmp["სხვაობა (ჩვენი − 1C)"].abs() > 0.001]

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("შედარებული SKU", f"{len(cmp):,}")
    k2.metric("ემთხვევა", f"{len(cmp) - len(diffs):,}")
    k3.metric("სხვაობით", f"{len(diffs):,}")
    k4.metric("მხოლოდ 1C-შია (არ ემატება)", f"{len(export_only):,}")

    only_diff = st.toggle("მხოლოდ სხვაობების ჩვენება", value=True)
    view = diffs if only_diff else cmp
    st.dataframe(view.sort_values("სხვაობა (ჩვენი − 1C)", key=lambda s: s.abs(), ascending=False),
                 use_container_width=True, hide_index=True)

    with st.expander(f"არტიკულები, რომლებიც მხოლოდ 1C რეპორტშია ({len(export_only)}) — ინფორმაციისთვის, არ ემატება"):
        st.dataframe(export_only, use_container_width=True, hide_index=True)

    rep = engine.build_comparison_file(cmp, export_only, export_up.name)
    st.download_button("⬇️ შედარების რეპორტი (xlsx)", data=rep,
                       file_name=f"შედარება_1C_{datetime.now():%Y-%m-%d_%H%M}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

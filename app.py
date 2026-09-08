import streamlit as st
from scraper import run_scraper

st.set_page_config(page_title="India INX Scraper", page_icon="📊")
st.title("India INX Member Scraper")
st.caption("One-click scraping → one combined Excel workbook.")

mode = st.radio("Scraping mode", ["All Members", "Selected Pages"], horizontal=True)

c1, c2 = st.columns(2)
with c1:
    page_from = st.number_input("From page", 1, 9999, 1, disabled=mode=="All Members")
with c2:
    page_to = st.number_input("To page", 1, 9999, 1, disabled=mode=="All Members")

st.info("Chrome runs in headless mode. No Chrome popup will appear.")

if st.button("▶ Run Scraper", type="primary", use_container_width=True):
    if mode == "Selected Pages" and page_from > page_to:
        st.error("From page must be less than or equal to To page.")
    else:
        progress = st.progress(0)
        status = st.empty()
        counts = st.empty()

        def update(done, total, message):
            progress.progress(min(100, int(done / total * 100)) if total else 0)
            status.write(message)
            counts.write(f"Members processed: {done} / {total}")

        try:
            result = run_scraper(
                page_from=int(page_from) if mode == "Selected Pages" else None,
                page_to=int(page_to) if mode == "Selected Pages" else None,
                progress_callback=update,
            )
            progress.progress(100)
            status.success("Completed successfully.")
            counts.write(
                f"Members: {result['members']} | Directors: {result['directors']} | "
                f"Designated: {result['designated']} | Others: {result['others']}"
            )
            with open(result["file"], "rb") as f:
                st.download_button(
                    "⬇ Download Excel",
                    f.read(),
                    file_name=result["filename"],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
        except Exception as e:
            st.error(f"Scraper failed: {e}")

st.caption("Output filename: India_INX_Members_DD-MM-YY.xlsx")

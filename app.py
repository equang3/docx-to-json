import json
import pandas as pd
import streamlit as st

from extractor import docx_bytes_to_paras

st.set_page_config(page_title="DOCX → JSON", layout="wide")
st.title("DOCX → JSON converter")

uploaded_files = st.file_uploader(
    "Upload one or more .docx files",
    type=["docx"],
    accept_multiple_files=True,
)

rows = []

if uploaded_files:
    for f in uploaded_files:
        data = f.getvalue()

        paras = docx_bytes_to_paras(data)
        payload = {"paras": [{"id": i, "text": t} for i, t in enumerate(paras)]}
        # Store as a JSON *string* so each CSV cell contains valid JSON
        rows.append(json.dumps(payload, ensure_ascii=False))

    df = pd.DataFrame({"json": rows})

    st.subheader("Output")
    st.dataframe(df, use_container_width=True)

    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download CSV",
        data=csv_bytes,
        file_name="outputs.csv",
        mime="text/csv",
    )
else:
    st.info("Upload .docx files to generate output.")

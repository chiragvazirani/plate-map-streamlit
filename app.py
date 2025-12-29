import streamlit as st
import tempfile
import os
from plate_map_core import generate_plate_maps

st.set_page_config(
    page_title="Subunit Pooling Plate Map Generator",
    layout="centered"
)

st.title("🧬 Subunit Pooling Plate Map Generator")

st.markdown("""
Upload a **subunit pooling CSV** to generate:
- 📄 Source plate map (PDF)
- 📊 Source plate map (Excel)

Files are processed temporarily and not stored.
""")

uploaded_file = st.file_uploader(
    "Upload subunit pooling CSV",
    type=["csv"]
)

if uploaded_file is not None:
    st.success(f"Uploaded: {uploaded_file.name}")

    if st.button("Generate plate maps"):
        with st.spinner("Generating plate maps…"):
            with tempfile.TemporaryDirectory() as tmp:
                input_path = os.path.join(tmp, uploaded_file.name)

                with open(input_path, "wb") as f:
                    f.write(uploaded_file.getbuffer())

                output_dir = os.path.join(tmp, "output")

                try:
                    pdf_path, xlsx_path = generate_plate_maps(
                        input_path,
                        output_dir
                    )

                    with open(pdf_path, "rb") as f:
                        st.download_button(
                            "⬇️ Download PDF",
                            f,
                            file_name=os.path.basename(pdf_path),
                            mime="application/pdf"
                        )

                    with open(xlsx_path, "rb") as f:
                        st.download_button(
                            "⬇️ Download Excel",
                            f,
                            file_name=os.path.basename(xlsx_path),
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )

                    st.success("Done! You can upload another file.")

                except Exception as e:
                    st.error("Something went wrong.")
                    st.exception(e)

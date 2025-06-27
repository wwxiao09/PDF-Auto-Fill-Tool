import streamlit as st
import pandas as pd
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName
from io import BytesIO
import zipfile

def fill_and_flatten_pdf(template_bytes, data_dict):
    reader = PdfReader(fdata=template_bytes.read())
    for page in reader.pages:
        if page.Annots:
            for annot in page.Annots:
                if annot.Subtype == PdfName.Widget and annot.T:
                    field_name = annot.T.to_unicode().strip("()")
                    if field_name in data_dict:
                        value = str(data_dict[field_name])
                        annot.update(PdfDict(V=value, Ff=1))  # Fill and make read-only
                        annot.update(PdfDict(AP=""))  # Flatten field appearance
    output = BytesIO()
    PdfWriter().write(output, reader)
    output.seek(0)
    return output

st.title("📄 Auto Fill PDF Forms from Excel")

pdf_file = st.file_uploader("Upload Fillable PDF Form", type="pdf")
excel_file = st.file_uploader("Upload Excel File (with matching column names)", type=["xlsx", "csv"])

if pdf_file and excel_file:
    # Read the Excel file
    df = pd.read_excel(excel_file) if excel_file.name.endswith("xlsx") else pd.read_csv(excel_file)

    st.write("🔍 Excel Preview:")
    st.dataframe(df)

    if st.button("Generate PDFs"):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zipf:
            for idx, row in df.iterrows():
                pdf_input = BytesIO(pdf_file.getvalue())
                filled_pdf = fill_and_flatten_pdf(pdf_input, row.to_dict())
                filename = f"{row['name']} New Pay Rate.pdf"
                zipf.writestr(filename, filled_pdf.read())

        zip_buffer.seek(0)
        st.success("✅ PDFs generated!")
        st.download_button(
            label="📥 Download All PDFs as ZIP",
            data=zip_buffer,
            file_name="Filled_Forms.zip",
            mime="application/zip"
        )

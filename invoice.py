import pandas as pd
from docxtpl import DocxTemplate
import streamlit as st
# import io  # For byte conversion
# from spire.doc import *
import subprocess


@st.cache_data
def generate_invoice(input_data, template):
    try:
        container = st.container(border=True)
        if input_data and template:
            df = pd.read_excel(input_data,engine='openpyxl')
            tpl = DocxTemplate(template)
            container.write("Your data is: ")
            container.dataframe(df)
            for i, row in df.iterrows():
                tpl.render(row.to_dict())
                tpl.save(f'invoice_{i}.docx')
                # with io.BytesIO() as buffer:
                #     tpl.save(buffer)
                #     invoice_bytes = buffer.getvalue()

                subprocess.run(["libreoffice", "--convert-to", "pdf", f"invoice_{i}.docx", "--headless", "--convert-to", "pdf"])


                # data = convert(f'invoice_{i}.docx')
                container.download_button(
                label=f"Download Invoice_{i}.pdf",
                data= open(f'invoice_{i}.pdf', 'rb').read(),
                file_name=f'invoice_{i}.pdf'
                )
            container.write('Invoices generated successfully!')
            # container.write('refresh to generate more invoices.')
        else:
            container.write('Please upload the Excel file and the Word template.')
    except Exception as e:
        print(e)
        st.write(f'An error occurred: {e}')

if __name__ == '__main__':

    st.header('Invoice Generator',divider='rainbow')
    st.write('This app generates invoices from an Excel file and a Word template.')
    input_data = st.file_uploader('Upload the Excel file', type='xlsx')
    template = st.file_uploader('Upload the Word template', type='docx')
    # st.write('Click the button below to generate the invoices.')
    # st.button('Generate Invoices', on_click=handleClick)
    generate_invoice(input_data, template)

    

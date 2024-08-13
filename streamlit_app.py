import streamlit as st
import pandas as pd
import statsmodels.api as sm
from io import BytesIO
import xlsxwriter
from fpdf import FPDF

# Upload CSV or Excel file
st.title("Regression Analysis App")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file is not None:
    # Read file
    if uploaded_file.name.endswith('.csv'):
        data = pd.read_csv(uploaded_file)
    else:
        data = pd.read_excel(uploaded_file)

    st.write("Preview of the Data")
    st.write(data.head())

    # Select variables
    columns = data.columns.tolist()
    dependent_var = st.selectbox("Select the dependent variable", columns)
    independent_vars = st.multiselect("Select independent variables", columns)

    if dependent_var and independent_vars:
        # Perform regression
        X = data[independent_vars]
        X = sm.add_constant(X)
        y = data[dependent_var]
        model = sm.OLS(y, X).fit()

        # Show regression summary
        st.write("Regression Summary")
        st.text(model.summary())

        # Function to generate PDF summary
        def generate_pdf_summary(summary_text):
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, summary_text)
            return pdf

        # Function to generate Excel summary
        def generate_excel_summary(summary_text):
            output = BytesIO()
            workbook = xlsxwriter.Workbook(output, {'in_memory': True})
            worksheet = workbook.add_worksheet()
            worksheet.write(0, 0, summary_text)
            workbook.close()
            output.seek(0)
            return output

        # Create downloadable PDF
        summary_text = model.summary().as_text()
        pdf = generate_pdf_summary(summary_text)
        pdf_output = BytesIO()
        pdf.output(pdf_output)
        pdf_output.seek(0)

        st.download_button(
            label="Download Regression Summary as PDF",
            data=pdf_output,
            file_name="regression_summary.pdf",
            mime="application/pdf"
        )

        # Create downloadable Excel
        excel_output = generate_excel_summary(summary_text)
        st.download_button(
            label="Download Regression Summary as Excel",
            data=excel_output,
            file_name="regression_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

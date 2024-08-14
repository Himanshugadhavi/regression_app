import streamlit as st
import pandas as pd
import statsmodels.api as sm
from statsmodels.formula.api import ols
import scikit_posthocs as sp
from io import BytesIO
import xlsxwriter
from fpdf import FPDF

# Function to perform ANOVA
def perform_anova(data, dependent_var, treatment_var):
    formula = f'{dependent_var} ~ C({treatment_var})'
    model = ols(formula, data=data).fit()
    anova_table = sm.stats.anova_lm(model, typ=2)
    return model, anova_table

# Function to perform LSD Test
def perform_lsd(data, treatment_var, dependent_var):
    mc = sm.stats.multicomp.MultiComparison(data[dependent_var], data[treatment_var])
    lsd_result = mc.tukeyhsd()
    return lsd_result.summary()

# Function to perform DMRT Test
def perform_dmrt(data, treatment_var, dependent_var):
    dmrt_result = sp.posthoc_dscf(data, val_col=dependent_var, group_col=treatment_var)
    return dmrt_result

# Function to generate PDF summary
def generate_pdf_summary(anova_table, lsd_result, dmrt_result):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    pdf.multi_cell(0, 10, "ANOVA Summary")
    pdf.multi_cell(0, 10, str(anova_table))
    
    pdf.add_page()
    pdf.multi_cell(0, 10, "LSD Test Summary")
    pdf.multi_cell(0, 10, str(lsd_result))
    
    pdf.add_page()
    pdf.multi_cell(0, 10, "DMRT Test Summary")
    pdf.multi_cell(0, 10, dmrt_result.to_string())
    
    return pdf

# Function to generate Excel summary
def generate_excel_summary(anova_table, lsd_result, dmrt_result):
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet("Summary")
    
    worksheet.write(0, 0, "ANOVA Summary")
    for i, line in enumerate(str(anova_table).split('\n')):
        worksheet.write(i+1, 0, line)
    
    worksheet.write(i+3, 0, "LSD Test Summary")
    for j, line in enumerate(str(lsd_result).split('\n')):
        worksheet.write(i+j+4, 0, line)
    
    worksheet.write(i+j+6, 0, "DMRT Test Summary")
    dmrt_result_as_str = dmrt_result.to_string()
    for k, line in enumerate(dmrt_result_as_str.split('\n')):
        worksheet.write(i+j+k+7, 0, line)
    
    workbook.close()
    output.seek(0)
    return output

# Streamlit app interface
st.title("CRD Analysis App")

uploaded_file = st.file_uploader("Upload CSV or Excel file", type=["csv", "xlsx"])

if uploaded_file is not None:
    if uploaded_file.name.endswith('.csv'):
        data = pd.read_csv(uploaded_file)
    else:
        data = pd.read_excel(uploaded_file)

    st.write("Preview of the Data")
    st.write(data.head())

    # Reshape the data from wide to long format
    data_long = data.melt(var_name='Replication', value_name='Response', ignore_index=False)
    data_long = data_long.reset_index().rename(columns={'index': 'Treatment'})

    st.write("Data in Long Format")
    st.write(data_long.head())

    treatment_var = 'Treatment'
    dependent_var = 'Response'

    if treatment_var and dependent_var:
        # Perform ANOVA
        model, anova_table = perform_anova(data_long, dependent_var, treatment_var)
        st.write("ANOVA Summary")
        st.write(anova_table)
        
        # Perform LSD Test
        lsd_result = perform_lsd(data_long, treatment_var, dependent_var)
        st.write("LSD Test Summary")
        st.text(lsd_result)
        
        # Perform DMRT Test
        dmrt_result = perform_dmrt(data_long, treatment_var, dependent_var)
        st.write("DMRT Test Summary")
        st.write(dmrt_result)

        # Generate PDF and Excel summaries
        pdf = generate_pdf_summary(anova_table, lsd_result, dmrt_result)
        pdf_output = BytesIO()
        pdf.output(pdf_output)
        pdf_output.seek(0)

        excel_output = generate_excel_summary(anova_table, lsd_result, dmrt_result)

        st.download_button(
            label="Download Summary as PDF",
            data=pdf_output,
            file_name="CRD_summary.pdf",
            mime="application/pdf"
        )

        st.download_button(
            label="Download Summary as Excel",
            data=excel_output,
            file_name="CRD_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

import pandas as pd
import numpy as np
import io
import streamlit as st
import openpyxl
import xlsxwriter
from openpyxl import Workbook
from io import BytesIO


def download_excel(dataframes):
    op = BytesIO()
    with pd.ExcelWriter(op, engine='xlsxwriter') as wr:
        for sheet_name, df in dataframes.items():
            df.to_excel(wr, index=False, sheet_name=sheet_name)
            workbook = wr.book
            worksheet = wr.sheets[sheet_name]
            # 1 
            if sheet_name == '1380_Raw':
                fm = workbook.add_format({'num_format': '0.00'})
                worksheet.set_column("A:A", None, fm)
            else:
                fm = workbook.add_format({'bold': True})
                worksheet.set_column("A:A", None, fm)
            if sheet_name == '1370_Raw':
                fm = workbook.add_format({'num_format': '0.00'})
                worksheet.set_column("A:A", None, fm)
            else:
                fm = workbook.add_format({'bold': True})
                worksheet.set_column("A:A", None, fm)
    op.seek(0)
    return op.getvalue()        
# 
def main():
    global raw, error, file_1, file_2
    st.title("Xiaomi Employee Reimbursement")

    # Use columns to organize the buttons and display
    col1, col2, col3 = st.columns(3)

    # Add buttons to the columns
    b1 = col1.button("Click Here for SAP Upload Template")
    b2 = col2.button("Click Here for Bank Upload Template")
    b3 = col3.button("Click Here for Master Upload")

    # Initialize session_state
    if "b1_clicked" not in st.session_state:
        st.session_state.b1_clicked = False

    if "b2_clicked" not in st.session_state:
        st.session_state.b2_clicked = False

    if "b3_clicked" not in st.session_state:
        st.session_state.b3_clicked = False

    # Check if the buttons have been clicked
    if b1:
        st.session_state.b1_clicked = True
        st.session_state.b2_clicked = False
        st.session_state.b3_clicked = False

    if b2:
        st.session_state.b1_clicked = False
        st.session_state.b2_clicked = True
        st.session_state.b3_clicked = False

    if b3:
        st.session_state.b1_clicked = False
        st.session_state.b2_clicked = False
        st.session_state.b3_clicked = True

    # Show different content based on button clicks
    if st.session_state.b1_clicked:
        st.markdown("SAP Upload Template")
        
        uploaded_data = st.file_uploader("Upload Raw File", type=["xlsx"])
        uploaded_master = st.file_uploader("Upload Master File", type=["xlsx"])
        
        if uploaded_data and uploaded_master:
            try:
                
                master = pd.read_excel(uploaded_master, dtype='object')  
                
                raw = pd.read_excel(uploaded_data, sheet_name='Reimbursement_2', dtype='object', skiprows=6)
                
                col1, col2 = st.columns(2)
                
                with col1:
                    st.subheader("Master Data")
                    st.dataframe(master)
                
                with col2:
                    st.subheader("Raw Data")
                    st.dataframe(raw)
                
                raw.dropna(subset=['Policy', 'Payment Status', 'Company Name', 'Company Code'], inplace=True)
                l1 = raw['Bank Location'].unique()
                l1 = list(l1)
                st.write("Unique Bank Locations:", l1)
            except ValueError as e:
                st.error(f"Error: {e}. Please check the sheet name in the Master File.")
                
            error1=raw[raw['Bank Location'].isin(['Nepal', 'nepal'])]
            error1['Remark'] = 'Nepal Case'
            st.dataframe(error1) 
            
            countries_to_delete = ['Nepal', 'nepal']
            raw = raw[~raw['Bank Location'].isin(countries_to_delete)]
            raw['Bank Location'] = 'India'
            
            error2 = raw[raw['IFSC'].str.len() != 11]
            error2['Remark'] = 'IFSC not equal to 11'
            st.dataframe(error2)
            
            raw = raw[~(raw['IFSC'].str.len() != 11)]
            raw['IFSC'] = raw['IFSC'].str[:4].str.upper()+raw['IFSC'].str[4:]
            
            error3 = pd.merge(master, raw, on='Report ID', how='inner')

            app_no_to_delete = error3['Report ID']
            error3 = raw[raw['Report ID'].isin(app_no_to_delete)]
            error3['Remark'] = 'Duplicate in Master'
            st.dataframe(error3)
            
            raw = raw[~raw['Report ID'].isin(app_no_to_delete)]
            Impure = pd.concat([error1,error2,error3],ignore_index=True, axis=0)
            st.dataframe(Impure)
            
            file_1 = raw[raw["Company Code"]== '1370']
            file_2 = raw[raw["Company Code"]== '1380']
            
            st.dataframe(file_1)
            st.dataframe(file_2)
            
        if st.button("Download Excel"):
            dataframes = {"1370_Raw": file_1 , "1380_Raw": file_2 }
            excel_data = download_excel(dataframes)
            st.session_state.excel_data = excel_data

            if excel_data is not None:
                st.download_button("Download Result", data=excel_data, file_name="result.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            else:
                st.warning("Please generate Excel data before downloading.")
                
                
if __name__ == "__main__":
    main()            


        
        
        

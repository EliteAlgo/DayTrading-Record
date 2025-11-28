import streamlit as st
import pandas as pd
import io

def process_data(df):
    # Ensure required columns exist
    required_columns = ['User ID', 'Portfolio Name', 'PNL Per Lot', 'Strategy Tag']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        return None, f"Missing columns: {missing_columns}"

    # Remove duplicates
    # Identify duplicates where User ID, Clean Portfolio Name (without _REX), and PNL Per Lot are identical
    df['CleanName'] = df['Portfolio Name'].astype(str).str.replace(r'_REX\d+$', '', regex=True)
    
    # Drop duplicates based on User ID, CleanName, and PNL Per Lot
    # We keep the first occurrence
    df = df.drop_duplicates(subset=['User ID', 'CleanName', 'PNL Per Lot'])

    # Create Portfolio Group
    df['Portfolio Group'] = df['Portfolio Name'].astype(str).str[:5]

    # Create Pivot Table
    # Rows: Portfolio Group, Strategy Tag
    # Columns: User ID
    # Values: PNL Per Lot
    pivot_df = df.pivot_table(index=['Portfolio Group', 'Strategy Tag'], 
                              columns='User ID', 
                              values='PNL Per Lot', 
                              aggfunc='sum')
    
    # Reset index to make Portfolio Group and Strategy Tag regular columns
    summary = pivot_df.reset_index()

    # Fill NaNs with empty string for display
    summary_display = summary.fillna('')
    
    return summary, summary_display, None

st.set_page_config(page_title="Portfolio Summary Processor", layout="wide")

st.title("Portfolio Summary Processor")

uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    try:
        # Read the Excel file
        # We assume the sheet name is 'Portfolios' based on previous context, 
        # but it's safer to let user choose or try default. 
        # For now, let's try to read 'Portfolios' or the first sheet.
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        
        if 'Portfolios' in sheet_names:
            df = pd.read_excel(uploaded_file, sheet_name='Portfolios')
        else:
            st.warning(f"Sheet 'Portfolios' not found. Using first sheet: {sheet_names[0]}")
            df = pd.read_excel(uploaded_file, sheet_name=sheet_names[0])



        summary, summary_display, error = process_data(df)

        if error:
            st.error(error)
        else:
            st.write("### Processed Summary")
            st.dataframe(summary_display)

            # Convert to Excel for download
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                summary.to_excel(writer, index=False, sheet_name='Summary')
            
            st.download_button(
                label="Download Summary as Excel",
                data=buffer.getvalue(),
                file_name="portfolio_summary.xlsx",
                mime="application/vnd.ms-excel"
            )

    except Exception as e:
        st.error(f"Error processing file: {e}")

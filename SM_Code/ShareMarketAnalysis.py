import streamlit as st
import pandas as pd
from datetime import datetime

def load_data(file):
    return pd.read_excel(file)

def save_data(df, filename):
    # Move 'Update Date' to the last position if it exists in the DataFrame
    if 'Update Date' in df.columns:
        # Reorder the columns to move 'Update Date' to the end
        new_column_order = [col for col in df.columns if col != 'Update Date'] + ['Update Date']
        df = df[new_column_order]
    
    # Sort the DataFrame case-insensitively
    if 'Company Name' in df.columns:
        df.sort_values('Company Name', inplace=True, key=lambda x: x.str.lower())

    # Save the DataFrame to an Excel file
    with pd.ExcelWriter(filename, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)

def main():
    st.title('Share Market Analysis')

    st.sidebar.image('niftylogo.png', width=100)
    uploaded_file = st.sidebar.file_uploader("Choose an Excel file", type=['xlsx'])

    if uploaded_file is not None:
        if 'data' not in st.session_state or st.session_state.uploaded_file != uploaded_file:
            st.session_state.data = load_data(uploaded_file)
            st.session_state.uploaded_file = uploaded_file

        data = st.session_state.data

        for column, default in [
            ('Daily Correction', 'No'), ('55 SMA Status', 'On 55 SMA'),
            ('RSI', 'Below 60'), ('Update Date', pd.NaT),
            ('Lot Size', pd.NA), ('Tradeable', 'No'), ('Comment', '')
        ]:
            if column not in data.columns:
                data[column] = default

        company_list = sorted(data['Company Name'].dropna().unique(), key=lambda x: x.lower())
        selected_company = st.sidebar.selectbox('Select a company to edit', options=company_list, key='selected_company')

        if 'last_selected_company' not in st.session_state or st.session_state.last_selected_company != selected_company:
            st.session_state.temp_comment = ''
            st.session_state.last_selected_company = selected_company
        else:
            if 'update_comment' in st.session_state:
                st.session_state.temp_comment = st.session_state.update_comment

        selected_index = data[data['Company Name'] == selected_company].index[0]

        # UI for displaying company info
        company_name = data.at[selected_index, 'Company Name']
        company_symbol = data.at[selected_index, 'Symbol']
        lot_size = data.at[selected_index, 'Lot Size']
        lot_display = f"Lot Size: {int(lot_size)}" if pd.notna(lot_size) and lot_size != '' else "Not available for Futures and Options trading"
        st.markdown(f"<h1 style='text-align: center;'>{company_name}</h1>", unsafe_allow_html=True)
        st.markdown(f"<h4 style='text-align: center;'>{company_symbol}-{lot_display}</h4>", unsafe_allow_html=True)

        # Input widgets
        with st.container():
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                correction = st.radio("Daily Correction", ['Yes', 'No'], key='correction')
            with col2:
                sma_status = st.radio("55 SMA Status", ['Below 55 SMA', 'On 55 SMA', 'Above 55 SMA'], key='sma_status')
            with col3:
                rsi = st.radio("RSI", ['Below 60', 'Above 60'], key='rsi')
            with col4:
                tradeable = st.radio("Tradeable", ['Yes', 'No'], key='tradeable')
            with col5:
                comment = st.text_input("Comment", value=st.session_state.temp_comment, on_change=None, key='comment')
                st.session_state.update_comment = comment  # Update the comment in session state only

        # Update and save actions
        if st.button('Update Record'):
            data.at[selected_index, 'Daily Correction'] = correction
            data.at[selected_index, '55 SMA Status'] = sma_status
            data.at[selected_index, 'RSI'] = rsi
            data.at[selected_index, 'Tradeable'] = tradeable
            data.at[selected_index, 'Comment'] = comment
            data.at[selected_index, 'Update Date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            st.session_state.data = data
            st.success('Record updated successfully!')

        if st.button('Save to Excel'):
            save_file = f'Updated_data_{datetime.now().strftime("%Y%m%d%H%M%S")}.xlsx'
            save_data(data, save_file)
            st.success('Data saved successfully!')
            with open(save_file, "rb") as file:
                st.download_button(
                    label="Download Updated Excel",
                    data=file,
                    file_name=save_file,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()

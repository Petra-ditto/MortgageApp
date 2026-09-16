import streamlit as st
from Mortgage_App_MVP import generate_amortization_schedule, export_to_excel, calculate_monthly_payment
import datetime
import pandas as pd

st.set_page_config(page_title="Mortgage Calculator", layout="wide")
st.title("🏠 Mortgage Calculator with Prepayments")

# --- User Inputs ---
with st.sidebar:
    st.header("Loan Details")

    # CURRENCY DROP-DOWN
    CCY_options = ['EUR', 'GBP', 'USD', 'HUF', 'CZK', 'NOK', 'SEK', 'DKK'] #need to order alphabetically here if needed

    # Create the selectbox
    selected_CCY = st.selectbox(
        'Currency',  # Label
        CCY_options  # Options
    )

    principal = st.number_input("Mortgage Amount", value=300000, step=1000)
    annual_rate = st.number_input("Annual Interest Rate (%)", value=4.0, step=0.1) / 100
    years = st.number_input("Loan Term (Years)", value=30, step=1)
    start_date = st.date_input("Start Date", value=datetime.date(2025, 9, 1))

    # TRYING OUT DROP-DOWN
    RE_options = ['Apartment', 'House', 'Garden', 'Plot']

    # Create the selectbox
    selected_color = st.selectbox(
        'Real Estate Type',  # Label
        RE_options  # Options
    )

    # END OF DROP-DOWN

    st.header("Prepayments")
    prepayment_entries = {}
    if 'prepayment_count' not in st.session_state:
        st.session_state['prepayment_count'] = 0

    add_prepayment = st.button("Add Prepayment")
    if add_prepayment:
        st.session_state['prepayment_count'] += 1

    for i in range(st.session_state['prepayment_count']):
        col1, col2 = st.sidebar.columns(2)
        with col1:
            prepay_date = st.date_input(f"Date {i+1}", key=f"date_{i}")
        with col2:
            prepay_amount = st.number_input(f"Amount {i+1}", key=f"amt_{i}", min_value=0)
        prepayment_entries[prepay_date.strftime("%Y-%m")] = prepay_amount

    #Delete Prepayment (my try)
    delete_prepayment = st.button("Delete Prepayment")

    if delete_prepayment:
        st.session_state['prepayment_count'] -= 1



    #Run Calculation
    generate_schedule = st.button("Generate Schedule")

# --- Calculate Schedule ---
##if st.button("Generate Schedule"):
if generate_schedule:
    with st.spinner("Calculating..."):
        schedule_df = generate_amortization_schedule(
            principal, annual_rate, years, start_date.strftime("%Y-%m-%d"), prepayment_entries
        )

        st.success("Calculation Complete ✅")

        # Show table
        st.subheader("📅 Amortization Schedule")
        st.dataframe(schedule_df, use_container_width=True)

        # Line chart
        st.subheader("📊 Installment Comparison")
        chart_df = schedule_df[["Payment Date", "Original Installment", "New Installment"]]
        chart_df["Payment Date"] = pd.to_datetime(chart_df["Payment Date"])
        chart_df.set_index("Payment Date", inplace=True)
        st.line_chart(chart_df)

        # Excel export
        st.subheader("📁 Export")
        filename = "Mortgage_Schedule.xlsx"
        export_to_excel(principal, annual_rate, years, start_date.strftime("%Y-%m-%d"), prepayment_entries, schedule_df, filename)

        with open(filename, "rb") as f:
            st.download_button(
                label="Download Excel File",
                data=f,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

import streamlit as st
import pandas as pd
import re
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

st.title("📊 Excel Sheet Consolidator")
st.write("Upload an Excel file with sheets named like 'March 2025'.")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    month_year_pattern = re.compile(
        r"^(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}$",
        re.IGNORECASE
    )

    dfs = []
    for sheet_name in xls.sheet_names:
        if month_year_pattern.match(sheet_name):
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            df["Sheet Name"] = sheet_name
            dfs.append(df)

    if dfs:
        consolidated_df = pd.concat(dfs, ignore_index=True)
        st.success(f"✅ Consolidated {len(dfs)} sheets.")
        st.dataframe(consolidated_df)

        # === INCIDENT TYPE ANALYSIS ===
        if "Incident Type" in consolidated_df.columns:
            st.subheader("📌 Top Incident Types")
            incident_counts = consolidated_df["Incident Type"].value_counts().reset_index()
            incident_counts.columns = ["Incident Type", "Count"]

            st.dataframe(incident_counts)
            st.bar_chart(incident_counts.set_index("Incident Type"))
        else:
            st.warning("⚠️ Column 'Incident Type' not found in the data.")

        # === ⏳ AVERAGE DELAY IN FORWARDING ===
        if "Incident Received by QA on" in consolidated_df.columns and "Date" in consolidated_df.columns:
            # Convert columns to datetime
            consolidated_df["Incident Received by QA on"] = pd.to_datetime(
                consolidated_df["Incident Received by QA on"], errors='coerce'
            )
            consolidated_df["Date"] = pd.to_datetime(consolidated_df["Date"], errors='coerce')

            # Calculate delay in days
            consolidated_df["Forwarding Delay (Days)"] = (
                consolidated_df["Incident Received by QA on"] - consolidated_df["Date"]
            ).dt.days

            # Overall average delay
            avg_delay = consolidated_df["Forwarding Delay (Days)"].mean(skipna=True)
            st.subheader("⏳ Average Delay in Forwarding to QA")
            st.metric(label="Average Delay (Days)", value=round(avg_delay, 2))

            # === MONTH-WISE AVERAGE DELAY ===
            consolidated_df["Month_Year"] = consolidated_df["Date"].dt.to_period("M").dt.to_timestamp()
            monthly_avg_delay = consolidated_df.groupby("Month_Year")["Forwarding Delay (Days)"].mean().reset_index()

            # Plot line chart
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(monthly_avg_delay["Month_Year"], monthly_avg_delay["Forwarding Delay (Days)"], marker="o")
            ax.set_title("Month-wise Average Delay (Days)")
            ax.set_xlabel("Month-Year")
            ax.set_ylabel("Average Delay (Days)")
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
            plt.xticks(rotation=45)
            plt.grid(True)
            st.pyplot(fig)

            # === FORWARDING DELAY TO SHAREHOLDERS ===

        if "Incident forwarded on" in consolidated_df.columns and "Incident Received by QA on" in consolidated_df.columns:
            consolidated_df["Incident forwarded on"] = pd.to_datetime(consolidated_df["Incident forwarded on"], errors='coerce')
            consolidated_df["Incident Received by QA on"] = pd.to_datetime(consolidated_df["Incident Received by QA on"], errors='coerce')

            consolidated_df["Forwarding Delay to Shareholders (Days)"] = (
                consolidated_df["Incident forwarded on"] - consolidated_df["Incident Received by QA on"]
            ).dt.days

            avg_shareholder_delay = consolidated_df["Forwarding Delay to Shareholders (Days)"].mean(skipna=True)
            st.subheader("⏳ Average Delay in Forwarding to Shareholders")
            st.metric(label="Average Delay (Days)", value=round(avg_shareholder_delay, 2))

            consolidated_df["Month_Year"] = consolidated_df["Incident Received by QA on"].dt.to_period("M").dt.to_timestamp()
            monthly_shareholder_delay = consolidated_df.groupby("Month_Year")["Forwarding Delay to Shareholders (Days)"].mean().reset_index()

            st.subheader("📈 Month-wise Average Forwarding Delay to Shareholders")
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.plot(monthly_shareholder_delay["Month_Year"], monthly_shareholder_delay["Forwarding Delay to Shareholders (Days)"], marker="o", color="orange")
            ax.set_title("Month-wise Average Delay to Shareholders")
            ax.set_xlabel("Month-Year")
            ax.set_ylabel("Average Delay (Days)")
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%b %Y'))
            plt.xticks(rotation=45)
            st.pyplot(fig)

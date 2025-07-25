import pandas as pd
import numpy as np
import streamlit as st
import io
import re
import altair as alt

# Initialize session state variables if they don't exist
if "use_demo" not in st.session_state:
    st.session_state.use_demo = False
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None
if "uploaded_data" not in st.session_state:
    st.session_state.uploaded_data = {}
if "uploaded_df_tab1" not in st.session_state:
    st.session_state.uploaded_df_tab1 = None
if "generated_graph" not in st.session_state:
    st.session_state.generated_graph = None

st.set_page_config(layout="wide")

# --------- HEADER UI STYLING ---------
st.markdown("""
<style>
.center { display: flex; justify-content: center; text-align: center; }
</style>
""", unsafe_allow_html=True)

st.markdown("<h2 class='center' style='color:rgb(70, 130, 255);'>An EsteStyle Streamlit Page<br>Where Python Wiz Meets Data Viz!</h2>", unsafe_allow_html=True)
st.markdown("<img src='https://1drv.ms/i/s!ArWyPNkF5S-foZspwsary83MhqEWiA?embed=1&width=307&height=307' width='300' style='display: block; margin: 0 auto;'>", unsafe_allow_html=True)
st.markdown("<h3 class='center' style='color: rgb(135, 206, 250);'>🏨 Originally created for Best Western at Firestone 🛎️</h3>", unsafe_allow_html=True)
st.markdown("<h3 class='center' style='color: rgb(135, 206, 250);'>🤖 By Esteban C Loetz 📟</h3>", unsafe_allow_html=True)
st.header("")

def parse_money(val):
                    return float(str(val).replace('$', '').replace(',', '').strip())

# --- Function to convert DataFrames to an in-memory Excel file ---
# Using st.cache_data to prevent re-generating the file
# on every minor interaction if the data hasn't changed.
@st.cache_data
def to_excel_bytes(dfs_dict):
    output = io.BytesIO()
    # Use pandas ExcelWriter to write multiple sheets
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for sheet_name, df_to_write in dfs_dict.items():
            # index=False prevents pandas from writing the DataFrame index as a column
            df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
    processed_data = output.getvalue()
    return processed_data

st.markdown("""
    <style>
        div[data-testid="stTabs"] button {
            display: flex;
            justify-content: center;
            flex-grow: 1;
        }
    </style>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["Single Room Pop Excel File Analysis", "Multi Room Pop Excel File Analysis by Date"])

with tab1:

    col1, col2, col3 = st.columns([1, 7, 1], gap="large")
    with col2:

        st.header("Single Room Pop Excel File Analysis")
        st.markdown("---")  # Just a divider line for UX
        # --------- FILE UPLOADER ---------
        st.subheader("📥 Download file for analysis:")
        st.write('')

        # Handle file upload
        uploaded_file = st.file_uploader("📄 Select Visual Matrix Output Excel File to Analyze", type=["xlsx", "xls"])

        # Handle demo file button
        use_demo = st.button("📂 Use Demo File")

        # Logic to handle file selection
        if use_demo:
            # Reset uploaded file state
            st.session_state.uploaded_file = None
            st.session_state.use_demo = True
            st.session_state.generated_graph = None  # Reset the graph

            # Load demo file
            demo_data_path = "2024-02 Room_Type_Popularity.xls"
            df = pd.read_excel(demo_data_path, sheet_name="Sheet1")
            st.session_state.uploaded_df_tab1 = df
            st.success("✅ Demo file loaded!")

        elif uploaded_file is not None:
            # Reset demo file state
            st.session_state.use_demo = False
            st.session_state.uploaded_file = uploaded_file
            st.session_state.generated_graph = None  # Reset the graph
            
            # Load uploaded file
            df = pd.read_excel(uploaded_file, sheet_name="Sheet1")
            st.session_state.uploaded_df_tab1 = df
            st.success("✅ Excel file uploaded and stored!")

        # Ensure only one file source is active
        if st.session_state.uploaded_df_tab1 is not None:
            st.info(
                "Using the demo file." if st.session_state.use_demo
                else "Using the uploaded file. Selecting the demo file will reset this option."
            )
            

            if st.button("📊 Generate Graphs"):
                # Reset the Excel bytes in session state
                st.session_state.excel_bytes = None  # Clear the previous Excel file
            
                df = st.session_state.uploaded_df_tab1

                cell_value = df.iloc[21, 25]  # Get cell value

                # Convert to int safely (handling NaN)
                if pd.isna(cell_value):  # Checks if it's NaN
                    cell_value = 0  # Replace NaN with a default value
                else:
                    cell_value = int(cell_value)  # Convert float to integer

                room_totals = {
                    "KH": df.iloc[21, 9],  # Example: Row 5, Column 2
                    "K": df.iloc[36, 9],  
                    "Q": df.iloc[50, 9], 
                    "QH": df.iloc[53, 9],
                    "QQ": df.iloc[73, 9],
                    "SQ": df.iloc[86, 9]  
                }

                # List of room types and their corresponding rental counts
                data_for_rent_totals_chart = {
                    "Room Type": list(room_totals.keys()),
                    "Total Rentals": list(room_totals.values())
                }
                #print(f'rent totals {rent_totals}')

                # Now, convert this into a pandas DataFrame
                rent_tot_chart_df = pd.DataFrame(data_for_rent_totals_chart)

                pct_totals = {
                    "KH": df.iloc[21, 25],  # Example: Row 5, Column 2
                    "K": df.iloc[36, 25],  
                    "Q": df.iloc[50, 25], 
                    "QH": df.iloc[53, 25],
                    "QQ": df.iloc[73, 25],
                    "SQ": df.iloc[86, 25]  
                }

                data_for_rent_totals_chart = {
                    "Room Type": list(room_totals.keys()),
                    "Room Percents": list(pct_totals.values())
                }
                #print(f'rent totals {rent_totals}')

                # Convert into a pandas DataFrame
                room_pct_chart_df = pd.DataFrame(data_for_rent_totals_chart)



                rev_totals = {
                    "KH": parse_money(df.iloc[21, 29]),  # Example: Row 5, Column 2
                    "K": parse_money(df.iloc[36, 29]),  
                    "Q": parse_money(df.iloc[50, 29]), 
                    "QH": parse_money(df.iloc[53, 29]),
                    "QQ": parse_money(df.iloc[73, 29]),
                    "SQ": parse_money(df.iloc[86, 29])  
                }

                data_for_rev_totals_chart = {
                    "Room Type": list(room_totals.keys()),
                    "Rev Totals": list(rev_totals.values())
                }

                # Convert into a pandas DataFrame
                rev_totals_chart_df = pd.DataFrame(data_for_rev_totals_chart)



                adr_totals = {
                    "KH": parse_money(df.iloc[21, 34]),  # Example: Row 5, Column 2
                    "K": parse_money(df.iloc[36, 34]),  
                    "Q": parse_money(df.iloc[50, 34]), 
                    "QH": parse_money(df.iloc[53, 34]),
                    "QQ": parse_money(df.iloc[73, 34]),
                    "SQ": parse_money(df.iloc[86, 34])  
                }

                data_for_adr_totals_chart = {
                    "Room Type": list(room_totals.keys()),
                    "ADR Totals": list(adr_totals.values())
                }

                # Convert into a pandas DataFrame
                adr_totals_chart_df = pd.DataFrame(data_for_adr_totals_chart)

                totals = {
                    "rm tot": parse_money(df.iloc[87, 9]),  # Example: Row 5, Column 2 
                    "% tot": parse_money(df.iloc[87, 25]), 
                    "rev tot": parse_money(df.iloc[87, 29]),
                    "adr tot": parse_money(df.iloc[87, 34]) 
                }

                data_for_total_totals_chart = {
                    "Category": list(totals.keys()),
                    "Totals": list(totals.values())
                }

                # Convert this dictionary directly into a pandas DataFrame
                total_totals_df = pd.DataFrame(data_for_total_totals_chart)

                # Set 'Room Type' as the index if you prefer, then just add columns
                combined_df = pd.DataFrame.from_dict(room_totals, orient='index', columns=['Total Rentals'])
                combined_df['Room Percents'] = pd.Series(pct_totals)
                combined_df['Total Revenue'] = pd.Series(rev_totals)
                combined_df['ADR Totals'] = pd.Series(adr_totals)
                combined_df.index.name = "Room Type" # Give the index a name if it's not a regular column
                combined_df = combined_df.reset_index() # If you want "Room Type" to be a regular column again

                col_left, col_right = st.columns([1, 1],  gap="large")

                # Display the bar chart
                with col_left:
                    
                    st.write("### Room Rentals by Type")
                    st.bar_chart(rent_tot_chart_df, x="Room Type", y="Total Rentals")

                    st.write("### Room Percents by Type")
                    st.bar_chart(room_pct_chart_df, x="Room Type", y="Room Percents")

                    # --- Displaying the Combined Table ---
                    st.write("### All Room Data at a Glance")
                    st.dataframe(combined_df)

                with col_right:
                    
                    st.write("### Revenue Totals by Type")
                    st.bar_chart(rev_totals_chart_df, x="Room Type", y="Rev Totals")

                    # Display the bar chart using Streamlit!
                    st.write("### Average Daily Rate by Type")
                    st.bar_chart(adr_totals_chart_df, x="Room Type", y="ADR Totals")

                    # --- Displaying the Table! ---
                    st.write("### Summary Totals Table")
                    st.dataframe(total_totals_df)

                # --- Prepare the DataFrames for export ---
                # Create a dictionary where keys are the desired sheet names and values are the DataFrames
                dataframes_to_export = {
                    'Combined Room Data': combined_df,
                    'Overall Totals Summary': total_totals_df # Changed sheet name for clarity
                }

                # Generate the Excel file in memory
                excel_data_bytes = to_excel_bytes(dataframes_to_export)

                # --- Display the Download Button ---
                st.download_button(
                    label="⬇️ Download Summary Data as Excel", # Text displayed on the button
                    data=excel_data_bytes, # The actual bytes of the Excel file
                    file_name="Room_Analysis_Report.xlsx", # The name of the file when downloaded
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" # The MIME type for .xlsx files
                )

                st.info("Click the button above to download your analysis data in a single Excel file!")

with tab2:

    col1, col2, col3 = st.columns([1, 3, 1])

    with col2:
        st.header("Multi Excel File Room Pop Analysis by Date")
        st.markdown("---")  # Just a divider line for UX
        # --------- FILE UPLOADER ---------
        st.subheader("📥 Download multiple files for analysis:")
        st.write('Minimum of 2 files must be loaded & file names must contain either "YYYY-MM" or "YYYY-MM-DD"')

        # Add a unique key for the file uploader
        file_uploader_key = "file_uploader_default"

        use_demo = st.button("📂 Use Demo Files")

        if use_demo:
            # Reset uploaded files state
            st.session_state.uploaded_data = {}  # Clear any previously uploaded files
            st.session_state.use_demo = True  # Mark demo mode as active
            file_uploader_key = "file_uploader_reset"  # Change the key to reset the file uploader
            successful_uploads = []  # List to store successfully uploaded filenames

            # Load demo files
            demo_files = {
                "2024-01 Room_Type_Popularity.xls": "2024-01 Room_Type_Popularity.xls",
                "2024-02 Room_Type_Popularity.xls": "2024-02 Room_Type_Popularity.xls",
                "2024-03 Room_Type_Popularity.xls": "2024-03 Room_Type_Popularity.xls",
                "2025-01 Room_Type_Popularity.xls": "2025-01 Room_Type_Popularity.xls",
                "2025-02 Room_Type_Popularity.xls": "2025-02 Room_Type_Popularity.xls",
                "2025-03 Room_Type_Popularity.xls": "2025-03 Room_Type_Popularity.xls",
            }
            for filename, path in demo_files.items():
                date_match_ymd = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
                date_match_ym = re.search(r"(\d{4}-\d{2})", filename)
                if date_match_ymd:
                    extracted_date = date_match_ymd.group(1)
                elif date_match_ym:
                    extracted_date = date_match_ym.group(1)
                else:
                    extracted_date = None
                    st.warning(f"⚠️ Could not extract date from filename: {filename} (demo file).")
                    continue
                try:
                    df = pd.read_excel(path)
                    st.session_state.uploaded_data[filename] = {"df": df, "date": extracted_date}
                    successful_uploads.append(filename)  # Add to the list of successful uploads
                except Exception as e:
                    st.error(f"🚨 Error reading demo file '{filename}': {e}")

            # Display a single success message with the total count of uploaded files
            if successful_uploads:
                success_message = f"✅ **{len(successful_uploads)} Demo files successfully uploaded:**\n\n" + "\n".join(f"- {file}" for file in successful_uploads)
                st.markdown(success_message)

        # File uploader with dynamic key
        uploaded_files = st.file_uploader(
            "📄 Select Multiple Visual Matrix output Excel files to analyze",
            type="xls",
            accept_multiple_files=True,
            key=file_uploader_key,  # Use the dynamic key
        )

        if uploaded_files:
            # Reset demo file state
            st.session_state.use_demo = False  # Mark demo mode as inactive
            st.session_state.uploaded_data = {}  # Clear any previously loaded demo files

            successful_uploads = []  # List to store successfully uploaded filenames

            for uploaded_file in uploaded_files:
                filename = uploaded_file.name
                date_match_ymd = re.search(r"(\d{4}-\d{2}-\d{2})", filename)
                date_match_ym = re.search(r"(\d{4}-\d{2})", filename)

                if date_match_ymd:
                    extracted_date = date_match_ymd.group(1)
                elif date_match_ym:
                    extracted_date = date_match_ym.group(1)
                else:
                    extracted_date = None
                    st.warning(f"⚠️ Could not extract date from filename: {filename}. This file will not be used for the time-based graph.")
                    continue

                try:
                    df = pd.read_excel(uploaded_file).reset_index(drop=True)
                    st.session_state.uploaded_data[filename] = {"df": df, "date": extracted_date}
                    successful_uploads.append(filename)  # Add to the list of successful uploads
                except Exception as e:
                    st.error(f"🚨 Error reading Excel file '{filename}': {e}")

            # Display a single success message with the total count of uploaded files
            if successful_uploads:
                success_message = f"✅ **{len(successful_uploads)} files successfully uploaded:**\n\n" + "\n".join(f"- {file}" for file in successful_uploads)
                st.markdown(success_message)

        # --------- UI OPTIONS ---------
        with col2:
            if st.session_state.uploaded_data:
                if st.button("📊 Generate Graphs"):
                    if len(st.session_state.uploaded_data) < 2:
                        st.warning("Please upload a minimum of 2 files to enable comparison charts.")
                    else:
                        st.markdown("---")
                        st.subheader("📈 Time-Based Comparison Graphs")

                        all_room_totals_data = []
                        all_pct_totals_data = []
                        all_rev_totals_data = []
                        all_adr_totals_data = []

                        sorted_uploaded_data = sorted(
                            st.session_state.uploaded_data.items(),
                            key=lambda item: pd.to_datetime(item[1]['date']) if item[1]['date'] else pd.Timestamp.min
                        )

                        for filename, file_info in sorted_uploaded_data:
                            current_df = file_info["df"]
                            extracted_date_str = file_info["date"]

                            if extracted_date_str:
                                current_date = pd.to_datetime(extracted_date_str)

                                # Re-extract the individual data points from each file
                                room_totals = {
                                    "KH": current_df.iloc[21, 9],
                                    "K": current_df.iloc[36, 9],
                                    "Q": current_df.iloc[50, 9],
                                    "QH": current_df.iloc[53, 9],
                                    "QQ": current_df.iloc[73, 9],
                                    "SQ": current_df.iloc[86, 9]
                                }
                                pct_totals = {
                                    "KH": (current_df.iloc[21, 25]),
                                    "K": (current_df.iloc[36, 25]),
                                    "Q": (current_df.iloc[50, 25]),
                                    "QH": (current_df.iloc[53, 25]),
                                    "QQ": (current_df.iloc[73, 25]),
                                    "SQ": (current_df.iloc[86, 25])
                                }
                                rev_totals = {
                                    "KH": parse_money(current_df.iloc[21, 29]),
                                    "K": parse_money(current_df.iloc[36, 29]),
                                    "Q": parse_money(current_df.iloc[50, 29]),
                                    "QH": parse_money(current_df.iloc[53, 29]),
                                    "QQ": parse_money(current_df.iloc[73, 29]),
                                    "SQ": parse_money(current_df.iloc[86, 29])
                                }
                                adr_totals = {
                                    "KH": parse_money(current_df.iloc[21, 34]),
                                    "K": parse_money(current_df.iloc[36, 34]),
                                    "Q": parse_money(current_df.iloc[50, 34]),
                                    "QH": parse_money(current_df.iloc[53, 34]),
                                    "QQ": parse_money(current_df.iloc[73, 34]),
                                    "SQ": parse_money(current_df.iloc[86, 34])
                                }


                                # Append data for each metric in a "long" format
                                for room_type, value in room_totals.items():
                                    all_room_totals_data.append({"Date": current_date, "Room Type": room_type, "Total Rentals": value})

                                for room_type, value in pct_totals.items():
                                    all_pct_totals_data.append({"Date": current_date, "Room Type": room_type, "Room Percents": value})

                                for room_type, value in rev_totals.items():
                                    all_rev_totals_data.append({"Date": current_date, "Room Type": room_type, "Total Revenue": value})

                                for room_type, value in adr_totals.items():
                                    all_adr_totals_data.append({"Date": current_date, "Room Type": room_type, "ADR": value})

                        # Create DataFrames for plotting
                        df_room_totals_trends = pd.DataFrame(all_room_totals_data)
                        df_pct_trends = pd.DataFrame(all_pct_totals_data)
                        df_rev_trends = pd.DataFrame(all_rev_totals_data)
                        df_adr_trends = pd.DataFrame(all_adr_totals_data)

                        # --- Generate the grouped bar charts using Altair ---
                        # We need to explicitly define the grouping for side-by-side bars

                        if not df_room_totals_trends.empty:
                            df_room_totals_trends["Date_str"] = df_room_totals_trends["Date"].dt.strftime("%Y-%m")
                            st.write("#### Total Rentals by Room Type Across Dates")
                            chart = alt.Chart(df_room_totals_trends).mark_bar().encode(
                                # Primary X-axis: Room Type
                                x=alt.X('Room Type:N', axis=alt.Axis(title="Room Type")),
                                # Offset bars within each Room Type group by Date
                                xOffset=alt.XOffset('Date:N'),
                                # Y-axis: The metric value
                                y=alt.Y('Total Rentals:Q', axis=alt.Axis(title="Total Rentals")),
                                # Color bars by Date to distinguish time points
                                color=alt.Color('Date_str:N', legend=alt.Legend(title="Date")),
                                tooltip=['Room Type', alt.Tooltip('Date_str:N', title="Date"), 'Total Rentals']
                            ).properties(
                                title='Total Rentals by Room Type Over Time'
                            ).interactive() # Allows zooming and panning
                            st.altair_chart(chart, use_container_width=True)
                        else:
                            st.info("No data available to plot Total Rentals trends.")


                        if not df_pct_trends.empty:
                            df_pct_trends["Date_str"] = df_pct_trends["Date"].dt.strftime("%Y-%m")
                            st.write("#### Room Percents by Room Type Across Dates")
                            chart = alt.Chart(df_pct_trends).mark_bar().encode(
                                x=alt.X('Room Type:N', axis=alt.Axis(title="Room Type")),
                                xOffset=alt.XOffset('Date:N'),
                                y=alt.Y('Room Percents:Q', axis=alt.Axis(title="Room Percents")),
                                color=alt.Color('Date_str:N', legend=alt.Legend(title="Date")),
                                tooltip=['Room Type', alt.Tooltip('Date_str:N', title="Date"), alt.Tooltip('Room Percents', format=".1%")] # Format as percentage
                            ).properties(
                                title='Room Percentages by Room Type Over Time'
                            ).interactive()
                            st.altair_chart(chart, use_container_width=True)
                        else:
                            st.info("No data available to plot Room Percents trends.")


                        if not df_rev_trends.empty:
                            df_rev_trends["Date_str"] = df_rev_trends["Date"].dt.strftime("%Y-%m")
                            st.write("#### Total Revenue by Room Type Across Dates")
                            chart = alt.Chart(df_rev_trends).mark_bar().encode(
                                x=alt.X('Room Type:N', axis=alt.Axis(title="Room Type")),
                                xOffset=alt.XOffset('Date:N'),
                                y=alt.Y('Total Revenue:Q', axis=alt.Axis(title="Total Revenue")),
                                color=alt.Color('Date_str:N', legend=alt.Legend(title="Date")),
                                tooltip=['Room Type', alt.Tooltip('Date_str:N', title="Date"), alt.Tooltip('Total Revenue', format="$,.2f")] # Format as currency
                            ).properties(
                                title='Total Revenue by Room Type Over Time'
                            ).interactive()
                            st.altair_chart(chart, use_container_width=True)
                        else:
                            st.info("No data available to plot Total Revenue trends.")


                        if not df_adr_trends.empty:
                            df_adr_trends["Date_str"] = df_adr_trends["Date"].dt.strftime("%Y-%m")
                            st.write("#### Average Daily Rate (ADR) by Room Type Across Dates")
                            chart = alt.Chart(df_adr_trends).mark_bar().encode(
                                x=alt.X('Room Type:N', axis=alt.Axis(title="Room Type")),
                                xOffset=alt.XOffset('Date:N'),
                                y=alt.Y('ADR:Q', axis=alt.Axis(title="ADR")),
                                color=alt.Color('Date_str:N', legend=alt.Legend(title="Date")),
                                tooltip=['Room Type', alt.Tooltip('Date_str:N', title="Date"), alt.Tooltip('ADR', format="$,.2f")] # Format as currency
                            ).properties(
                                title='ADR by Room Type Over Time'
                            ).interactive()
                            st.altair_chart(chart, use_container_width=True)
                        else:
                            st.info("No data available to plot ADR trends.")

                        # Display and Download Summary Data for Multi-File Analysis
                        # Only show download button if trend DataFrames have data
                        if not df_room_totals_trends.empty:
                            st.markdown("---")
                            st.subheader("⬇️ Download All Trends Data")

                            # Optional: Display the DataFrames for user to preview
                            with st.expander("🔍 Click to view detailed trend data tables"):
                                st.write("#### Total Rentals Trend Data")
                                # Create a display-friendly version of the DataFrame for this specific view
                                # Select 'Date_str', 'Room Type', and 'Total Rentals'
                                # Then rename 'Date_str' to 'Date' for better readability
                                df_display_rentals = df_room_totals_trends[['Date_str', 'Room Type', 'Total Rentals']].rename(columns={'Date_str': 'Date'})
                                st.dataframe(df_display_rentals, use_container_width=True)

                                st.write("#### Room Percent Trend Data")
                                # Do the same for the percents DataFrame
                                df_display_pct = df_pct_trends[['Date_str', 'Room Type', 'Room Percents']].rename(columns={'Date_str': 'Date'})
                                st.dataframe(df_display_pct, use_container_width=True)

                                st.write("#### Total Revenue Trend Data")
                                # And for the revenue DataFrame
                                df_display_rev = df_rev_trends[['Date_str', 'Room Type', 'Total Revenue']].rename(columns={'Date_str': 'Date'})
                                st.dataframe(df_display_rev, use_container_width=True)

                                st.write("#### Average Daily Rate (ADR) Trend Data")
                                # Finally, for the ADR DataFrame
                                df_display_adr = df_adr_trends[['Date_str', 'Room Type', 'ADR']].rename(columns={'Date_str': 'Date'})
                                st.dataframe(df_display_adr, use_container_width=True)

                            # Prepare the DataFrames for export
                            dataframes_to_export_multi_file = {
                                'Room Rentals Trends': df_room_totals_trends[['Date_str', 'Room Type', 'Total Rentals']].rename(columns={'Date_str': 'Date'}),
                                'Room Percent Trends': df_pct_trends[['Date_str', 'Room Type', 'Room Percents']].rename(columns={'Date_str': 'Date'}),
                                'Revenue Totals Trends': df_rev_trends[['Date_str', 'Room Type', 'Total Revenue']].rename(columns={'Date_str': 'Date'}),
                                'ADR Totals Trends': df_adr_trends[['Date_str', 'Room Type', 'ADR']].rename(columns={'Date_str': 'Date'})
                            }
                            excel_data_bytes_multi_file = to_excel_bytes(dataframes_to_export_multi_file)

                            st.download_button(
                                label="⬇️ Download Multi-File Trends as Excel",
                                data=excel_data_bytes_multi_file,
                                file_name="Multi_File_Room_Trends_Report.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            st.info("Click the button above to download a single Excel file with all trend data on separate sheets!")
                        else:
                            st.info("Upload more files to generate downloadable trend data.")

                elif not st.session_state.uploaded_data and not st.session_state.use_demo:
                    st.info("Please upload Excel files or use the demo files to see the analysis.")

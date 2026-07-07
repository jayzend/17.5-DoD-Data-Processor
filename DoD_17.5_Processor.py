import pandas as pd
import numpy as np
import openpyxl
import streamlit as st
import io

# --- 1. Set up the Web App UI ---
st.set_page_config(page_title="Battery Data Processor", layout="centered")
st.title("17.5% DoD Battery Data Processor")
st.write("Upload your raw test data to generate the processed Excel file and charts.")

# Create the file upload box on the web page
uploaded_file = st.file_uploader("Select Raw Data Excel File", type=["xlsx"])

# --- 2. Process the Data (Only runs if a file is uploaded) ---
if uploaded_file is not None:
    
    # Show a loading spinner while the math runs in the background
    with st.spinner("Processing data... This may take a minute for large files!"):
        try:
            # Extract Battery Name from B4
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            sheet = wb.active
            battery_name = str(sheet['B4'].value)  
            wb.close()

            # IMPORTANT: Reset the file reader back to the top so Pandas can read it!
            uploaded_file.seek(0)

            # Load and prepare Sheet 1 Data
            raw_df = pd.read_excel(uploaded_file, skiprows=29)

            df = raw_df.copy()
            df.rename(columns={df.columns[1]: 'Status'}, inplace=True)

            df['Voltage per Cell'] = df['Voltage'] / 6
            df['Current per Cell'] = df['Current'] / 6
            df['SoC'] = 0.00

            # Calculate SoC row-by-row
            for i in range(len(df)):
                if i == 0:
                    first_ah = 0 if pd.isna(df.loc[i, 'AhStep']) else df.loc[i, 'AhStep']
                    if df.loc[i, 'Status'] == 'CHA':
                        df.loc[i, 'SoC'] = (60 + first_ah) / 60
                    else:
                        df.loc[i, 'SoC'] = (60 - first_ah) / 60
                else:
                    previous_SoC = df.loc[i-1, 'SoC']
                    ahstep = df.loc[i, 'AhStep']
                    ahstep_before = df.loc[i-1, 'AhStep']
                    diff_ahstep = ahstep - ahstep_before

                    if pd.isna(ahstep) or pd.isna(ahstep_before) or ahstep == ahstep_before:
                        diff_ahstep = 0

                    if df.loc[i, 'Status'] == 'CHA':
                        df.loc[i, 'SoC'] = previous_SoC + (diff_ahstep / 60)    
                    else:
                        df.loc[i, 'SoC'] = previous_SoC - (diff_ahstep / 60)

            df['SoC%'] =  df['SoC'] / 1.0 * 100

            start_discharge = (df['SoC%'] < 100).idxmax()
            full_charge = (df.loc[start_discharge:, 'SoC%'] >= 100).idxmax()

            if full_charge > start_discharge:
                df = df.loc[:full_charge]

            # Prepare Sheet 2 Data
            df2 = raw_df.copy()
            df2.rename(columns={df2.columns[1]: 'Status'}, inplace=True)

            df2 = df2[df2['Status'] == 'DCH']
            df2 = df2[df2['Step'].isin([3,7])]
            df2 = df2[(df2['Step Time'] >= '00:30:00.000') & (df2['Step Time']<= '00:30:00.1')]
            df2 = df2.dropna(subset=['Temperature_1'])

            df2 = df2[df2['Cycle'] != df2['Cycle'].shift()]
            df2['Total Cycle'] = range(0, len(df2))

            # --- 3. Create the Output File in Memory ---
            output_buffer = io.BytesIO()

            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Sheet 1', index=False)
                df2.to_excel(writer, sheet_name='Sheet 2', index= False)

                workbook = writer.book
                worksheet1 = writer.sheets['Sheet 1']
                worksheet2 = writer.sheets['Sheet 2']   

                # CHART 1
                chart1 = workbook.add_chart({'type': 'line'})
                chart1.add_series({
                    'categories': ['Sheet 1', 1, df.columns.get_loc('Prog Time'), len(df), df.columns.get_loc('Prog Time')],
                    'values':     ['Sheet 1', 1, df.columns.get_loc('Voltage per Cell'), len(df), df.columns.get_loc('Voltage per Cell')],
                    'line' : {'width': 0.01, 'color': 'blue'}
                })
                chart1.set_title({'name': f'{battery_name} - Voltage vs Time'})
                chart1.set_x_axis({'name': 'Prog Time', 'label_position': 'low'})
                chart1.set_y_axis({'name': 'Voltage per Cell', 'min': 1.5, 'max': 3})
                chart1.set_size({'width': 1500, 'height': 600})
                worksheet1.insert_chart('V2', chart1)

                # CHART 2
                chart2 = workbook.add_chart({'type': 'line'})
                chart2.add_series({
                    'categories': ['Sheet 1', 1, df.columns.get_loc('Prog Time'), len(df), df.columns.get_loc('Prog Time')],
                    'values':     ['Sheet 1', 1, df.columns.get_loc('Current per Cell'), len(df), df.columns.get_loc('Current per Cell')],
                    'line' : {'width': 0.01, 'color': 'blue'}
                })
                chart2.set_title({'name': f'{battery_name} - Current vs Time'})
                chart2.set_x_axis({'name': 'Prog Time', 'label_position': 'low'})
                chart2.set_y_axis({'name': 'Current per Cell', 'min': -4, 'max': 4})
                chart2.set_size({'width': 1500, 'height': 600})
                worksheet1.insert_chart('V33', chart2)

                # CHART 3
                chart3 = workbook.add_chart({'type': 'line'})  
                chart3.add_series({
                    'categories': ['Sheet 1', 1, df.columns.get_loc('Prog Time'), len(df), df.columns.get_loc('Prog Time')],
                    'values':     ['Sheet 1', 1, df.columns.get_loc('SoC%'), len(df), df.columns.get_loc('SoC%')],
                    'line' : {'width': 0.01, 'color': 'blue'}
                })
                chart3.set_title({'name': f'{battery_name} - State of Charge (SoC%) vs Time'})
                chart3.set_x_axis({'name': 'Prog Time', 'label_position': 'low'})
                chart3.set_y_axis({'name': 'SoC%', 'min': 0, 'max': 100})
                chart3.set_size({'width': 1500, 'height': 600})
                worksheet1.insert_chart('V64', chart3)

                # CHART 4
                chart4 = workbook.add_chart({'type': 'scatter', 'subtype': 'straight_with_markers'})
                chart4.add_series({
                    'categories': ['Sheet 2', 1, df2.columns.get_loc('Total Cycle'), len(df2), df2.columns.get_loc('Total Cycle')],
                    'values':     ['Sheet 2', 1, df2.columns.get_loc('Voltage'), len(df2), df2.columns.get_loc('Voltage')],
                    'line': {'width': 0.01, 'color': 'blue'}
                })
                chart4.set_title({'name': f'17.5% DoD of {battery_name}'})  
                chart4.set_x_axis({'name': 'Cycle', 'min': 0, 'label_position': 'low', 'major_unit': 85})
                chart4.set_y_axis({'name': 'Voltage', 'min': 10, 'max': 12.5})
                chart4.set_size({'width': 1500, 'height': 600})
                worksheet2.insert_chart('S2', chart4)

            # --- 4. Provide the Download Button ---
            st.success("✅ File successfully processed!")
            
            st.download_button(
                label=f"📥 Download Processed Battery Results of {battery_name}.xlsx",
                data=output_buffer.getvalue(),
                file_name=f'Processed Battery Results of {battery_name}.xlsx',
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"An error occurred while processing the file: {e}")
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import io
import numpy as np
import os

st.set_page_config(page_title="Account Data Formatter", page_icon="📊", layout="wide")

st.title("📊 Account Data Formatter")
st.write("Upload a file with TID, TX_DATE, TXN AMT, MDR AMT, GST columns to transform it into the required format.")

# Load store.xlsx file for TID to siteid mapping
store_lookup = {}
store_lookup_numeric = {}  # Also create numeric lookup for better matching
store_file_path = None

# Try to find store file - check multiple possible locations
workspace_path = r"c:\Users\APL41051\c drive progrm\accounts"
possible_dirs = [
    workspace_path,
    os.path.dirname(os.path.abspath(__file__)) if '__file__' in globals() else None,
    os.getcwd()
]
possible_dirs = [d for d in possible_dirs if d and os.path.exists(d)]

possible_store_files = ['STORE.xlsx']

for current_dir in possible_dirs:
    for store_file in possible_store_files:
        store_path = os.path.join(current_dir, store_file)
        if os.path.exists(store_path):
            store_file_path = store_path
            break
    if store_file_path:
        break

if store_file_path:
    try:
        store_df = pd.read_excel(store_file_path)
        # Find TID NUMBER and Pharmacy Code columns (case-insensitive)
        # Updated column names: TID NUMBER and Pharmacy Code
        tid_col = None
        pharmacy_code_col = None
        
        for col in store_df.columns:
            col_upper = str(col).upper().strip()
            # Look for "TID NUMBER" (new column name)
            if col_upper in ['TID NUMBER', 'TIDNUMBER', 'TID_NUMBER', 'TID-NUMBER']:
                tid_col = col
            # Also check for old "TID" for backward compatibility
            elif tid_col is None and col_upper == 'TID':
                tid_col = col
            # Look for "Pharmacy Code" (new column name)
            elif col_upper in ['PHARMACY CODE', 'PHARMACYCODE', 'PHARMACY_CODE', 'PHARMACY-CODE']:
                pharmacy_code_col = col
            # Fallback to old "siteid" for backward compatibility
            elif pharmacy_code_col is None and col_upper in ['SITEID', 'SITE ID', 'SITE_ID', 'SITE-ID']:
                pharmacy_code_col = col
        
        if tid_col and pharmacy_code_col:
            # Create lookup dictionary: TID NUMBER -> Pharmacy Code
            # Remove any NaN values
            store_df_clean = store_df[[tid_col, pharmacy_code_col]].dropna(subset=[tid_col, pharmacy_code_col])
            
            # Create both string and numeric lookups for better matching
            for _, row in store_df_clean.iterrows():
                tid_val = row[tid_col]
                pharmacy_code_val = row[pharmacy_code_col]
                
                # Store as string
                store_lookup[str(tid_val).strip()] = str(pharmacy_code_val).strip()
                
                # Also store as numeric if possible
                try:
                    tid_num = float(tid_val)
                    store_lookup_numeric[tid_num] = str(pharmacy_code_val).strip()
                except:
                    pass
            
            st.info(f"✅ Store file loaded: {len(store_lookup)} TID mappings found from {os.path.basename(store_file_path)}")
        else:
            missing_cols = []
            if not tid_col:
                missing_cols.append('TID NUMBER (or TID)')
            if not pharmacy_code_col:
                missing_cols.append('Pharmacy Code (or siteid)')
            st.warning(f"⚠️ Store file found but missing required columns ({', '.join(missing_cols)}). Available columns: {', '.join(store_df.columns.tolist())}")
    except Exception as e:
        st.warning(f"⚠️ Could not load store file: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
else:
    st.warning("⚠️ Store file (STORE.xlsx) not found. Offset Store Code/CC will be empty.")

# File uploader - explicitly allow Excel and CSV files
uploaded_file = st.file_uploader(
    "Upload CSV or Excel file", 
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=False,
    help="Supported formats: CSV (.csv), Excel (.xlsx, .xls)"
)

if uploaded_file is not None:
    try:
        # Validate file type
        file_extension = uploaded_file.name.split('.')[-1].lower()
        allowed_extensions = ['csv', 'xlsx', 'xls']
        
        if file_extension not in allowed_extensions:
            st.error(f"❌ Unsupported file type: .{file_extension}. Please upload a CSV or Excel file (.csv, .xlsx, or .xls)")
        else:
            # Read the file based on extension
            if file_extension == 'csv':
                df = pd.read_csv(uploaded_file)
            else:
                # For Excel files, try to read with openpyxl engine
                try:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                except Exception as e:
                    # Fallback to default engine
                    st.warning(f"⚠️ Trying alternative Excel reader: {str(e)}")
                    df = pd.read_excel(uploaded_file)
            
            st.success(f"✅ File loaded successfully! {len(df)} rows found.")
            
            # Display original data preview
            with st.expander("📋 Preview Original Data"):
                st.dataframe(df.head(10))
            
            # Check if required columns exist
            required_columns = ['TID', 'TX_DATE', 'TXN AMT', 'MDR AMT','NET AMT','GST']
            missing_columns = [col for col in required_columns if col not in df.columns]
            
            if missing_columns:
                st.error(f"❌ Missing required columns: {', '.join(missing_columns)}")
                st.info(f"Available columns: {', '.join(df.columns.tolist())}")
            else:
                # Set default account numbers (hidden from frontend)
                credit_account_number = "126011200"
                debit_account_number = "762021010"
                additional_columns = []
                
                # Process the data
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Initialize the output dataframe with all required columns
                # Base columns structure - Account number in Account column, code in empty column
                base_columns = [
                    'Voucher Date', 'Account Type', 'Account', '', 'Debit', 'Credit'
                ]
                
                # Insert additional columns after 'Credit' if specified
                if additional_columns:
                    base_columns.extend(additional_columns)
                
                # Continue with remaining columns
                remaining_columns = [
                    'Offset account type', 'Offset account', 'Currency code', 'Offset Location',
                    'Offset Store Code/CC', 'Chennal', 'Line Narration', 'Note Description',
                    'Invoice No.', 'Document Date', 'Due Date', 'Exchange Rate',
                    'Payment Referance', 'Payment method', 'HSN Code', 'SAC CODE',
                    'Exempt', 'ITC Category', 'TDS Group', 'Posting Profile',
                    'Assessable Value', 'Adjustment Tax_CGST', 'Adjustment Tax_SGST',
                    'Adjustment Tax_IGST', 'TDS Amount'
                ]
                
                output_columns = base_columns + remaining_columns
                
                # Collect data for grouping - store by Offset Store Code/CC or TID
                grouped_data = {}  # Key: (offset_code, account_type), Value: {amounts, dates, etc}
                matches_found = 0
                matches_not_found = 0
                
                for idx, row in df.iterrows():
                    # Parse TX_DATE and add 1 day
                    try:
                        # Try different date formats
                        tx_date = pd.to_datetime(row['TX_DATE'], errors='coerce')
                        if pd.isna(tx_date):
                            # Try common date formats
                            for fmt in ['%d-%m-%Y', '%Y-%m-%d', '%d/%m/%Y', '%Y/%m/%d']:
                                try:
                                    tx_date = datetime.strptime(str(row['TX_DATE']), fmt)
                                    break
                                except:
                                    continue
                        
                        if pd.isna(tx_date):
                            st.warning(f"⚠️ Could not parse date in row {idx + 1}: {row['TX_DATE']}")
                            voucher_date = ""
                            original_date_str = str(row['TX_DATE'])
                        else:
                            # Add 1 day
                            voucher_date = (tx_date + timedelta(days=1)).strftime('%d-%m-%Y')
                            original_date_str = tx_date.strftime('%d-%m-%Y')
                    except Exception as e:
                        st.warning(f"⚠️ Date parsing error in row {idx + 1}: {str(e)}")
                        voucher_date = ""
                        original_date_str = str(row['TX_DATE'])
                    
                    # Get MDR AMT, TXN AMT, and GST
                    try:
                        mdr_amt = float(row['MDR AMT']) if pd.notna(row['MDR AMT']) else 0
                    except:
                        mdr_amt = 0
                    
                    try:
                        txn_amt = float(row['TXN AMT']) if pd.notna(row['TXN AMT']) else 0
                    except:
                        txn_amt = 0
                    
                    try:
                        gst_amt = float(row['GST']) if pd.notna(row['GST']) else 0
                    except:
                        gst_amt = 0
                    
                    # Calculate debit amount: MDR AMT + GST
                    debit_amount = mdr_amt + gst_amt
                    
                    # Get siteid from store lookup using TID
                    tid_value = row['TID']
                    siteid_value = ''
                    
                    if pd.notna(tid_value):
                        # Try multiple matching strategies
                        tid_str = str(tid_value).strip()
                        
                        # First try string match
                        if tid_str in store_lookup:
                            siteid_value = store_lookup[tid_str]
                            matches_found += 1
                        else:
                            # Try numeric match
                            try:
                                tid_num = float(tid_value)
                                if tid_num in store_lookup_numeric:
                                    siteid_value = store_lookup_numeric[tid_num]
                                    matches_found += 1
                            except:
                                pass
                            
                            # If still not found, try without leading/trailing spaces and case variations
                            if not siteid_value:
                                for key, value in store_lookup.items():
                                    if str(key).strip() == tid_str or str(key).strip() == tid_str.strip():
                                        siteid_value = value
                                        matches_found += 1
                                        break
                                
                                if not siteid_value:
                                    matches_not_found += 1
                                    # Show first few unmatched TIDs for debugging
                                    if matches_not_found <= 3:
                                        st.caption(f"⚠️ TID '{tid_str}' not found in store file")
                    else:
                        matches_not_found += 1
                    
                    # Use Offset Store Code/CC for grouping, fallback to TID if empty
                    group_key = siteid_value if siteid_value else str(tid_value) if pd.notna(tid_value) else 'UNKNOWN'
                    
                    # Group credit rows (CCCA)
                    if txn_amt > 0:
                        credit_key = (group_key, 'Credit Card Control Account')
                        if credit_key not in grouped_data:
                            grouped_data[credit_key] = {
                                'account_number': credit_account_number,
                                'account_code': 'Credit Card Control Account',
                                'debit': 0,
                                'credit': 0,
                                'offset_code': siteid_value,
                                'voucher_date': voucher_date,
                                'original_date': original_date_str,
                                'tid': tid_value
                            }
                        grouped_data[credit_key]['credit'] += txn_amt
                    
                    # Group debit rows (BCCA) - Debit = MDR AMT + GST
                    if debit_amount > 0:
                        debit_key = (group_key, 'Bank Charges - Credit Card')
                        if debit_key not in grouped_data:
                            grouped_data[debit_key] = {
                                'account_number': debit_account_number,
                                'account_code': 'Bank Charges - Credit Card',
                                'debit': 0,
                                'credit': 0,
                                'offset_code': siteid_value,
                                'voucher_date': voucher_date,
                                'original_date': original_date_str,
                                'tid': tid_value
                            }
                        grouped_data[debit_key]['debit'] += debit_amount
                    
                    # Update progress
                    if len(df) > 100:
                        progress = (idx + 1) / len(df)
                        progress_bar.progress(progress)
                        status_text.text(f"Processing row {idx + 1} of {len(df)}...")
                
                # Helper function to create a base row
                def create_base_row(account_number, account_code, debit_val, credit_val, offset_code, voucher_date, original_date_str):
                    row = {
                        'Voucher Date': voucher_date,
                        'Account Type': 'Ledger',
                        'Account': account_number,  # Account number only
                        '': account_code,  # Account code in empty column (no header)
                        'Debit': debit_val if debit_val else '',
                        'Credit': credit_val if credit_val else '',
                    }
                    
                    # Add empty values for additional columns if they exist
                    for col in additional_columns:
                        row[col] = ''
                    
                    # Add remaining columns
                    narration_text = (
                        f"BEING CREDIT CARD SALES {original_date_str} COLLECTION "
                        f"FROM SBI BANK-1585 FOR THE DT: {voucher_date}"
                    )

                    row.update({
                        'Offset account type': '',
                        'Offset account': '',
                        'Currency code': 'INR',
                        'Offset Location': '2X221',
                        'Offset Store Code/CC': offset_code,
                        'Chennal': 'offline',
                        'Line Narration': narration_text,
                        'Note Description': narration_text,
                        'Invoice No.': '',
                        'Document Date': '',
                        'Due Date': '',
                        'Exchange Rate': '',
                        'Payment Referance': '',
                        'Payment method': '',
                        'HSN Code': '',
                        'SAC CODE': '',
                        'Exempt': '',
                        'ITC Category': '',
                        'TDS Group': '',
                        'Posting Profile': '',
                        'Assessable Value': '',
                        'Adjustment Tax_CGST': '',
                        'Adjustment Tax_SGST': '',
                        'Adjustment Tax_IGST': '',
                        'TDS Amount': ''
                    })
                    
                    return row
                
                # Create output rows from grouped data
                # Display mappings to match the example image
                display_account_map = {
                    'Credit Card Control Account': '126011200',
                    'Bank Charges - Credit Card': '762021010'
                }
                display_account_code_map = {
                    'Credit Card Control Account': 'Credit Card Control Account',
                    'Bank Charges - Credit Card': 'Bank Charges - Credit Card',
                    'IUB': 'IUB'
                }
                
                # Create output rows from grouped data
                # First combine credit and debit totals per group_key so we can consolidate
                combined_by_group = {}
                for (group_key, account_type), data in grouped_data.items():
                    if group_key not in combined_by_group:
                        combined_by_group[group_key] = {
                            'credit': 0,
                            'debit': 0,
                            'offset_code': data.get('offset_code', ''),
                            'voucher_date': data.get('voucher_date', ''),
                            'original_date': data.get('original_date', '')
                        }
                    combined_by_group[group_key]['credit'] += data.get('credit', 0)
                    combined_by_group[group_key]['debit'] += data.get('debit', 0)

                output_data = []
                for group_key, agg in combined_by_group.items():
                    offset_code = str(agg.get('offset_code', '')).strip()

                    # If offset_code contains alphabetic characters, produce a single consolidated row
                    if any(c.isalpha() for c in offset_code):
                        net_amount = agg.get('credit', 0) - agg.get('debit', 0)
                        # Keep two decimals (currency style)
                        net_amount = round(net_amount, 2)

                        consolidated_row = create_base_row(
                            '115011060',
                            'Inter Unit Balance : APL TamilNadu',
                            '',
                            net_amount,
                            offset_code,
                            agg.get('voucher_date', ''),
                            agg.get('original_date', '')
                        )

                        # Ensure short display as in image
                        consolidated_row['Account'] = '115011060'
                        consolidated_row[''] = 'Inter Unit Balance : APL TamilNadu'

                        # Override narration specifically for IUB consolidated rows
                        try:
                            voucher_dt = agg.get('voucher_date', '')
                            iub_narration = f"BEING CREDIT CARD COLLECTION REC SBI - 1585 BANK TRF TO Inter Unit Balance : APL TamilNadu FOR THE DT:{voucher_dt}"
                            consolidated_row['Line Narration'] = iub_narration
                            consolidated_row['Note Description'] = iub_narration
                        except Exception:
                            pass

                        output_data.append(consolidated_row)
                    else:
                        # For non-alphabetic offset codes, create separate rows for credit and debit as before
                        # Credit row
                        if agg.get('credit', 0) > 0:
                            credit_row = create_base_row(
                                display_account_map.get('Credit Card Control Account', credit_account_number),
                                'Credit Card Control Account',
                                '',
                                agg.get('credit', 0),
                                offset_code,
                                agg.get('voucher_date', ''),
                                agg.get('original_date', '')
                            )
                            credit_row['Account'] = display_account_map.get('Credit Card Control Account', credit_row['Account'])
                            credit_row[''] = display_account_code_map.get('Credit Card Control Account', credit_row[''])
                            output_data.append(credit_row)

                        # Debit row (bank charges)
                        if agg.get('debit', 0) > 0:
                            debit_row = create_base_row(
                                display_account_map.get('Bank Charges - Credit Card', debit_account_number),
                                'Bank Charges - Credit Card',
                                agg.get('debit', 0),
                                '',
                                offset_code,
                                agg.get('voucher_date', ''),
                                agg.get('original_date', '')
                            )
                            debit_row['Account'] = display_account_map.get('Bank Charges - Credit Card', debit_row['Account'])
                            debit_row[''] = display_account_code_map.get('Bank Charges - Credit Card', debit_row[''])
                            output_data.append(debit_row)

                # Create output dataframe
                output_df = pd.DataFrame(output_data, columns=output_columns)

                # Ensure Debit/Credit columns are uniformly formatted strings to avoid
                # pyarrow/Streamlit serialization issues when types are mixed.
                def fmt_amount(x):
                    if x is None or x == '' or (isinstance(x, float) and np.isnan(x)):
                        return ''
                    try:
                        val = float(x)
                        return f"{val:.2f}"
                    except:
                        return str(x)

                if 'Debit' in output_df.columns:
                    output_df['Debit'] = output_df['Debit'].apply(fmt_amount)
                if 'Credit' in output_df.columns:
                    output_df['Credit'] = output_df['Credit'].apply(fmt_amount)

                # Append final summary row only for Bank account type, using NET AMNT (fallback to TXN AMT)
                def _parse_num(val):
                    try:
                        if pd.isna(val):
                            return 0.0
                        s = str(val).strip()
                        # Remove common thousands separators and currency symbols
                        s = s.replace(',', '').replace('₹', '').replace('$', '')
                        # Handle parentheses for negatives
                        if s.startswith('(') and s.endswith(')'):
                            s = '-' + s[1:-1]
                        return float(s) if s not in ['', 'nan', 'None'] else 0.0
                    except Exception:
                        return 0.0

                net_sum = 0.0
                try:
                    if 'NET AMNT' in df.columns:
                        net_sum = df['NET AMNT'].apply(_parse_num).sum()
                    elif 'NET_AMNT' in df.columns:
                        net_sum = df['NET_AMNT'].apply(_parse_num).sum()
                    elif 'TXN AMT' in df.columns:
                        # Fallback: use per-row (TXN AMT - MDR AMT-GST) to compute net per your requirement
                        def _row_net(r):
                            txn = _parse_num(r.get('TXN AMT', r.get('TXN_AMT', 0)))
                            mdr = _parse_num(r.get('MDR AMT', r.get('MDR_AMT', 0)))
                            gst = _parse_num(r.get('GST', r.get('GST', 0)))
                            return txn - mdr - gst

                        net_sum = df.apply(_row_net, axis=1).sum()
                    else:
                        net_sum = 0.0
                except Exception:
                    net_sum = 0.0

                # Append summary row whenever net_sum is non-zero (ensure last row is always added)
                if net_sum and net_sum != 0:
                    summary_row = {col: '' for col in output_columns}

                    # Voucher date: use last parsable TX_DATE +1 day if available
                    voucher_date_val = ''
                    try:
                        if 'TX_DATE' in df.columns:
                            td = pd.to_datetime(df['TX_DATE'], errors='coerce')
                            if not td.isna().all():
                                voucher_date_val = (td.dropna().iloc[-1] + timedelta(days=1)).strftime('%d-%m-%Y')
                    except Exception:
                        voucher_date_val = ''

                    summary_row['Voucher Date'] = voucher_date_val
                    summary_row['Account Type'] = 'BANK'
                    # Set requested account and account code
                    summary_row['Account'] = 'KASBI1585'
                    summary_row[''] = 'SBI BANK -  39632211585'
                    # Debit is the net_sum (formatted)
                    summary_row['Debit'] = fmt_amount(net_sum)
                    summary_row['Credit'] = ''
                    summary_row['Currency code'] = 'INR'
                    summary_row['Offset Location'] = '2X221'
                    summary_row['Chennal'] = 'offline'
                    # Set Offset Store Code/CC to requested value
                    if 'Offset Store Code/CC' in summary_row:
                        summary_row['Offset Store Code/CC'] = '9009'
                    else:
                        summary_row['Offset Store Code/CC'] = '9009'

                    # Narration: mirror the same narration style used in create_base_row
                    narration_text = (
                        f"BEING CREDIT CARD SALES {voucher_date_val} COLLECTION "
                        f"FROM SBI BANK-1585 FOR THE DT: {voucher_date_val}"
                    )
                    summary_row['Line Narration'] = narration_text
                    summary_row['Note Description'] = narration_text

                    output_df = pd.concat([output_df, pd.DataFrame([summary_row])], ignore_index=True)

                # Clear progress indicators
                progress_bar.empty()
                status_text.empty()

                st.success(f"✅ Transformation complete! {len(output_df)} rows processed.")

                # Show matching statistics
                if store_lookup:
                    st.info(f"📊 TID Matching: {matches_found} matches found, {matches_not_found} not found")

                # Display preview
                with st.expander("📋 Preview Transformed Data"):
                    st.dataframe(output_df.head(10))
                    # Show Offset Store Code/CC column specifically
                    if 'Offset Store Code/CC' in output_df.columns:
                        st.caption("Sample Offset Store Code/CC values:")
                        sample_values = output_df['Offset Store Code/CC'].head(10).tolist()
                        st.write(sample_values)

                # Download button
                output_buffer = io.BytesIO()

                # Save as Excel (better for this format)
                with pd.ExcelWriter(output_buffer, engine='openpyxl') as writer:
                    output_df.to_excel(writer, index=False, sheet_name='Formatted Data')

                output_buffer.seek(0)

                st.download_button(
                    label="📥 Download Formatted Excel File",
                    data=output_buffer,
                    file_name=f"formatted_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                # Also provide CSV option
                csv_buffer = io.StringIO()
                output_df.to_csv(csv_buffer, index=False)
                csv_data = csv_buffer.getvalue().encode('utf-8')

                st.download_button(
                    label="📥 Download Formatted CSV File",
                    data=csv_data,
                    file_name=f"formatted_data_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            
    except Exception as e:
        st.error(f"❌ Error processing file: {str(e)}")
        st.exception(e)

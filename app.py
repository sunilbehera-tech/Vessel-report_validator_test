import streamlit as st
import pandas as pd
import numpy as np
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# NOTE: This version uses standard SMTP (no OAuth dependencies)
# For OAuth support, install: google-auth-oauthlib, google-api-python-client, pywin32

@st.cache_data
def calculate_report_hours_from_data(start_dates, end_dates, start_times, end_times, time_shifts):
    """Calculate Report Hours from Start Date/Time, End Date/Time and Time Shift"""
    report_hours = []
    
    for idx in range(len(start_dates)):
        try:
            start_date = pd.to_datetime(start_dates[idx], errors='coerce')
            end_date = pd.to_datetime(end_dates[idx], errors='coerce')
            
            start_time = str(start_times[idx] if idx < len(start_times) else "00:00:00").strip()
            end_time = str(end_times[idx] if idx < len(end_times) else "00:00:00").strip()
            
            time_shift = time_shifts[idx] if idx < len(time_shifts) else 0
            if pd.isna(time_shift):
                time_shift = 0
            else:
                time_shift = float(time_shift)
            
            if pd.notna(start_date) and pd.notna(end_date):
                try:
                    start_time_obj = pd.to_datetime(start_time, format='%H:%M:%S').time()
                except:
                    try:
                        start_time_obj = pd.to_datetime(start_time, format='%H:%M').time()
                    except:
                        start_time_obj = datetime.strptime("00:00:00", '%H:%M:%S').time()
                
                try:
                    end_time_obj = pd.to_datetime(end_time, format='%H:%M:%S').time()
                except:
                    try:
                        end_time_obj = pd.to_datetime(end_time, format='%H:%M').time()
                    except:
                        end_time_obj = datetime.strptime("00:00:00", '%H:%M:%S').time()
                
                start_datetime = datetime.combine(start_date.date(), start_time_obj)
                end_datetime = datetime.combine(end_date.date(), end_time_obj)
                
                time_diff = end_datetime - start_datetime
                hours_diff = time_diff.total_seconds() / 3600
                total_hours = hours_diff + time_shift
                
                report_hours.append(round(total_hours, 2))
            else:
                report_hours.append(0)
                
        except:
            report_hours.append(0)
    
    return report_hours

def calculate_report_hours(df):
    """Calculate Report Hours from Start Date/Time, End Date/Time and Time Shift"""
    df = df.copy()
    start_dates = df.get("Start Date", pd.Series([None]*len(df))).tolist()
    end_dates = df.get("End Date", pd.Series([None]*len(df))).tolist()
    start_times = df.get("Start Time", pd.Series(["00:00:00"]*len(df))).tolist()
    end_times = df.get("End Time", pd.Series(["00:00:00"]*len(df))).tolist()
    time_shifts = df.get("Time Shift", pd.Series([0]*len(df))).tolist()
    
    return calculate_report_hours_from_data(
        tuple(start_dates), tuple(end_dates), tuple(start_times), tuple(end_times), tuple(time_shifts)
    )

@st.cache_data(show_spinner=False, hash_funcs={pd.DataFrame: lambda x: x.to_json()})
def validate_reports(df):
    """Validate ship reports with all 6 rules including SCOC"""
    df = df.copy()
    
    numeric_cols = [
        "Average Load [kW]", "ME Rhrs (From Last Report)", "Avg. Speed",
        "Fuel Cons. [MT] (ME Cons 1)", "Fuel Cons. [MT] (ME Cons 2)", "Fuel Cons. [MT] (ME Cons 3)",
        "Time Shift", "Average Load [%]",
        "A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)", "A.E. 2 Last Report [Rhrs] (Aux Engine Unit 2)",
        "A.E. 3 Last Report [Rhrs] (Aux Engine Unit 3)", "A.E. 4 Total [Rhrs] (Aux Engine Unit 4)",
        "A.E. 5 Last Report [Rhrs] (Aux Engine Unit 5)", "A.E. 6 Last Report [Rhrs] (Aux Engine Unit 6)",
        "Tank Cleaning [MT]", "Cargo Transfer [MT]", "Maintaining Cargo Temp. [MT]",
        "Shaft Gen. Propulsion [MT]", "Raising Cargo Temp. [MT]", "Burning Sludge [MT]",
        "Ballast Transfer [MT]", "Fresh Water Prod. [MT]", "Others [MT]", "EGCS Consumption [MT]",
        "Cyl. Oil Cons. [Ltrs]"
    ]
    
    for col in numeric_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace(",", "").str.strip().replace(["", "nan", "None"], np.nan)
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Report Hours"] = calculate_report_hours(df)
    
    df["SFOC"] = (
        (df["Fuel Cons. [MT] (ME Cons 1)"] + df["Fuel Cons. [MT] (ME Cons 2)"] + df["Fuel Cons. [MT] (ME Cons 3)"]) * 1_000_000
        / (df["Average Load [kW]"].replace(0, np.nan) * df["ME Rhrs (From Last Report)"].replace(0, np.nan))
    ).fillna(0)

    df["SCOC"] = (
        df["Cyl. Oil Cons. [Ltrs]"] * 1000
        / (df["Average Load [kW]"].replace(0, np.nan) * df["ME Rhrs (From Last Report)"].replace(0, np.nan))
    ).fillna(0)

    reasons = []
    fail_columns = set()

    for idx, row in df.iterrows():
        reason = []
        report_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0)
        report_hours = row.get("Report Hours", 0)
        sfoc = row.get("SFOC", 0)
        scoc = row.get("SCOC", 0)
        avg_speed = row.get("Avg. Speed", 0)

        # Rule 1: SFOC
        if report_type == "At Sea" and ME_Rhrs > 12 and not (150 <= sfoc <= 200):
            reason.append("SFOC out of 150-200 at sea with ME Rhrs > 12")
            fail_columns.add("SFOC")

        # Rule 2: Avg Speed
        if report_type == "At Sea" and ME_Rhrs > 12 and not (0 <= avg_speed <= 20):
            reason.append("Avg. Speed out of 0-20 at sea with ME Rhrs > 12")
            fail_columns.add("Avg. Speed")

        # Rule 3: Exhaust Temp deviation
        if report_type == "At Sea" and ME_Rhrs > 12:
            exhaust_cols = [f"Exh. Temp [°C] (Main Engine Unit {j})" for j in range(1, 17) if f"Exh. Temp [°C] (Main Engine Unit {j})" in df.columns]
            temps = [row[c] for c in exhaust_cols if pd.notna(row[c]) and row[c] != 0]
            if temps:
                avg_temp = np.mean(temps)
                for j, c in enumerate(exhaust_cols, start=1):
                    val = row[c]
                    if pd.notna(val) and val != 0 and abs(val - avg_temp) > 50:
                        reason.append(f"Exhaust temp deviation > ±50 from avg at Unit {j}")
                        fail_columns.add(c)

        # Rule 4: ME Rhrs vs Report Hours
        if report_hours > 0 and (ME_Rhrs - report_hours) > 1.0:
            hours_diff = ME_Rhrs - report_hours
            reason.append(f"ME Rhrs ({ME_Rhrs:.2f}) exceeds Report Hours ({report_hours:.2f}) by {hours_diff:.2f}h")
            fail_columns.update(["ME Rhrs (From Last Report)", "Report Hours"])

        # Rule 5: Aux Engines & Sub-Consumers
        if report_type == "At Sea" and row.get("Average Load [%]", 0) > 40:
            ae_rhrs_sum = sum([row.get(f"A.E. {i} Last Report [Rhrs] (Aux Engine Unit {i})", 0) for i in range(1, 4)])
            ae_rhrs_sum += sum([row.get(f"A.E. {i} {'Total' if i==4 else 'Last Report'} [Rhrs] (Aux Engine Unit {i})", 0) for i in range(4, 7)])
            
            ae_ratio = ae_rhrs_sum / report_hours if report_hours > 0 else 0
            sub_consumers_sum = sum([row.get(f"{c} [MT]", 0) for c in ["Tank Cleaning", "Cargo Transfer", "Maintaining Cargo Temp.", "Shaft Gen. Propulsion", "Raising Cargo Temp.", "Burning Sludge", "Ballast Transfer", "Fresh Water Prod.", "Others", "EGCS Consumption"]])
            
            if ae_ratio > 1.25 and sub_consumers_sum == 0:
                reason.append(f"Multiple AEs (ratio={ae_ratio:.2f}) with ME Load>40% but no sub-consumers")
                fail_columns.update(["Average Load [%]", "A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)", "Tank Cleaning [MT]"])

        # Rule 6: SCOC
        if report_type == "At Sea" and ME_Rhrs > 12 and scoc > 0:
            if scoc < 0.8:
                reason.append(f"SCOC ({scoc:.2f} g/kWh) below normal range (0.8-1.5)")
                fail_columns.update(["SCOC", "Cyl. Oil Cons. [Ltrs]"])
            elif scoc > 1.5:
                reason.append(f"SCOC ({scoc:.2f} g/kWh) above normal range (0.8-1.5)")
                fail_columns.update(["SCOC", "Cyl. Oil Cons. [Ltrs]"])

        reasons.append("; ".join(reason))

    df["Reason"] = reasons
    failed = df[df["Reason"] != ""].copy()

    exhaust_cols = [f"Exh. Temp [°C] (Main Engine Unit {j})" for j in range(1, 17) if f"Exh. Temp [°C] (Main Engine Unit {j})" in df.columns]
    context_cols = ["Ship Name", "IMO_No", "Report Type", "Start Date", "Start Time", "End Date", "End Time", "Voyage Number", "Time Zone", "Distance - Ground [NM]", "Time Shift", "Distance - Sea [NM]", "Average Load [kW]", "Average RPM", "Average Load [%]", "ME Rhrs (From Last Report)", "Report Hours", "Cyl. Oil Cons. [Ltrs]", "SCOC"]

    cols_to_keep = context_cols + exhaust_cols + list(fail_columns) + ["Reason"]
    cols_to_keep_unique = []
    seen = set()
    for col in cols_to_keep:
        if col not in seen and col in failed.columns:
            seen.add(col)
            cols_to_keep_unique.append(col)
    
    if "Ship Name" in cols_to_keep_unique:
        cols_to_keep_unique.remove("Ship Name")
        cols_to_keep_unique = ["Ship Name"] + cols_to_keep_unique

    return failed[cols_to_keep_unique] if not failed.empty else failed, df


def send_email(smtp_server, smtp_port, sender_email, sender_password, recipient_emails, subject, body, attachment_data=None, attachment_name="Failed_Validation.xlsx", cc_emails=None):
    """Send email via SMTP"""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        
        recipient_list = [e.strip() for e in recipient_emails.split(',')] if isinstance(recipient_emails, str) else recipient_emails
        msg['To'] = ', '.join(recipient_list)
        
        cc_list = []
        if cc_emails:
            cc_list = [e.strip() for e in cc_emails.split(',')] if isinstance(cc_emails, str) else cc_emails
            if cc_list:
                msg['Cc'] = ', '.join(cc_list)
        
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))
        
        if attachment_data:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_data.getvalue())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment_name}')
            msg.attach(part)
        
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipient_list + cc_list, msg.as_string())
        server.quit()
        
        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Failed: {str(e)}"


def create_email_body(ship_name, failed_count, reasons_summary):
    return f"""<html><body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
        <h2 style="color: #2c3e50;">Vessel Report Validation Alert</h2>
        <p>Dear Captain and C/E of <strong>{ship_name}</strong>,</p>
        <p>Automated notification regarding validation failures in your vessel reports.</p>
        <div style="background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0;">
            <h3 style="margin-top: 0; color: #856404;">Validation Summary</h3>
            <p><strong>Failed Reports:</strong> {failed_count}</p>
        </div>
        <h3>Common Issues Detected:</h3>
        <ul>{reasons_summary}</ul>
        <p>Review the attached Excel file for details.</p>
        <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
        <p style="color: #7f8c8d; font-size: 0.9em;">Contact: <strong><a href="mailto:smartapp@enginelink.blue">smartapp@enginelink.blue</a></strong></p>
        <p style="color: #7f8c8d; font-size: 0.85em;">Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S UTC')}</p>
    </body></html>"""


@st.cache_data(show_spinner=False)
def process_excel_file(file_bytes, file_name):
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="All Reports")
    failed, df_with_calcs = validate_reports(df)
    return (df.to_dict('records'), df.columns.tolist(), failed.to_dict('records') if not failed.empty else [], failed.columns.tolist() if not failed.empty else [], df_with_calcs.to_dict('records'), df_with_calcs.columns.tolist())


def main():
    st.set_page_config(page_title="Ship Report Validator", page_icon="🚢", layout="wide")
    
    for key in ['validation_done', 'failed_df', 'df_with_calcs', 'original_df']:
        if key not in st.session_state:
            st.session_state[key] = False if key == 'validation_done' else None
    
    st.title("🚢 Ship Report Validation System")
    st.markdown("Upload Excel file to validate ship reports (includes SCOC validation)")
    
    with st.sidebar:
        st.header("📋 Validation Rules")
        st.markdown("""
        **Rule 1:** SFOC 150–200 g/kWh (At Sea, ME Rhrs > 12)
        **Rule 2:** Avg Speed 0–20 knots (At Sea, ME Rhrs > 12)
        **Rule 3:** Exhaust Temp ±50°C deviation (At Sea, ME Rhrs > 12)
        **Rule 4:** ME Rhrs ≤ Report Hours + 1h
        **Rule 5:** AE/Sub-consumers check (At Sea, ME Load > 40%)
        **Rule 6:** SCOC 0.8–1.5 g/kWh (At Sea, ME Rhrs > 12)
        """)
        st.divider()
        st.header("📧 SMTP Settings")
        with st.expander("Configure Email"):
            smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
            smtp_port = st.number_input("Port", value=587, min_value=1)
            sender_email = st.text_input("Sender Email")
            sender_password = st.text_input("Password", type="password", help="Use App Password for Gmail")
    
    uploaded_file = st.file_uploader("Choose Excel file", type=["xlsx", "xls"])
    
    if uploaded_file:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if 'current_file_id' not in st.session_state or st.session_state.current_file_id != file_id:
            st.session_state.update({'current_file_id': file_id, 'validation_done': False, 'failed_df': None, 'df_with_calcs': None, 'original_df': None})
    
    if uploaded_file and not st.session_state.validation_done:
        try:
            with st.spinner("Validating..."):
                df_data, df_cols, failed_data, failed_cols, calc_data, calc_cols = process_excel_file(uploaded_file.read(), uploaded_file.name)
                st.session_state.original_df = pd.DataFrame(df_data, columns=df_cols)
                st.session_state.failed_df = pd.DataFrame(failed_data, columns=failed_cols) if failed_data else pd.DataFrame()
                st.session_state.df_with_calcs = pd.DataFrame(calc_data, columns=calc_cols)
                st.session_state.validation_done = True
            st.success(f"✅ Validated! Total rows: {len(st.session_state.original_df)}")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
    
    if st.session_state.validation_done:
        df, failed = st.session_state.original_df, st.session_state.failed_df
        
        st.header("📈 Results")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total", len(df))
        col2.metric("Failed", len(failed))
        col3.metric("Pass Rate", f"{((len(df)-len(failed))/len(df)*100):.1f}%")
        
        if not failed.empty:
            st.warning(f"⚠️ {len(failed)} failed")
            st.dataframe(failed, use_container_width=True, height=400)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                failed.to_excel(writer, index=False, sheet_name="Failed")
            output.seek(0)
            st.download_button("📥 Download Failed", output, "Failed_Validation.xlsx")
            
            st.divider()
            if "Ship Name" in failed.columns:
                vessels = failed["Ship Name"].unique()
                with st.form("email_form"):
                    vessel = st.selectbox("Select Vessel", vessels)
                    to_email = st.text_area("To:", placeholder="email@company.com")
                    cc_email = st.text_area("CC (optional):", placeholder="cc@company.com")
                    if st.form_submit_button("📤 Send Email"):
                        if not sender_email or not sender_password:
                            st.error("Configure SMTP settings")
                        elif to_email:
                            vessel_failed = failed[failed["Ship Name"] == vessel]
                            v_output = io.BytesIO()
                            with pd.ExcelWriter(v_output, engine='openpyxl') as w:
                                vessel_failed.to_excel(w, index=False)
                            v_output.seek(0)
                            
                            reasons_html = "\n".join([f"<li>{r} ({c}x)</li>" for r, c in pd.Series([r for rs in vessel_failed["Reason"] for r in rs.split("; ") if r]).value_counts().items()])
                            
                            success, msg = send_email(smtp_server, smtp_port, sender_email, sender_password, to_email, f"Validation Alert - {vessel}", create_email_body(vessel, len(vessel_failed), reasons_html), v_output, f"Failed_{vessel}.xlsx", cc_email)
                            st.success(msg) if success else st.error(msg)
        else:
            st.success("🎉 All passed!")
            st.balloons()
    
    elif not uploaded_file:
        st.info("👆 Upload Excel file to begin")


if __name__ == "__main__":
    main()

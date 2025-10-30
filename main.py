import win32com.client
import os
from datetime import datetime, timedelta
import pandas as pd
from config import BASE_FOLDER, COUNTRY_RULES, SNIPPET_LEN, MAX_EMAILS, DEFAULT_DISPLAY_ID, KEEP_COLS
from logger import logger

def send_email_notification(output_folder, filename, row_count, minutes_threshold=2):
    try:
        full_path = os.path.join(output_folder, filename)

        if not os.path.exists(full_path):
            logger.info(f"❌ File not found: {full_path}")
            return

        created_time = os.path.getctime(full_path)
        file_age_minutes = (
            datetime.now() - datetime.fromtimestamp(created_time)
        ).total_seconds() / 60

        if file_age_minutes > minutes_threshold:
            logger.info(f"⏩ Skipping email — file is {file_age_minutes:.1f} minutes old")
            return

        current_user = os.environ.get("USERNAME", "")
        cleaned_path = full_path.replace(f"C:\\Users\\{current_user}\\", "")

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.Subject = f"File Processed - {filename}"
        mail.Body = (
            f"Hello,\n\n"
            f"A new file has been processed.\n\n"
            f"File Path: {cleaned_path}\n"
            f"Rows in File: {row_count}\n\n"
        )
        mail.To = "Luc.visser@ab-inbev.com"; "kiran.kaur@ab-inbev.com"
        mail.Send()

        logger.info(f"✅ Email sent for new file: {full_path}")

    except Exception as e:
        logger.error(f"[ERROR line 43] {e}")


def standardize_date_column(df, col_name):
    if col_name not in df.columns:
        return df
    df[col_name] = pd.to_datetime(df[col_name], errors="coerce", dayfirst=True)
    df[col_name] = df[col_name].dt.strftime("%Y-%m-%d").fillna("")
    return df


def standardize_all_date_columns(df):
    for col in df.columns:
        if "DATE" in col.upper():
            df = standardize_date_column(df, col)
    return df


def replace_column_values(df, col_name, old_val, new_val):
    if col_name in df.columns:
        df[col_name] = df[col_name].replace(old_val, new_val)
    return df


def keep_only_columns(df, keep_cols):
    return df[[col for col in keep_cols if col in df.columns]]


def process_file(file_path, country, email_date_string, output_folder=None, skip_if_exists=True):
    try:
        df = pd.read_excel(file_path)
        df.columns = [col.upper() for col in df.columns]

        if country == "FR":
            df.rename(columns={"ABI_SFA_MECHANISM": "ABI_SFA_MECHANISM__C"}, inplace=True)
            df["ABI_SFA_DISPLAY__C"] = DEFAULT_DISPLAY_ID
            df = replace_column_values(df, "ABI_SFA_PRODUCT_SET__C", "FR_OFF_Bud Can 4x50xl", "FR_OFF_Bud Can 4x50cl")
            df = replace_column_values(df, "ABI_SFA_PRODUCT_SET__C", "FR_OFF_Hoegaarden Ros e 0,0 Bottle 6x25 cl", "FR_OFF_Hoegaarden Rosée 0,0 Bottle 6x25 cl")

        elif country == "IT":
            df = df.iloc[2:]
            df = df[df.iloc[:, 0] != "Example"]
            df["ABI_SFA_DISPLAY__C"] = DEFAULT_DISPLAY_ID
            df = replace_column_values(df, "ABI_SFA_PRODUCT_SET__C", "IT01_OFF_DIS_LEFFE BLONDE BOTTLE ONE WAY 1X0.50 L", "IT01_OFF_LEFFE BLONDE BOTTLE ONE WAY 1x50CL")
            df = replace_column_values(df, "ABI_SFA_PRODUCT_SET__C", "IT01_OFF_DIS_CORONA EXTRA BOTTLE ONE WAY 1x0.50 L", "IT01_OFF_CORONA EXTRABOTTLE ONE WAY 1X0.500 L")

        elif country == "BE":
            df["ABI_SFA_DISPLAY__C"] = DEFAULT_DISPLAY_ID
            df.rename(columns={
                "ABI_SFA_TYPE__C": "ABI_SFA_DISPLAY_TYPE__C",
                "ABI_SFA_PRODUCT_SET_FOR_SURVEY__C": "ABI_SFA_PRODUCT_SET__C",
                "ABI_SFA_PERSON_ON_REGISTERED__C": "ABI_SFA_PERSON_REGISTERED__C",
            }, inplace=True)

        elif country == "NL":
            df.rename(columns={"ABI_SFA_PRODUCT_SET_FOR_SURVEY__C": "ABI_SFA_PRODUCT_SET__C"}, inplace=True)
            df = replace_column_values(df, "ABI_SFA_PRODUCT_SET__C", "NL01_HERT JAN CAN 7+1 PROMO 3x8 05L DP NL", "NL01_HERTOG JAN PILS 0,0 BOTTLE 24x30CL")
            df["ABI_SFA_DISPLAY__C"] = DEFAULT_DISPLAY_ID

        else:
            raise ValueError(f"Country not supported: {country}")

        df = standardize_all_date_columns(df)
        df = keep_only_columns(df, KEEP_COLS)
        df = df.applymap(lambda x: str(x).strip() if pd.notnull(x) else "")

        if output_folder is None:
            output_folder = os.path.dirname(file_path)
        os.makedirs(output_folder, exist_ok=True)

        out_file = os.path.join(output_folder, f"Processed_{country}_{email_date_string}.csv")
        if skip_if_exists and os.path.exists(out_file):
            logger.info(f"⚠️ File already exists, skipping: {out_file}")
            return out_file, len(df)

        df.to_csv(out_file, index=False, sep=",", encoding="utf-8")
        logger.info(f"✅ Exported: {out_file} | Rows: {len(df)}")
        return out_file, len(df)

    except Exception as e:
        logger.error(f"[ERROR in process_file] {e}")
        return None, 0



outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
messages = inbox.Items
messages.Sort("[ReceivedTime]", True)

last_week = datetime.now() - timedelta(days=7)
messages = [
    m for m in messages
    if hasattr(m, "ReceivedTime")
    and datetime.fromtimestamp(m.ReceivedTime.timestamp()) >= last_week
]

matches_found = False

for i, message in enumerate(messages):
    if i >= MAX_EMAILS:
        break
    try:
        subject_lower = str(message.Subject).lower()
        if "re:" in subject_lower:
            continue
        if not hasattr(message, "Attachments") or message.Attachments.Count == 0:
            continue

        body_lower = str(message.Body).lower()
        sender_lower = str(message.SenderEmailAddress).lower()
        snippet = (subject_lower + " " + sender_lower + " " + body_lower[:SNIPPET_LEN]).lower()

        for country, rule in COUNTRY_RULES.items():
            if any(keyword.lower() in snippet for keyword in rule["keywords"]):
                matches_found = True
                email_date_string = message.ReceivedTime.strftime("%Y-%m-%d")
                dated_folder = os.path.join(rule["folder"], email_date_string)
                os.makedirs(dated_folder, exist_ok=True)

                for att in message.Attachments:
                    filename = getattr(att, "FileName", "")
                    att_type = getattr(att, "Type", None)

                    if not filename or not filename.lower().endswith(".xlsx") or att_type != 1:
                        continue

                    raw_filename = f"Raw_{rule['country_code']}_Data_{email_date_string}.xlsx"
                    raw_path = os.path.join(dated_folder, raw_filename)

                    if not os.path.exists(raw_path):
                        att.SaveAsFile(raw_path)
                        logger.info(f"⬇️ Saved raw file: {raw_path}")

                    processed_file, row_count = process_file(
                        raw_path,
                        rule["country_code"],
                        email_date_string,
                        output_folder=dated_folder,
                    )

                    if processed_file:
                        send_email_notification(dated_folder, os.path.basename(processed_file), row_count)

                break

    except Exception as e:
        logger.error(f"[ERROR] {e}")

if not matches_found:
    logger.info("No matching emails with .xlsx attachments.")

logger.info("Script execution completed.")

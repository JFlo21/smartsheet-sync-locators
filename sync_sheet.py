import smartsheet
import os
import requests
import time
import threading
import hashlib

# API Key
API_KEY = "wCQ53EjJ5LncpdIkuHH0ZC23nH3SEHDQnZSuN"

# Sheet IDs
SOURCE_SHEET_ID = "7261271259828100"
TARGET_SHEET_ID = "4417312359665540"

COLUMN_MAPPING = {
    7525747076059012: 5333430655209348,  # FOREMAN -> Foreman
    488872658292612: 8148180422315908,   # WORK REQUEST # -> WR #
    2740672471977860: 829831027838852    # LOCATION -> City
}

WR_NUMBER_COLUMN_ID = 8148180422315908  # Used as unique identifier
FOREMAN_COLUMN_ID = 5333430655209348

VALID_FOREMEN = [
    "Victor Duran",
    "Armando Garcia",
    "Paul Watson",
    "Chris Solomon",
    "Ricardo Martinez"
]

client = smartsheet.Smartsheet(API_KEY)
DOWNLOAD_FOLDER = "C:/Users/juflores/OneDrive - Centuri Group, Inc/Smartsheet-download-attachment-automation"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def validate_column_mapping(target_sheet, column_map):
    target_column_ids = [column.id for column in target_sheet.columns]
    for source_column_id, target_column_id in column_map.items():
        if target_column_id not in target_column_ids:
            raise ValueError(f"Target column ID {target_column_id} does not exist in the target sheet.")

def handle_rate_limit(response):
    retry_count = 0
    while response.status_code == 429 and retry_count < 5:
        wait_time = (2 ** retry_count) * 10
        print(f"⚠ Rate limit exceeded. Retrying in {wait_time} seconds...")
        time.sleep(wait_time)
        retry_count += 1
    return retry_count < 5

def download_attachment(attachment_name, attachment_url):
    file_path = os.path.join(DOWNLOAD_FOLDER, attachment_name)
    try:
        response = requests.get(attachment_url, stream=True, timeout=60)
        if response.status_code == 200:
            sha256_hash = hashlib.sha256()
            total_bytes = 0
            with open(file_path, "wb") as file:
                for chunk in response.iter_content(chunk_size=65536):
                    if chunk:
                        file.write(chunk)
                        sha256_hash.update(chunk)
                        total_bytes += len(chunk)
            print(f"✅ Downloaded: {attachment_name} ({total_bytes / (1024 * 1024):.2f} MB)")
            return file_path
        else:
            print(f"❌ Failed to download '{attachment_name}'. HTTP {response.status_code}")
    except Exception as e:
        print(f"❌ Error downloading '{attachment_name}': {e}")
    return None

def list_attachment_names(sheet_id, row_id):
    """Returns a set of attachment names for a given row."""
    try:
        attachments = client.Attachments.list_row_attachments(sheet_id, row_id).data
        return set(a.name for a in attachments if a.attachment_type == "FILE")
    except Exception:
        return set()

def copy_attachments(source_row_id, target_row_id):
    try:
        attachments = client.Attachments.list_row_attachments(SOURCE_SHEET_ID, source_row_id).data
        for attachment in attachments:
            if attachment.attachment_type != "FILE":
                continue
            file_attachment = client.Attachments.get_attachment(SOURCE_SHEET_ID, attachment.id)
            file_path = download_attachment(attachment.name, file_attachment.url)
            if file_path:
                with open(file_path, "rb") as file_content:
                    client.Attachments.attach_file_to_row(
                        TARGET_SHEET_ID, target_row_id,
                        (attachment.name, file_content, 'application/octet-stream')
                    )
                print(f"🔁 Uploaded to target: {attachment.name}")
                os.remove(file_path)
    except Exception as e:
        print(f"❌ Error copying attachments: {e}")

def sync_back_attachments(target_rows, source_rows):
    """Sync new attachments from target back to source rows based on WR #."""
    source_lookup = {
        next((c.value for c in r.cells if c.column_id == COLUMN_MAPPING[488872658292612]), None): r
        for r in source_rows
    }

    for target_row in target_rows:
        wr_value = next((c.value for c in target_row.cells if c.column_id == COLUMN_MAPPING[488872658292612]), None)
        if not wr_value or wr_value not in source_lookup:
            continue

        source_row = source_lookup[wr_value]
        source_names = list_attachment_names(SOURCE_SHEET_ID, source_row.id)
        target_names = list_attachment_names(TARGET_SHEET_ID, target_row.id)

        new_attachments = target_names - source_names
        if not new_attachments:
            continue

        for name in new_attachments:
            print(f"🔄 Syncing new attachment '{name}' to source row {source_row.id}")
            for attachment in client.Attachments.list_row_attachments(TARGET_SHEET_ID, target_row.id).data:
                if attachment.name == name and attachment.attachment_type == "FILE":
                    file_attachment = client.Attachments.get_attachment(TARGET_SHEET_ID, attachment.id)
                    file_path = download_attachment(attachment.name, file_attachment.url)
                    if file_path:
                        with open(file_path, "rb") as file_content:
                            client.Attachments.attach_file_to_row(
                                SOURCE_SHEET_ID, source_row.id,
                                (attachment.name, file_content, 'application/octet-stream')
                            )
                        print(f"✅ Synced back to source: {attachment.name}")
                        os.remove(file_path)

def copy_rows_with_mapping(source_rows, target_wr_numbers, target_sheet_id):
    for row in source_rows:
        try:
            if getattr(row, "locked", False):
                continue

            foreman = next((c.value for c in row.cells if c.column_id == FOREMAN_COLUMN_ID), None)
            wr_value = next((c.value for c in row.cells if c.column_id == COLUMN_MAPPING[488872658292612]), None)

            if not foreman or foreman not in VALID_FOREMEN or not wr_value:
                continue

            if wr_value in target_wr_numbers:
                continue  # ✅ Skip existing WR #

            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            for cell in row.cells:
                if cell.column_id in COLUMN_MAPPING and cell.value:
                    new_row.cells.append(smartsheet.models.Cell({
                        "column_id": COLUMN_MAPPING[cell.column_id],
                        "value": cell.value
                    }))

            created_row = client.Sheets.add_rows(target_sheet_id, [new_row]).result[0]
            print(f"✅ Row copied: WR #{wr_value}")
            copy_attachments(row.id, created_row.id)

        except Exception as e:
            print(f"❌ Error copying row {row.id}: {e}")

def get_wr_number_map(rows, wr_column_id):
    return {
        row.id: next((c.value for c in row.cells if c.column_id == wr_column_id), None)
        for row in rows
    }

def main():
    try:
        print("📥 Loading source + target sheets...")
        source_sheet = client.Sheets.get_sheet(SOURCE_SHEET_ID, include=["attachments"])
        target_sheet = client.Sheets.get_sheet(TARGET_SHEET_ID, include=["attachments"])

        print("🔍 Validating columns...")
        validate_column_mapping(target_sheet, COLUMN_MAPPING)

        # Build WR # lookup for deduplication
        target_wr_map = get_wr_number_map(target_sheet.rows, COLUMN_MAPPING[488872658292612])
        target_wr_values = set(filter(None, target_wr_map.values()))

        # ✅ Copy only new rows from source to target
        print("🚀 Copying new rows...")
        copy_rows_with_mapping(source_sheet.rows, target_wr_values, TARGET_SHEET_ID)

        # 🔁 Reverse sync: push new attachments from target → source
        print("🔁 Syncing attachments back to source...")
        sync_back_attachments(target_sheet.rows, source_sheet.rows)

        print("🎉 All done!")

    except smartsheet.exceptions.ApiError as e:
        print(f"❌ Smartsheet API error: {e.message}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")

if __name__ == "__main__":
    main()
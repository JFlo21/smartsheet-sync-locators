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

SOURCE_WR_NUMBER_COLUMN_ID = 488872658292612
TARGET_WR_NUMBER_COLUMN_ID = 8148180422315908
FOREMAN_COLUMN_ID = 5333430655209348

VALID_FOREMEN = [
    "Victor Duran", "Armando Garcia", "Paul Watson", "Chris Solomon", "Ricardo Martinez"
]

client = smartsheet.Smartsheet(API_KEY)
DOWNLOAD_FOLDER = "C:/Users/juflores/OneDrive - Centuri Group, Inc/Smartsheet-download-attachment-automation"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

def validate_column_mapping(target_sheet, column_map):
    target_column_ids = [col.id for col in target_sheet.columns]
    for src_col, tgt_col in column_map.items():
        if tgt_col not in target_column_ids:
            raise ValueError(f"Missing target column ID: {tgt_col}")

def download_attachment(name, url):
    path = os.path.join(DOWNLOAD_FOLDER, name)
    try:
        r = requests.get(url, stream=True, timeout=60)
        if r.status_code == 200:
            with open(path, "wb") as f:
                for chunk in r.iter_content(65536):
                    if chunk:
                        f.write(chunk)
            return path
        else:
            print(f"❌ Download failed: {name} (status {r.status_code})")
    except Exception as e:
        print(f"❌ Error downloading {name}: {e}")
    return None

def copy_attachments(source_row_id, target_row_id):
    try:
        attachments = client.Attachments.list_row_attachments(SOURCE_SHEET_ID, source_row_id).data
        for att in attachments:
            if att.attachment_type != "FILE":
                continue
            file_obj = client.Attachments.get_attachment(SOURCE_SHEET_ID, att.id)
            path = download_attachment(att.name, file_obj.url)
            if path:
                with open(path, "rb") as f:
                    client.Attachments.attach_file_to_row(
                        TARGET_SHEET_ID, target_row_id,
                        (att.name, f, 'application/octet-stream'))
                os.remove(path)
                print(f"📤 Uploaded to target: {att.name}")
    except Exception as e:
        print(f"❌ Attachment sync error: {e}")

def copy_rows_with_mapping(source_rows, existing_wr_keys, target_sheet_id):
    for row in source_rows:
        try:
            if getattr(row, "locked", False):
                continue

            foreman = next((c.value for c in row.cells if c.column_id == FOREMAN_COLUMN_ID), None)
            if not foreman or foreman not in VALID_FOREMEN:
                continue

            wr = next((c.value for c in row.cells if c.column_id == SOURCE_WR_NUMBER_COLUMN_ID), None)
            if not wr or str(wr).strip() == "":
                print(f"⚠️ Skipping row {row.id}: blank WR #")
                continue
            try:
                wr_key = int(str(wr).split('.')[0])
            except ValueError:
                print(f"⚠️ Skipping row {row.id}: invalid WR # format")
                continue

            if wr_key in existing_wr_keys:
                print(f"⏭️ WR #{wr_key} already in target — skipping.")
                continue

            new_row = smartsheet.models.Row()
            new_row.to_bottom = True
            for c in row.cells:
                if c.column_id in COLUMN_MAPPING and c.value is not None:
                    new_row.cells.append(smartsheet.models.Cell({
                        "column_id": COLUMN_MAPPING[c.column_id],
                        "value": c.value
                    }))

            created = client.Sheets.add_rows(target_sheet_id, [new_row]).result[0]
            print(f"✅ Copied new row: WR #{wr_key}")
            copy_attachments(row.id, created.id)

        except Exception as e:
            print(f"❌ Error copying row {row.id}: {e}")

def update_changed_rows(source_rows, target_rows, column_map):
    print("\n🔧 Checking for updates...")
    src_map, tgt_map = {}, {}

    for r in source_rows:
        for c in r.cells:
            if c.column_id == SOURCE_WR_NUMBER_COLUMN_ID and c.value:
                try:
                    src_map[int(str(c.value).split('.')[0])] = r
                except: pass

    for r in target_rows:
        for c in r.cells:
            if c.column_id == TARGET_WR_NUMBER_COLUMN_ID and c.value:
                try:
                    tgt_map[int(str(c.value).split('.')[0])] = r
                except: pass

    for wr_key, src_row in src_map.items():
        if wr_key not in tgt_map:
            continue
        tgt_row = tgt_map[wr_key]
        tgt_cell_map = {c.column_id: c.value for c in tgt_row.cells}
        updates = []

        for sc in src_row.cells:
            if sc.column_id in column_map:
                tgt_col = column_map[sc.column_id]
                source_val = sc.value
                target_val = tgt_cell_map.get(tgt_col)

                if source_val is not None and (target_val is None or source_val != target_val):
                    updates.append(smartsheet.models.Cell({
                        "column_id": tgt_col,
                        "value": source_val
                    }))

        if updates:
            row_update = smartsheet.models.Row()
            row_update.id = tgt_row.id
            row_update.cells = updates
            try:
                client.Sheets.update_rows(TARGET_SHEET_ID, [row_update])
                print(f"🔁 Updated WR #{wr_key}")
            except Exception as e:
                print(f"❌ Update failed for WR #{wr_key}: {e}")

def sync_target_attachments_to_source(source_rows, target_rows):
    print("\n🔁 Syncing attachments from target → source...")
    src_map = {}
    for r in source_rows:
        for c in r.cells:
            if c.column_id == SOURCE_WR_NUMBER_COLUMN_ID and c.value:
                try:
                    src_map[int(str(c.value).split('.')[0])] = r.id
                except: pass

    for row in target_rows:
        try:
            wr = next((c.value for c in row.cells if c.column_id == TARGET_WR_NUMBER_COLUMN_ID), None)
            wr_key = int(str(wr).split('.')[0]) if wr else None
            if not wr_key or wr_key not in src_map:
                continue
            source_row_id = src_map[wr_key]

            target_attachments = client.Attachments.list_row_attachments(TARGET_SHEET_ID, row.id).data
            existing = client.Attachments.list_row_attachments(SOURCE_SHEET_ID, source_row_id).data
            existing_names = {a.name for a in existing if a.attachment_type == "FILE"}

            for att in target_attachments:
                if att.attachment_type != "FILE" or att.name in existing_names:
                    continue

                file_obj = client.Attachments.get_attachment(TARGET_SHEET_ID, att.id)
                path = download_attachment(att.name, file_obj.url)
                if path:
                    with open(path, "rb") as f:
                        client.Attachments.attach_file_to_row(
                            SOURCE_SHEET_ID, source_row_id, (att.name, f, 'application/octet-stream'))
                    os.remove(path)
                    print(f"📎 Synced back to source: {att.name}")
        except Exception as e:
            print(f"❌ Error syncing back row {row.id}: {e}")

def get_wr_number_map(rows, column_id):
    result = {}
    for row in rows:
        for c in row.cells:
            if c.column_id == column_id and c.value:
                try:
                    result[row.id] = int(str(c.value).split('.')[0])
                except: pass
    return result

def main():
    try:
        print("📥 Loading source & target sheets...")
        src = client.Sheets.get_sheet(SOURCE_SHEET_ID, include=["attachments"])
        tgt = client.Sheets.get_sheet(TARGET_SHEET_ID, include=["attachments"])

        print("🔍 Validating columns...")
        validate_column_mapping(tgt, COLUMN_MAPPING)

        target_wr_keys = set(get_wr_number_map(tgt.rows, TARGET_WR_NUMBER_COLUMN_ID).values())

        print("📤 Copying new rows...")
        copy_rows_with_mapping(src.rows, target_wr_keys, TARGET_SHEET_ID)

        print("🛠️ Updating changed rows...")
        update_changed_rows(src.rows, tgt.rows, COLUMN_MAPPING)

        print("🔁 Syncing attachments back to source...")
        sync_target_attachments_to_source(src.rows, tgt.rows)

        print("🎉 Sync complete.")
    except Exception as e:
        print(f"❌ Fatal error: {e}")

if __name__ == "__main__":
    main()
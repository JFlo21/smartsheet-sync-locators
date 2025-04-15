import smartsheet
import os
import requests
from datetime import datetime

# === CONFIGURATION ===

API_KEY = "wCQ53EjJ5LncpdIkuHH0ZC23nH3SEHDQnZSuN"
SOURCE_SHEET_ID = "7261271259828100"
TARGET_SHEET_ID = "1105533072265092"

COLUMN_MAPPING = {
    7525747076059012: 1209764350742404,  # FOREMAN -> Foreman
    488872658292612: 2335664257585028,   # WORK REQUEST # -> WR #
    2740672471977860: 6839263884955524,  # LOCATION -> City
}

SOURCE_WR_NUMBER_COLUMN_ID = 488872658292612
TARGET_WR_NUMBER_COLUMN_ID = 2335664257585028
FOREMAN_COLUMN_ID = 7525747076059012
COMPLETED_DATE_COLUMN_ID = 1051822611713924

VALID_FOREMEN = [
    "Ramon Perez", "Christopher Tiner", "Alphonso Flores", "Joe Hatman",
    "Dylan Hester", "Kyle Wagner", "Jimmy Adames", "Cody Tipps",
    "Walker Moody", "Travis McConnell", "Paul Watson"
]

DOWNLOAD_FOLDER = "C:/Users/juflores/OneDrive - Centuri Group, Inc/Smartsheet-download-attachment-automation"
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

client = smartsheet.Smartsheet(API_KEY)

# === VALIDATION ===

def validate_column_mapping(source_sheet, target_sheet, column_map):
    source_columns = {col.id: col.title for col in source_sheet.columns}
    target_columns = {col.id: col.title for col in target_sheet.columns}
    for src_id, tgt_id in column_map.items():
        if src_id not in source_columns:
            raise ValueError(f"❌ Missing SOURCE column ID: {src_id}")
        if tgt_id not in target_columns:
            raise ValueError(f"❌ Missing TARGET column ID: {tgt_id}")
    print("✅ Column mapping validated.")

# === UTILITIES ===

def get_wr_number_map(rows, column_id):
    result = {}
    for row in rows:
        for c in row.cells:
            if c.column_id == column_id and c.value:
                try:
                    result[row.id] = int(str(c.value).split('.')[0])
                except: pass
    return result

def get_completed_wr_keys(rows, wr_col_id, completed_col_id):
    result = set()
    for row in rows:
        wr = next((c.value for c in row.cells if c.column_id == wr_col_id), None)
        completed = next((c.value for c in row.cells if c.column_id == completed_col_id), None)
        if wr and completed:
            try:
                result.add(int(str(wr).split('.')[0]))
            except: pass
    return result

# === ATTACHMENT SYNC (Updated) ===

def copy_attachments(source_row_id, target_row_id):
    try:
        src_attachments = client.Attachments.list_row_attachments(SOURCE_SHEET_ID, source_row_id).data
        tgt_existing = client.Attachments.list_row_attachments(TARGET_SHEET_ID, target_row_id).data
        tgt_existing_names = {att.name for att in tgt_existing if att.attachment_type == "FILE"}

        for att in src_attachments:
            if att.attachment_type != "FILE" or att.name in tgt_existing_names:
                continue

            attachment_meta = client.Attachments.get_attachment(SOURCE_SHEET_ID, att.id)
            file_url = attachment_meta.url
            file_path = os.path.join(DOWNLOAD_FOLDER, att.name)

            try:
                response = requests.get(file_url, stream=True)
                if response.status_code == 200:
                    with open(file_path, 'wb') as f:
                        for chunk in response.iter_content(chunk_size=8192):
                            f.write(chunk)

                    with open(file_path, "rb") as f:
                        client.Attachments.attach_file_to_row(
                            TARGET_SHEET_ID, target_row_id, (att.name, f, 'application/octet-stream'))

                    os.remove(file_path)
                    print(f"📤 Uploaded to target: {att.name}")
                else:
                    print(f"❌ Download failed: {att.name} (HTTP {response.status_code})")
            except Exception as e:
                print(f"❌ Error downloading {att.name}: {e}")
    except Exception as e:
        print(f"❌ Attachment sync error: {e}")

def sync_target_attachments_to_source(source_rows, target_rows):
    print("\n🔁 Syncing attachments and Completed Date back to source...")
    src_map = {}
    for r in source_rows:
        for c in r.cells:
            if c.column_id == SOURCE_WR_NUMBER_COLUMN_ID and c.value:
                try:
                    src_map[int(str(c.value).split('.')[0])] = r
                except: pass

    for row in target_rows:
        try:
            wr = next((c.value for c in row.cells if c.column_id == TARGET_WR_NUMBER_COLUMN_ID), None)
            wr_key = int(str(wr).split('.')[0]) if wr else None
            if not wr_key or wr_key not in src_map:
                continue
            source_row = src_map[wr_key]
            source_row_id = source_row.id

            # Sync Completed Date
            tgt_completed = next((c.value for c in row.cells if c.column_id == COMPLETED_DATE_COLUMN_ID), None)
            src_completed = next((c.value for c in source_row.cells if c.column_id == COMPLETED_DATE_COLUMN_ID), None)
            if tgt_completed and not src_completed:
                row_update = smartsheet.models.Row()
                row_update.id = source_row_id
                row_update.cells = [smartsheet.models.Cell({
                    "column_id": COMPLETED_DATE_COLUMN_ID,
                    "value": tgt_completed
                })]
                client.Sheets.update_rows(SOURCE_SHEET_ID, [row_update])
                print(f"🗓️ Synced Completed Date for WR #{wr_key}")

            # Sync attachments from target to source
            target_attachments = client.Attachments.list_row_attachments(TARGET_SHEET_ID, row.id).data
            existing = client.Attachments.list_row_attachments(SOURCE_SHEET_ID, source_row_id).data
            existing_names = {a.name for a in existing if a.attachment_type == "FILE"}

            for att in target_attachments:
                if att.attachment_type != "FILE" or att.name in existing_names:
                    continue
                file_path = os.path.join(DOWNLOAD_FOLDER, att.name)
                try:
                    file_meta = client.Attachments.get_attachment(TARGET_SHEET_ID, att.id)
                    file_url = file_meta.url
                    response = requests.get(file_url, stream=True)
                    if response.status_code == 200:
                        with open(file_path, 'wb') as f:
                            for chunk in response.iter_content(chunk_size=8192):
                                f.write(chunk)

                        with open(file_path, "rb") as f:
                            client.Attachments.attach_file_to_row(
                                SOURCE_SHEET_ID, source_row_id, (att.name, f, 'application/octet-stream'))
                        os.remove(file_path)
                        print(f"🔁 Synced back to source: {att.name}")
                    else:
                        print(f"❌ Download failed: {att.name} (HTTP {response.status_code})")
                except Exception as e:
                    print(f"❌ Error syncing attachment {att.name}: {e}")
        except Exception as e:
            print(f"❌ Error syncing back row {row.id}: {e}")

# === ROW OPERATIONS ===

def copy_rows_with_mapping(source_rows, blocked_wr_keys, target_sheet_id):
    for row in source_rows:
        try:
            if getattr(row, "locked", False):
                continue

            foreman = next((c.value for c in row.cells if c.column_id == FOREMAN_COLUMN_ID), None)
            if not foreman or foreman not in VALID_FOREMEN:
                continue

            wr = next((c.value for c in row.cells if c.column_id == SOURCE_WR_NUMBER_COLUMN_ID), None)
            if not wr or str(wr).strip() == "":
                continue
            try:
                wr_key = int(str(wr).split('.')[0])
            except ValueError:
                continue

            if wr_key in blocked_wr_keys:
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
            copy_attachments(row.id, created.id)
            print(f"✅ Copied new WR #{wr_key}")

        except Exception as e:
            print(f"❌ Error copying row {row.id}: {e}")

def update_changed_rows(source_rows, target_rows, column_map, completed_sources, completed_targets):
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
        if wr_key in completed_sources or wr_key in completed_targets:
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

# === MAIN EXECUTION ===

def main():
    try:
        print("📥 Loading source & target sheets...")
        src = client.Sheets.get_sheet(SOURCE_SHEET_ID, include=["attachments"])
        tgt = client.Sheets.get_sheet(TARGET_SHEET_ID, include=["attachments"])

        print("🔍 Validating columns...")
        validate_column_mapping(src, tgt, COLUMN_MAPPING)

        print("🔎 Gathering WR numbers...")
        target_wr_keys = set(get_wr_number_map(tgt.rows, TARGET_WR_NUMBER_COLUMN_ID).values())
        completed_sources = get_completed_wr_keys(src.rows, SOURCE_WR_NUMBER_COLUMN_ID, COMPLETED_DATE_COLUMN_ID)
        completed_targets = get_completed_wr_keys(tgt.rows, TARGET_WR_NUMBER_COLUMN_ID, COMPLETED_DATE_COLUMN_ID)

        print("📤 Copying new rows...")
        blocked_wr_keys = target_wr_keys.union(completed_targets)
        copy_rows_with_mapping(src.rows, blocked_wr_keys, TARGET_SHEET_ID)

        print("🛠️ Updating changed rows...")
        update_changed_rows(src.rows, tgt.rows, COLUMN_MAPPING, completed_sources, completed_targets)

        print("🔁 Syncing target ➜ source...")
        sync_target_attachments_to_source(src.rows, tgt.rows)

        print("🎉 Sync complete.")
    except Exception as e:
        print(f"❌ Fatal error: {e}")

if __name__ == "__main__":
    main()

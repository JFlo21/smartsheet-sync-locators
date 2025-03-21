import smartsheet

# API Key
API_KEY = "wCQ53EjJ5LncpdIkuHH0ZC23nH3SEHDQnZSuN"

# Source and Target Sheet IDs
SOURCE_SHEET_ID = "cFQC6WGjXJvVwXqRMFfpRQQ24433cRPfqqvQxRW1"
TARGET_SHEET_ID = "rWjVj7g4M72fFjmjfFC8FfR8cFW8jxwP9QVCX431"

# Column Mapping
COLUMN_MAPPING = {
    "FOREMAN":"Foreman",
    "WORK REQUEST #":"WR #",
    "LOCATION":"City"
}

ROW_VALUES = ["vduran@ltspower.com", "armagarcia@ltspower.com", "pwatson@ltspower.com", "chsolomon@ltspower.com", "ricamartinez@ltspower.com"]

# Initialize Smartsheet client
client = smartsheet.Smartsheet(API_KEY)

def get_sheet(sheet_id):
    """Fetch the sheet details."""
    return client.Sheets.get_sheet(sheet_id)

def get_column_map(sheet, column_mapping):
    """Generate column ID mapping based on column names."""
    column_map = {}
    for col in sheet.columns:
        if col.title in column_mapping:
            column_map[col.id] = column_mapping[col.title]
    return column_map

def copy_rows_with_mapping_and_attachments(source_sheet, target_sheet, source_rows, column_map):
    """
    Copy rows from the source sheet to the target sheet with mapping
    and include attachments for each row.
    """
    for row in source_rows:
        # Create a new row for the target sheet
        new_row = smartsheet.models.Row()
        new_row.to_bottom = True

        for cell in row.cells:
            if cell.column_id in column_map:
                target_column_id = column_map[cell.column_id]

                # Handle contact columns (e.g., FOREMAN) and extract email if necessary
                if isinstance(cell.value, dict) and "email" in cell.value:
                    value = cell.value["email"]  # Extract email for contact list columns
                else:
                    value = cell.value  # Use the existing value for other columns

                # Add mapped cell to the new row
                new_row.cells.append(
                    smartsheet.models.Cell({
                        "column_id": target_column_id,
                        "value": value
                    })
                )

        # Add the new row to the target sheet
        created_row = client.Sheets.add_rows(target_sheet.id, [new_row]).result[0]

        # Copy attachments from the source row to the target row
        copy_attachments(row.id, created_row.id)


def copy_attachments(source_row_id, target_row_id):
    """Copy attachments from source row to target row."""
    attachments = client.Attachments.list_row_attachments(SOURCE_SHEET_ID, source_row_id).data
    for attachment in attachments:
        client.Attachments.attach_url_to_row(
            TARGET_SHEET_ID, target_row_id,
            smartsheet.models.Attachment({
                'name': attachment.name,
                'attachmentType': attachment.attachment_type,
                'url': attachment.url
            })
        )

def sync_attachments_from_target_to_source(target_sheet, column_map):
    """Synchronize attachments from target rows to source rows."""
    for row in target_sheet.rows:
        for cell in row.cells:
            if cell.column_id in column_map and cell.value in ROW_VALUES:
                attachments = client.Attachments.list_row_attachments(TARGET_SHEET_ID, row.id).data
                for attachment in attachments:
                    client.Attachments.attach_url_to_row(
                        SOURCE_SHEET_ID, row.id,
                        smartsheet.models.Attachment({
                            'name': attachment.name,
                            'attachmentType': attachment.attachment_type,
                            'url': attachment.url
                        })
                    )

def main():
    """Main function to handle synchronization."""
    source_sheet = get_sheet(SOURCE_SHEET_ID)
    target_sheet = get_sheet(TARGET_SHEET_ID)

    column_map = get_column_map(source_sheet, COLUMN_MAPPING)
    
    copy_rows_with_mapping(source_sheet, target_sheet, source_sheet.rows, column_map)
    sync_attachments_from_target_to_source(target_sheet, column_map)

if __name__ == "__main__":
    main()
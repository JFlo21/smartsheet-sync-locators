import smartsheet

# API Key
API_KEY = "wCQ53EjJ5LncpdIkuHH0ZC23nH3SEHDQnZSuN"

# Sheet IDs
SOURCE_SHEET_ID = "cFQC6WGjXJvVwXqRMFfpRQQ24433cRPfqqvQxRW1"
TARGET_SHEET_ID = "rWjVj7g4M72fFjmjfFC8FfR8cFW8jxwP9QVCX431"

# Column mapping between Source and Target Sheets
COLUMN_MAPPING = {
    "FOREMAN": "Foreman",
    "WORK REQUEST #": "WR #",
    "LOCATION": "City"
}

# Valid foreman values (names)
VALID_FOREMEN = [
    "Victor Duran",
    "Armando Garcia",
    "Paul Watson",
    "Chris Solomon",
    "Ricardo Martinez"
]

# Initialize Smartsheet client
client = smartsheet.Smartsheet(API_KEY)


def get_column_mapping(sheet, mapping):
    """
    Maps the source sheet columns to target sheet columns using their titles.
    """
    column_map = {}
    for column in sheet.columns:
        if column.title in mapping:
            column_map[column.id] = mapping[column.title]
    return column_map


def copy_rows_with_mapping_and_attachments(source_rows, column_map, target_sheet_id):
    """
    Copy rows from the source sheet to the target sheet, ensuring correct mapping and including attachments.
    """
    for row in source_rows:
        # Filter rows based on the FOREMAN column values
        foreman_cell_value = next((cell.value for cell in row.cells if cell.column_id in column_map and column_map[cell.column_id] == "Foreman"), None)
        if foreman_cell_value not in VALID_FOREMEN:
            continue  # Skip rows where FOREMAN is not in the valid list

        # Create a new row for the target sheet
        new_row = smartsheet.models.Row()
        new_row.to_bottom = True

        # Create new cells for the target row
        for cell in row.cells:
            if cell.column_id in column_map:
                target_column_id = column_map[cell.column_id]
                value = cell.value  # Directly use the string value for Text/Number columns

                # Append the cell to the new row
                new_row.cells.append(
                    smartsheet.models.Cell({
                        "column_id": target_column_id,
                        "value": value
                    })
                )

        # Add the row to the target sheet
        created_row = client.Sheets.add_rows(target_sheet_id, [new_row]).result[0]

        # Copy attachments from the source row to the target row
        copy_attachments(row.id, created_row.id)


def copy_attachments(source_row_id, target_row_id):
    """
    Copy attachments from a row in the source sheet to the corresponding row in the target sheet.
    """
    # Retrieve attachments from the source row
    attachments = client.Attachments.list_row_attachments(SOURCE_SHEET_ID, source_row_id).data

    # Add each attachment to the target row
    for attachment in attachments:
        client.Attachments.attach_url_to_row(
            TARGET_SHEET_ID, target_row_id,
            smartsheet.models.Attachment({
                "name": attachment.name,
                "attachmentType": attachment.attachment_type,
                "url": attachment.url
            })
        )


def main():
    """
    Main function to handle synchronization.
    """
    # Get source and target sheets
    source_sheet = client.Sheets.get_sheet(SOURCE_SHEET_ID)
    target_sheet = client.Sheets.get_sheet(TARGET_SHEET_ID)

    # Map the columns
    column_map = get_column_mapping(source_sheet, COLUMN_MAPPING)

    # Copy rows and their attachments
    copy_rows_with_mapping_and_attachments(source_sheet.rows, column_map, TARGET_SHEET_ID)


if __name__ == "__main__":
    main()

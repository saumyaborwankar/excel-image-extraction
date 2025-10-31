# from openpyxl import load_workbook
# import os


# def emu_to_pixels(emu):
#     return emu / 9525 if emu else 0


# file_path = "sample.xlsx"
# output_dir = "images_with_positions"
# os.makedirs(output_dir, exist_ok=True)

# wb = load_workbook(file_path, data_only=True)

# for sheet_name in wb.sheetnames:
#     ws = wb[sheet_name]

#     for idx, image in enumerate(ws._images, start=1):
#         # --- 1️⃣ Save the raw image ---
#         img_bytes = image._data()
#         img_filename = f"{sheet_name}_image_{idx}.png"
#         img_path = os.path.join(output_dir, img_filename)
#         with open(img_path, "wb") as f:
#             f.write(img_bytes)

#         # --- 2️⃣ Safely extract anchor info ---
#         anchor = image.anchor
#         cell = "Unknown"
#         offset_x = offset_y = 0

#         try:
#             if hasattr(anchor, "_from"):  # for OneCellAnchor
#                 start = anchor._from
#                 cell = ws.cell(row=start.row + 1, column=start.col + 1).coordinate
#                 offset_x = emu_to_pixels(start.colOff)
#                 offset_y = emu_to_pixels(start.rowOff)

#             elif hasattr(anchor, "from_"):  # alternate attribute name
#                 start = anchor.from_
#                 cell = ws.cell(row=start.row + 1, column=start.col + 1).coordinate
#                 offset_x = emu_to_pixels(start.colOff)
#                 offset_y = emu_to_pixels(start.rowOff)

#             elif isinstance(anchor, str):  # Fallback (older openpyxl)
#                 cell = anchor
#         except Exception as e:
#             print(f"[Warning] Could not parse anchor for {img_filename}: {e}")

#         # --- 3️⃣ Print image position info ---
#         print(f"Sheet: {sheet_name}")
#         print(f"  ➜ Image: {img_filename}")
#         print(f"  ➜ Cell: {cell}")
#         print(f"  ➜ Offset (pixels): x={offset_x:.1f}, y={offset_y:.1f}")
#         print("-" * 50)

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
import os


def emu_to_pixels(emu):
    return emu / 9525 if emu else 0


def calculate_bottom_right_cell(
    ws,
    start_row,
    start_col,
    start_offset_x,
    start_offset_y,
    img_width_px,
    img_height_px,
):
    """
    Calculate the bottom-right cell based on image dimensions.
    Returns (cell_coordinate, offset_x, offset_y)
    """
    current_row = start_row
    current_col = start_col
    remaining_width = start_offset_x + img_width_px
    remaining_height = start_offset_y + img_height_px

    # Traverse columns
    while remaining_width > 0:
        col_width_px = ws.column_dimensions[get_column_letter(current_col)].width
        if col_width_px is None:
            col_width_px = 8.43
        col_width_px *= 7

        if remaining_width <= col_width_px:
            break
        remaining_width -= col_width_px
        current_col += 1

    # Traverse rows
    while remaining_height > 0:
        row_height_px = ws.row_dimensions[current_row].height
        if row_height_px is None:
            row_height_px = 15
        row_height_px *= 1.33

        if remaining_height <= row_height_px:
            break
        remaining_height -= row_height_px
        current_row += 1

    cell_coord = ws.cell(row=current_row, column=current_col).coordinate
    return cell_coord, remaining_width, remaining_height


file_path = "sample.xlsx"
output_dir = "images_with_positions"
os.makedirs(output_dir, exist_ok=True)

wb = load_workbook(file_path, data_only=True)

for sheet_name in wb.sheetnames:
    ws = wb[sheet_name]

    for idx, image in enumerate(ws._images, start=1):
        anchor = image.anchor
        top_left_cell = "Unknown"
        bottom_right_cell = "Unknown"
        offset_x = offset_y = 0
        end_offset_x = end_offset_y = 0

        try:
            if hasattr(anchor, "_from") and hasattr(anchor, "to"):
                start = anchor._from
                top_left_cell = ws.cell(
                    row=start.row + 1, column=start.col + 1
                ).coordinate
                offset_x = emu_to_pixels(start.colOff)
                offset_y = emu_to_pixels(start.rowOff)

                end = anchor.to
                bottom_right_cell = ws.cell(
                    row=end.row + 1, column=end.col + 1
                ).coordinate
                end_offset_x = emu_to_pixels(end.colOff)
                end_offset_y = emu_to_pixels(end.rowOff)

            elif hasattr(anchor, "_from"):
                start = anchor._from
                top_left_cell = ws.cell(
                    row=start.row + 1, column=start.col + 1
                ).coordinate
                offset_x = emu_to_pixels(start.colOff)
                offset_y = emu_to_pixels(start.rowOff)

                img_width_px = (
                    emu_to_pixels(anchor.ext.cx) if hasattr(anchor, "ext") else 0
                )
                img_height_px = (
                    emu_to_pixels(anchor.ext.cy) if hasattr(anchor, "ext") else 0
                )

                if img_width_px > 0 and img_height_px > 0:
                    bottom_right_cell, end_offset_x, end_offset_y = (
                        calculate_bottom_right_cell(
                            ws,
                            start.row + 1,
                            start.col + 1,
                            offset_x,
                            offset_y,
                            img_width_px,
                            img_height_px,
                        )
                    )
                else:
                    bottom_right_cell = "Unknown"

            elif isinstance(anchor, str):
                top_left_cell = anchor

        except Exception as e:
            print(f"[Warning] Could not parse anchor for image {idx}: {e}")

        if top_left_cell != "Unknown" and bottom_right_cell != "Unknown":
            img_filename = f"{top_left_cell}-{bottom_right_cell}.png"
        else:
            img_filename = f"{sheet_name}_image_{idx}.png"

        img_bytes = image._data()
        img_path = os.path.join(output_dir, img_filename)
        with open(img_path, "wb") as f:
            f.write(img_bytes)

        print(f"Sheet: {sheet_name}")
        print(f"  ➜ Image: {img_filename}")
        print(f"  ➜ Top-Left Cell: {top_left_cell}")
        print(f"  ➜ Top-Left Offset (pixels): x={offset_x:.1f}, y={offset_y:.1f}")
        print(f"  ➜ Bottom-Right Cell: {bottom_right_cell}")
        print(
            f"  ➜ Bottom-Right Offset (pixels): x={end_offset_x:.1f}, y={end_offset_y:.1f}"
        )
        print("-" * 50)

print(f"\nAll images saved to: {output_dir}")

import re
import pandas as pd
from pathlib import Path

from common import load_json
from common import create_floor_plans_dict
from common import create_tag_keys_dict
from common import create_simulated_radios_dict
from common import create_notes_dict
from common import create_antenna_types_dict

nl = '\n'


def adjust_column_widths(df, writer):
    """Adjust column widths in the Excel sheet and apply text wrap to the 'Notes' column."""
    worksheet = writer.sheets['AP List']
    # Create a format for wrapping text
    wrap_format = writer.book.add_format({'text_wrap': True})

    for idx, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).map(len).max(), len(col)) + 5
        # Check if the current column is 'Notes' to apply text wrap format
        if col == 'Notes':
            worksheet.set_column(idx, idx, column_len * 1.2, wrap_format)
        else:
            worksheet.set_column(idx, idx, column_len * 1.2)


def format_headers(df, writer):
    """Format header row in the Excel sheet."""
    worksheet = writer.sheets['AP List']
    header_format = writer.book.add_format(
        {'bold': True, 'valign': 'center', 'font_size': 16, 'border': 0})

    for idx, col in enumerate(df.columns):
        # Write the header with custom format
        worksheet.write(0, idx, col, header_format)


def create_ap_list(project_object):

    message_callback = project_object.append_message

    message_callback(f'Generating BoM XLSX for: {project_object.project_name}\n')
    project_dir = Path(project_object.working_directory) / project_object.project_name

    # Load JSON data
    floor_plans_json = load_json(project_dir, 'floorPlans.json', message_callback)
    access_points_json = load_json(project_dir, 'accessPoints.json', message_callback)
    simulated_radios_json = load_json(project_dir, 'simulatedRadios.json', message_callback)
    antenna_types_json = load_json(project_dir, 'antennaTypes.json', message_callback)
    tag_keys_json = load_json(project_dir, 'tagKeys.json', message_callback)
    notes_json = load_json(project_dir, 'notes.json', message_callback)

    # Process data
    floor_plans_dict = create_floor_plans_dict(floor_plans_json)
    tag_keys_dict = create_tag_keys_dict(tag_keys_json)
    simulated_radio_dict = create_simulated_radios_dict(simulated_radios_json)
    antenna_types_dict = create_antenna_types_dict(antenna_types_json)
    notes_dict = create_notes_dict(notes_json)

    custom_ap_list = project_object.current_profile_ap_list_module.create_custom_ap_list(access_points_json, floor_plans_dict, tag_keys_dict, simulated_radio_dict, antenna_types_dict, notes_dict)

    # Create a pandas dataframe and export to Excel
    df = pd.DataFrame(custom_ap_list)

    # Extract version number (e.g., "v1.3") if present
    match = re.search(r'v\d+\.\d+', project_object.project_name)
    if match:
        version = match.group(0)

        # Remove version from project_name
        project_name_cleaned = re.sub(r' - predictive design v\d+\.\d+', '', project_object.project_name)

        # Construct the new filename format
        output_filename = f'{project_name_cleaned} - AP List {version}.xlsx' if version else f'{project_name_cleaned} - AP List.xlsx'

    else:
        message_callback('### ERROR: Unable to find expected pattern in project file name, substituting with default output name ###')
        output_filename = f'{project_object.project_name} - AP List.xlsx'

    try:
        with pd.ExcelWriter(Path(project_object.working_directory / output_filename), engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='AP List', index=False)
            adjust_column_widths(df, writer)
            format_headers(df, writer)

        message_callback(f'{nl}"{Path(output_filename).name}" created successfully{nl}{nl}### PROCESS COMPLETE ###')

    except Exception as e:
        error_msg = f"Error: {str(e)}"
        message_callback(f'{nl}### ERROR: Failed to create output file ###{nl}{error_msg}{nl}### PROCESS INCOMPLETE ###')

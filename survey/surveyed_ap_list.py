import pandas as pd
from pathlib import Path

from common import load_json
from common import create_floor_plans_dict
from common import create_tag_keys_dict
from common import create_access_point_measurements_dict
from common import create_measured_radios_dict
from common import create_notes_dict
from common import adjust_column_widths
from common import format_headers
from common import flatten_picture_notes_hierarchical

from common import nl


channel_bands = ['2.4', '5', '6']

right_align_cols =\
    (
        [f'{band} SSIDs' for band in channel_bands] +
        [f'{band} GHz' for band in channel_bands] +
        [f'{band} Supported Rates' for band in channel_bands]
    )

narrow_fixed_width_cols =\
    (
        [f'{band} Ch Primary' for band in channel_bands] +
        [f'{band} Width' for band in channel_bands] +
        [f'{band} Tx Power' for band in channel_bands] +
        [f'{band} WiFi Band' for band in channel_bands] +
        ['Colour', 'hidden']
    )

wide_fixed_width_cols =\
    (
        [f'{band} Security / Standards' for band in channel_bands] +
        [f'{band} Channel from IEs' for band in channel_bands] +
        ['flagged as My AP', 'manually positioned']
    )

# Define SSID columns
ssid_columns = ['2.4 SSIDs', '5 SSIDs', '6 SSIDs']


def create_surveyed_ap_list(self):
    message_callback = self.append_message

    message_callback(f'Generating surveyed AP list for: {self.project_name}\n')
    project_dir = Path(self.working_directory) / self.project_name

    # Load JSON data
    floor_plans_json = load_json(project_dir, 'floorPlans.json', message_callback)
    access_points_json = load_json(project_dir, 'accessPoints.json', message_callback)
    access_point_measurements_json = load_json(project_dir, 'accessPointMeasurements.json', message_callback)
    measured_radios_json = load_json(project_dir, 'measuredRadios.json', message_callback)
    tag_keys_json = load_json(project_dir, 'tagKeys.json', message_callback)
    notes_json = load_json(project_dir, 'notes.json', message_callback)

    # Process data
    floor_plans_dict = create_floor_plans_dict(floor_plans_json)
    tag_keys_dict = create_tag_keys_dict(tag_keys_json)
    access_point_measurements_dict = create_access_point_measurements_dict(access_point_measurements_json)
    measured_radios_dict = create_measured_radios_dict(measured_radios_json, access_point_measurements_dict)
    notes_dict = create_notes_dict(notes_json)

    surveyed_ap_list = self.current_profile_ap_list_module.create_custom_measured_ap_list(access_points_json, floor_plans_dict, tag_keys_dict, measured_radios_dict, notes_dict)

    # Create a pandas dataframe and export to Excel
    df = pd.DataFrame(surveyed_ap_list)

    # Helper function to clean SSID Series
    def clean_ssids(series):
        return (
            series
            .dropna()
            .str.split('\n')
            .explode()
            .dropna()
            .str.strip()
            .str.replace(r'\s*\(.*?\)', '', regex=True)
            .drop_duplicates()
            .sort_values()
        )

    # Separate cleaned SSIDs
    ssids_24 = clean_ssids(df['2.4 SSIDs']).to_frame(name='SSID')
    ssids_5 = clean_ssids(df['5 SSIDs']).to_frame(name='SSID')
    ssids_6 = clean_ssids(df['6 SSIDs']).to_frame(name='SSID')

    # Merged cleaned SSIDs
    all_ssids = pd.concat([ssids_24, ssids_5, ssids_6]).drop_duplicates().sort_values(by='SSID')
    all_ssids = all_ssids.reset_index(drop=True)

    map_note_df = None

    # Check if pictureNotes.json exists
    picture_notes_json = load_json(project_dir, 'pictureNotes.json', message_callback)

    if picture_notes_json is not None:
        map_notes = flatten_picture_notes_hierarchical(picture_notes_json, notes_dict, floor_plans_dict)
        map_note_df = pd.DataFrame(map_notes)

    # Create directory to hold output
    output_dir = self.working_directory / 'OUTPUT'
    output_dir.mkdir(parents=True, exist_ok=True)

    output_filename = output_dir / f'{self.project_name} - Surveyed AP List.xlsx'

    try:
        sheet_name = 'Surveyed AP List'
        writer = pd.ExcelWriter(str(output_filename), engine='xlsxwriter')
        # Sheet 1: Surveyed AP List
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        adjust_column_widths(df, writer, sheet_name, right_align_cols, narrow_fixed_width_cols, wide_fixed_width_cols)
        format_headers(df, writer, sheet_name, freeze_row=True, freeze_col=True)

        # Separate SSID Sheets
        sheet_name = '2.4GHz SSIDs'
        ssids_24.to_excel(writer, sheet_name=sheet_name, index=False)
        adjust_column_widths(ssids_24, writer, sheet_name)
        format_headers(ssids_24, writer, sheet_name)

        sheet_name = '5GHz SSIDs'
        ssids_5.to_excel(writer, sheet_name='5GHz SSIDs', index=False)
        adjust_column_widths(ssids_5, writer, sheet_name)
        format_headers(ssids_5, writer, sheet_name)

        sheet_name = '6GHz SSIDs'
        ssids_6.to_excel(writer, sheet_name='6GHz SSIDs', index=False)
        adjust_column_widths(ssids_6, writer, sheet_name)
        format_headers(ssids_6, writer, sheet_name)

        sheet_name = 'All SSIDs'
        all_ssids.to_excel(writer, sheet_name='All SSIDs', index=False)
        adjust_column_widths(all_ssids, writer, sheet_name)
        format_headers(all_ssids, writer, sheet_name)

        if map_note_df is not None and not map_note_df.empty:
            sheet_name = 'Map Notes'
            map_note_df.to_excel(writer, sheet_name=sheet_name, index=False)
            adjust_column_widths(map_note_df, writer, sheet_name)
            format_headers(map_note_df, writer, sheet_name)

        writer.close()
        message_callback(f'{nl}"{output_filename.name}" created successfully{nl}{nl}### PROCESS COMPLETE ###')
    except Exception as e:
        print(e)
        message_callback(f'{nl}### ERROR: Unable to create "{output_filename}" ###{nl}file could be open in another application{nl}### PROCESS INCOMPLETE ###')

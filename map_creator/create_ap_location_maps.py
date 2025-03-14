# create_custom_ap_location_maps.py

import wx
import shutil
import threading
from pathlib import Path
from PIL import Image

from common import nl
from common import load_json
from common import create_floor_plans_dict
from common import create_simulated_radios_dict
from common import ERROR, PROCESS_COMPLETE, PROCESS_ABORTED

from map_creator.map_creator_comon import vector_source_check
from map_creator.map_creator_comon import crop_assessment
from map_creator.map_creator_comon import annotate_map
from map_creator.map_creator_comon import oversize_map_check
from map_creator.map_creator_comon import add_project_filename_to_map


CUSTOM_AP_ICON_SIZE_ADJUSTER = 4.87


def create_custom_ap_location_maps_threaded(working_directory, project_name, message_callback, custom_ap_icon_size, ap_name_label_size, stop_event):
    # Wrapper function to run insert_images in a separate thread
    def run_in_thread():
        create_ap_location_maps(working_directory, project_name, message_callback, custom_ap_icon_size, ap_name_label_size, stop_event)
    # Start the long-running task in a separate thread
    threading.Thread(target=run_in_thread).start()


def create_ap_location_maps(working_directory, project_name, message_callback, custom_ap_icon_size, ap_name_label_size, stop_event):
    wx.CallAfter(message_callback, f'Creating custom AP location maps for: {project_name}{nl}'
                                   f'Custom AP icon size: {custom_ap_icon_size}{nl}')

    custom_ap_icon_size = int(custom_ap_icon_size * CUSTOM_AP_ICON_SIZE_ADJUSTER)

    project_dir = Path(working_directory) / project_name

    # Load JSON data
    floor_plans_json = load_json(project_dir, 'floorPlans.json', message_callback)
    access_points_json = load_json(project_dir, 'accessPoints.json', message_callback)
    simulated_radios_json = load_json(project_dir, 'simulatedRadios.json', message_callback)

    # Process data
    floor_plans_dict = create_floor_plans_dict(floor_plans_json)
    simulated_radio_dict = create_simulated_radios_dict(simulated_radios_json)

    # Create directory to hold output directories
    output_dir = working_directory / 'OUTPUT'
    output_dir.mkdir(parents=True, exist_ok=True)

    # Create subdirectory for Blank floor plans
    blank_plan_dir = output_dir / 'blank'
    blank_plan_dir.mkdir(parents=True, exist_ok=True)

    # Create subdirectory for Custom AP Location maps
    custom_ap_location_maps = output_dir / 'AP location maps'
    custom_ap_location_maps.mkdir(parents=True, exist_ok=True)

    # Create subdirectory for temporary files
    temp_dir = output_dir / 'temp'
    temp_dir.mkdir(parents=True, exist_ok=True)

    for floor in sorted(floor_plans_json['floorPlans'], key=lambda i: i['name']):
        if stop_event.is_set():
            wx.CallAfter(message_callback, PROCESS_ABORTED)
            return

        wx.CallAfter(message_callback, f"{nl}{nl}Processing floor: {floor['name']}{nl}")

        floor_id = vector_source_check(floor, message_callback)

        # Move floor plan to temp_dir
        shutil.copy(project_dir / ('image-' + floor_id), temp_dir / floor_id)

        # Open the floor plan to be used for AP placement activities
        source_floor_plan_image = Image.open(temp_dir / floor_id)

        # Check if the map is oversized
        oversize_map_check(source_floor_plan_image, message_callback)

        # Ensure the map_image is in 'RGBA' mode
        if source_floor_plan_image.mode != 'RGBA':
            wx.CallAfter(message_callback, f'Converting {floor_id} to RGBA colour space')
            source_floor_plan_image = source_floor_plan_image.convert('RGBA')

        map_cropped_within_ekahau, scaling_ratio, crop_bitmap = crop_assessment(floor, source_floor_plan_image, project_dir, floor_id, blank_plan_dir)

        aps_on_this_floor = []

        for ap in sorted(access_points_json['accessPoints'], key=lambda i: i['name']):
            if stop_event.is_set():
                wx.CallAfter(message_callback, PROCESS_ABORTED)
                return

            if ap['location']['floorPlanId'] == floor['id']:
                aps_on_this_floor.append(ap)

        current_map_image = source_floor_plan_image.copy()

        # Initialize all_aps to None
        all_aps = None

        if not aps_on_this_floor:
            wx.CallAfter(message_callback, f"No APs on this floor, generating a blank floor plan.")

            # Create a blank floor plan image to save
            blank_floor_plan = source_floor_plan_image.copy()

            # Crop it if Ekahau cropping applies
            if map_cropped_within_ekahau:
                blank_floor_plan = blank_floor_plan.crop(crop_bitmap)

            # Add project filename to blank map
            blank_floor_plan = add_project_filename_to_map(blank_floor_plan, ap_name_label_size, project_name)
            wx.CallAfter(message_callback, "Blank map stamped with project filename")

            # Save the blank floor plan
            blank_floor_plan.save(Path(custom_ap_location_maps / floor['name']).with_suffix('.png'))

            # Continue to the next floor instead of skipping
            continue

        else:
            # Generate the all_aps map
            for ap in aps_on_this_floor:
                if stop_event.is_set():
                    wx.CallAfter(message_callback, PROCESS_ABORTED)
                    return
                all_aps = annotate_map(current_map_image, ap, scaling_ratio, custom_ap_icon_size, ap_name_label_size, simulated_radio_dict, message_callback, floor_plans_dict)

        # If map was cropped within Ekahau, crop the all_AP map
        if map_cropped_within_ekahau:
            all_aps = all_aps.crop(crop_bitmap)

        # add project filename to the output image
        all_aps = add_project_filename_to_map(all_aps, ap_name_label_size, project_name)
        wx.CallAfter(message_callback, "map stamped with project filename")

        # Save the output images
        try:
            all_aps.save(Path(custom_ap_location_maps / floor['name']).with_suffix('.png'))
            wx.CallAfter(message_callback, f"Custom AP location map for {floor['name']} saved successfully")
        except Exception as e:
            wx.CallAfter(message_callback, ERROR)
            wx.CallAfter(message_callback, "Failure Attempting to save the OUTPUT images")
            print(e)

        # source_floor_plan_image.close()
        # current_map_image.close()

    try:
        shutil.rmtree(temp_dir)
        wx.CallAfter(message_callback, PROCESS_COMPLETE)
    except Exception as e:
        wx.CallAfter(message_callback, ERROR)
        wx.CallAfter(message_callback, str(e))
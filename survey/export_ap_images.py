#!/usr/bin/env python3

"""
Credit to Francois Verges (@VergesFrancois) for the original script and inspiration.
Adapted, modified, mangled by Nick Turner (@nickjvturner)
"""

import wx
import shutil

from common import load_json
from common import nl
from common import sanitize_string


def export_ap_images(project_object):
    message_callback = project_object.append_message

    project_dir = project_object.working_directory / project_object.project_name

    access_points_json = load_json(project_dir, 'accessPoints.json', message_callback)
    notes_json = load_json(project_dir, 'notes.json', message_callback)

    if not notes_json:
        wx.CallAfter(message_callback, f'No notes found in the project{nl}')
        return

    if not access_points_json:
        wx.CallAfter(message_callback, f'No access points found in the project{nl}')
        return

    wx.CallAfter(message_callback, f'Extracting AP Images from: {project_object.project_name}{nl}')

    # Create directory to hold output directories
    output_dir = project_object.working_directory / 'OUTPUT'
    output_dir.mkdir(parents=True, exist_ok=True)

    # Create subdirectory for note images
    ap_images_dir = output_dir / 'AP images'
    ap_images_dir.mkdir(parents=True, exist_ok=True)

    image_extraction_counter = []

    # Loop through all the APs in the project
    for ap in access_points_json['accessPoints']:
        image_count = 1

        # Check if the AP has any notes
        if 'noteIds' in ap.keys():

            # Check if the AP is placed on a map
            if 'location' in ap.keys():

                # Determine if there are multiple notes with images for this AP
                image_notes_count = sum(1 for ap_note in ap['noteIds'] for note in notes_json['notes'] if note['id'] == ap_note and len(note['imageIds']) > 0)

                for ap_note in ap['noteIds']:

                    # Loop through all the notes stored within the project
                    for note in notes_json['notes']:

                        # Skip notes that do not contain images
                        # imageIds exists and is a non-empty list
                        if 'imageIds' in note and note['imageIds']:

                            # Check if the note id matches the current AP note id
                            if note['id'] == ap_note:

                                # Loop through all the images attached to the note
                                for image in note['imageIds']:
                                    source_image_file = 'image-' + image
                                    source_image_full_path = project_dir / source_image_file

                                    # Prepare output image name
                                    ap_image_name = sanitize_string(ap['name'], message_callback)

                                    # Determine the output image name
                                    if image_notes_count > 1 or len(note['imageIds']) > 1:
                                        # Add image count starting from 1 if there are multiple images associated with this AP
                                        ap_image_name = f"{ap_image_name}-{image_count}.png"
                                    else:
                                        # Only one note with images, so no suffix for the first image
                                        ap_image_name = f"{ap_image_name}.png"

                                    output_destination = ap_images_dir / ap_image_name

                                    # Count total number of APs extracted
                                    image_extraction_counter.append(source_image_file)

                                    shutil.copy(source_image_full_path, output_destination)
                                    wx.CallAfter(message_callback, f"{ap_image_name} Image extracted")

                                    image_count += 1

    wx.CallAfter(message_callback, f'{nl}{len(image_extraction_counter)} images extracted{nl}')

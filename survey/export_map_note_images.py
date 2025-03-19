#!/usr/bin/env python3

"""
The original, well written script by Francois Verges (@VergesFrancois)
Adapted, modified, mangled by Nick Turner (@nickjvturner)

This script will extract all image notes attached to an AP objects of an
Ekahau project file (.esx)
"""

import shutil
from pathlib import Path
from datetime import datetime

from common import load_json
from common import nl


def export_map_note_images(project_object):
	message_callback = project_object.append_message

	project_dir = Path(project_object.working_directory) / project_object.project_name

	access_points_json = load_json(project_dir, 'accessPoints.json', message_callback)
	notes_json = load_json(project_dir, 'notes.json', message_callback)

	if not notes_json:
		message_callback(f'No notes found in the project{nl}')
		return

	# Create an empty list to store noteIds that are associated with an AP
	ap_note_ids = []

	# Check that access_points_json is not empty
	if access_points_json:
		for ap in access_points_json['accessPoints']:
			# If AP has notes and they contain images, add the noteIds to the list
			if 'noteIds' in ap.keys() and len(ap['noteIds']) > 0:
				for noteId in ap['noteIds']:
					ap_note_ids.append(noteId)

	else:
		message_callback(f'No access points found in the project{nl}')

	map_note_ids = []

	for note in notes_json['notes']:
		# Collect notes that are not associated with an AP and contain images
		if note['id'] not in ap_note_ids and len(note['imageIds']) > 0:
			map_note_ids.append(note)

	if not map_note_ids:
		message_callback(f'No map notes containing images found in the project{nl}')
		return

	# Create directory to hold output directories
	output_dir = project_object.working_directory / 'OUTPUT'
	output_dir.mkdir(parents=True, exist_ok=True)

	# Create subdirectory for note images
	map_note_image_dir = output_dir / 'map note images'
	map_note_image_dir.mkdir(parents=True, exist_ok=True)

	message_callback(f'Extracting Images from: {project_object.project_name} notes')

	image_extraction_counter = []

	for map_note in map_note_ids:
		if len(map_note['imageIds']) > 0:
			image_count = 1

			for image in map_note['imageIds']:
				image = 'image-' + image
				image_full_path = project_dir / image

				# Process the createdAt stamp to make it filename friendly
				created_at = datetime.fromisoformat((map_note['history']['createdAt']).replace('Z', '+00:00')).strftime(f"%Y-%m-%d__%H-%M-%S")

				if len(map_note['imageIds']) > 1:
					# there must be more than 1 image, add '-1', '-2', '-3', etc
					map_note_image_name = f"{created_at}-{str(image_count)}.png"
				else:
					map_note_image_name = f"{created_at}.png"

				dst = map_note_image_dir / map_note_image_name

				# count total number of APs extracted
				image_extraction_counter.append(map_note_image_name)

				shutil.copy(image_full_path, dst)
				message_callback(f"{image} extracted as {map_note_image_name}")

				image_count += 1

	message_callback(f'{nl}{len(image_extraction_counter)} images extracted{nl}')

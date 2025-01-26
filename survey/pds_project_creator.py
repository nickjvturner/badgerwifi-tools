import wx
import re
import shutil
import json

from common import load_json
from common import nl
from esx_actions.rebundle_esx import rebundle_project


def process_pds_maps(floor_plans_json, pds_maps_dir, temp_project_dir, message_callback):
    # Process and replace images in the temporary directory
    for floor in sorted(floor_plans_json['floorPlans'], key=lambda i: i['name']):
        image_id = floor['imageId']
        floor_name = floor['name']
        pds_map_path = pds_maps_dir / f'{floor_name}.png'

        if not pds_map_path.exists():
            wx.CallAfter(message_callback, f"{nl}WARNING: Missing PDS map for {floor_name}. Skipping {nl}")
            continue

        dest_path = temp_project_dir / f'image-{image_id}'
        shutil.copy(pds_map_path, dest_path)
        wx.CallAfter(message_callback, f"Copied PDS map for {floor_name} into temp_dir")
    # add a newline after the last message to make the output more readable
    wx.CallAfter(message_callback, "")


def remove_unwanted_json_assets(self, temp_project_dir, message_callback):
    # Remove unnecessary JSON files in the temporary directory
    if hasattr(self, 'project_profile_module'):
        # wx.CallAfter(message_callback, f"")
        for file in getattr(self.project_profile_module, 'predictive_json_asset_deletion', []):
            file_path = temp_project_dir / f"{file}.json"
            if file_path.exists():
                file_path.unlink()
                wx.CallAfter(message_callback, f"Removed: {file} from {temp_project_dir.parts[-2]}")
        # add a newline after the last message to make the output more readable
        wx.CallAfter(message_callback, "")

    else:
        wx.CallAfter(message_callback, f"{nl}Selected project profile does not contain json asset removal instructions{nl}"
                                       f"PDS maps have been swapped in, project will be rebundled with predictive design elements still present.{nl}")


def name_check(target_name, module, message_callback):
    """
    Checks if the target_name matches any 'name' defined in dictionaries within the module.

    :param target_name: The name to check against.
    :param module: The module containing the list of dictionaries.
    :param message_callback: Function to handle messages.
    :return: True if a match is found, False otherwise.
    """
    # Retrieve the names from the module (assuming it has an attribute with the list of dictionaries)
    names_to_check = [
        profile.get('name') for profile in getattr(module, 'profiles', [])
    ]

    # Check if the target_name matches any existing name in the list
    if target_name in names_to_check:
        wx.CallAfter(message_callback, f"Match found for name: {target_name}")
        return True
    else:
        return False


def install_post_deployment_survey_coverage_requirements(self, requirements_json, temp_project_dir, message_callback):
    # look for pre-existing name matches in the project requirements_json
    for requirement in requirements_json['requirements']:
        if name_check(requirement.get('name'), self.project_profile_module.post_deployment_survey_coverage_requirements, message_callback):
            wx.CallAfter(message_callback, f"Requirement {requirement.get('name')} already exists in project file, this will be overwritten.")
            # remove the existing requirement
            requirements_json['requirements'].remove(requirement)

    # remove default from existing profiles
    for requirement in requirements_json['requirements']:
        if requirement.get('isDefault') is True:
            # change the default to false
            requirement['isDefault'] = False

    # install defined coverage requirements from project profile
    for requirement in getattr(self.project_profile_module, 'post_deployment_survey_coverage_requirements', []):
        # add the requirement to the requirements.json file
        requirements_json['requirements'].append(requirement)

    # save the updated requirements.json file
    requirements_json_path = temp_project_dir / 'requirements.json'
    with open(requirements_json_path, 'w') as f:
        json.dump(requirements_json, f, indent=4)

    wx.CallAfter(message_callback, f"post-deployment coverage requirement(s) installed.json")


def configure_existing_coverage_area_requirements(self, areas_json, temp_project_dir, message_callback):
    # Set all areas defined within project to the newly configured default requirement profile
    for requirement in getattr(self.project_profile_module, 'post_deployment_survey_coverage_requirements', []):
        if requirement.get('isDefault') is True:
            coverage_requirement_name = requirement.get('name')
            coverage_requirement_id = requirement.get('id')
            break

    for area in areas_json['areas']:
        area['requirementId'] = coverage_requirement_id

    # save the updated areas.json file
    areas_json_path = temp_project_dir / 'areas.json'
    with open(areas_json_path, 'w') as f:
        json.dump(areas_json, f, indent=4)

    wx.CallAfter(message_callback, f"Configured all areas to use Coverage Requirement: {coverage_requirement_name}")


def create_pds_project_esx(self, message_callback):
    # Validate directories
    pds_maps_dir = self.working_directory / 'OUTPUT' / 'PDS AP location maps'
    project_dir = self.working_directory / self.project_name

    if not pds_maps_dir.exists():
        wx.CallAfter(message_callback, f"PDS maps directory not found. Run the PDS map creator first.")
        return

    # Load and validate JSON
    floor_plans_json = load_json(project_dir, 'floorPlans.json', message_callback)
    if not floor_plans_json:
        return
    requirements_json = load_json(project_dir, 'requirements.json', message_callback)
    if not requirements_json:
        return
    areas_json = load_json(project_dir, 'areas.json', message_callback)
    if not areas_json:
        return

    # Create a temporary directory inside the working directory
    temp_dir_name = f"temp_dir"
    temp_dir = self.working_directory / temp_dir_name
    try:
        if temp_dir.exists():
            shutil.rmtree(temp_dir)  # Clean up any existing temp directory
        temp_dir.mkdir()

        # Copy the project directory into the temporary directory
        temp_project_dir = temp_dir / self.project_name
        shutil.copytree(project_dir, temp_project_dir)

        wx.CallAfter(message_callback, f"Temporary duplicate project directory created:{nl}{temp_project_dir}{nl}")

        process_pds_maps(floor_plans_json, pds_maps_dir, temp_project_dir, message_callback)
        remove_unwanted_json_assets(self, temp_project_dir, message_callback)
        install_post_deployment_survey_coverage_requirements(self, requirements_json, temp_project_dir, message_callback)
        configure_existing_coverage_area_requirements(self, areas_json, temp_project_dir, message_callback)

        # Rebundle project from the temporary directory
        rebundle_project(temp_dir, self.project_name, self.append_message)

        # Rename and move the rebundled file to the working directory
        try:
            rebundled_filename = f"{self.project_name}_re-zip.esx"
            rebundled_file = temp_dir / rebundled_filename

            if rebundled_file.exists():
                # Check if the expected pattern is found in the filename
                if re.search(r' - predictive design v(\d+\.\d+)', self.project_name):
                    # If pattern is found, apply the new naming convention
                    post_deployment_filename = re.sub(
                        r' - predictive design v(\d+\.\d+)',  # Match the version pattern
                        r' - post-deployment v0.1',  # Replace with the new pattern
                        self.project_name
                    ) + '.esx'
                else:
                    # If pattern is not found, leave the rebundled filename unchanged
                    wx.CallAfter(message_callback, f"'predictive design vx.x' pattern NOT found in source filename")
                    post_deployment_filename = f"{self.project_name}_re-zip.esx"

                destination_path = self.working_directory / post_deployment_filename

                # Move and rename the rebundled file
                shutil.move(rebundled_file, destination_path)
                wx.CallAfter(message_callback, f"{nl}Rebundled file renamed and moved to:{nl}{destination_path}{nl}")
            else:
                wx.CallAfter(message_callback, f"Error: Rebundled file {rebundled_file} not found")
        except Exception as e:
            wx.CallAfter(message_callback, f"Unexpected error while renaming file: {e}")

    finally:
        # Clean up the temporary directory
        if temp_dir.exists():
            shutil.rmtree(temp_dir)
            wx.CallAfter(message_callback, f"Temporary directory {temp_dir} has been deleted")

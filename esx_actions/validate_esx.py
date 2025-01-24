# validate_esx.py

from common import load_json
from common import create_floor_plans_dict
from common import create_tag_keys_dict
from common import create_simulated_radios_dict
from common import offender_constructor
from common import create_custom_ap_dict
from common import acceptable_antenna_tilt_angles

from common import FIVE_GHZ_RADIO_ID

from common import nl, SPACER, PASS, FAIL, CAUTION, HASH_BAR


# def project_filename_compliance(esx, message_callback):
#     if esx.
#     message_callback(f"{SPACER}### ESX PROJECT FILE NAME FORMATTING ###")
#




def validate_ap_name_formatting(offenders, total_ap_count, message_callback):
    message_callback(f"{SPACER}### AP NAME FORMATTING ###")
    if len(offenders.get('ap_name_format', [])) > 0:
        message_callback(f"{FAIL}The following {len(offenders.get('ap_name_format', []))} APs have a non-conforming name")
        for ap in offenders['ap_name_format']:
            message_callback(ap)
        return False
    message_callback(f"{PASS}All {total_ap_count} APs have a conforming name format{nl}")
    return True


def validate_ap_name_uniqueness(offenders, total_ap_count, message_callback):
    message_callback(f"{SPACER}### AP NAME UNIQUENESS ###")
    if len(offenders.get('ap_name_duplication', [])) > 0:
        message_callback(f"{FAIL}The following {len(offenders.get('ap_name_duplication', []))} APs have been automatically renamed, please check the original AP names")
        for ap in offenders['ap_name_duplication']:
            message_callback(ap)
        return False
    message_callback(f"{PASS}All {total_ap_count} APs have a unique name{nl}")
    return True


def validate_color_assignment(offenders, total_ap_count, message_callback):
    message_callback(f"{SPACER}### COLOUR ASSIGNMENT ###")
    if len(offenders.get('color', [])) > 0:
        message_callback(f"{FAIL}The following {len(offenders.get('color', []))} APs have been assigned no color")
        for ap in offenders['color']:
            message_callback(ap)
        return False
    message_callback(f"{PASS}All {total_ap_count} APs have a non-default colour{nl}")
    return True


def validate_height_manipulation(offenders, total_ap_count, message_callback):
    message_callback(f"{SPACER}### ANTENNA HEIGHT ###")
    if len(offenders.get('antennaHeight', [])) > 0:
        message_callback(f"{CAUTION}The following {len(offenders.get('antennaHeight', []))} APs are configured with the Ekahau 'default' height of 2.4 meters, is this intentional?")
        for ap in offenders['antennaHeight']:
            message_callback(ap)
        return True
    message_callback(f"{PASS}All {total_ap_count} APs have an assigned height other than '2.4' metres")
    return True


def validate_required_tags(offenders, total_ap_count, total_required_tag_keys_count, required_tag_keys, message_callback):
    message_callback(f"{SPACER}### REQUIRED TAGS ###")
    # Initialize a list to store failed tag validations
    pass_required_tag_validation = []

    for missing_tag in offenders['missing_required_tags']:
        if len(offenders.get('missing_required_tags', {}).get(missing_tag, [])) > 0:
            message_callback(
                f"{FAIL}There is a problem! The following {len(offenders.get('missing_required_tags', {}).get(missing_tag, []))} APs are missing the '{missing_tag}' tag")
            for ap in sorted(offenders['missing_required_tags'][missing_tag]):
                message_callback(ap)
            pass_required_tag_validation.append(False)
        pass_required_tag_validation.append(True)

    if all(pass_required_tag_validation):
        message_callback(f"{PASS}{total_required_tag_keys_count} tag keys are defined:")
        for tagKey in required_tag_keys:
            message_callback(f"{tagKey}")
        message_callback(f"All {total_ap_count} APs have the required {total_required_tag_keys_count} tag keys assigned")
        return True
    return False


def validate_antenna_tilt(offenders, total_ap_count, message_callback):
    message_callback(f"{SPACER}### ANTENNA TILT ###")
    if len(offenders.get('antennaTilt', [])) > 0:
        message_callback(f"{FAIL}The following {len(offenders.get('antennaTilt', []))} APs have an antenna tilt that will cause problems when generating per AP installer documentation")
        for ap in offenders['antennaTilt']:
            message_callback(ap)
        return False
    message_callback(f"{PASS}All {total_ap_count} APs have an antenna tilt value that will work with the per AP installer documentation generation process{nl}")
    return True


def validate_antenna_mounting_and_tilt_mismatch(offenders, total_ap_count, message_callback, custom_ap_dict):
    message_callback(f"{SPACER}### ANTENNA MOUNTING AND TILT ###")
    if len(offenders.get('antennaMounting_and_antennaTilt_mismatch', [])) > 0:
        message_callback(f"{CAUTION}The following {len(offenders.get('antennaMounting_and_antennaTilt_mismatch', []))} APs may be configured incorrectly{nl}These APs are WALL mounted with 0 degrees of tilt, is this intentional?")
        for ap in offenders['antennaMounting_and_antennaTilt_mismatch']:
            message_callback(f"{ap} | {custom_ap_dict[ap]['model']}")
        return True
    message_callback(f"{PASS}All {total_ap_count} APs have a conforming antenna mounting and tilt")
    return True


def validate_view_as_mobile_disabled(project_configuration_json, message_callback):
    view_as_mobile = None
    for item in project_configuration_json["projectConfiguration"]["displayOptions"]:
        if item["key"] == "view_as_mobile_device_selected":
            view_as_mobile = item["value"]
            break

    if view_as_mobile == "true":
        message_callback(f"{SPACER}### VIEW AS MOBILE ###")
        message_callback(f"{FAIL}View as mobile is enabled, this is a DISASTER")
        return False
    return True


# check if the floor plan has been cropped within Ekahau
def validate_ekahau_crop(floor_plans_json, message_callback):
    for floor in floor_plans_json['floorPlans']:
        if floor.get('cropMinX') != 0 or floor.get('cropMinY') != 0 or floor.get('cropMaxX') != floor.get('width') or floor.get('cropMaxY') != floor.get('height'):
            message_callback(f"{SPACER}### MAP CROPPED WITHIN EKAHAU ###")
            message_callback(f"{FAIL}{floor.get('name')} has been cropped within Ekahau")
            message_callback(f"This may prevent or complicate post-deployment map creation and seamless map swaps in later phases of the project")
            return False
        return True


def check_duplicate_coverage_requirement_names(requirements_json, message_callback):
    """
    Checks for duplicate names in the requirements JSON.

    :param requirements_json: The loaded requirements.json data
    :param message_callback: Function to handle messages
    :return: True if no duplicates found, False otherwise
    """

    # Extract all names
    coverage_requirement_names = [requirement.get('name') for requirement in requirements_json.get('requirements', [])]

    # Identify duplicates
    duplicates = {name for name in coverage_requirement_names if coverage_requirement_names.count(name) > 1}

    if duplicates:
        message_callback(f"{SPACER}### COVERAGE REQUIREMENT NAME UNIQUENESS ###")
        message_callback(f"{FAIL}Duplicate Coverage Requirement names found:")
        for name in duplicates:
            message_callback(f"  - {name}")
        return False

    return True


def coverage_requirement_name_match_check(esx, requirements_json, message_callback):
    for requirement in requirements_json['requirements']:
        if requirement.get('name') == esx.predictive_design_coverage_requirements.get('name'):
            message_callback(f"  PASS  - Coverage Requirement '{esx.predictive_design_coverage_requirements.get('name')}' is defined")
            return requirement

    message_callback(f"{FAIL}Coverage Requirement '{esx.predictive_design_coverage_requirements.get('name')}' is not defined\n")
    return False


def default_check(esx, requirement, message_callback):
    # check if default match
    if requirement.get('isDefault') == esx.predictive_design_coverage_requirements.get('isDefault'):
        if esx.predictive_design_coverage_requirements.get('isDefault'):
            message_callback(f"  PASS  - Coverage Requirement is correctly configured as the 'default'")
        else:
            message_callback(f"  PASS  - Coverage Requirement is correctly configured as 'non-default'")
        return True

    if esx.predictive_design_coverage_requirements.get('isDefault'):
        message_callback(f"# FAIL  - Coverage Requirement is NOT configured as 'default'\n")
    else:
        message_callback(f"# FAIL  - Coverage Requirement is NOT configured as 'non-default'\n")
    return False


def extract_value_from_criteria(criteria_list, match_criteria):
    """
    Extracts the value from a criteria list where radioTechnology, frequencyBand, and type match.

    :param criteria_list: List of criteria dictionaries to search through
    :param match_criteria: Dictionary containing the criteria to match (radioTechnology, frequencyBand, type)
    :return: The value of the matched criteria, or None if not found
    """
    for criteria in criteria_list:
        if (criteria.get('radioTechnology') == match_criteria.get('radioTechnology') and
                criteria.get('frequencyBand') == match_criteria.get('frequencyBand') and
                criteria.get('type') == match_criteria.get('type')):
            return criteria.get('value')
    return None


def specific_criteria_check(frequency_band, type, text_descriptor, esx, requirement, message_callback):
    # check if specific criteria is a match
    match_criteria = {
        'radioTechnology': 'IEEE802_11',
        'frequencyBand': frequency_band,
        'type': type
    }

    BWT_value = extract_value_from_criteria(esx.predictive_design_coverage_requirements.get('criteria'), match_criteria)
    ESX_value = extract_value_from_criteria(requirement.get('criteria'), match_criteria)

    if ESX_value == BWT_value:
        message_callback(f"  PASS  - {text_descriptor} is correctly configured as '{int(BWT_value)}'")
        return True

    message_callback(f"\n# FAIL  - {text_descriptor} is NOT configured correctly! Current value: '{int(ESX_value)}', should be: '{int(BWT_value)}'\n")
    return False


def validate_predictive_design_coverage_requirements(esx, requirements_json, message_callback):
    message_callback(f"{SPACER}### PREDICTIVE DESIGN COVERAGE REQUIREMENTS ###")
    if esx.predictive_design_coverage_requirements is None:
        message_callback(f"Selected Project Profile does not define 'Predictive Design Coverage Requirements'")
        return True

    requirement = coverage_requirement_name_match_check(esx, requirements_json, message_callback)

    if not requirement:
        return False

    predictive_design_coverage_requirements_pass = (
        default_check(esx, requirement, message_callback),
        specific_criteria_check('FIVE', 'SIGNAL_STRENGTH', '5GHz Primary Signal Strength', esx, requirement, message_callback),
        specific_criteria_check('FIVE', 'SECONDARY_SIGNAL_STRENGTH', '5GHz Secondary Signal Strength', esx, requirement, message_callback),
        specific_criteria_check('FIVE', 'SIGNAL_TO_NOISE_RATIO', '5GHz Signal to Noise Ratio', esx, requirement, message_callback),
        specific_criteria_check('FIVE', 'DATA_RATE', '5GHz Data Rate', esx, requirement, message_callback),
        specific_criteria_check('FIVE', 'CHANNEL_OVERLAP', '5GHz Channel Interference', esx, requirement, message_callback),
    )

    if all(predictive_design_coverage_requirements_pass):
        message_callback(f"{PASS}All predictive design coverage requirements are configured correctly.")
        return True

    message_callback(f"{FAIL}One or more predictive design coverage requirements are not configured correctly.")
    return False


def requirementId_getter(esx, requirements_json):
    for requirement in requirements_json['requirements']:
        if requirement.get('name') == esx.predictive_design_coverage_requirements.get('name'):
            return requirement.get('requirementId')

    return False


def validate_area_requirement_assignment(esx, areas_json, requirements_json, message_callback):
    # check if all areas are assigned the correct coverage requirement
    message_callback(f"{SPACER}### AREA REQUIREMENT ASSIGNMENT ###")
    if esx.predictive_design_coverage_requirements is None:
        message_callback(f"Selected Project Profile does not define 'Predictive Design Coverage Requirements' unable to validate area requirement assignment")
        return False

    target_requirementId = requirementId_getter(esx, requirements_json)

    for area in areas_json['areas']:
        if area.get('requirementID') != target_requirementId:
            message_callback(f"{FAIL}Area '{area.get('name')}' is not assigned the correct coverage requirement")
            return False

        message_callback(f"""  PASS  - Area '{area.get('name')}' is assigned the correct coverage requirement""")
            
    message_callback(f"""{PASS}All defined areas are correctly assigned '{esx.predictive_design_coverage_requirements.get("name")}' coverage requirement""")
    return True


def validate_esx(esx, message_callback):
    message_callback(f'Performing Validation for: {esx.project_name}')

    project_dir = esx.working_directory / esx.project_name

    # Load JSON data
    floor_plans_json = load_json(project_dir, 'floorPlans.json', message_callback)
    access_points_json = load_json(project_dir, 'accessPoints.json', message_callback)
    simulated_radios_json = load_json(project_dir, 'simulatedRadios.json', message_callback)
    tag_keys_json = load_json(project_dir, 'tagKeys.json', message_callback)
    project_configuration_json = load_json(project_dir, 'projectConfiguration.json', message_callback)
    requirements_json = load_json(project_dir, 'requirements.json', message_callback)
    areas_json = load_json(project_dir, 'areas.json', message_callback)

    # Process data
    floor_plans_dict = create_floor_plans_dict(floor_plans_json)
    tag_keys_dict = create_tag_keys_dict(tag_keys_json)
    simulated_radio_dict = create_simulated_radios_dict(simulated_radios_json)

    # Process access points
    custom_ap_dict = create_custom_ap_dict(access_points_json, floor_plans_dict, simulated_radio_dict)

    for ap in access_points_json['accessPoints']:
        for tag in ap['tags']:
            custom_ap_dict[ap['name']]['tags'][tag_keys_dict.get(tag['tagKeyId'])] = tag['value']

    offenders = offender_constructor(esx.required_tag_keys, esx.optional_tag_keys)

    # Count occurrences of each
    for ap in custom_ap_dict.values():

        if not ap['name'].startswith('AP-') or not ap['name'][3:].isdigit():
            offenders['ap_name_format'].append(ap['name'])

        # if an AP name contains string '_BW_DUPLICATE_AP_NAME_' it is not unique
        if '_BW_DUPLICATE_AP_NAME_' in ap['name']:
            offenders['ap_name_duplication'].append(ap['name'])

        if ap['color'] == 'none':
            offenders['color'].append(ap['name'])

        if ap['antennaHeight'] == 2.4:
            offenders['antennaHeight'].append(ap['name'])

        for radio in ap['radios'].values():
            if radio.get('radioTechnology') == 'BLUETOOTH' and radio.get('enabled', False):
                offenders['bluetooth'].append(ap['name'])

        if ap.get('radios', {}).get(FIVE_GHZ_RADIO_ID, {}).get('antennaTilt') not in acceptable_antenna_tilt_angles:
            offenders['antennaTilt'].append(ap['name'])

        if ap.get('antennaMounting') == 'WALL' and ap.get('radios', {}).get(FIVE_GHZ_RADIO_ID, {}).get('antennaTilt') == 0:
            offenders['antennaMounting_and_antennaTilt_mismatch'].append(ap['name'])

        for tagKey in esx.required_tag_keys:
            if tagKey not in ap['tags']:
                offenders['missing_required_tags'][tagKey].append(ap['name'])

    total_ap_count = len(custom_ap_dict)
    total_required_tag_keys_count = len(esx.required_tag_keys)

    # Perform all validations
    validations = [
        validate_ap_name_formatting(offenders, total_ap_count, message_callback),
        validate_ap_name_uniqueness(offenders, total_ap_count, message_callback),
        validate_color_assignment(offenders, total_ap_count, message_callback),
        validate_height_manipulation(offenders, total_ap_count, message_callback),
        validate_required_tags(offenders, total_ap_count, total_required_tag_keys_count, esx.required_tag_keys, message_callback),
        validate_antenna_tilt(offenders, total_ap_count, message_callback),
        validate_antenna_mounting_and_tilt_mismatch(offenders, total_ap_count, message_callback, custom_ap_dict),
        validate_view_as_mobile_disabled(project_configuration_json, message_callback),
        validate_ekahau_crop(floor_plans_json, message_callback),
        check_duplicate_coverage_requirement_names(requirements_json, message_callback),
        validate_predictive_design_coverage_requirements(esx, requirements_json, message_callback),
        validate_area_requirement_assignment(esx, areas_json, requirements_json, message_callback)
    ]

    # Print pass/fail states
    if all(validations):
        message_callback(f"{HASH_BAR}### VALIDATION PASSED ###{HASH_BAR}")
    else:
        message_callback(f"{HASH_BAR}### VALIDATION FAILED ###{HASH_BAR}")
    return

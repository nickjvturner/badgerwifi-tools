# my_frame.py

import wx
import os
import json
import threading
import webbrowser
import subprocess
import importlib.util
import platform
import random

from pathlib import Path
from importlib.machinery import SourceFileLoader

from drop_target import DropTarget
from common import file_or_dir_exists

from esx_actions.validate_esx import validate_esx
from esx_actions.unpack_esx import unpack_esx_file
from esx_actions.backup_esx import backup_esx
from esx_actions.ap_list_creator import create_ap_list
from esx_actions.rebundle_esx import rebundle_project

from project_detail.Summarise import run as summarise_esx

from rename_aps.rename_visualiser import visualise_ap_renaming
from rename_aps.ap_renamer import ap_renamer

from survey import export_map_note_images, export_ap_images

from map_creator.extract_blank_maps import extract_blank_maps
from map_creator.create_ap_location_maps import create_custom_ap_location_maps_threaded
from map_creator.create_zoomed_ap_location_maps import create_zoomed_ap_location_maps_threaded
from map_creator.create_pds_maps import create_pds_maps_threaded

from common import nl
from common import CONFIGURATION_DIR
from common import PROJECT_PROFILES_DIR
from common import RENAME_APS_DIR
from common import PROJECT_DETAIL_DIR
from common import ADMIN_ACTIONS_DIR
from common import BOUNDARY_SEPARATION_WIDGET
from common import WHIMSY_WELCOME_MESSAGES
from common import CALL_TO_DONATE_MESSAGE
from common import DIR_STRUCTURE_PROFILES_DIR

from common import discover_available_scripts
from common import import_module_from_path
from common import tracked_project_profile_check_for_update
from common import example_project_profile_names

from common import parse_project_metadata
from common import cleanup_unpacked_project_folder

from admin import check_for_updates
from admin.dir_creator import select_root_and_create_directory_structure
from admin.dir_creator import preview_directory_structure

from survey.surveyed_ap_list import create_surveyed_ap_list
from survey.pds_project_creator import create_pds_project_esx


class MyFrame(wx.Frame):
    def __init__(self, parent, title):
        wx.Frame.__init__(self, parent, title=title, size=(1000, 800))
        self.set_window()
        self.panel = wx.Panel(self)
        self.initialize_variables()
        self.setup_list_box()
        self.setup_display_log()
        self.setup_tabs()
        self.setup_text_labels()
        self.setup_buttons()
        self.setup_dropdowns()
        self.setup_text_input_boxes()
        self.setup_tab1()
        self.setup_tab2()
        self.setup_tab3()
        self.setup_tab4()
        self.setup_panel_rows()
        self.setup_main_sizer()
        self.create_menu()
        self.setup_drop_target()
        self.load_application_state()
        self.Center()
        self.Show()
        self.check_for_updates_on_startup()
        self.display_welcome_message()

    def set_window(self):
        self.SetMinSize((500, 600))
        self.widget_margin = 5

        if platform.system() == 'Windows':
            # Set the frame size to the minimum size
            self.SetSize((800, 700))
            self.sizer_edge_margin = 0
            self.notebook_margin = 3
            self.row_sizer_margin = 0
        if platform.system() == 'Darwin':
            # Set the frame size to the minimum size
            self.sizer_edge_margin = 5
            self.notebook_margin = 5
            self.row_sizer_margin = -5

    def initialize_variables(self):
        self.esx_project_unpacked = False  # Initialize the state variable
        self.working_directory = None
        self.project_name = None
        self.filepath = None
        self.current_project_profile_module = None
        self.rename_aps_boundary_separator = 200  # Initialize the boundary separator variable
        self.required_tag_keys = {}
        self.optional_tag_keys = {}
        self.project_filename_expected_pattern = None
        self.predictive_design_coverage_requirements = {}
        self.post_deployment_survey_coverage_requirements = []

        self.ignored_extensions = ('.DS_Store')
        self.project_metadata = None

        # Define the configuration directory path
        self.config_dir = Path(__file__).resolve().parent / CONFIGURATION_DIR

        # Ensure the configuration directory exists
        self.config_dir.mkdir(parents=True, exist_ok=True)

        # Define the path for the application state file
        self.app_state_file_path = self.config_dir / 'app_state.json'

        # Create a thread control variable
        self.stop_event = threading.Event()  # Initialize the stop event

    def setup_list_box(self):
        # Set up your list box here
        self.list_box = wx.ListBox(self.panel, style=wx.LB_EXTENDED)

        self.list_box.Bind(wx.EVT_KEY_DOWN, self.on_delete_key)

    def setup_display_log(self):
        # Setup display log here
        self.display_log = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY)

        # Set a monospaced font for display_log
        monospace_font = wx.Font(14, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.display_log.SetFont(monospace_font)

    def setup_dropdowns(self):
        """
        Setup all dropdown elements
        """
        # Discover available AP renaming scripts in 'rename_aps' directory
        self.available_ap_rename_scripts = discover_available_scripts(RENAME_APS_DIR)

        # Create a dropdown to select an AP renaming script
        self.ap_rename_script_dropdown = wx.Choice(self.tab1, choices=self.available_ap_rename_scripts)
        self.ap_rename_script_dropdown.SetSelection(0)  # Set default selection
        self.ap_rename_script_dropdown.Bind(wx.EVT_CHOICE, self.on_ap_rename_script_dropdown_selection)

        # Discover available project Profiles 'project_profiles' directory
        self.available_project_profiles = discover_available_scripts(PROJECT_PROFILES_DIR)

        # Create a dropdown to select a Project Profile on the Predictive Design tab
        self.project_profile_dropdown = wx.Choice(self.tab1, choices=self.available_project_profiles)
        self.project_profile_dropdown.SetSelection(0)  # Set default selection
        self.project_profile_dropdown.Bind(wx.EVT_CHOICE, self.on_design_project_profile_dropdown_selection)

        # Create a dropdown to select a Project Profile on the Survey tab
        self.survey_project_profile_dropdown = wx.Choice(self.tab3, choices=self.available_project_profiles)
        self.survey_project_profile_dropdown.SetSelection(0)  # Set default selection
        self.survey_project_profile_dropdown.Bind(wx.EVT_CHOICE, self.on_survey_project_profile_dropdown_selection)

        # Discover available Project Detail Views
        self.available_project_detail_views = discover_available_scripts(PROJECT_DETAIL_DIR)

        # Create a dropdown to select a Project Detail View
        self.project_detail_dropdown = wx.Choice(self.tab4, choices=self.available_project_detail_views)
        self.project_detail_dropdown.SetSelection(0)  # Set default selection
        self.project_detail_dropdown.Bind(wx.EVT_CHOICE, self.on_project_detail_dropdown_selection)

        # Discover available Admin action scripts
        self.available_admin_actions = discover_available_scripts(ADMIN_ACTIONS_DIR)

        # Create a dropdown to select an Admin action
        self.admin_actions_dropdown = wx.Choice(self.tab4, choices=self.available_admin_actions)
        self.admin_actions_dropdown.SetSelection(0)  # Set default selection
        self.admin_actions_dropdown.Bind(wx.EVT_CHOICE, self.on_admin_actions_dropdown_selection)

        # Discover available directory structure profiles
        self.available_dir_structure_profiles = discover_available_scripts(DIR_STRUCTURE_PROFILES_DIR)

        # Create a dropdown to select a directory structure profile
        self.dir_structure_profile_dropdown = wx.Choice(self.tab4, choices=self.available_dir_structure_profiles)
        self.dir_structure_profile_dropdown.SetSelection(0)  # Set default selection
        self.dir_structure_profile_dropdown.Bind(wx.EVT_CHOICE, self.on_dir_structure_profile_dropdown_selection)


    def setup_buttons(self):
        # Create add file button
        self.add_files_button = wx.Button(self.panel, label="Add Files")
        self.add_files_button.Bind(wx.EVT_BUTTON, self.on_add_file)
        self.add_files_button.SetToolTip(wx.ToolTip("Add .esx files to the file list"))

        # Create open working directory button
        self.open_working_directory_button = wx.Button(self.panel, label="Open Working Directory")
        self.open_working_directory_button.Bind(wx.EVT_BUTTON, self.on_open_working_directory)
        self.open_working_directory_button.SetToolTip(wx.ToolTip("Open the .esx working directory in your file manager"))

        # Create reset button
        self.reset_button = wx.Button(self.panel, label="Reset")
        self.reset_button.Bind(wx.EVT_BUTTON, self.on_reset)
        self.reset_button.SetToolTip(wx.ToolTip("Clear the file list and reset the application state"))

        # Create copy log button
        self.copy_log_button = wx.Button(self.panel, label="Copy Log")
        self.copy_log_button.Bind(wx.EVT_BUTTON, self.on_copy_log)
        self.copy_log_button.SetToolTip(wx.ToolTip("Copy the log to the clipboard"))

        # Create clear log button
        self.clear_log_button = wx.Button(self.panel, label="Clear Log")
        self.clear_log_button.Bind(wx.EVT_BUTTON, self.on_clear_log)
        self.clear_log_button.SetToolTip(wx.ToolTip("Clear the log"))

        self.display_project_detail_button = wx.Button(self.tab4, label="Display")
        self.display_project_detail_button.Bind(wx.EVT_BUTTON, self.on_display_project_detail)
        self.display_project_detail_button.SetToolTip(wx.ToolTip("Display detailed information about the current .esx project"))

        # Create unpack esx file button
        self.unpack_button = wx.Button(self.panel, label="Unpack .esx")
        self.unpack_button.Bind(wx.EVT_BUTTON, self.on_unpack)
        self.unpack_button.SetToolTip(wx.ToolTip("Unpack the selected .esx file"))

        # Create re-bundle esx file button
        self.rebundle_button = wx.Button(self.panel, label="Re-bundle .esx")
        self.rebundle_button.Bind(wx.EVT_BUTTON, self.on_rebundle_esx)
        self.rebundle_button.SetToolTip(wx.ToolTip("Re-bundle the unpacked project into a new .esx file"))

        # Create backup esx file button
        self.backup_button = wx.Button(self.panel, label="Backup .esx")
        self.backup_button.Bind(wx.EVT_BUTTON, self.on_backup)
        self.backup_button.SetToolTip(wx.ToolTip("Make a backup the of .esx file currently in the file list"))

        # Create a button to execute the selected AP renaming script
        self.rename_aps_button = wx.Button(self.tab1, label="Rename APs")
        self.rename_aps_button.Bind(wx.EVT_BUTTON, self.on_rename_aps)
        self.rename_aps_button.SetToolTip(wx.ToolTip("Execute the selected AP renaming script"))

        self.visualise_ap_renaming_button = wx.Button(self.tab1, label="Rename Pattern Visualiser")
        self.visualise_ap_renaming_button.Bind(wx.EVT_BUTTON, self.on_visualise_ap_renaming)
        self.visualise_ap_renaming_button.SetToolTip(wx.ToolTip("Preview the AP renaming pattern"))

        # Create a button to create an AP List Excel file in accordance with the selected project profile
        self.create_ap_list = wx.Button(self.tab1, label="AP List")
        self.create_ap_list.Bind(wx.EVT_BUTTON, self.on_create_ap_list)
        self.create_ap_list.SetToolTip(wx.ToolTip("Export AP data to Excel in accordance with the selected project profile"))

        self.validate_button = wx.Button(self.tab1, label="Validate")
        self.validate_button.Bind(wx.EVT_BUTTON, self.on_validate)
        self.validate_button.SetToolTip(wx.ToolTip("Validate the .esx project in accordance with the selected project profile"))

        self.summarise_button = wx.Button(self.tab1, label="Summarise")
        self.summarise_button.Bind(wx.EVT_BUTTON, self.on_summarise)
        self.summarise_button.SetToolTip(wx.ToolTip("Summarise the contents of the .esx project"))

        self.export_ap_images_button = wx.Button(self.tab3, label="AP Images")
        self.export_ap_images_button.Bind(wx.EVT_BUTTON, self.on_export_ap_images)
        self.export_ap_images_button.SetToolTip(wx.ToolTip("Images from within AP notes"))

        self.export_map_note_images_button = wx.Button(self.tab3, label="Note Images")
        self.export_map_note_images_button.Bind(wx.EVT_BUTTON, self.on_export_map_note_images)
        self.export_map_note_images_button.SetToolTip(wx.ToolTip("Images from within map notes"))

        self.extract_blank_maps_button = wx.Button(self.tab2, label="Blank Maps")
        self.extract_blank_maps_button.Bind(wx.EVT_BUTTON, self.on_export_blank_maps)
        self.extract_blank_maps_button.SetToolTip(wx.ToolTip("No APs, no walls, nothing, just the map image as it was imported"))

        self.create_ap_location_maps_button = wx.Button(self.tab2, label="AP Location Maps")
        self.create_ap_location_maps_button.Bind(wx.EVT_BUTTON, self.on_create_ap_location_maps)
        self.create_ap_location_maps_button.SetToolTip(wx.ToolTip("Generate AP location maps with Ekahau style AP icons"))

        self.create_zoomed_ap_maps_button = wx.Button(self.tab2, label="Zoomed AP Maps")
        self.create_zoomed_ap_maps_button.Bind(wx.EVT_BUTTON, self.on_create_zoomed_ap_maps)
        self.create_zoomed_ap_maps_button.SetToolTip(wx.ToolTip("Generate zoomed per AP location maps with Ekahau style AP icons"))

        self.export_pds_maps_button = wx.Button(self.tab2, label="PDS Maps")
        self.export_pds_maps_button.Bind(wx.EVT_BUTTON, self.on_export_pds_maps)
        self.export_pds_maps_button.SetToolTip(wx.ToolTip("Generate maps with red circle AP markers for use during Post Deployment Surveys"))

        self.create_pds_project_button = wx.Button(self.tab3, label="Create PDS Project")
        self.create_pds_project_button.Bind(wx.EVT_BUTTON, self.on_create_pds_project)
        self.create_pds_project_button.SetToolTip(wx.ToolTip("Create a PDS project from the current .esx project"))

        self.create_surveyed_ap_list_button = wx.Button(self.tab3, label="Surveyed AP List")
        self.create_surveyed_ap_list_button.Bind(wx.EVT_BUTTON, self.on_create_surveyed_ap_list)
        self.create_surveyed_ap_list_button.SetToolTip(wx.ToolTip("Dump surveyed AP detail to XLSX"))

        self.perform_admin_action_button = wx.Button(self.tab4, label="Perform Action")
        self.perform_admin_action_button.Bind(wx.EVT_BUTTON, self.on_perform_admin_action)
        self.perform_admin_action_button.SetToolTip(wx.ToolTip("Perform the selected Admin action"))

        self.check_for_updates_button = wx.Button(self.tab4, label="Check for Updates")
        self.check_for_updates_button.Bind(wx.EVT_BUTTON, self.on_check_for_updates)
        self.check_for_updates_button.SetToolTip(wx.ToolTip("Check for new commits on GitHub"))

        self.feedback_button = wx.Button(self.tab4, label="Feedback")
        self.feedback_button.Bind(wx.EVT_BUTTON, self.on_feedback)
        self.feedback_button.SetToolTip(wx.ToolTip("Send feedback to the developer"))

        self.contribute_button = wx.Button(self.tab4, label="Contribute")
        self.contribute_button.Bind(wx.EVT_BUTTON, self.on_contribute)
        self.contribute_button.SetToolTip(wx.ToolTip("Buy the developer a coffee"))

        self.preview_directory_structure_button = wx.Button(self.tab4, label="Preview")
        self.preview_directory_structure_button.Bind(wx.EVT_BUTTON, self.on_preview_directory_structure)
        self.preview_directory_structure_button.SetToolTip(wx.ToolTip("Preview selected directory structure profile"))

        self.create_directory_structure_button = wx.Button(self.tab4, label="Create Directory Structure")
        self.create_directory_structure_button.Bind(wx.EVT_BUTTON, self.on_create_directory_structure)
        self.create_directory_structure_button.SetToolTip(wx.ToolTip("Specify a root dir, create selected directory structure, presumably for a new project"))

        # Create an abort thread button
        self.abort_thread_button = wx.Button(self.panel, label="Abort Current Process")
        self.abort_thread_button.Bind(wx.EVT_BUTTON, self.on_abort_thread)

        # Create exit button
        self.exit_button = wx.Button(self.panel, label="Exit")
        self.exit_button.Bind(wx.EVT_BUTTON, self.on_exit)

    def setup_text_input_boxes(self):
        # Create a text input box for rename function start number
        self.rename_start_number_text_box = wx.TextCtrl(self.tab1, value="1", style=wx.TE_PROCESS_ENTER)

        # Create a text input box for the AP icon size
        self.ap_icon_size_text_box = wx.TextCtrl(self.tab2, value="25", style=wx.TE_PROCESS_ENTER)

        # Create a text input box for the AP name label size
        self.ap_name_label_size_text_box = wx.TextCtrl(self.tab2, value="30", style=wx.TE_PROCESS_ENTER)

        # Create a text input box for the zoomed AP image crop size
        self.zoomed_ap_crop_text_box = wx.TextCtrl(self.tab2, value="2000", style=wx.TE_PROCESS_ENTER)

    def setup_text_labels(self):
        # Create a text label for the drop target with custom position
        self.drop_target_label = wx.StaticText(self.panel, label="Drag and Drop files here", pos=(22, 17))

        # Create a text label for the rename start number text box
        self.rename_start_number_label = wx.StaticText(self.tab1, label="Start Number:")

        # Create a text label for the Create Simulated AP List function
        self.create_ap_list_label = wx.StaticText(self.tab1, label="Export to Excel:")

        # Create a text label for the AP icon size
        self.ap_icon_size_label = wx.StaticText(self.tab2, label="AP Icon Size:")
        self.ap_icon_size_label.SetToolTip(wx.ToolTip("Enter the size of the AP icon in pixels"))

        # Create a text label for the ap name label size
        self.ap_name_label_size_label = wx.StaticText(self.tab2, label="Text Size:")
        self.ap_name_label_size_label.SetToolTip(wx.ToolTip("Enter the font size for the AP name label"))

        # Create a text label for the zoomed AP crop size text box
        self.zoomed_ap_crop_label = wx.StaticText(self.tab2, label="Zoomed AP Crop Size:")
        self.zoomed_ap_crop_label.SetToolTip(wx.ToolTip("Enter the size of the zoomed AP crop in pixels"))

        # Create a text label for the Create Surveyed AP List function
        self.create_surveyed_ap_list_label = wx.StaticText(self.tab3, label="Export to Excel:")

    def setup_panel_rows(self):
        self.button_row1_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.button_row1_sizer.Add(self.add_files_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)
        self.button_row1_sizer.AddStretchSpacer(1)
        self.button_row1_sizer.Add(self.open_working_directory_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)
        self.button_row1_sizer.Add(self.reset_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)

        self.button_row2_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.button_row2_sizer.Add(self.copy_log_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)
        self.button_row2_sizer.Add(self.clear_log_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)
        self.button_row2_sizer.AddStretchSpacer(1)
        self.button_row2_sizer.Add(self.unpack_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)
        self.button_row2_sizer.Add(self.rebundle_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)
        self.button_row2_sizer.Add(self.backup_button, 0, wx.LEFT | wx.RIGHT | wx.BOTTOM, self.widget_margin)

        self.button_exit_row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.button_exit_row_sizer.AddStretchSpacer(1)
        self.button_exit_row_sizer.Add(self.abort_thread_button, 0, wx.ALL, self.widget_margin)
        self.button_exit_row_sizer.Add(self.exit_button, 0, wx.ALL, self.widget_margin)

    def setup_main_sizer(self):
        main_sizer = wx.BoxSizer(wx.VERTICAL)
        main_sizer.Add(self.list_box, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        main_sizer.Add(self.button_row1_sizer, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, self.sizer_edge_margin)
        main_sizer.Add(self.display_log, 1, wx.EXPAND | wx.ALL, self.widget_margin)
        main_sizer.Add(self.button_row2_sizer, 0, wx.EXPAND | wx.ALL, self.sizer_edge_margin)
        main_sizer.Add(self.notebook, 0, wx.EXPAND | wx.ALL, self.notebook_margin)
        main_sizer.Add(self.button_exit_row_sizer, 0, wx.EXPAND | wx.ALL, self.sizer_edge_margin)

        self.panel.SetSizer(main_sizer)

    def setup_tabs(self):
        self.notebook = wx.Notebook(self.panel)
        self.tab1 = wx.Panel(self.notebook)
        self.tab2 = wx.Panel(self.notebook)
        self.tab3 = wx.Panel(self.notebook)
        self.tab4 = wx.Panel(self.notebook)

        self.notebook.AddPage(self.tab1, "Predictive Design")
        self.notebook.AddPage(self.tab2, "Asset Creator")
        self.notebook.AddPage(self.tab3, "Survey")
        self.notebook.AddPage(self.tab4, "Admin")

        self.tab1_sizer = wx.BoxSizer(wx.VERTICAL)
        self.tab2_sizer = wx.BoxSizer(wx.VERTICAL)
        self.tab3_sizer = wx.BoxSizer(wx.VERTICAL)
        self.tab4_sizer = wx.BoxSizer(wx.VERTICAL)

        # Bind the tab change event
        self.notebook.Bind(wx.EVT_NOTEBOOK_PAGE_CHANGED, self.on_tab_changed)

    def setup_tab1(self):
        self.tab1_sizer = wx.BoxSizer(wx.VERTICAL)
        self.tab1_row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.setup_design_project_profile_section()
        self.setup_rename_aps_section()

        self.tab1_row_sizer.Add(self.project_profile_sizer, 1, wx.EXPAND | wx.ALL, 2)
        self.tab1_row_sizer.Add(self.rename_aps_sizer, 1, wx.EXPAND | wx.ALL, 2)

        self.tab1_sizer.Add(self.tab1_row_sizer, 0, wx.EXPAND | wx.TOP, self.sizer_edge_margin)
        self.tab1.SetSizer(self.tab1_sizer)

    def setup_design_project_profile_section(self):
        self.project_profile_box = wx.StaticBox(self.tab1, label="Project Profile")
        self.project_profile_sizer = wx.StaticBoxSizer(self.project_profile_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.project_profile_dropdown, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        row_sizer.Add(self.validate_button, 0, wx.ALL, self.widget_margin)
        row_sizer.Add(self.summarise_button, 0, wx.ALL, self.widget_margin)
        self.project_profile_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

        # Row 2
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.create_ap_list_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.create_ap_list, 0, wx.ALL, self.widget_margin)
        self.project_profile_sizer.Add(row_sizer, 0, wx.LEFT, self.row_sizer_margin)

    def setup_rename_aps_section(self):
        self.rename_aps_box = wx.StaticBox(self.tab1, label="Rename APs")
        self.rename_aps_sizer = wx.StaticBoxSizer(self.rename_aps_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.ap_rename_script_dropdown, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        row_sizer.Add(self.rename_start_number_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.rename_start_number_text_box, 1, wx.EXPAND | wx.ALL, self.widget_margin)
        row_sizer.Add(self.rename_aps_button, 0, wx.ALL, self.widget_margin)
        row_sizer.AddStretchSpacer()
        self.rename_aps_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

        # Row 2
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.visualise_ap_renaming_button, 0, wx.ALL, self.widget_margin)
        self.rename_aps_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

    def setup_tab2(self):
        self.tab2_sizer = wx.BoxSizer(wx.VERTICAL)
        self.tab2_row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.setup_export_section()
        self.setup_create_section()

        self.tab2_row_sizer.Add(self.export_sizer, 1, wx.EXPAND | wx.ALL, 2)
        self.tab2_row_sizer.Add(self.create_sizer, 1, wx.EXPAND | wx.ALL, 2)

        self.tab2_sizer.Add(self.tab2_row_sizer, 0, wx.EXPAND | wx.TOP, self.sizer_edge_margin)
        self.tab2.SetSizer(self.tab2_sizer)

    def setup_export_section(self):
        self.export_box = wx.StaticBox(self.tab2, label="Export")
        self.export_sizer = wx.StaticBoxSizer(self.export_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.extract_blank_maps_button, 0, wx.ALL, self.widget_margin)
        self.export_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

    def setup_create_section(self):
        self.create_box = wx.StaticBox(self.tab2, label="Create Map Assets")
        self.create_sizer = wx.StaticBoxSizer(self.create_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.create_ap_location_maps_button, 0, wx.ALL, self.widget_margin)
        row_sizer.Add(self.create_zoomed_ap_maps_button, 0, wx.ALL, self.widget_margin)
        row_sizer.Add(self.export_pds_maps_button, 0, wx.ALL, self.widget_margin)
        self.create_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

        # Row 2
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.ap_icon_size_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.ap_icon_size_text_box, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.ap_name_label_size_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.ap_name_label_size_text_box, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.zoomed_ap_crop_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.zoomed_ap_crop_text_box, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        self.create_sizer.Add(row_sizer, 0, wx.EXPAND, wx.LEFT, self.row_sizer_margin)


    def setup_tab3(self):
        self.tab3_sizer = wx.BoxSizer(wx.VERTICAL)
        self.tab3_row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.setup_survey_project_profile_section()
        self.setup_survey_export_section()

        self.tab3_row_sizer.Add(self.survey_project_profile_section_sizer, 1, wx.EXPAND | wx.ALL, 2)
        self.tab3_row_sizer.Add(self.survey_export_section_sizer, 1, wx.EXPAND | wx.ALL, 2)

        self.tab3_sizer.Add(self.tab3_row_sizer, 0, wx.EXPAND | wx.TOP, self.sizer_edge_margin)
        self.tab3.SetSizer(self.tab3_sizer)

    def setup_survey_project_profile_section(self):
        self.survey_project_profile_box = wx.StaticBox(self.tab3, label="Project Profile")
        self.survey_project_profile_section_sizer = wx.StaticBoxSizer(self.survey_project_profile_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.survey_project_profile_dropdown, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        row_sizer.Add(self.create_pds_project_button, 0, wx.ALL, self.widget_margin)
        self.survey_project_profile_section_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

        # Row 2
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.create_surveyed_ap_list_label, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, self.widget_margin)
        row_sizer.Add(self.create_surveyed_ap_list_button, 0, wx.ALL, self.widget_margin)
        self.survey_project_profile_section_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

    def setup_survey_export_section(self):
        self.survey_export_box = wx.StaticBox(self.tab3, label="Export")
        self.survey_export_section_sizer = wx.StaticBoxSizer(self.survey_export_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.survey_export_section_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

        # Row 2
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.export_ap_images_button, 0, wx.ALL, self.widget_margin)
        row_sizer.Add(self.export_map_note_images_button, 0, wx.ALL, self.widget_margin)
        self.survey_export_section_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

    def setup_tab4(self):
        self.tab4_sizer = wx.BoxSizer(wx.VERTICAL)
        self.tab4_row_sizer = wx.BoxSizer(wx.HORIZONTAL)

        self.setup_application_section()
        self.setup_misc_section()

        self.tab4_row_sizer.Add(self.application_section_sizer, 1, wx.EXPAND | wx.ALL, 2)
        self.tab4_row_sizer.Add(self.misc_sizer, 1, wx.EXPAND | wx.ALL, 2)

        self.tab4_sizer.Add(self.tab4_row_sizer, 0, wx.EXPAND | wx.TOP, self.sizer_edge_margin)
        self.tab4.SetSizer(self.tab4_sizer)

    def setup_application_section(self):
        self.application_box = wx.StaticBox(self.tab4, label="Application")
        self.application_section_sizer = wx.StaticBoxSizer(self.application_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.check_for_updates_button, 0, wx.ALL, self.widget_margin)
        row_sizer.AddStretchSpacer(1)
        row_sizer.Add(self.feedback_button, 0, wx.ALL, self.widget_margin)
        row_sizer.Add(self.contribute_button, 0, wx.ALL, self.widget_margin)
        self.application_section_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

        # Row 2
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.admin_actions_dropdown, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        row_sizer.Add(self.perform_admin_action_button, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        self.application_section_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

    def setup_misc_section(self):
        self.misc_box = wx.StaticBox(self.tab4, label="Misc")
        self.misc_sizer = wx.StaticBoxSizer(self.misc_box, wx.VERTICAL)

        # Row 1
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.dir_structure_profile_dropdown, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        row_sizer.Add(self.preview_directory_structure_button, 0, wx.ALL, self.widget_margin)
        row_sizer.Add(self.create_directory_structure_button, 0, wx.ALL, self.widget_margin)
        self.misc_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

        # Row 2
        row_sizer = wx.BoxSizer(wx.HORIZONTAL)
        row_sizer.Add(self.project_detail_dropdown, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        row_sizer.Add(self.display_project_detail_button, 0, wx.EXPAND | wx.ALL, self.widget_margin)
        self.misc_sizer.Add(row_sizer, 0, wx.EXPAND | wx.LEFT, self.row_sizer_margin)

    def create_menu(self):
        menubar = wx.MenuBar()

        # File Menu
        file_menu = wx.Menu()
        file_menu.Append(wx.ID_ADD, "&Add Files", "Add files to the list")
        file_menu.Append(wx.ID_SAVE, "&Save", "Save the current configuration")
        file_menu.AppendSeparator()
        file_menu.Append(wx.ID_EXIT, "&Exit", "Exit the application")
        menubar.Append(file_menu, "&File")

        # Help Menu
        help_menu = wx.Menu()
        contribute_menu_item = help_menu.Append(wx.ID_ANY, "&Contribute", "Go to the ko-fi contribution page")
        documentation_menu_item = help_menu.Append(wx.ID_ANY, "&Documentation", "View the documentation")
        view_release_notes_menu_item = help_menu.Append(wx.ID_ANY, "&View 'Release Notes'", "View the GitHub commit messages")
        feedback_menu_item = help_menu.Append(wx.ID_ANY, "&Feedback", "Send feedback to the developer")
        help_menu.Append(wx.ID_ABOUT, '&About')
        menubar.Append(help_menu, '&Help')

        self.SetMenuBar(menubar)

        # Bind the menu items to their respective functions
        self.Bind(wx.EVT_MENU, self.on_about, id=wx.ID_ABOUT)
        self.Bind(wx.EVT_MENU, self.on_add_file, id=wx.ID_ADD)
        self.Bind(wx.EVT_MENU, self.on_save, id=wx.ID_SAVE)
        self.Bind(wx.EVT_MENU, self.on_exit, id=wx.ID_EXIT)

        self.Bind(wx.EVT_MENU, self.on_contribute, contribute_menu_item)
        self.Bind(wx.EVT_MENU, self.on_view_documentation, documentation_menu_item)
        self.Bind(wx.EVT_MENU, self.on_view_release_notes, view_release_notes_menu_item)
        self.Bind(wx.EVT_MENU, self.on_feedback, feedback_menu_item)

    @staticmethod
    def on_contribute(event):
        webbrowser.open("https://ko-fi.com/badgerwifitools")

    @staticmethod
    def on_view_documentation(event):
        webbrowser.open("https://badgerwifi.co.uk")

    @staticmethod
    def on_view_release_notes(event):
        webbrowser.open("https://github.com/nickjvturner/badgerwifi-tools/activity")

    def on_feedback(self, event):
        self.append_message(f"Opening feedback page... {nl}Please leave your feedback on the GitHub page.")
        webbrowser.open("https://github.com/nickjvturner/badgerwifi-tools/issues")

    @staticmethod
    def on_about(event):
        # Implement the About dialog logic
        wx.MessageBox("This is a wxPython GUI application created by Nick Turner. Intended to make the lives of Wi-Fi engineers making reports a little bit easier. ", "About")

    def display_welcome_message(self):
        if random.random() < 0.5:  # 50% chance of displaying the welcome message
            welcome_message = random.choice(WHIMSY_WELCOME_MESSAGES)  # Select a random welcome message

            # Append the welcome message to the display log
            self.append_message(welcome_message)

        if random.random() < 0.3:
            self.append_message(CALL_TO_DONATE_MESSAGE)

    def display_message_on_reset(self):
        if random.random() < 0.1:  # 10% chance of displaying a welcome message
            welcome_message = random.choice(WHIMSY_WELCOME_MESSAGES)  # Select a random welcome message

            # Append the welcome message to the display log
            self.append_message(welcome_message)

    def load_module(self, module_subdir, module_name):
        module_path = Path(__file__).resolve().parent / module_subdir / f"{module_name}.py"
        spec = importlib.util.spec_from_file_location(module_name, str(module_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def on_create_surveyed_ap_list(self, event):
        if not hasattr(self.current_profile_ap_list_module, 'create_custom_measured_ap_list'):
            self.append_message("Currently selected project profile has no surveyed ap list export definition.")
            return
        elif not self.basic_checks():
            return
        else:
            create_surveyed_ap_list(self)

    def on_admin_actions_dropdown_selection(self, event):
        selected_index = self.admin_actions_dropdown.GetSelection()
        action_module = self.available_admin_actions[selected_index]
        self.current_admin_action_module = self.load_module(ADMIN_ACTIONS_DIR, action_module)

    def on_perform_admin_action(self, event):
        self.current_admin_action_module.run(self)

    def on_check_for_updates(self, event):
        check_for_updates.check_for_updates(self.append_message)

    def on_project_detail_dropdown_selection(self, event):
        selected_index = self.project_detail_dropdown.GetSelection()
        project_detail_module = self.available_project_detail_views[selected_index]
        self.current_project_detail_module = self.load_module(PROJECT_DETAIL_DIR, project_detail_module)

    def on_display_project_detail(self, event):
        if not self.basic_checks():
            return
        self.current_project_detail_module.run(self.working_directory, self.project_name, self.append_message)

    def on_dir_structure_profile_dropdown_selection(self, event):
        selected_index = self.dir_structure_profile_dropdown.GetSelection()
        dir_structure_profile = self.available_dir_structure_profiles[selected_index]
        self.current_dir_structure_profile = self.load_module(DIR_STRUCTURE_PROFILES_DIR, dir_structure_profile)

    def on_save(self, event):
        self.save_application_state(event)
        message = f'Application state saved on exit, file list and dropdown options should be the same next time you launch the app.'
        self.append_message(message)
        print(message)

    def setup_drop_target(self):
        """Set up the drop target for the list box."""
        self.allowed_extensions = (".esx")  # Define allowed file extensions
        self.drop_target = DropTarget(self)
        self.list_box.SetDropTarget(self.drop_target)

    def append_message(self, message):
        # Append a message to the message display area.
        self.display_log.AppendText(message + '\n')

    def update_last_message(self, message):
        content = self.display_log.GetValue()

        # Find the last occurrence of a newline character
        last_newline_index = content.rfind('\n')

        # If there's at least one newline, replace the text after the last newline
        if last_newline_index != -1:
            # Calculate the start position for replacement (after the newline character)
            start_pos = last_newline_index + 1  # Start replacing after the newline

            # Use Replace method to replace the last line
            self.display_log.Replace(start_pos, self.display_log.GetLastPosition(), message)
        else:
            # If there's no newline, this means there's only one line, so we can directly set the value
            self.display_log.SetValue(message)

    def save_application_state(self, event):
        """Save the application state to the defined path."""
        state = {
            'list_box_contents': [self.list_box.GetString(i) for i in range(self.list_box.GetCount())],
            'selected_ap_rename_script_index': self.ap_rename_script_dropdown.GetSelection(),
            'selected_project_profile_index': self.project_profile_dropdown.GetSelection(),
            'selected_project_detail_index': self.project_detail_dropdown.GetSelection(),
            'selected_admin_actions_index': self.admin_actions_dropdown.GetSelection(),
            'selected_dir_structure_profile_index': self.dir_structure_profile_dropdown.GetSelection(),
            'selected_tab_index': self.notebook.GetSelection(),
            'ap_icon_size_text_box': self.ap_icon_size_text_box.GetValue(),
            'zoomed_ap_crop_text_box': self.zoomed_ap_crop_text_box.GetValue(),
            'boundary_separator_value': self.rename_aps_boundary_separator
        }
        # Save the state to the defined path
        with open(self.app_state_file_path, 'w') as f:
            json.dump(state, f)

    def load_application_state(self):
        """Load the application state from the defined path."""
        try:
            with open(self.app_state_file_path, 'r') as f:
                state = json.load(f)

                # Restore list box contents
                for item in state.get('list_box_contents', []):
                    if not file_or_dir_exists(item):
                        self.append_message(f"** WARNING ** {nl}The file {item} does not exist.")
                    self.list_box.Append(item)
                    self.drop_target_label.Hide()  # Hide the drop target label

                # Restore the boundary separator value
                self.rename_aps_boundary_separator = state.get('boundary_separator_value', 400)

                # Restore selected ap rename script index
                self.ap_rename_script_dropdown.SetSelection(state.get('selected_ap_rename_script_index', 0))
                self.on_ap_rename_script_dropdown_selection(None)

                # Restore selected project profile index
                self.project_profile_dropdown.SetSelection(state.get('selected_project_profile_index', 0))
                self.survey_project_profile_dropdown.SetSelection(state.get('selected_project_profile_index', 0))
                self.on_project_profile_dropdown_selection(None)

                # Restore selected project detail index
                self.project_detail_dropdown.SetSelection(state.get('selected_project_detail_index', 0))
                self.on_project_detail_dropdown_selection(None)

                # Restore selected admin action index
                self.admin_actions_dropdown.SetSelection(state.get('selected_admin_actions_index', 0))
                self.on_admin_actions_dropdown_selection(None)

                # Restore selected tab index
                self.notebook.SetSelection(state.get('selected_tab_index', 0))

                # Restore the text box values
                self.ap_icon_size_text_box.SetValue(state.get('ap_icon_size_text_box', "25"))
                self.zoomed_ap_crop_text_box.SetValue(state.get('zoomed_ap_crop_text_box', "2000"))

                # Restore the directory structure profile index
                self.dir_structure_profile_dropdown.SetSelection(state.get('selected_dir_structure_profile_index', 0))
                self.on_dir_structure_profile_dropdown_selection(None)

        except FileNotFoundError:
            self.on_ap_rename_script_dropdown_selection(None)
            self.on_project_profile_dropdown_selection(None)
            pass  # It's okay if the state file doesn't exist on first run

    def on_reset(self, event):
        self.list_box.Clear()  # Reset list_box contents
        self.display_log.SetValue("")  # Clear the contents of the display_log
        self.display_message_on_reset()
        self.esx_project_unpacked = False  # Reset project_unpacked state
        self.drop_target_label.Show()  # Show the drop target label
        self.ap_icon_size_text_box.SetValue("25")  # Reset the AP icon size
        self.zoomed_ap_crop_text_box.SetValue("2000")  # Reset the zoomed AP crop size
        self.rename_aps_boundary_separator = 200  # Reset the boundary separator value
        self.stop_event.clear()  # Clear the stop event

    def on_clear_log(self, event):
        self.display_log.SetValue("")  # Clear the contents of the display_log

    def on_add_file(self, event):
        wildcard = "Ekahau Project file (*.esx)|*.esx"
        dlg = wx.FileDialog(self, "Choose a file", wildcard=wildcard,
                            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE)
        if dlg.ShowModal() == wx.ID_OK:
            filepaths = dlg.GetPaths()
            for filepath in filepaths:
                self.drop_target.process_file(filepath)
        dlg.Destroy()

    def on_delete_key(self, event):
        keycode = event.GetKeyCode()
        if keycode in [wx.WXK_DELETE, wx.WXK_BACK]:
            selected_indices = self.list_box.GetSelections()
            for index in reversed(selected_indices):
                self.list_box.Delete(index)

    def on_unpack(self, event):
        if not self.esx_project_unpacked:
            self.unpack_esx()
            return
        # Create a message dialog with Yes and No buttons
        dlg = wx.MessageDialog(None, "Project has already been unpacked. Would you like to re-unpack it?",
                               "Re-unpack project?", wx.YES_NO | wx.ICON_QUESTION)
        # Show the dialog and check the response
        result = dlg.ShowModal()

        if result == wx.ID_YES:
            # User wants to re-unpack the project
            self.esx_project_unpacked = False
            self.unpack_esx()
        else:
            # User chose not to re-unpack the project
            self.append_message(f"Re-unpack operation aborted{nl}")
        # Destroy the dialog after using it
        dlg.Destroy()

    def on_backup(self, event):
        if not self.filepath:
            if not self.get_single_specific_file_type('.esx'):
                return
        backup_esx(self.working_directory, self.project_name, self.filepath, self.append_message)

    def on_validate(self, event):
        if not self.basic_checks():
            return
        validate_esx(self, self.append_message)

    def on_summarise(self, event):
        if not self.basic_checks():
            return
        summarise_esx(self.working_directory, self.project_name, self.append_message)

    def on_create_ap_list(self, event):
        if not self.basic_checks():
            return
        if hasattr(self, 'current_project_profile_module'):
            create_ap_list(self)

    def on_copy_log(self, event):
        if wx.TheClipboard.Open():
            wx.TheClipboard.SetData(wx.TextDataObject(self.display_log.GetValue()))
            wx.TheClipboard.Close()
            self.append_message("Log copied to clipboard.")
        else:
            wx.MessageBox("Unable to access the clipboard.", "Error", wx.OK | wx.ICON_ERROR)

    def on_exit(self, event):
        # Save the application state before exiting
        self.save_application_state(None)
        print(f'Application state saved on exit, file list and dropdown options should be the same next time you launch the application')
        cleanup_unpacked_project_folder(self)
        self.Close()
        self.Destroy()

    def get_single_specific_file_type(self, extension):
        for filepath in self.list_box.GetStrings():
            if filepath.lower().endswith(extension):
                if extension == '.esx':
                    self.filepath = Path(filepath)
                    if not self.filepath.exists():
                        self.append_message(f'The file {self.filepath} does not exist.')
                        return False

                    self.project_name = self.filepath.stem  # Set the project name based on the file stem
                    self.append_message(f'Project name: {self.project_name}')

                    self.working_directory = self.filepath.parent
                    self.append_message(f'Working directory: {self.working_directory}{nl}')

                    self.project_metadata = parse_project_metadata(self.project_name, self.project_filename_expected_pattern)
                    self.site_id = self.project_metadata.get('site_id', None)
                    self.site_location = self.project_metadata.get('site_location', None)
                    self.project_phase = self.project_metadata.get('project_phase', None)
                    self.project_version = self.project_metadata.get('project_version', None)
                return True

        self.append_message(f"No file with {extension} present in file list.")
        return False

    def unpack_esx(self):
        if not self.esx_project_unpacked:
            if not self.get_single_specific_file_type('.esx'):
                return
            unpack_esx_file(self.working_directory, self.project_name, self.filepath, self.append_message)
            self.esx_project_unpacked = True
        return True

    def load_project_profile(self, profile_name):
        profile_path = Path(__file__).resolve().parent / PROJECT_PROFILES_DIR / f"{profile_name}.py"
        spec = importlib.util.spec_from_file_location(profile_name, str(profile_path))
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module

    def on_project_profile_dropdown_selection(self, event):
        selected_profile = self.project_profile_dropdown.GetStringSelection()

        # Load the selected project profile module for simulated AP list creation
        project_profile_module = self.load_project_profile(selected_profile)
        self.project_profile_module = project_profile_module
        self.current_profile_ap_list_module = project_profile_module

        self.on_clear_log(None)
        self.append_message(f"Project profile {selected_profile} loaded.")

        # Temporary check to ensure everybody gets the latest project profiles from SharePoint
        if selected_profile not in example_project_profile_names:
            if not hasattr(self.project_profile_module, 'project_profile_id'):
                self.append_message(f"Selected project profile does not have a project_profile_id attribute. Please update the project profile module.")

            if not hasattr(self.project_profile_module, 'project_profile_version'):
                self.append_message(f"Selected project profile does not have a project_profile_version attribute. Please update the project profile module.")

            if hasattr(self.project_profile_module, 'project_profile_version'):
                if float(self.project_profile_module.project_profile_version) > 2:
                    self.append_message(f"Selected project profile has an outdated versioning scheme. Please update the project profile module.")

            else:
                self.append_message(f"Selected project profile does not have a project_profile_version attribute. Please update the project profile module.")

        # Update the object variables with the configuration from the selected module
        self.required_tag_keys = getattr(project_profile_module, 'requiredTagKeys', None)
        self.optional_tag_keys = getattr(project_profile_module, 'optionalTagKeys', None)
        self.project_filename_expected_pattern = getattr(project_profile_module, 'PROJECT_FILENAME_EXPECTED_PATTERN', None)
        self.predictive_design_coverage_requirements = getattr(project_profile_module, 'predictive_design_coverage_requirements', None)
        self.post_deployment_survey_coverage_requirements = getattr(project_profile_module, 'post_deployment_survey_coverage_requirements', None)
        if hasattr(project_profile_module, 'preferred_ap_rename_script'):
            self.ap_rename_script_dropdown.SetStringSelection(project_profile_module.preferred_ap_rename_script)
            self.on_ap_rename_script_dropdown_selection(None)
        if hasattr(project_profile_module, 'project_profile_id'):
            tracked_project_profile_check_for_update(project_profile_module, self.append_message)
        self.save_application_state(None)

    def on_design_project_profile_dropdown_selection(self, event):
        self.survey_project_profile_dropdown.SetSelection(self.project_profile_dropdown.GetSelection())
        self.on_project_profile_dropdown_selection(event)

    def on_survey_project_profile_dropdown_selection(self, event):
        self.project_profile_dropdown.SetSelection(self.survey_project_profile_dropdown.GetSelection())
        self.on_project_profile_dropdown_selection(event)

    def get_rename_start_number(self):
        """
        Retrieve and validate the start number from the text box.
        Returns:
            int: The validated start number.
        Raises:
            ValueError: If the input is not a valid number.
        """
        start_number_str = self.rename_start_number_text_box.GetValue()
        try:
            return int(start_number_str)
        except ValueError:
            raise ValueError("Please enter a valid number for the start number.")

    def on_rename_aps(self, event):
        if not self.basic_checks():
            return

        try:
            # Use the helper function to get the start number
            rename_start_number = self.get_rename_start_number()
        except ValueError as e:
            wx.MessageBox(str(e), "Error", wx.OK | wx.ICON_ERROR)
            return

        # Load the selected script
        selected_script = self.available_ap_rename_scripts[self.ap_rename_script_dropdown.GetSelection()]
        script_path = str(Path(__file__).resolve().parent / RENAME_APS_DIR / (selected_script + ".py"))

        # Load module from selected script
        script_module = SourceFileLoader(selected_script, script_path).load_module()

        # Pass the start number to the rename function
        ap_renamer(self.working_directory, self.project_name, script_module, self.append_message, self.rename_aps_boundary_separator, rename_start_number)

    def on_ap_rename_script_dropdown_selection(self, event):
        """Handle rename script selection change."""
        selected_script = self.ap_rename_script_dropdown.GetStringSelection()
        path_to_module = Path(__file__).resolve().parent / RENAME_APS_DIR / f'{selected_script}.py'

        _, short_description = self.get_ap_rename_tooltips(selected_script)  # Ignore script_name
        self.current_sorting_module = import_module_from_path(selected_script, path_to_module)

        self.ap_rename_script_dropdown.SetToolTip(wx.ToolTip(short_description))

        if event:
            self.display_boundary_separator_message()
            self.save_application_state(None)

    def display_boundary_separator_message(self):
        if hasattr(self.current_sorting_module, BOUNDARY_SEPARATION_WIDGET):
            self.display_log.SetValue("")  # Clear the contents of the display_log
            self.append_message(f"Selected AP renaming script contains a configurable boundary parameter{nl}Boundary separator value: {self.rename_aps_boundary_separator}")
        else:
            self.display_log.SetValue("")  # Clear the contents of the display_log

    def get_ap_rename_tooltips(self, script_name):
        script_path = str(Path(__file__).resolve().parent / RENAME_APS_DIR / f"{script_name}.py")
        spec = importlib.util.spec_from_file_location(script_name, script_path)
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        short_description = getattr(module, 'SHORT_DESCRIPTION', "No short description available.")
        return script_name, short_description

    def placeholder(self, event):
        self.append_message(f'No action implemented yet')

    def on_export_ap_images(self, event):
        if not self.basic_checks():
            return
        export_ap_images.export_ap_images(self)

    def on_export_map_note_images(self, event):
        if not self.basic_checks():
            return
        export_map_note_images.export_map_note_images(self)

    def on_export_pds_maps(self, event):
        if not self.basic_checks():
            return

        # Retrieve the number from the custom AP icon size text box
        ap_icon_size = self.ap_icon_size_text_box.GetValue()
        ap_name_label_size = self.ap_name_label_size_text_box.GetValue()

        # Clear the stop event flag before starting the thread
        self.stop_event.clear()

        try:
            ap_icon_size = int(ap_icon_size)  # Convert the input to a float
            ap_name_label_size = int(ap_name_label_size)  # Ensure the AP name label size value is an integer
            create_pds_maps_threaded(self.working_directory, self.project_name, self.append_message, ap_icon_size, ap_name_label_size, self.stop_event)

        except ValueError:
            # Handle the case where the input is not a valid number
            wx.MessageBox("Please enter a valid number", "Error", wx.OK | wx.ICON_ERROR)

    def on_create_ap_location_maps(self, event):
        if not self.basic_checks():
            return

        # Retrieve the numbers from the custom size text boxes as an integers
        self.ap_icon_size = int(self.ap_icon_size_text_box.GetValue())
        self.ap_name_label_size = int(self.ap_name_label_size_text_box.GetValue())

        # Clear the stop event flag before starting the thread
        self.stop_event.clear()

        try:
            create_custom_ap_location_maps_threaded(self)

        except ValueError:
            # Handle the case where the input is not a valid number
            wx.MessageBox("Please enter a valid number", "Error", wx.OK | wx.ICON_ERROR)

    def on_create_zoomed_ap_maps(self, event):
        if not self.basic_checks():
            return

        self.stop_event.clear()

        # Retrieve the number from the zoomed AP crop size text box
        zoomed_ap_crop_size = self.zoomed_ap_crop_text_box.GetValue()
        ap_icon_size = self.ap_icon_size_text_box.GetValue()
        ap_name_label_size = int(self.ap_name_label_size_text_box.GetValue())

        try:
            zoomed_ap_crop_size = int(zoomed_ap_crop_size)  # Convert the input to a float
            custom_ap_icon_size = int(ap_icon_size)  # Convert the input to a float
            create_zoomed_ap_location_maps_threaded(self.working_directory, self.project_name, self.append_message, zoomed_ap_crop_size, custom_ap_icon_size, ap_name_label_size, self.stop_event)
        except ValueError:
            # Handle the case where the input is not a valid number
            wx.MessageBox("Please enter a valid number", "Error", wx.OK | wx.ICON_ERROR)

    def on_export_blank_maps(self, event):
        if not self.basic_checks():
            return
        extract_blank_maps(self.working_directory, self.project_name, self.append_message)

    def on_create_pds_project(self, event):
        if not self.basic_checks():
            return
        create_pds_project_esx(self, self.append_message)

    def on_tab_changed(self, event):
        # Get the index of the newly selected tab
        self.save_application_state(None)
        # new_tab_index = event.GetSelection()
        # print(f"Tab changed to index {new_tab_index}")

        # It's important to call event.Skip() to ensure the event is not blocked
        event.Skip()

    def on_rebundle_esx(self, event):
        if not self.esx_project_unpacked:
            self.get_single_specific_file_type('.esx')
        # Check that working directory and project name directories exist
        if self.working_directory and (self.working_directory / self.project_name).exists():
            rebundle_project(self.working_directory, self.project_name, self.append_message)

    def basic_checks(self):
        if not self.esx_project_unpacked:
            if not self.unpack_esx():
                return False
        self.on_clear_log(None)
        return True

    def on_abort_thread(self, event):
        self.stop_event.set()

    def on_open_working_directory(self, event):
        if not self.esx_project_unpacked:
            self.append_message("Project has not been unpacked yet, the working directory is not defined.")
            return
        try:
            # Open the directory using the operating system's file navigator
            if os.name == 'nt':  # Windows
                os.startfile(self.working_directory)
            elif os.name == 'posix':  # macOS or Linux
                # Directly attempt to use 'open' on macOS
                subprocess.Popen(['open', self.working_directory], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except Exception as e:
            wx.MessageBox(f"Error opening directory: {str(e)}", "Error", wx.OK | wx.ICON_ERROR)

    def show_replace_dialog(self, filepath):
        # Dialog asking if the user wants to replace the existing .esx file
        dlg = wx.MessageDialog(self.panel,
                               f"A .esx file is already present. Do you want to replace it with {Path(filepath).name}?",
                               "Replace File?", wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal() == wx.ID_YES
        dlg.Destroy()
        return result

    def on_visualise_ap_renaming(self, event):
        if not self.basic_checks():
            return

        visualise_ap_renaming(self.working_directory, self.project_name, self.append_message, self)

    def update_boundary_separator_value(self, value):
        """Callback function to update the boundary_separator variable."""
        self.rename_aps_boundary_separator = value
        self.append_message(f"Boundary separator value changed to: {value}")

    def update_ap_rename_script_dropdown_selection(self, index):
        self.ap_rename_script_dropdown.SetSelection(index)
        self.on_ap_rename_script_dropdown_selection(None)

    def check_for_updates_on_startup(self):
        try:
            latest_sha = check_for_updates.get_latest_commit_sha()
            local_commit_sha = check_for_updates.get_git_commit_sha()
            if latest_sha != local_commit_sha:
                self.append_message(f"** Update available **{nl}")
        except Exception as e:
            print(e)  # Check for updates failed

    def on_create_directory_structure(self, event):
        self.on_clear_log(None)
        select_root_and_create_directory_structure(self)

    def on_preview_directory_structure(self, event):
        self.on_clear_log(None)
        preview_directory_structure(self)

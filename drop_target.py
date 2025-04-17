# drop_target.py

import wx
import re
from common import nl
from pathlib import Path


class DropTarget(wx.FileDropTarget):
    def __init__(self, frame):
        super(DropTarget, self).__init__()
        self.frame = frame
        self.window = frame.list_box
        self.message_callback = frame.append_message
        # self.drop_target_label_callback = frame.drop_target_label_callback

    def OnDragOver(self, x, y, d):
        # This method can be used to provide feedback while moving over the target
        return wx.DragCopy

    def process_file(self, filepath):
        filepath = Path(filepath)
        filename = filepath.name

        if "re-zip" in filename:
            self.message_callback(f"{filename} cannot be added because it contains 're-zip' in the name.")
            return

        if filepath.suffix in self.frame.ignored_extensions:
            return

        if not filepath.suffix.lower() in self.frame.allowed_extensions:
            self.message_callback(f"{filename} has an unsupported extension.")
            return

        existing_files = self.window.GetStrings()
        if str(filepath) in existing_files:
            self.message_callback(f"{filename} is already in the list.")
            return

        if not filepath.exists():
            self.message_callback(f"File not found: {filepath}")
            return

        if filepath.suffix.lower() == '.esx':
            existing_esx_in_list = False
            for index, existing_file in enumerate(existing_files):
                if existing_file.lower().endswith('.esx'):
                    existing_esx_in_list = True
                    if not Path(existing_file).exists():
                        self.window.Delete(index)
                        self.window.Append(str(filepath))
                        self.message_callback(f"{existing_file} replaced with {filename}{nl}")
                        self.frame.esx_project_unpacked = False
                    if self.show_replace_dialog(filepath):
                        self.message_callback(f"{Path(existing_file).name} removed.")
                        self.window.Delete(index)
                        self.window.Append(str(filepath))
                        self.message_callback(f"{filename} added to the list.{nl}")
                        self.frame.esx_project_unpacked = False
                    else:
                        return
                    break
            if not existing_esx_in_list:
                self.window.Append(str(filepath))
                self.message_callback(f"{filename} added to the list.")

            project_name = filepath.stem

            self.frame.project_phase = None
            self.frame.project_version = None
            self.frame.site_reference = None

            if self.frame.predictive_design_expected_pattern or self.frame.post_deployment_survey_expected_pattern:

                pd_pattern = self.frame.predictive_design_expected_pattern
                pds_pattern = self.frame.post_deployment_survey_expected_pattern

                if pd_pattern and (match := re.search(pd_pattern, project_name)):
                    self.frame.site_id = match.groupdict().get("site_id")
                    self.frame.site_location = match.groupdict().get("site_location")
                    self.frame.project_phase = 'Predictive Design'
                    self.frame.project_version = match.groupdict().get("version")
                elif pds_pattern and (match := re.search(pds_pattern, project_name)):
                    self.frame.site_id = match.groupdict().get("site_id")
                    self.frame.site_location = match.groupdict().get("site_location")
                    self.frame.project_phase = 'Post-Deployment Survey'
                    self.frame.project_version = match.groupdict().get("version")

                else:
                    self.message_callback(f"Filename does not follow required convention for phase and version detection, you may proceed, but automated export filename generation will not be available.")

            if self.frame.site_id:
                self.message_callback(f"Detected Site ID: {self.frame.site_id}")
            if self.frame.site_location:
                self.message_callback(f"Detected Site Location: {self.frame.site_location}")
            if self.frame.project_phase:
                self.message_callback(f"Detected Phase: {self.frame.project_phase}")
            if self.frame.project_version:
                self.message_callback(f"Detected Version: {self.frame.project_version}")

        elif filepath.suffix.lower() == '.xlsx':
            self.window.Append(str(filepath))
            self.message_callback(f".xlsx File: {filename}")

        else:
            self.message_callback(f"Unsupported file type: {filename}")
            return

        if self.window.GetCount() > 0:
            self.frame.drop_target_label.Hide()
            self.frame.panel.Refresh()

    def OnDropFiles(self, x, y, filenames):
        for dropped_item in filenames:
            dropped_path = Path(dropped_item)
            if dropped_path.is_dir():
                # Recursively process all files in the dropped directory
                for file in dropped_path.rglob('*'):
                    self.process_file(str(file))
            else:
                self.process_file(dropped_item)

        return True

    def show_replace_dialog(self, filepath):
        # Dialog asking if the user wants to replace the existing .esx file
        dlg = wx.MessageDialog(self.window,
                               f"An .esx file is already present. Do you want to replace it with {Path(filepath).name}?",
                               "Replace File?", wx.YES_NO | wx.ICON_QUESTION)
        result = dlg.ShowModal()
        dlg.Destroy()
        return result == wx.ID_YES

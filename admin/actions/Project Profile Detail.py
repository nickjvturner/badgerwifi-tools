
def run(self):
    message_callback = self.append_message
    message_callback(f'Current Project Profile: {self.project_profile_module.project_profile_name}')
    if hasattr(self.project_profile_module, 'project_profile_version'):
        message_callback(f'Current Project Profile Version: {self.project_profile_module.project_profile_version}')

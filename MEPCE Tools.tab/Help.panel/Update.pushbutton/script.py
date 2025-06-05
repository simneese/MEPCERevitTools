# -*- coding: utf-8 -*-
__title__ = "Update"
__doc__ = """Version = 1.0
Date    = 2025.06.05
_________________________________________________________________
Description:
This button will update the pyRevit environment.
_________________________________________________________________
How-to:
-> Click button
-> Update pyRevit
-> Done!
_________________________________________________________________
Last update:
- [2025.06.05] - 1.0 RELEASE
_________________________________________________________________
To-Do:
-
_________________________________________________________________
Author: Simeon Neese"""

import os
from Autodesk.Revit.UI import TaskDialog, TaskDialogCommonButtons, TaskDialogResult

def find_in_path(filename):
    """Search for an executable in the system PATH (Python 2.7 compatible)."""
    paths = os.environ.get("PATH", "").split(os.pathsep)
    for path in paths:
        full_path = os.path.join(path, filename)
        if os.path.isfile(full_path) and os.access(full_path, os.X_OK):
            return full_path
    return None

def find_pyrevit_cli():
    # Common pyRevit CLI paths
    possible_paths = []

    # %APPDATA%\pyRevit\bin\pyrevit.exe
    appdata = os.environ.get("APPDATA")
    if appdata:
        possible_paths.append(os.path.join(appdata, "pyRevit", "bin", "pyrevit.exe"))

    # %LOCALAPPDATA%\pyRevit\bin\pyrevit.exe
    localappdata = os.environ.get("LOCALAPPDATA")
    if localappdata:
        possible_paths.append(os.path.join(localappdata, "pyRevit", "bin", "pyrevit.exe"))

    for path in possible_paths:
        if os.path.exists(path):
            return path

    # Fallback: check in system PATH
    return find_in_path("pyrevit.exe")

# === UI Dialog: Confirm update ===
dialog = TaskDialog("Update pyRevit")
dialog.MainInstruction = "Do you want to update pyRevit now?"
dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.No
dialog.DefaultButton = TaskDialogResult.Yes
result = dialog.Show()

if result == TaskDialogResult.Yes:
    cli_path = find_pyrevit_cli()

    if not cli_path:
        error_dialog = TaskDialog("Update Result")
        error_dialog.MainInstruction = "Could not locate pyrevit.exe"
        error_dialog.MainContent = "Check if pyRevit is installed and CLI is available."
        error_dialog.Show()
    else:
        # Build and run the command
        command = '"{}" update --all'.format(cli_path)
        exit_code = os.system(command)

        result_dialog = TaskDialog("Update Result")
        if exit_code == 0:
            result_dialog.MainInstruction = "pyRevit update completed successfully!"
        else:
            result_dialog.MainInstruction = "pyRevit update failed."
            result_dialog.MainContent = "Exit Code: {}".format(exit_code)
        result_dialog.CommonButtons = TaskDialogCommonButtons.Close
        result_dialog.Show()
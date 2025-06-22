import zipfile
import os
import re
import shutil
import traceback

# Consistent log file path in the desired directory
log_path = os.path.join(
    os.environ.get("APPDATA", ""),
    "Ace Centre",
    "AACSpeakHelper",
    "creategridset_log.txt",
)


def write_log(message):
    """Writes a message to the log file."""
    with open(log_path, "a") as log_file:
        log_file.write(f"{message}\n")


def create_folder_shortcut(target_folder, shortcut_name):
    try:
        desktop = os.path.join(os.path.join(os.environ["USERPROFILE"]), "Desktop")
        shortcut_path = os.path.join(desktop, f"{shortcut_name}.lnk")
        write_log(f"Creating shortcut at: {shortcut_path}")

        vbs_script = f"""
        Set oWS = WScript.CreateObject("WScript.Shell")
        sLinkFile = "{shortcut_path}"
        Set oLink = oWS.CreateShortcut(sLinkFile)
        oLink.TargetPath = "{target_folder}"
        oLink.Save
        """

        vbs_script_path = os.path.join(desktop, "create_shortcut.vbs")
        with open(vbs_script_path, "w") as file:
            file.write(vbs_script)
        write_log(f"VBS script created at: {vbs_script_path}")

        result = os.system(f'cscript //Nologo "{vbs_script_path}"')
        write_log(f"Executed VBS script with result: {result}")
        os.remove(
            vbs_script_path
        )  # Clean up the VBS script file after creating the shortcut
        write_log("VBS script removed successfully.")

    except Exception as e:
        write_log(f"Failed to create shortcut: {e}")
        write_log(traceback.format_exc())


def modify_gridset(gridset_path, LocalAppPath):
    if not os.path.exists(gridset_path):
        write_log(f"Error: The gridset file does not exist: {gridset_path}")
        return  # Exit the function if the file doesn't exist

    try:
        temp_dir = "temp_gridset"
        os.makedirs(temp_dir, exist_ok=True)
        write_log(f"Created temporary directory: {temp_dir}")

        with zipfile.ZipFile(gridset_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
        write_log(f"Extracted gridset: {gridset_path}")

        for foldername, _, filenames in os.walk(temp_dir):
            for filename in filenames:
                if filename.endswith(".xml"):
                    xml_path = os.path.join(foldername, filename)

                    with open(xml_path, "r") as f:
                        filedata = f.read()

                    local_app_data_path = os.environ.get("LOCALAPPDATA", "")
                    full_path_to_exe = os.path.join(
                        local_app_data_path,
                        "Programs",
                        "Ace Centre",
                        "AACSpeakHelper",
                        "client.exe",
                    )
                    full_path_to_exe_escaped = full_path_to_exe.replace("\\", "\\\\")
                    new_data = re.sub(
                        "%FILEPATHTOREPLACE%", full_path_to_exe_escaped, filedata
                    )

                    with open(xml_path, "w") as f:
                        f.write(new_data)
                    write_log(f"Modified XML file: {xml_path}")

        # Updated to keep everything under Ace Centre\AACSpeakHelper
        new_gridset_dir = os.path.join(
            LocalAppPath, "Ace Centre", "AACSpeakHelper", "Example AAC Helper Pages"
        )
        new_gridset_path = os.path.join(new_gridset_dir, "AAC Helper Tool Demo.gridset")

        os.makedirs(new_gridset_dir, exist_ok=True)
        write_log(f"Created new gridset directory: {new_gridset_dir}")

        with zipfile.ZipFile(new_gridset_path, "w") as zipf:
            for root, _, files in os.walk(temp_dir):
                for file in files:
                    zipf.write(
                        os.path.join(root, file),
                        os.path.relpath(os.path.join(root, file), temp_dir),
                    )
        write_log(f"Created new gridset file: {new_gridset_path}")

        # Clean up
        shutil.rmtree(temp_dir)
        write_log(f"Removed temporary directory: {temp_dir}")

        # Prevent deleting the original file
        # os.remove(gridset_path)
        write_log(f"Retained original gridset file: {gridset_path}")

        # Create a shortcut on the Desktop
        create_folder_shortcut(new_gridset_dir, "Example AAC Pages")

    except Exception as e:
        write_log(f"Error during modify_gridset: {e}")
        write_log(traceback.format_exc())


if __name__ == "__main__":
    try:
        write_log("Script started.")
        app_data_path = os.environ.get("APPDATA", "")
        gridset_location = os.path.join(
            app_data_path,
            "Ace Centre",
            "AACSpeakHelper",
            "TranslateAndTTS DemoGridset.gridset",
        )

        write_log(f"Gridset location: {gridset_location}")
        modify_gridset(gridset_location, app_data_path)
        write_log("Script completed successfully.")

    except Exception as e:
        write_log(f"Critical error in main execution: {e}")
        write_log(traceback.format_exc())

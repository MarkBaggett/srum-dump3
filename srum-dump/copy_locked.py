import os
import win32com.client
import subprocess
import pathlib

from ui_tk import ProgressWindow


def create_shadow_copy(volume_path):
    wmi_service = win32com.client.GetObject("winmgmts:\\\\.\\root\\cimv2")
    shadow_copy_class = wmi_service.Get("Win32_ShadowCopy")
    in_params = shadow_copy_class.Methods_("Create").InParameters.SpawnInstance_()
    in_params.Volume = volume_path
    in_params.Context = "ClientAccessible"
    out_params = wmi_service.ExecMethod("Win32_ShadowCopy", "Create", in_params)
    if out_params.ReturnValue == 0:
        shadow_id = out_params.ShadowID
        shadow_copy = wmi_service.ExecQuery(f"SELECT * FROM Win32_ShadowCopy WHERE ID='{shadow_id}'")[0]
        shadow_path = shadow_copy.DeviceObject.replace("\\\\?\\", "\\\\.\\", 1)
        return shadow_path
    else:
        raise Exception("Unable to create VSS.")


def extract_live_file(source, destination):
    esentutl_path = pathlib.Path(os.environ.get("COMSPEC")).parent.joinpath("esentutl.exe")
    if not esentutl_path.is_file():
        raise FileNotFoundError("esentutl.exe not found")
    if not pathlib.Path(source).is_file():
        raise FileNotFoundError("Source file not found")

    cmdline = f"{str(esentutl_path)} /y {source} /vss /d {destination}"
    result = subprocess.run(cmdline.split(), shell=True, capture_output=True)
    if result.returncode != 0:
        raise Exception(f"Failed to extract file {result.stderr.decode()}")
    return result.stdout.decode()


def copy_locked_files(destination_folder: pathlib.Path, repair_srum=True):
    """
    Copies a locked file using Volume Shadow Copy Service (VSS) and copy.
    
    :param source: Full path to the locked file (e.g., C:\\Windows\\System32\\SRU\\srudb.dat)
    :param destination: Path to save the copied file
    """
    ui_window = ProgressWindow("Extracting Locked files")
    ui_window.hide_record_stats()
    ui_window.start(6)
    ui_window.set_current_table("Creating Volume Shadow Copy")
    volume = pathlib.Path(os.environ["SystemRoot"]).drive
    ui_window.log_message(f"Creating a volume shadow copy for {volume}\n")
    success = True
    
    # Create the shadow copy

    shadow_path = create_shadow_copy(f"{volume}\\")
    if isinstance(shadow_path, int):
        ui_window.log_message(f"[-] Failed to create shadow copy: {shadow_path}\n")
        success = False
    ui_window.log_message(f"[+] Shadow Copy Device: {shadow_path}\n")

    if success:
        # Copy the sru directory to destination
        file_path = shadow_path + r"\Windows\system32\sru\*"
        ui_window.set_current_table("Copying SRU Folder")
        cmd_copy = f'copy /V "{file_path}" "{destination_folder}" '
        res = subprocess.run(cmd_copy, shell=True, capture_output=True)
        success = success and (res.returncode == 0)
        ui_window.log_message(res.stdout.decode() + res.stderr.decode())

        if repair_srum and success:
            # Repair srum based on log files
            ui_window.set_current_table("Recover SRU Folder")
            cmd_copy = f'esentutl.exe /r sru /i '
            ui_window.log_message(cmd_copy)
            res = subprocess.run(cmd_copy, shell=True, cwd=destination_folder, capture_output=True)
            success = success and (res.returncode == 0)
            ui_window.log_message(res.stdout.decode() + res.stderr.decode())
            # Repair srum based on log files
            ui_window.set_current_table("Repair SRUM Database")
            cmd_copy = f'esentutl.exe /p SRUDB.dat'
            ui_window.log_message(cmd_copy)
            res = subprocess.run(cmd_copy, shell=True, cwd=destination_folder, capture_output=True)
            success = success and (res.returncode == 0)
            ui_window.log_message(res.stdout.decode() + res.stderr.decode())
        

        # Copy the SOFTWARE key to destination
        file_path = shadow_path + r"\Windows\system32\config\SOFTWARE"
        ui_window.set_current_table("Copying SOFTWARE")
        cmd_copy = f'copy "{file_path}" "{destination_folder}" '
        ui_window.log_message(cmd_copy)
        res = subprocess.run(cmd_copy, shell=True, capture_output=True)
        success = success and (res.returncode == 0)
        ui_window.log_message(res.stdout.decode() + res.stderr.decode())

    if success:
        ui_window.close()
    else:
        ui_window.finished()
        ui_window.root.mainloop()

    return success


#copy_locked_files(r"c:\Users\mark\Desktop\output")
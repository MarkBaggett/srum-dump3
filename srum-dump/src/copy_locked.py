import os
import win32com.client
import subprocess
import re



def create_shadow_copy(volume_path):
    try:
        wmi_service = win32com.client.GetObject("winmgmts:\\\\.\\root\\cimv2")
        shadow_copy_class = wmi_service.Get("Win32_ShadowCopy")
        in_params = shadow_copy_class.Methods_("Create").InParameters.SpawnInstance_()
        in_params.Volume = volume_path
        in_params.Context = "ClientAccessible"
        out_params = wmi_service.ExecMethod("Win32_ShadowCopy", "Create", in_params)
        if out_params.ReturnValue == 0:
            return f"\\?\GLOBALROOT\Device\HarddiskVolumeShadowCopy{out_params.ShadowID}"
            shadow_id = out_params.ShadowID
            shadow_copy = wmi_service.ExecQuery(f"SELECT * FROM Win32_ShadowCopy WHERE ID='{shadow_id}'")[0]
            return shadow_copy.DeviceObject
        else:
            return out_params.ReturnValue
    except Exception as e:
        print(f"An error occurred: {e}")
        return None



def copy_locked_file(source, destination):
    """
    Copies a locked file using Volume Shadow Copy Service (VSS) and robocopy.
    
    :param source: Full path to the locked file (e.g., C:\\Windows\\System32\\SRU\\srudb.dat)
    :param destination: Path to save the copied file
    """
    volume = os.path.splitdrive(source)[0]  # Extract drive letter (e.g., C:)
    
    # Create the shadow copy
    shadow_path = create_shadow_copy(f"{volume}\\")
    if isinstance(shadow_path, int):
        print(f"[-] Failed to create shadow copy: {shadow_path}")
        return
     
    print(f"[+] Shadow Copy Device: {shadow_path}")

    # Map the source file to the shadow copy
    snapshot_file_path = shadow_path.replace("\\\\?\\", "\\\\.\\", 1)  + source.replace(volume, "", 1)

    # Use robocopy in backup mode to copy the file
    cmd_copy = f'robocopy "{os.path.dirname(snapshot_file_path)}" "{os.path.dirname(destination)}" "{os.path.basename(snapshot_file_path)}" /b /nfl /ndl /np'
    res = subprocess.run(cmd_copy, shell=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    print(f"[+] Successfully copied: {source} -> {destination}")

# Example usage:
copy_locked_file(r"C:\Windows\System32\SRU\srudb.dat", r"srudb_copy.dat")
copy_locked_file(r"C:\Windows\System32\config\SOFTWARE", r"SOFTWARE_copy")
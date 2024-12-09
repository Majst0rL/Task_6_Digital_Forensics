import ctypes
import os
import time
from ctypes import wintypes, windll
from PIL import Image
import piexif

def edit_png_metadata(file_path, creation_date):
    temp_file = file_path + ".temp"
    with Image.open(file_path) as img:
        exif_dict = {"Exif": {piexif.ExifIFD.DateTimeOriginal: creation_date.encode()}}
        exif_bytes = piexif.dump(exif_dict)

        img.save(temp_file, "png", exif=exif_bytes)

    os.replace(temp_file, file_path)

def set_file_times(file_path, unified_time):
    unified_timestamp = time.mktime(time.strptime(unified_time, "%Y-%m-%d %H:%M:%S"))
    file_time = int(unified_timestamp * 10**7 + 116444736000000000)
    file_time_struct = wintypes.FILETIME(file_time & 0xFFFFFFFF, file_time >> 32)

    file_handle = windll.kernel32.CreateFileW(
        file_path,
        0x40000000,
        0,
        None,
        3,
        0x02000000,
        None,
    )
    if file_handle == -1:
        raise OSError("Failed to open file handle.")

    windll.kernel32.SetFileTime(
        file_handle,
        ctypes.pointer(file_time_struct),  # Creation time
        ctypes.pointer(file_time_struct),  # Access time
        ctypes.pointer(file_time_struct),  # Modification time
    )
    windll.kernel32.CloseHandle(file_handle)

if __name__ == "__main__":
    file_path = r"D:\FAX\FAX\IPT 2024_25\DF\Vaje\Naloga 6\Podatki7\SlikaDF.png"
    unified_time = "2024-12-09 20:30:00"
    creation_date_exif = "2024:12:09 20:30:00"

    edit_png_metadata(file_path, creation_date_exif)

    set_file_times(file_path, unified_time)

    print(f"All timestamps for {file_path} are setted on {unified_time}.")

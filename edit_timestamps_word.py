import ctypes
import os
import time
from ctypes import wintypes, windll
import zipfile
import xml.etree.ElementTree as ET

def edit_core_metadata(file_path, creation_date):
    temp_file_path = file_path + '.temp'
    with zipfile.ZipFile(file_path, 'r') as zf_in, zipfile.ZipFile(temp_file_path, 'w') as zf_out:
        for item in zf_in.infolist():
            content = zf_in.read(item.filename)
            if item.filename == "docProps/core.xml":
                root = ET.fromstring(content)
                created_element = root.find('.//{http://purl.org/dc/terms/}created')
                modified_element = root.find('.//{http://purl.org/dc/terms/}modified')

                # Posodobi ali dodaj 'created' element
                if created_element is not None:
                    created_element.text = creation_date
                else:
                    dcterms_ns = "{http://purl.org/dc/terms/}"
                    xsi_ns = "{http://www.w3.org/2001/XMLSchema-instance}"
                    created_element = ET.Element(dcterms_ns + "created")
                    created_element.set(xsi_ns + "type", "dcterms:W3CDTF")
                    created_element.text = creation_date
                    root.append(created_element)

                # Posodobi ali dodaj 'modified' element
                if modified_element is not None:
                    modified_element.text = creation_date
                else:
                    dcterms_ns = "{http://purl.org/dc/terms/}"
                    xsi_ns = "{http://www.w3.org/2001/XMLSchema-instance}"
                    modified_element = ET.Element(dcterms_ns + "modified")
                    modified_element.set(xsi_ns + "type", "dcterms:W3CDTF")
                    modified_element.text = creation_date
                    root.append(modified_element)

                content = ET.tostring(root, encoding='utf-8', xml_declaration=True)
            zf_out.writestr(item, content, zf_in.getinfo(item.filename).compress_type)
    os.replace(temp_file_path, file_path)


def set_file_times(file_path, unified_time):
    unified_timestamp = time.mktime(time.strptime(unified_time, "%Y-%m-%d %H:%M:%S"))
    file_time = int(unified_timestamp * 10 ** 7 + 116444736000000000)
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
    file_path = r"D:\FAX\FAX\IPT 2024_25\DF\Vaje\Naloga 6\Podatki7\Clanek_1.docx"
    unified_time = "2024-12-09 20:30:00"
    creation_date_iso = "2024-12-09T20:30:00Z"

    edit_core_metadata(file_path, creation_date_iso)

    set_file_times(file_path, unified_time)

    print(f"All timestamps for {file_path} are setted on {unified_time}.")

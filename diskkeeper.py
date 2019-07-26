import os
import re
import win32api
import win32file
import csv
import time
from datetime import datetime

import zipfile
import zlib

def get_drive_letters():
    drives_found = win32api.GetLogicalDriveStrings().split('\000')[:-1]
    drives_returned = []
    drive_types = [win32file.DRIVE_FIXED, win32file.DRIVE_REMOVABLE]

    for drive in drives_found:
        print(win32file.GetDriveType(drive))
        if (win32file.GetDriveType(drive) in drive_types):
            drives_returned.append(drive);

    return drives_returned

def get_files_on_drive(drive):
    file_list = []
    for root, dirs, files in os.walk(drive):
        for f in files + dirs:
            file = {}
            file["name"] = os.path.join(root, f)

            if (os.path.isdir(file["name"])):
                file["folder"] = "Y"
                print("{}".format(file["name"]))
            else:
                file["folder"] = ""

            try:
                file["size"] = os.path.getsize(file["name"]) 
            except:
                file["size"] = 0

            #print("{}\t{}".format(file["name"], file["size"]))

            try:
                file["modified_raw"] = os.path.getmtime(file["name"])
            except:
                file["modified_raw"] = ""
            
            try:
                file["modified"] = time.ctime(file["modified_raw"])
            except:
                file["modified"] = ""

            try:
                file["created_raw"] = os.path.getctime(file["name"])
            except:
                file["created_raw"] = ""

            try:
                file["created"] = time.ctime(file["created_raw"])
            except:
                file["created"] = ""         

            file_list.append(file)

    # for file in file_list:
    #     if (os.path.isdir(file["name"])):
    #         size = 0
    #         for file2 in file_list:
    #             if (file["name"] in file2["name"]):
    #                 size += file2["size"]
    #         file["size"] = size

    return file_list

def write_filelist_to_csv(filelist, csvname):

    keys = filelist[0].keys()

    with open(csvname, 'w', newline='') as csvFile:
        writer = csv.DictWriter(csvFile, keys, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        writer.writeheader()
        writer.writerows(filelist)

log_path = "G:\\Dropbox\\Stuff\\Backups\\HDD Logs\\"

if __name__ == '__main__':
    print("Getting available drives...")
    drives = get_drive_letters()
    print(drives)

    for drive in drives:
        print("Scanning " + drive + "...")
        files = get_files_on_drive(drive)
        filename = "{}_{}_{}.csv".format(os.environ['COMPUTERNAME'], drive[0], datetime.now().strftime('%Y_%m_%d_%H_%M_%S')) 

        print("Writing to CSV...")
        write_filelist_to_csv(files, filename + ".csv")

        print("Zipping CSV...")
        with zipfile.ZipFile(log_path + filename + ".zip", mode='w') as newzip:
            newzip.write(filename + ".csv", compress_type=zipfile.ZIP_DEFLATED)

        print("Deleting CSV...")
        os.remove(filename + ".csv")

    print("Done!")



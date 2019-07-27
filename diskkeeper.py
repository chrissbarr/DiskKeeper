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

log_path = "G:\\Dropbox\\Stuff\\Backups\\HDD Logs\\"

if __name__ == '__main__':
    print("Getting available drives...")
    drives = get_drive_letters()
    print(drives)

    #drives = ["C:\\"]

    for drive in drives:

        filename = "{}_{}_{}".format(os.environ['COMPUTERNAME'], drive[0], datetime.now().strftime('%Y_%m_%d_%H_%M_%S')) 
        filename_csv = filename + ".csv"

        with open(filename_csv, 'w', newline='', encoding='utf-8') as csvFile:
            writer = csv.writer(csvFile, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
            writer.writerow(['Filename', 'Filesize', 'Modified Timestamp', 'Created Timestamp', 'Modified Readable', 'Created Readable'])

            count = 0
            extf = set(('$RECYCLE.BIN','$Recycle.Bin','System Volume Information'))

            for root, dirs, files in os.walk(drive):
                dirs[:] = [d for d in dirs if d not in extf]
                for f in files + dirs:
                    file = {}
                    file["name"] = os.path.join(root, f)
                    file["size"] = ""
                    file["modified"] = ""
                    file["created"] = ""
                    file["modified_r"] = ""
                    file["created_r"] = ""

                    try:
                        statinfo = os.stat(file["name"])
                    except:
                        pass
                    else:
                        file["size"] = statinfo.st_size
                        file["modified"] = statinfo.st_mtime
                        file["created"] = statinfo.st_ctime

                        try:
                            file["modified_r"] = time.strftime('%Y-%m-%d %H-%M-%S', time.localtime(file["modified"]))
                        except:
                            file["modified_r"] = ""

                        try:
                            file["created_r"] = time.strftime('%Y-%m-%d %H-%M-%S', time.localtime(file["created"]))
                        except:
                            file["created_r"] = ""

                    count += 1
                    if (count % 10000 == 0):
                        print("{}-{}".format(count, file))

                    writer.writerow([
                        file["name"], 
                        file["size"],
                        #file["modified"],
                        file["modified_r"],
                        #file["created"],
                        file["created_r"]
                    ])

        print("Zipping CSV...")
        with zipfile.ZipFile(log_path + filename + ".zip", mode='w') as newzip:
            newzip.write(filename_csv, compress_type=zipfile.ZIP_DEFLATED)

        print("Deleting CSV...")
        os.remove(filename_csv)

    print("Done!")



import os
import re
import win32api
import win32file
import csv
import time
from datetime import datetime
import argparse
import sys

import zipfile
import zlib

def get_drive_letters(drive_types):
    drives_found = win32api.GetLogicalDriveStrings().split('\000')[:-1]
    drives_returned = []

    for drive in drives_found:
        if (win32file.GetDriveType(drive) in drive_types):
            drives_returned.append(drive);

    return drives_returned

if __name__ == '__main__':

    parser = argparse.ArgumentParser(description="DiskKeeper")
    parser.add_argument('output_dir', help='Output directory for records')
    parser.add_argument("-z", action="store_true", help="Zip output records")
    parser.add_argument("--drive_fixed", action="store_true", help="Include all fixed drives")
    parser.add_argument("--drive_removable", action="store_true", help="Include all removable drives")
    parser.add_argument("--drive_remote", action="store_true", help="Include all remote drives")
    parser.add_argument("--drive", action="store", help="Include ONLY a specific drive")

    args = parser.parse_args()

    if args.z:
        zip_output = True
    else:
        zip_output = False

    log_path = args.output_dir

    drive_list = []

    if args.drive:
        print("Drive specified ({}), checking only this drive.")
        drive_list.append(args.drive)
    else:
        print("Getting available drives...")
        drive_types = []
        if args.drive_fixed:
            drive_types.append(win32file.DRIVE_FIXED)
        if args.drive_removable:
            drive_types.append(win32file.DRIVE_REMOVABLE)
        if args.drive_remote:
            drive_types.append(win32file.DRIVE_REMOTE)

        if len(drive_types) == 0:
            print("Warning - no drive types specified. Please specify one or more using the --drive_fixed, --drive_removable or --drive_remote arguments.")
            sys.exit()

        print("Only including the following drive types: {}".format(drive_types))
        drive_list = get_drive_letters(drive_types)

    print("The following drives will be checked:")
    print(drive_list)

    for drive in drive_list:

        print("Beginning check for drive: {}".format(drive))

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
                        file["modified"],
                        file["modified_r"],
                        file["created"],
                        file["created_r"]
                    ])

        if (count > 0):
            if zip_output:
                print("Zipping CSV...")
                with zipfile.ZipFile(log_path + filename + ".zip", mode='w') as newzip:
                    newzip.write(filename_csv, compress_type=zipfile.ZIP_DEFLATED)
                print("Deleting CSV...")
                os.remove(filename_csv)
        else:
            print("No record created for {}, discarding...".format(drive))
            os.remove(filename_csv)

        print("Completed check for drive: {}".format(drive))

    print("Done!")

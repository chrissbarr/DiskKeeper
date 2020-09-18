import argparse
import csv
import logging
import os
import re
import shutil
import sys
import time
import zipfile
import zlib
from datetime import datetime

import win32api
import win32file

format_linebreak_width = 80
excluded_directories = set(("$RECYCLE.BIN", "$Recycle.Bin", "System Volume Information"))

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.FileHandler(os.path.basename(__file__) + ".log"), logging.StreamHandler(),],
)


def get_drive_letters(drive_types):
    drives_found = win32api.GetLogicalDriveStrings().split("\000")[:-1]
    drives_returned = []

    for drive in drives_found:
        if win32file.GetDriveType(drive) in drive_types:
            drives_returned.append(drive)

    return drives_returned


def write_filelist_to_csv(filelist, filename):

    with open(filename, "w", newline="", encoding="utf-8") as csvFile:
        writer = csv.writer(csvFile, quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        writer.writerow(
            [
                "Filename",
                "Filesize",
                "Modified Timestamp",
                "Modified Readable",
                "Created Timestamp",
                "Created Readable",
            ]
        )

        for file in filelist:
            writer.writerow(
                [file["name"], file["size"], file["modified"], file["modified_r"], file["created"], file["created_r"],]
            )


if __name__ == "__main__":

    logging.info("-" * format_linebreak_width)
    logging.info("DiskKeeper START")
    logging.info("-" * format_linebreak_width)

    parser = argparse.ArgumentParser(description="DiskKeeper")
    parser.add_argument("output_dir", help="Output directory for records")
    parser.add_argument("-z", action="store_true", help="Zip output records")
    parser.add_argument("--drive_fixed", action="store_true", help="Include all fixed drives")
    parser.add_argument("--drive_removable", action="store_true", help="Include all removable drives")
    parser.add_argument("--drive_remote", action="store_true", help="Include all remote drives")
    parser.add_argument("--drive", action="store", help="Include ONLY a specific drive")

    args = parser.parse_args()

    logging.info("Arguments passed: ")

    output_dir = args.output_dir
    logging.info("  output_dir = {}".format(output_dir))

    if args.z:
        zip_output = True
        logging.info("  -z = TRUE (output CSV will be zipped)")
    else:
        zip_output = False
        logging.info("  -z = FALSE (output CSV will not be zipped)")

    drive_list = []

    if args.drive:
        logging.info("  --drive = {} (this drive is added to the list to check)".format(args.drive))
        drive_list.append(args.drive)
    else:
        logging.info("  --drive = NONE (no drive specifically added to list to check)")

    drive_types = []
    if args.drive_fixed:
        logging.info("  --drive_fixed = TRUE (drives of this type will be checked)")
        drive_types.append(win32file.DRIVE_FIXED)
    else:
        logging.info("  --drive_fixed = FALSE (drives of this type will not be checked)")

    if args.drive_removable:
        logging.info("  --drive_removable = TRUE (drives of this type will be checked)")
        drive_types.append(win32file.DRIVE_REMOVABLE)
    else:
        logging.info("  --drive_removable = FALSE (drives of this type will not be checked)")

    if args.drive_remote:
        logging.info("  --drive_remote = TRUE (drives of this type will be checked)")
        drive_types.append(win32file.DRIVE_REMOTE)
    else:
        logging.info("  --drive_remote = FALSE (drives of this type will not be checked)")

    logging.info("Assembling list of drives to check: ")

    if len(drive_types) != 0:
        logging.info("  Adding the following drive types to list: {}".format(drive_types))
        drive_list.append(get_drive_letters(drive_types))

    if len(drive_list) == 0:
        logging.error("  No drives in list of drives to check!")
        logging.error("  Please either explicitly specify a drive (--drive) to check, or nominate drive")
        logging.error("  types to be checked using the --drive_fixed, --drive_removable or --drive_remote flags.")
        sys.exit()

    logging.info("  The following drives will be checked:")
    logging.info("  {}".format(drive_list))

    for drive in drive_list:

        logging.info("Beginning check for drive: {}".format(drive))
        filelist = []
        count = 0

        for root, dirs, files in os.walk(drive):
            dirs[:] = [d for d in dirs if d not in excluded_directories]
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
                        file["modified_r"] = time.strftime("%Y-%m-%d %H-%M-%S", time.localtime(file["modified"]))
                    except:
                        file["modified_r"] = ""

                    try:
                        file["created_r"] = time.strftime("%Y-%m-%d %H-%M-%S", time.localtime(file["created"]))
                    except:
                        file["created_r"] = ""

                filelist.append(file)

                count += 1
                if count % 10000 == 0:
                    logging.info("  {0:08} - {1}".format(count, file["name"]))

        if len(filelist) > 0:
            logging.info("  Filelist contains entries - saving to CSV.")

            base_filename = "{}_{}_{}".format(
                os.environ["COMPUTERNAME"], drive[0], datetime.now().strftime("%Y_%m_%d_%H_%M_%S")
            )

            logging.info("  Filename will be: {}".format(base_filename))

            filename_csv = base_filename + ".csv"
            write_filelist_to_csv(filelist, filename_csv)
            output_filename = filename_csv

            if zip_output:

                filename_zip = base_filename + ".zip"

                logging.info("  Zipping CSV...")
                with zipfile.ZipFile(filename_zip, mode="w") as newzip:
                    newzip.write(filename_csv, compress_type=zipfile.ZIP_DEFLATED)
                logging.info("  Done!")

                logging.info("  Deleting CSV...")
                os.remove(filename_csv)
                logging.info("  Done!")

                output_filename = filename_zip

            logging.info("  Moving output file to output directory: ")
            shutil.move(output_filename, os.path.join(output_dir, output_filename))
            logging.info("  Done!")

        else:
            logging.info("  Filelist is empty - no records to save.")

        logging.info("  Completed check for drive: {}".format(drive))

    logging.info("Completed checks for all drives in list!")
    logging.info("-" * format_linebreak_width)

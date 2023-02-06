#!/usr/bin/env python
"""Script to crawl a path and write a CSV file with information about all files in that path

Typical usage example:
    python path_to_csv.py --dir "C:\\Users\\MyUser\\Documents\\TheseDocuments" --recursive
"""

import sys
import os
import typing
import logging
import argparse
import csv
from math import ceil
import win32com.client
import epub_meta

N_FILES = [0]
N_DIRS = [0]
N_EBOOK_FAILURS = [0]


def transform_to_mb(size: str) -> str:
    """Transforms a string representing a size to MB

    Takes a string of the type "60 Bytes", "1,90 KB" or "1,80 GB"
    to MB format.

    Args:
        size (str): Current format of the file size

    Returns:
        str: Filesize in MB
    """
    conversion_dict = {
        "KB": 1 / 1024,
        "Bytes": 1 / (1024) ** 2,
        "MB": 1,
        "GB": 1024,
        "TB": 1024**2,
    }
    value, unit = size.split(" ")
    if unit in conversion_dict:
        # Turn "1,90" into the number 1.9
        number_value = float(value.replace(",", "."))
        # get value for MB
        number_value = number_value * conversion_dict[unit]
        # Round UP to two decimals
        number_value = ceil(number_value * 100) / 100.0
        # Turn back to a string and replace "." with ","
        # to get original formatting back
        return_value = str(number_value).replace(".", ",")
        # Add new unit and return
        return " ".join([return_value, "MB"])
    return size


def go_recursive(crawl_path: str) -> typing.Iterator[str]:
    """Crawl a path recursively and return all directories

    Args:
        crawl_path (str): Path to crawl

    Yields:
        str: The path of the next (sub)directory in the path
    """
    crawl_path = os.path.abspath(crawl_path)
    # Do not forget to return the original directory
    yield crawl_path
    for root, dirs, _ in os.walk(crawl_path, topdown=True):
        for directory in dirs:
            yield_path = os.path.join(root, directory)
            logging.debug("Yielding path: %s!", yield_path)
            # Return every subdirectory
            yield yield_path


def get_information(dir_path: str, dispatch: win32com.client.dynamic.CDispatch) -> list[dict[str, str]]:
    """Get information about all files in a directory

    Args:
        dir_path (str): Path of the directory
        dispatch (win32com.client.CDispatch): "Microsoft Shell Controls And Automation.IShellDispatch6"
            Dispatch used to read file metadata.

    Returns:
        list[dict[str, str]]: List that for each file contains a dictionary mapping
            metadata keys to their values for that file

    Raises:
        FileNotFoundError: If the given path does not exist
        FileNotFoundError: If the given path is not a directory
    """
    logging.info("In directory %s", dir_path)
    if not os.path.exists(dir_path):
        raise FileNotFoundError("Could not find the given directory!")
    if not os.path.isdir(dir_path):
        raise FileNotFoundError("Path has to be for a directory!")
    N_DIRS[0] += 1
    folder_files = []
    folder = dispatch.NameSpace(dir_path)
    columns = []
    # Skip empty columns (up to 296) and useless information
    # GesamtGröße = 57 (DiskSize); Dateiname = 165 (gedoppelt);
    # Freier Speicherplatz = 169; OrdnerInfo = 190,191,192 (gedoppelt);
    # Typ = 196 (gedoppelt); Verwendeter Speicherplatz = 254
    skip_set = {
        37,
        38,
        39,
        40,
        41,
        59,
        170,
        171,
        205,
        206,
        207,
        218,
        296,  # Empty up to here
        57,  # Total size
        165,  # Filename
        169,  # Space free
        190,  # Folder name
        191,  # Folder path
        192,  # Folder
        196,  # Type
        254,  # Space used
    }
    # There are currently columns up to
    # 320 for Windows 10
    # https://stackoverflow.com/a/62279888/7895542
    for colnum in set(range(321)) - skip_set:
        colname = folder.GetDetailsOf(None, colnum)
        columns.append((colnum, colname))

    item_index = 0
    for file_name in os.listdir(dir_path):
        item = folder.ParseName(file_name)
        # Do not care about directories
        if os.path.isdir(item.Path):
            continue
        N_FILES[0] += 1
        if item_index % 100 == 0:
            logging.info("Checking file number %s in the current folder.", item_index)
        this_file = {}
        this_file["Pfad"] = item.Path
        for colnum, column in columns:
            colval = folder.GetDetailsOf(item, colnum)
            if colval:
                # Column 1 is "Size"
                if colnum == 1:
                    this_file[column] = transform_to_mb(colval)
                else:
                    this_file[column] = colval
        if "epub" in os.path.splitext(file_name)[1]:
            logging.debug("Found epub file. Parsing additional metadata!")
            try:
                pub_meta_data = epub_meta.get_epub_metadata(os.path.join(dir_path, file_name), read_cover_image=False)
                for pub_key in pub_meta_data:
                    if pub_meta_data[pub_key]:
                        if pub_key == "toc":
                            this_file["epub_chapters"] = [entry["title"] for entry in pub_meta_data[pub_key]]
                        else:
                            column_name = pub_key if "epub" in pub_key else "epub_" + pub_key
                            this_file[column_name] = pub_meta_data[pub_key]
            except epub_meta.EPubException as e:
                N_EBOOK_FAILURS[0] += 1
                logging.debug("Failed to parse ebook. Got error message %s", e)
        folder_files.append(this_file)
        item_index += 1
    return folder_files


def main(args):
    """Crawl a path and write a CSV file with file information"""
    parser = argparse.ArgumentParser("Crawl a path and write a CSV file with file information")
    parser.add_argument("-d", "--debug", action="store_true", default=False, help="Enable debug output.")
    parser.add_argument(
        "--dir",
        default=r"C:\Users\MyUser\Documents\TheseDocuments",
        help="Directory that should be crawled",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        default=False,
        help="Also recursively parse all subdirectories",
    )
    options = parser.parse_args(args)

    # options.dir = input("Enter the path to the directory you want to crawl: ")
    # response = None
    # while response not in ["Y", "N"]:
    #     try:
    #         response = input("Do you want to also check all subdirectories? [Y/N]: ").upper()
    #     except (EOFError, KeyboardInterrupt):
    #         print("Bye")
    #         sys.exit()
    #     except (KeyError, ValueError, AttributeError):
    #         print("Bad choice")
    # options.recursive = response == "Y"

    # Check if the requested directory even exists
    if not os.path.exists(options.dir):
        raise FileNotFoundError("Could not find the path to be crawled!")
    # And that it is a directory
    if not os.path.isdir(options.dir):
        raise FileNotFoundError("The given path does not point to a directory!")

    if options.debug:
        logfile = os.path.join(options.dir, "crawl.log")
        logging.basicConfig(
            filename=logfile,
            encoding="utf-8",
            level=logging.DEBUG,
            filemode="w",
            format="%(asctime)s %(levelname)-8s %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )
    else:
        logging.basicConfig(
            encoding="utf-8",
            level=logging.INFO,
            filemode="w",
            format="%(asctime)s %(levelname)-8s %(message)s",
            datefmt="%Y-%m-%d %H:%M:%S",
        )

    logging.info(
        "Running with search directory: %s. Searching %s.",
        options.dir,
        "recursively" if options.recursive else "non recursively",
    )
    dispatch = win32com.client.gencache.EnsureDispatch("Shell.Application", 0)
    all_files = []

    if options.recursive:
        for dir_path in go_recursive(crawl_path=options.dir):
            all_files.extend(get_information(dir_path, dispatch))
    else:
        all_files.extend(get_information(options.dir, dispatch))

    # Header has to contain any field that shows
    # up for any file
    field_names = ["Pfad"]
    # Use set for faster lookup
    # Memory should not be constraining at all
    field_set = {"Pfad"}
    for file_entry in all_files:
        for key in file_entry:
            if key not in field_set:
                field_set.add(key)
                field_names.append(key)

    csv_path = os.path.join(options.dir, "contents.csv")
    logging.info("Writing results to %s", csv_path)
    with open(csv_path, "w", encoding="utf-8", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=field_names)
        writer.writeheader()
        for data in all_files:
            writer.writerow(data)

    logging.info("Analyzed a total of %s files in %s (sub)directories.", N_FILES[0], N_DIRS[0])
    if N_EBOOK_FAILURS[0] > 0:
        logging.info("Errors occured when parsing %s .epub files.", N_EBOOK_FAILURS[0])
    # input("Finished! Press any key to exit.")


if __name__ == "__main__":
    main(sys.argv[1:])

#!/usr/bin/env python
r"""Script to crawl a path and write a CSV file with information about all files.

Typical usage example:
python path_to_csv.py --dir "C:\\Users\\MyUser\\Documents\\TheseDocuments" --recursive
"""
# pylint: disable=import-error

import csv
import logging
import os
import sys
from collections.abc import Iterator
from math import ceil
from typing import Any

import epub_meta  # pyright: ignore [reportMissingTypeStubs]
import win32com.client
from gooey import Gooey, GooeyParser  # pyright: ignore [reportMissingTypeStubs]


def transform_to_mb(size: str) -> str:
    """Transforms a string representing a size to MB.

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
        number_value *= conversion_dict[unit]
        # Round UP to two decimals
        number_value = ceil(number_value * 100) / 100.0
        # Turn back to a string and replace "." with ","
        # to get original formatting back
        return_value = str(number_value).replace(".", ",")
        # Add new unit and return
        return f"{return_value} MB"
    return size


def go_recursive(crawl_path: str) -> Iterator[str]:
    """Crawl a path recursively and return all directories.

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


class InformationExtractor:
    """Class to extract information out of a given path.

    Attributes:
        n_files (int): Running count of how many files were encountered.
        n_dirs (int): Running count of how many directories were encountered.
        n_ebook_failures (int): Running count of ebooks encountered.
    """

    def __init__(self) -> None:
        """Instantiate an InformationExtractor."""
        self.n_files: int = 0
        self.n_dirs: int = 0
        self.failed_ebooks: list[str] = []
        self.dispatch: win32com.client.dynamic.CDispatch = (
            win32com.client.gencache.EnsureDispatch("Shell.Application", 0)
        )

    def get_columns_to_parse(self, folder: Any) -> list[tuple[int, str]]:
        """Collect all the columns to parse as well as their ids.

        Args:
            folder (Any): Folder to extract columns from.
                win32com.client.dynamic.CDispatch.NameSpace("path")

        Returns:
            list[tuple[int, str]]: Columns represented by their number and name.
        """
        columns: list[tuple[int, str]] = []
        # Skip empty columns (up to 296) and useless information
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
        return columns

    def extract_general_information(
        self,
        columns: list[tuple[int, str]],
        folder: Any,
        this_file: dict[str, str],
        item: Any,
    ) -> None:
        """Extract general information about the file.

        Extract information via folder.GetDetailsOf from win32com.client.
        The information is stored in `this_file` inplace.

        Args:
            columns (list[tuple[int, str]]): Columns of interest.
            folder (Any): Folder in which the item of interest lies.
                win32com.client.dynamic.CDispatch.NameSpace("path")
            this_file (dict[str, str]): Dictionary storing information about each file.
            item (Any): File to get information about.
        """
        for colnum, column in columns:
            if colval := folder.GetDetailsOf(item, colnum):
                # Column 1 is "Size"
                this_file[column] = transform_to_mb(colval) if colnum == 1 else colval

    def extract_epub_information(
        self, file_path: str, this_file: dict[str, Any]
    ) -> None:
        """Extract additional information from epub file.

        The information is stored in `this_file` inplace.

        Args:
            file_path (str): Path of the current directory
            this_file (dict[str, Any]): Dictionary storing information about each file.
        """
        logging.debug("Found epub file %s. Parsing additional metadata!", file_path)
        try:
            pub_meta_data = epub_meta.get_epub_metadata(
                file_path, read_cover_image=False
            )
            for pub_key in pub_meta_data:
                if pub_meta_data[pub_key]:
                    if pub_key == "toc":
                        this_file["epub_chapters"] = [
                            entry["title"] for entry in pub_meta_data[pub_key]
                        ]
                    else:
                        column_name = (
                            pub_key if "epub" in pub_key else f"epub_{pub_key}"
                        )
                        this_file[column_name] = pub_meta_data[pub_key]
        except Exception as e:  # pylint: disable=broad-except  # noqa: BLE001
            self.failed_ebooks.append(file_path)
            logging.info("Failed to parse ebook. Got error message %s", e)

    def get_information(self, dir_path: str) -> list[dict[str, str]]:
        """Get information about all files in a directory.

        Args:
            dir_path (str): Path of the directory
            dispatch (win32com.client.CDispatch): "Microsoft Shell Controls And
                Automation.IShellDispatch6" Dispatch used to read file metadata.

        Returns:
            list[dict[str, str]]: List that for each file contains a dictionary mapping
                metadata keys to their values for that file

        Raises:
            FileNotFoundError: If the given path does not exist
            FileNotFoundError: If the given path is not a directory
        """
        logging.info("In directory %s", dir_path)
        if not os.path.exists(dir_path):
            msg = "Could not find the given directory!"
            raise FileNotFoundError(msg)
        if not os.path.isdir(dir_path):
            msg = "Path has to be for a directory!"
            raise FileNotFoundError(msg)
        self.n_dirs += 1
        folder_files = []
        folder = self.dispatch.NameSpace(dir_path)

        columns = self.get_columns_to_parse(folder)

        for file_name in os.listdir(dir_path):
            item = folder.ParseName(file_name)
            # Do not care about directories
            if os.path.isdir(item.Path):
                continue
            self.n_files += 1
            if self.n_files % 1000 == 1:
                logging.info("Checking file number %s.", self.n_files)
            this_file: dict[str, str] = {"Pfad": item.Path}
            self.extract_general_information(columns, folder, this_file, item)

            if "epub" in os.path.splitext(file_name)[1]:
                self.extract_epub_information(
                    os.path.join(dir_path, file_name), this_file
                )
            folder_files.append(this_file)
        return folder_files


def get_field_names(all_files: list[dict[str, str]]) -> list[str]:
    """Extract all field names from parsed file dict.

    Args:
        all_files (list[dict[str, str]]): List of dictionaries containing information
            about each field name in each file.

    Returns:
        list[str]: List of all field names that appear in at least one file.
    """
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
    return field_names


def write_csv(
    csv_path: str, all_files: list[dict[str, str]], field_names: list[str]
) -> None:
    """Write the information in all_files to a csv file.

    Args:
        csv_path (str): Path to write the file to.
        all_files (list[dict[str, str]]): Information to write to file.
        field_names (list[str]): List of headers to use.
    """
    logging.info("Writing results to %s", csv_path)
    with open(csv_path, "w", encoding="utf-8", newline="") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=field_names)
        writer.writeheader()
        for data in all_files:
            writer.writerow(data)


@Gooey
def main(args: list[str]) -> None:
    """Crawl a path and write a CSV file with file information."""
    parser = GooeyParser(
        description="Crawl a path and write a CSV file with file information"
    )
    parser.add_argument(
        "-d", "--debug", action="store_true", default=False, help="Enable debug output."
    )
    parser.add_argument(
        "--dir",
        default=r"C:\Users\MyUser\Documents\TheseDocuments",
        help="Directory that should be crawled",
        widget="DirChooser",
    )
    parser.add_argument(
        "-r",
        "--recursive",
        action="store_true",
        default=False,
        help="Also recursively parse all subdirectories",
    )
    options = parser.parse_args(args)

    # Check if the requested directory even exists
    if not os.path.exists(options.dir):
        msg = "Could not find the path to be crawled!"
        raise FileNotFoundError(msg)
    # And that it is a directory
    if not os.path.isdir(options.dir):
        msg = "The given path does not point to a directory!"
        raise FileNotFoundError(msg)

    if options.debug:
        logging.basicConfig(
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
    all_files: list[dict[str, str]] = []

    information_extractor = InformationExtractor()

    if options.recursive:
        for dir_path in go_recursive(crawl_path=options.dir):
            all_files.extend(information_extractor.get_information(dir_path))
    else:
        all_files.extend(information_extractor.get_information(options.dir))

    field_names = get_field_names(all_files)

    write_csv(os.path.join(options.dir, "contents.csv"), all_files, field_names)

    logging.info(
        "Analyzed a total of %s files in %s (sub)directories.",
        information_extractor.n_files,
        information_extractor.n_dirs,
    )
    if information_extractor.failed_ebooks:
        logging.info(
            "Errors occurred when parsing %s .epub files.",
            len(information_extractor.failed_ebooks),
        )
        logging.debug(information_extractor.failed_ebooks)


if __name__ == "__main__":
    main(sys.argv[1:])

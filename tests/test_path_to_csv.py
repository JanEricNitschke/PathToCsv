"""Tests download_pdfs.py"""

# pylint: disable=attribute-defined-outside-init

import os
import shutil
import csv
import pytest
import win32com.client
from path_to_csv import go_recursive, get_information, transform_to_mb, main


class TestHanserDownload:
    """Class to test download_pdfs.py"""

    def setup_class(self):
        """Create file structure to test and Shell application"""
        self.dispatch = win32com.client.gencache.EnsureDispatch("Shell.Application", 0)
        self.test_folder_level1_1 = os.path.abspath("test_folder_level1_1")
        self.csv_path = os.path.join(self.test_folder_level1_1, "contents.csv")
        self.test_folder_level2_1_1 = os.path.join(
            self.test_folder_level1_1, "test_folder_level2_1_1"
        )
        self.test_folder_level2_1_2 = os.path.join(
            self.test_folder_level1_1, "test_folder_level2_1_2"
        )
        self.test_folder_level3_1_2_1 = os.path.join(
            self.test_folder_level2_1_2, "test_folder_level3_1_2_1"
        )
        os.makedirs(self.test_folder_level1_1)
        os.makedirs(self.test_folder_level2_1_1)
        os.makedirs(self.test_folder_level2_1_2)
        os.makedirs(self.test_folder_level3_1_2_1)
        self.file1_path = os.path.join(self.test_folder_level1_1, "file_level1.txt")
        with open(
            self.file1_path,
            "w",
            encoding="utf-8",
        ) as file:
            file.write("file_level1")
        self.file2_path = os.path.join(self.test_folder_level2_1_2, "file_level2.log")
        with open(
            self.file2_path,
            "w",
            encoding="utf-8",
        ) as file:
            file.write("file_level2")
        self.file3_path = os.path.join(self.test_folder_level3_1_2_1, "file_level3.mp4")
        with open(
            self.file3_path,
            "w",
            encoding="utf-8",
        ) as file:
            file.write("file_level3")

    def teardown_class(self):
        """Delete testing files"""
        shutil.rmtree(self.test_folder_level1_1)

    def test_go_recursive(self):
        "Tests go_recursive"
        assert set(go_recursive(self.test_folder_level1_1)) == {
            os.path.abspath(test_path)
            for test_path in (
                self.test_folder_level1_1,
                self.test_folder_level2_1_1,
                self.test_folder_level2_1_2,
                self.test_folder_level3_1_2_1,
            )
        }

    def test_get_information(self):
        """Tests get_information"""
        test_information = get_information(self.test_folder_level1_1, self.dispatch)
        assert isinstance(test_information, list)
        assert len(test_information) == 1
        assert isinstance(test_information[0], dict)
        assert "Pfad" in test_information[0]
        assert os.path.basename(test_information[0]["Pfad"]) == "file_level1.txt"

        test_information = get_information(self.test_folder_level2_1_1, self.dispatch)
        assert isinstance(test_information, list)
        assert len(test_information) == 0

        with pytest.raises(FileNotFoundError):
            get_information("non_existent_path", self.dispatch)
        with pytest.raises(FileNotFoundError):
            get_information(self.file1_path, self.dispatch)

    def test_transform_to_mb(self):
        """Tests transform_to_mb"""
        assert transform_to_mb("1,90 KB") == "0,01 MB"
        assert transform_to_mb("2,5 TB") == "2621440,0 MB"
        assert transform_to_mb("3292429 Bytes") == "3,14 MB"
        assert transform_to_mb("156 PB") == "156 PB"

    def test_main(self):
        """Tests the full script"""
        main(["--dir", self.test_folder_level1_1])
        assert os.path.exists(self.csv_path)
        with open(self.csv_path, newline="", encoding="utf-8") as csvfile:
            reader = list(csv.DictReader(csvfile))
            assert len(reader) == 1
            assert reader[0]["Pfad"] == self.file1_path
        os.remove(self.csv_path)
        main(["--dir", self.test_folder_level1_1, "-r"])
        with open(self.csv_path, newline="", encoding="utf-8") as csvfile:
            reader = list(csv.DictReader(csvfile))
            assert len(reader) == 3
            assert reader[0]["Pfad"] == self.file1_path
            assert reader[1]["Pfad"] == self.file2_path
            assert reader[2]["Pfad"] == self.file3_path

        with pytest.raises(FileNotFoundError):
            main(["--dir", self.file1_path])
        with pytest.raises(FileNotFoundError):
            main(["--dir", "non_existent_path"])

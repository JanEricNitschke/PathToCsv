"""Tests path_to_csv.py."""

# pylint: disable=attribute-defined-outside-init

import os
import shutil

import pytest

from path_to_csv import (
    InformationExtractor,
    get_field_names,
    go_recursive,
    transform_to_mb,
)


class TestHanserDownload:
    """Class to test path_to_csv.py."""

    def setup_class(self):
        """Create file structure to test and Shell application."""
        self.extractor = InformationExtractor()
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
        """Delete testing files."""
        shutil.rmtree(self.test_folder_level1_1)

    def test_go_recursive(self):
        """Tests go_recursive."""
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
        """Tests get_information."""
        test_information = self.extractor.get_information(self.test_folder_level1_1)
        assert isinstance(test_information, list)
        assert len(test_information) > 0
        assert isinstance(test_information[0], dict)
        assert "Pfad" in test_information[0]
        assert os.path.basename(test_information[0]["Pfad"]) == "file_level1.txt"

        test_information = self.extractor.get_information(self.test_folder_level2_1_1)
        assert isinstance(test_information, list)
        assert len(test_information) == 0

        with pytest.raises(FileNotFoundError):
            self.extractor.get_information("non_existent_path")
        with pytest.raises(FileNotFoundError):
            self.extractor.get_information(self.file1_path)

    def test_ebook(self):
        """Try if parsing ebooks works."""
        ebook_information = self.extractor.get_information(
            os.path.abspath(os.path.join("tests", "ref"))
        )
        assert len(ebook_information) == 8
        assert "epub_description" in ebook_information[0]
        assert (
            ebook_information[0]["epub_description"]
            == "Font rendering for multiple languages in a single ePub 3"
        )

    def test_transform_to_mb(self):
        """Tests transform_to_mb."""
        assert transform_to_mb("1,90 KB") == "0,01 MB"
        assert transform_to_mb("2,5 TB") == "2621440,0 MB"
        assert transform_to_mb("3292429 Bytes") == "3,14 MB"
        assert transform_to_mb("156 PB") == "156 PB"

    def test_get_field_names(self):
        """Tests get_field_names."""
        all_files = [{"a": 1, "b": 2, "c": 3}, {"a": 1, "affe": 2}]
        field_names = get_field_names(all_files)
        assert len(field_names) == len(set(field_names))
        assert set(field_names) == {"Pfad", "a", "b", "c", "affe"}

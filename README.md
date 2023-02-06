# PathToCsv
Script to crawl a path and write a CSV file with information about all files in that path.

Install dependencies with:
```bash
pip install -r requirements.txt
```

Run python version with:
```bash
python path_to_csv.py --dir "C:\\Users\\MyUser\\Documents\\TheseDocuments" --recursive
```

Run exe version by double clicking it or with:
```bash
path_to_csv.exe
```
It will then prompt you for the directory and whether subdirectories should be checked recursively.

Exe file produced with [PyInstaller](https://pyinstaller.org/en/stable/) via
```bash
pyinstaller --onefile path_to_csv.py
```

Test with:
```bash
coverage run -m pytest
coverage report -m
coverage html
```
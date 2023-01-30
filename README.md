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

Run exe version with:
```bash
path_to_csv.exe --dir "C:\\Users\\MyUser\\Documents\\TheseDocuments" --recursive
```

Test with:
```bash
coverage run -m pytest
coverage report -m
coverage html
```
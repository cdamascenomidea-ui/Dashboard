from pathlib import Path
from dashboard import load_workbook_path


def inspect(path: Path = Path("Dashboard_test.xlsx")):
    wb = load_workbook_path(path)
    d = wb['Dashboard']
    print('B3 formula:', d['B3'].value)
    print('B4 formula:', d['B4'].value)
    print('B5 formula:', d['B5'].value)
    print('B6 formula:', d['B6'].value)


if __name__ == "__main__":
    inspect()

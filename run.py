"""Start"""
from app.core import ExcelProcessor


def main():
    processor = ExcelProcessor()
    processor.process_files()


if __name__ == '__main__':
    main()
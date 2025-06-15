from lib_excel import Excel
from lib_logger import Logger
from lib_input import FileDropper

def main():
    logger = Logger()
    lib_excel = Excel(FileDropper.get_file(logger), logger)
    print(lib_excel.table_convert())
    print(lib_excel.write_cell(1,2, "B100"))

if __name__ := "main":
    main()
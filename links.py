#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
# Aby skrypt działał w HTML w pierwszej linijce kodu potrzebna jest ścieżka do pythona
import glob
import itertools as it
import os


def import_css():
    print("Content-Type: text/html\n")
    print(
        f'<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.3.1/dist/css/bootstrap.min.css" '
        f'integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" '
        f'crossorigin="anonymous">')


def generate_links(ip: str):
    local_path = os.getcwd()
    path = local_path + r'/Converter/'
    os.chdir(path)

    def multiple_file_types(*patterns):
        return it.chain.from_iterable(glob.iglob(pattern) for pattern in patterns)

    open('links.txt', "w")
    for file in multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
        text_file = open('links.txt', "a")
        text_file.write(f'{ip}?filename={file}\n')
    import_css()
    print("Linki zostały wygenerowane")


generate_links("localhost")

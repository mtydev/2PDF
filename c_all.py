#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
import glob
import itertools as it
import os

from win32com import client


def import_css():
    print("Content-Type: text/html\n")
    print(
        f'<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.3.1/dist/css/bootstrap.min.css" '
        f'integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" '
        f'crossorigin="anonymous">')


def button_printer(self, script_path):
    print(f'<h3 class="title pt-xl-5"><a href={script_path}>{self}</a></h3>')


# Funkcja odpowiadająca za konwersję wszystkich plików excela znajdujących się w folderze Converter.
def convert_all():
    local_path = os.getcwd()
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False
    path = local_path + r'/Converter/'
    os.chdir(path)

    def multiple_file_types(*patterns):
        return it.chain.from_iterable(glob.iglob(pattern) for pattern in patterns)

    if not multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
        import_css()
        print(
            "W folderze nie ma plików excela przeznaczonych do konwersji. Proszę przenieść do folderu pliki "
            "przeznaczone do konwersji.")
    else:
        for file in multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
            input_file = os.path.join(path + file)
            if file in glob.glob("*.xls"):
                length = len(input_file)
                length -= 4
                output_file = input_file[:length]
                workbook = app.Workbooks.Open(input_file)
            else:
                length = len(input_file)
                length -= 5
                output_file = input_file[:length]
                workbook = app.Workbooks.Open(input_file)
            try:
                workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
            except Exception as e:
                print("Konwersja do PDF nie udała sie.")
                print(str(e))
            finally:
                workbook.Close()
                app.Quit()
        import_css()
        print(
            '<div class="jumbotron pt-5 text-center"><h1 class="title">Konwersja plików zostala zakończona</h1></div>')
        print('<div class="container text-center">')
        print('Kliknij w napis poniżej, aby wyświetlić listę plików')
        button_printer("Lista plików", "pdf_list.py")
        print("</div>")


convert_all()

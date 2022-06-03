#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
# Wymaga biblioteki pywin32 i zainstalowanego Excela
# Aby skrypt działał w HTML w pierwszej linijce kodu potrzebna jest ścieżka do pythona
import glob
import os


def show_file_list():
    local_path = os.getcwd()
    main_path = local_path + r'/Converter/'
    os.chdir(main_path)
    if not glob.glob("*.pdf"):
        print("W folderze nie ma żadnego pliku PDF")
    else:
        print('<div class="jumbotron pt-5 text-center">')
        print('<h1 class="title">Lista plików pdf: </h1>')
        print('</div>')
        print('<div class="container">')
        print('<p class="title text-center"><i>Kliknij na nazwę, aby wyświetlić lub pobrać plik</i></p>')
        print('<ul class="list-group">')
        for pdfFiles in glob.glob("*.pdf"):
            download_file = '/Converter/' + pdfFiles
            print(f'<li class="list-group-item"><a href={download_file}>{pdfFiles}</a></li>', end='\n')
        print('</ul>')
        print("</div>")


def import_css():
    print("Content-Type: text/html\n")
    print(
        f'<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.3.1/dist/css/bootstrap.min.css" '
        f'integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" '
        f'crossorigin="anonymous">')


import_css()
show_file_list()

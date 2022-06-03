#!C:\Users\kenno\AppData\Local\Programs\Python\Python38\python.exe
# Wymaga biblioteki pywin32 i zainstalowanego Excela
# Aby skrypt działał w HTML w pierwszej linijce kodu potrzebna jest ścieżka do pythona
import cgi
import glob
import itertools as it
import os

from win32com import client

# Przekierowanie zmiennych z JS do Pythona, przesłanych przez POST
form = cgi.FieldStorage()
fileName = form.getvalue('filename')
fileName = str(fileName)


# Funkcja odpowiedzialna za import CSS
def import_css():
    print("Content-Type: text/html\n")
    print(
        f'<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@4.3.1/dist/css/bootstrap.min.css" '
        f'integrity="sha384-ggOyR0iXCbMQv3Xipma34MD+dH/1fQ784/j6cY/iJTQUOhcWr7x9JvoRxT2MZw1T" '
        f'crossorigin="anonymous">')


# Funkcja odpowiedzialna za otwieranie instancji excela z danym plikiem
def conversion(workbook, view_file, output_file, app):
    try:
        workbook.ActiveSheet.ExportAsFixedFormat(0, output_file)
    except Exception as e:
        print(
            '<div class="d-flex flex-wrap align-content-center justify-content-center text-danger container '
            'text-center align-items-center bg-dark"><h1>Konwersja do PDF nie udała się.</h1></div>')
        f = open("error_log.log", "a")
        f.write(str(e))
        f.close()
    finally:
        workbook.Close()
        app.Quit()
        # Wyświetlanie pliku na stronie
        print("Content-Type: text/html\n")
        print("<head>")
        print('<script type="text/javascript" src="https://code.jquery.com/jquery-1.4.3.min.js"></script>')
        print('''
                <script type="text/javascript">
                $(window).load(function(){
                $('#pdf').attr('src','''f'{view_file}'''');

                });
                </script> ''')
        print('</head>')
        print('<body>')
        print(f'<iframe src="/Converter/{view_file}.pdf" height="100%" width="100%" id="pdf"></iframe>')
        print('</body>')


# Funkcja odpowiadająca za konwersję danego pliku z parametru w linku.
def convert_specific(self):
    local_path = os.getcwd()
    app = client.DispatchEx("Excel.Application")
    app.Interactive = False
    app.Visible = False
    path = local_path + "\\Converter\\"
    os.chdir(path)

    def multiple_file_types(*patterns):
        return it.chain.from_iterable(glob.iglob(pattern) for pattern in patterns)

    file = str(self)
    # Walidacja plików
    if not multiple_file_types("*.xlsx", "*.xlsm", "*.xls"):
        import_css()
        print(
            '<div class="d-flex flex-wrap align-content-center justify-content-center text-danger container '
            'text-center align-items-center"><h1>W folderze nie ma plików excela przeznaczonych do konwersji. Proszę '
            'przenieść do folderu pliki przeznaczone do konwersji.</h1></div>')
    else:
        # Usunięcie zbędnych rozszerzeń plików
        input_file = os.path.join(path + file)
        if file in glob.glob("*.xls"):
            length = len(input_file)
            length -= 4
            output_file = input_file[:length]
            workbook = app.Workbooks.Open(input_file)
            # Usunięcie zbędnych rozszerzeń plików dla wyświetlania na stronie
            if self.endswith('.xlsx') or self.endswith('.xlsm'):
                view_file = len(self)
                view_file -= 5
                view_file = self[:view_file]
                conversion(workbook, view_file, output_file, app)
            elif self.endswith('.xls'):
                view_file = len(self)
                view_file -= 4
                view_file = self[:view_file]
                view_file = str(view_file)
                conversion(workbook, view_file, output_file, app)
        elif file in glob.glob("*.xlsx") or file in glob.glob("*.xlsm"):
            length = len(input_file)
            length -= 5
            output_file = input_file[:length]
            workbook = app.Workbooks.Open(input_file)
            # Usunięcie zbędnych rozszerzeń plików dla wyświetlania na stronie
            if self.endswith('.xlsx') or self.endswith('.xlsm'):
                view_file = len(self)
                view_file -= 5
                view_file = self[:view_file]
                conversion(workbook, view_file, output_file, app)
            elif self.endswith('.xls'):
                view_file = len(self)
                view_file -= 4
                view_file = self[:view_file]
                view_file = str(view_file)
                conversion(workbook, view_file, output_file, app)
        else:
            import_css()
            print(
                '<div class="d-flex flex-wrap align-content-center justify-content-center text-danger container '
                'text-center align-items-center bg-dark"><h1>Podany link jest niepoprawny. Sprawdź poprawność '
                'linku.</h1></div>')


convert_specific(fileName)

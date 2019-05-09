
cd C:\Eigene Dateien\Documents\GitHub\Notenauswertung
pyinstaller --onefile --windowed --exclude matplotlib --exclude numpy --exclude pandas --exclude jupyter --exclude scipy --exclude virtualenv --exclude pip Notenauswertung.py
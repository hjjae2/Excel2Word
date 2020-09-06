from PyQt5 import QtWidgets

import sys
import handle_window
import datetime

if __name__ == '__main__':
    sys.stdout = open('./log.txt', 'w', encoding='utf-8')
    print(datetime.datetime.today(), ">>")

    # Create app
    app = QtWidgets.QApplication([])

    # Create window
    window = handle_window.App()
    window.setWindowTitle("v1")

    # Built-in styles (Fusion, Windows, WindowsVista, GTK+)
    app.setStyle('Fusion')
    app.exec_()
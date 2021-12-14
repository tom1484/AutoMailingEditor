from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMainWindow, QFileDialog, QMessageBox
from PyQt5 import QtWebEngineWidgets

import qt.editor as editor
import mail.mail as mail

from os import mkdir
from os.path import split, join, abspath, exists
import configparser as cp


CURRENT_DIR = split(abspath(__file__))[0]


class Main(QMainWindow, editor.Ui_MainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi(self)

        self.html_viewer = QtWebEngineWidgets.QWebEngineView(self.label.parentWidget())
        self.html_viewer.setGeometry(self.label.geometry())
        self.html_viewer.setStyleSheet(self.label.styleSheet())
        self.html_viewer.setObjectName("html_viewer")

        self.button_load.clicked.connect(lambda: self.load())
        self.button_save.clicked.connect(lambda: self.save(popup=True))
        self.button_send.clicked.connect(lambda: self.send())

    def load(self):
        if not hasattr(self, 'folder'):
            self.folder = str(QFileDialog.getExistingDirectory(self, "Select Directory"))
            self.edit_folder.setText(split(self.folder)[1])

            config = cp.ConfigParser()
            config.read(self.folder + "\\config.ini", encoding="utf-8")

            # read sender account information
            account_info = config["ACCOUNT"]
            acc_user = account_info["user"]
            acc_password = account_info["pw"]

            self.edit_user.setText(acc_user)
            self.edit_password.setText(acc_password)

            # read message information
            message_info = config["MESSAGE"]
            mes_from = message_info["from"]
            mes_subject = message_info["subject"]

            self.edit_from.setText(mes_from)
            self.edit_subject.setText(mes_subject)

            # read recipient list
            self.edit_recipient.setText(open(
                self.folder + "\\recipients.csv", 'r', encoding="utf-8"
            ).read())

            # read letter detail
            self.html_viewer.load(QtCore.QUrl().fromLocalFile(
                self.folder + "\\letter.html")
            )

        else:
            self.html_viewer.load(QtCore.QUrl().fromLocalFile(
                self.folder + "\\letter.html")
            )

    def save(self, popup):
        if self.edit_folder.text() == "":
            return

        if not hasattr(self, 'folder'):
            folder_name = self.edit_folder.text()
            self.folder = join(CURRENT_DIR, folder_name)

            if exists(self.folder):
                msg = QMessageBox()
                msg.setWindowTitle("Saving Result")
                msg.setText("Folder Exists!")
                msg.exec_()
            else:
                mkdir(self.folder)
                mkdir(self.folder + "\\attach")
                open(join(self.folder, "letter.html"), 'w').close()

                with open(join(self.folder, "recipients.csv"), 'w', encoding="utf-8") as recipients_file:
                    recipients_file.write(self.edit_recipient.toPlainText())

                with open(join(self.folder, "config.ini"), 'w') as config_file:
                    config = cp.ConfigParser()

                    config.add_section("ACCOUNT")
                    config.set("ACCOUNT", "user", self.edit_user.text())
                    config.set("ACCOUNT", "pw", self.edit_password.text())

                    config.add_section("MESSAGE")
                    config.set("MESSAGE", "from", self.edit_from.text())
                    config.set("MESSAGE", "subject", self.edit_subject.text())

                    config.write(config_file)

        else:
            with open(join(self.folder, "recipients.csv"), 'w', encoding="utf-8") as recipients_file:
                recipients_file.write(self.edit_recipient.toPlainText())

            config = cp.ConfigParser()
            config.read(join(self.folder, "config.ini"), encoding="utf-8")

            config.set("ACCOUNT", "user", self.edit_user.text())
            config.set("ACCOUNT", "pw", self.edit_password.text())

            config.set("MESSAGE", "from", self.edit_from.text())
            config.set("MESSAGE", "subject", self.edit_subject.text())

            with open(join(self.folder, "config.ini"), 'w', encoding="utf-8") as config_file:
                config.write(config_file)

        if popup:
            msg = QMessageBox()
            msg.setWindowTitle("Saving Result")
            msg.setText("Successful!")
            msg.exec_()

    def send(self):
        if hasattr(self, 'folder'):
            self.save(popup=False)

            msg = QMessageBox()
            msg.setWindowTitle("Sending Result")
            msg.setText(mail.send(self.folder))
            msg.exec_()


if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = Main()
    window.show()
    sys.exit(app.exec_())

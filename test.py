from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QVBoxLayout, QWidget, QLineEdit, QLabel, QDialog, \
    QDialogButtonBox


class InputDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Input Dialog")
        self.layout = QVBoxLayout()

        self.input_label = QLabel("Enter value:")
        self.layout.addWidget(self.input_label)

        self.input_field = QLineEdit(self)
        self.layout.addWidget(self.input_field)

        self.button_box = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        self.button_box.accepted.connect(self.accept)
        self.button_box.rejected.connect(self.reject)
        self.layout.addWidget(self.button_box)

        self.setLayout(self.layout)

    def get_value(self):
        return self.input_field.text()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Main Window")

        self.button = QPushButton("Редактировать компоненты")
        self.button.clicked.connect(self.open_input_dialog)

        self.layout = QVBoxLayout()
        self.layout.addWidget(self.button)

        self.container = QWidget()
        self.container.setLayout(self.layout)

        self.setCentralWidget(self.container)

    def open_input_dialog(self):
        dialog = InputDialog()
        if dialog.exec():
            value = dialog.get_value()
            self.add_input_value(value)

    def add_input_value(self, value):
        label = QLabel(f"Value: {value}")
        self.layout.addWidget(label)


if __name__ == "__main__":
    app = QApplication([])
    window = MainWindow()
    window.show()
    app.exec()

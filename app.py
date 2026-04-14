import sys
import os
from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLineEdit,
    QTextEdit,
    QProgressBar,
    QFileDialog,
    QLabel,
    QFrame,
    QMessageBox,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QTextCursor

from automation import run_automation, reset_login, AutomationWorker


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Sharjeel Web Automation")
        self.resize(600, 700)
        self.is_running = False
        self.worker = None
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout()

        title = QLabel("Sharjeel Web Automation")
        title.setStyleSheet("font-size: 20px; font-weight: bold;")
        layout.addWidget(title)
        layout.addWidget(self.make_divider())

        excel_row = QHBoxLayout()
        excel_label = QLabel("Excel File:")
        excel_label.setFixedWidth(100)
        self.excel_input = QLineEdit()
        self.excel_input.setPlaceholderText("C:\\path\\to\\file.xlsx")
        excel_browse_btn = QPushButton("Browse")
        excel_browse_btn.clicked.connect(self.pick_excel_file)
        excel_row.addWidget(excel_label)
        excel_row.addWidget(self.excel_input)
        excel_row.addWidget(excel_browse_btn)
        layout.addLayout(excel_row)

        column_row = QHBoxLayout()
        column_label = QLabel("Column:")
        column_label.setFixedWidth(100)
        self.column_input = QLineEdit()
        self.column_input.setPlaceholderText("e.g., Roll No")
        column_row.addWidget(column_label)
        column_row.addWidget(self.column_input)
        layout.addLayout(column_row)

        output_row = QHBoxLayout()
        output_label = QLabel("Output:")
        output_label.setFixedWidth(100)
        self.output_input = QLineEdit()
        self.output_input.setPlaceholderText("C:\\path\\to\\downloads")
        output_browse_btn = QPushButton("Browse")
        output_browse_btn.clicked.connect(self.pick_output_directory)
        output_row.addWidget(output_label)
        output_row.addWidget(self.output_input)
        output_row.addWidget(output_browse_btn)
        layout.addLayout(output_row)

        email_row = QHBoxLayout()
        email_label = QLabel("Email:")
        email_label.setFixedWidth(100)
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("user@example.com")
        email_row.addWidget(email_label)
        email_row.addWidget(self.email_input)
        layout.addLayout(email_row)

        password_row = QHBoxLayout()
        password_label = QLabel("Password:")
        password_label.setFixedWidth(100)
        self.password_input = QLineEdit()
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("••••••••")
        password_row.addWidget(password_label)
        password_row.addWidget(self.password_input)
        layout.addLayout(password_row)

        layout.addWidget(self.make_divider())

        self.reset_btn = QPushButton("Reset Login")
        self.reset_btn.clicked.connect(self.on_reset)
        layout.addWidget(self.reset_btn)

        self.continue_btn = QPushButton("Continue")
        self.continue_btn.setEnabled(False)
        self.continue_btn.clicked.connect(self.on_continue)
        layout.addWidget(self.continue_btn)

        layout.addWidget(self.make_divider())

        self.run_btn = QPushButton("Run Automation")
        self.run_btn.setFixedHeight(50)
        self.run_btn.clicked.connect(self.on_run)
        layout.addWidget(self.run_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.progress_label = QLabel("")
        self.progress_label.setVisible(False)
        layout.addWidget(self.progress_label)

        layout.addWidget(self.make_divider())

        log_title = QLabel("Log Output:")
        log_title.setStyleSheet("font-weight: bold;")
        layout.addWidget(log_title)

        self.log_area = QTextEdit()
        self.log_area.setReadOnly(True)
        self.log_area.setMaximumHeight(200)
        layout.addWidget(self.log_area)

        self.setLayout(layout)

    def make_divider(self):
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        return line

    def pick_excel_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx);;All Files (*)")
        if path:
            self.excel_input.setText(path)

    def pick_output_directory(self):
        path = QFileDialog.getExistingDirectory(self, "Select Output Directory", "")
        if path:
            self.output_input.setText(path)

    def log(self, msg):
        self.log_area.append(msg)
        self.log_area.moveCursor(QTextCursor.End)
        self.log_area.ensureCursorVisible()

    def on_reset(self):
        if reset_login():
            self.log("Login session reset successfully!")
        else:
            self.log("No login session found to reset.")

    def on_continue(self):
        if self.worker:
            self.worker.resolve_captcha()
            self.continue_btn.setEnabled(False)
            self.log("Continuing...")

    def on_run(self):
        if self.is_running:
            self.log("Automation is already running!")
            return

        excel_path = self.excel_input.text()
        column_name = self.column_input.text()
        output_dir = self.output_input.text()
        email = self.email_input.text()
        password = self.password_input.text()

        if not excel_path:
            self.log("Error: Please select an Excel file.")
            return
        if not column_name:
            self.log("Error: Please enter a column name.")
            return
        if not output_dir:
            self.log("Error: Please select an output directory.")
            return
        if not email:
            self.log("Error: Please enter an email address.")
            return
        if not password:
            self.log("Error: Please enter a password.")
            return

        if not os.path.exists(excel_path):
            self.log(f"Error: Excel file not found: {excel_path}")
            return

        self.is_running = True
        self.run_btn.setEnabled(False)
        self.reset_btn.setEnabled(False)
        self.continue_btn.setEnabled(False)

        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.progress_bar.setValue(0)

        self.log(f"Starting automation...")
        self.log(f"Excel: {excel_path}")
        self.log(f"Column: {column_name}")
        self.log(f"Output: {output_dir}")
        self.log("-" * 40)

        self.worker = AutomationWorker(excel_path, column_name, output_dir, email, password)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.on_progress)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error_signal.connect(self.on_error)
        self.worker.captcha_pending_signal.connect(self.on_captcha_pending)
        self.worker.captcha_resolved_signal.connect(self.on_captcha_resolved)
        self.worker.start()

    def on_captcha_pending(self):
        self.log("CAPTCHA detected! Please solve it and click Continue.")
        self.continue_btn.setEnabled(True)

    def on_captcha_resolved(self):
        self.log("CAPTCHA resolved. Continuing...")

    def on_progress(self, current, total, status):
        if total > 0:
            self.progress_bar.setValue(int(current / total * 100))
            self.progress_label.setText(f"{status} ({current}/{total})" if status else f"{current}/{total}")
        else:
            self.progress_bar.setValue(0)
            self.progress_label.setText(status)

    def on_finished(self):
        self.is_running = False
        self.run_btn.setEnabled(True)
        self.reset_btn.setEnabled(True)
        self.continue_btn.setEnabled(False)
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)
        self.log("-" * 40)
        self.log("Automation complete!")

    def on_error(self, error_msg):
        self.log(f"Error: {error_msg}")

    def closeEvent(self, event):
        if self.is_running and self.worker:
            self.worker.terminate()
            self.worker.wait()
        event.accept()


def main():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()

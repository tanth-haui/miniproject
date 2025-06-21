import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox
from GUI import Ui_MainWindow
from funtion_process import run_processing
from pathlib import Path

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.setWindowTitle("Mini Project GUI")

        # Kết nối nút
        self.ui.Input_button.clicked.connect(self.select_input_folder)
        self.ui.Output_button.clicked.connect(self.select_output_folder)
        self.ui.Start_button.clicked.connect(self.start_processing)

    def select_input_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục Input")
        if folder:
            self.ui.Input_Folder.setText(folder)

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục Output")
        if folder:
            self.ui.Output_Folder.setText(folder)

    def start_processing(self):
        input_path = self.ui.Input_Folder.text()
        output_path = self.ui.Output_Folder.text()
        start_time = self.ui.Input_Start_Time.text()
        end_time = self.ui.Input_End_Time.text()
        signal = self.ui.Signal.currentText()

        # Kiểm tra đầu vào
        if not input_path or not output_path or not start_time or not end_time:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng nhập đầy đủ thông tin.")
            return

        if not Path(input_path).is_dir():
            QMessageBox.critical(self, "Lỗi", "❌ Thư mục Input không tồn tại.")
            return

        # ✅ KHÔNG kiểm tra thư mục output — để cho phép tạo tự động
        print(f"[DEBUG] Output folder được nhập: {output_path}")  # Debug nhẹ nếu cần

        try:
            run_processing(input_path, output_path, signal, start_time, end_time)
            QMessageBox.information(self, "Thành công", "✅ Xử lý hoàn tất!")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi xử lý", f"❌ Đã xảy ra lỗi:\n{str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

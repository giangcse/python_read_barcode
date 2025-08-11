import sys
import os
import cv2
import sqlite3
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QFileDialog,
    QMessageBox,
    QDateEdit,
)
from PyQt5.QtCore import Qt, QTimer, QDate
from PyQt5.QtGui import QImage, QPixmap
from pyzbar import pyzbar
import winsound  # Thư viện phát âm thanh (Windows)
import time
import pyperclip  # Thư viện để sao chép vào clipboard
from openpyxl import Workbook

class BarcodeReaderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Barcode Reader")
        self.setGeometry(100, 100, 800, 600)

        # Tạo widget chính và layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.layout = QVBoxLayout(self.central_widget)

        # Vùng control: chọn ngày và xuất Excel
        controls_layout = QHBoxLayout()
        self.layout.addLayout(controls_layout)

        controls_layout.addWidget(QLabel("Từ ngày:"))
        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDisplayFormat("dd/MM/yyyy")
        self.from_date.setDate(QDate.currentDate())
        controls_layout.addWidget(self.from_date)

        controls_layout.addWidget(QLabel("Đến ngày:"))
        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDisplayFormat("dd/MM/yyyy")
        self.to_date.setDate(QDate.currentDate())
        controls_layout.addWidget(self.to_date)

        self.export_btn = QPushButton("Xuất Excel")
        self.export_btn.clicked.connect(self.export_to_excel)
        controls_layout.addWidget(self.export_btn)

        # Label để hiển thị video từ webcam
        self.video_label = QLabel()
        self.video_label.setAlignment(Qt.AlignCenter)
        self.layout.addWidget(self.video_label)

        # Ô input để hiển thị kết quả barcode
        self.result_input = QLineEdit()
        self.result_input.setReadOnly(True)
        self.result_input.setPlaceholderText("Barcode sẽ hiển thị ở đây...")
        self.layout.addWidget(self.result_input)

        # Khởi tạo webcam
        self.capture = cv2.VideoCapture(0)
        self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

        # Timer để cập nhật frame
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.timer.start(30)  # Cập nhật mỗi 30ms

        # Biến để theo dõi barcode cuối cùng được đọc
        self.last_barcode = None
        self.last_beep_time = 0

        # Khởi tạo cơ sở dữ liệu SQLite
        self.init_database()

    def init_database(self):
        """Khởi tạo file DB và bảng lưu lịch sử quét."""
        db_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "barcodes.db")
        self.conn = sqlite3.connect(db_path)
        # Tạo bảng nếu chưa có (kèm các cột mới)
        self.conn.execute(
            """
            CREATE TABLE IF NOT EXISTS scans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                content TEXT NOT NULL,
                scanned_at TEXT NOT NULL,
                scan_date TEXT,
                scan_time TEXT
            )
            """
        )
        # Đảm bảo các cột mới tồn tại nếu DB cũ
        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(scans)")
        existing_cols = {row[1] for row in cursor.fetchall()}
        if "scan_date" not in existing_cols:
            self.conn.execute("ALTER TABLE scans ADD COLUMN scan_date TEXT")
        if "scan_time" not in existing_cols:
            self.conn.execute("ALTER TABLE scans ADD COLUMN scan_time TEXT")

        # Backfill từ scanned_at nếu thiếu
        self.conn.execute(
            "UPDATE scans SET scan_date = substr(scanned_at, 1, 10) WHERE scan_date IS NULL OR scan_date = ''"
        )
        self.conn.execute(
            "UPDATE scans SET scan_time = substr(scanned_at, 12, 8) WHERE scan_time IS NULL OR scan_time = ''"
        )

        # Index phục vụ lọc theo ngày
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_scans_scan_date ON scans(scan_date)")
        self.conn.execute("CREATE INDEX IF NOT EXISTS idx_scans_scanned_at ON scans(scanned_at)")
        self.conn.commit()

    def save_scan(self, content: str) -> None:
        """Lưu một dòng quét vào DB với thời gian hệ thống."""
        now = datetime.now()
        scanned_at = now.strftime("%Y-%m-%d %H:%M:%S")
        scan_date = now.strftime("%Y-%m-%d")
        scan_time = now.strftime("%H:%M:%S")
        self.conn.execute(
            "INSERT INTO scans(content, scanned_at, scan_date, scan_time) VALUES(?, ?, ?, ?)",
            (content, scanned_at, scan_date, scan_time),
        )
        self.conn.commit()

    def export_to_excel(self) -> None:
        """Xuất dữ liệu từ DB theo khoảng ngày ra file Excel (.xlsx)."""
        start_qdate: QDate = self.from_date.date()
        end_qdate: QDate = self.to_date.date()

        start_date = start_qdate.toString("yyyy-MM-dd")
        end_date = end_qdate.toString("yyyy-MM-dd")

        # Đảm bảo khoảng ngày hợp lệ
        if start_qdate > end_qdate:
            QMessageBox.warning(self, "Khoảng ngày không hợp lệ", "'Từ ngày' phải nhỏ hơn hoặc bằng 'Đến ngày'.")
            return

        cursor = self.conn.cursor()
        cursor.execute(
            """
            SELECT id, content, scan_date, scan_time
            FROM scans
            WHERE scan_date >= ? AND scan_date <= ?
            ORDER BY scan_date ASC, scan_time ASC
            """,
            (start_date, end_date),
        )
        rows = cursor.fetchall()

        if not rows:
            QMessageBox.information(self, "Không có dữ liệu", "Không tìm thấy bản ghi nào trong khoảng ngày đã chọn.")
            return

        # Hộp thoại chọn nơi lưu
        default_filename = f"scans_{start_date}_to_{end_date}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu file Excel",
            os.path.join(os.path.expanduser("~"), default_filename),
            "Excel Files (*.xlsx)",
        )
        if not save_path:
            return

        # Ghi Excel bằng openpyxl
        wb = Workbook()
        ws = wb.active
        ws.title = "Scans"
        ws.append(["ID", "Nội dung", "Ngày", "Giờ"])
        for rid, content, scan_date, scan_time in rows:
            # Định dạng ngày hiển thị dd/MM/yyyy
            try:
                d_disp = datetime.strptime(scan_date, "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                d_disp = scan_date or ""
            ws.append([rid, content, d_disp, scan_time])

        try:
            wb.save(save_path)
            QMessageBox.information(self, "Thành công", f"Đã lưu Excel:\n{save_path}")
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lưu file: {e}")

    def update_frame(self):
        ret, frame = self.capture.read()
        if ret:
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            # Đọc barcode từ frame
            barcodes = pyzbar.decode(gray)
            
            # Nếu phát hiện barcode
            for barcode in barcodes:
                barcode_data = barcode.data.decode("utf-8")
                
                # Kiểm tra barcode có đúng định dạng không
                if True:
                    # Chỉ phát âm thanh, cập nhật và copy nếu barcode mới hoặc sau 1 giây
                    current_time = time.time()
                    if barcode_data != self.last_barcode or (current_time - self.last_beep_time > 1):
                        self.result_input.setText(barcode_data)
                        self.play_beep_sound()
                        pyperclip.copy(barcode_data)  # Sao chép barcode vào clipboard
                        self.save_scan(barcode_data)  # Lưu vào DB
                        self.last_barcode = barcode_data
                        self.last_beep_time = current_time
                
                # Vẽ khung bao quanh barcode
                (x, y, w, h) = barcode.rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)

            # Chuyển frame sang định dạng RGB
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            # Chuyển thành QImage
            h, w, ch = frame.shape
            bytes_per_line = ch * w
            image = QImage(frame.data, w, h, bytes_per_line, QImage.Format_RGB888)
            
            # Chuyển thành QPixmap và hiển thị
            pixmap = QPixmap.fromImage(image)
            self.video_label.setPixmap(pixmap.scaled(
                self.video_label.size(),
                Qt.KeepAspectRatio,
                Qt.SmoothTransformation
            ))

    def play_beep_sound(self):
        # Phát âm thanh bíp (Windows)
        winsound.Beep(1000, 200)  # Tần số 1000Hz, thời gian 200ms

    def closeEvent(self, event):
        # Giải phóng webcam khi đóng ứng dụng
        self.capture.release()
        try:
            if hasattr(self, "conn"):
                self.conn.close()
        except Exception:
            pass
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = BarcodeReaderApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()

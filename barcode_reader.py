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
    QFrame,
    QGroupBox,
    QGridLayout,
    QSpacerItem,
    QSizePolicy,
)
from PyQt5.QtCore import Qt, QTimer, QDate
from PyQt5.QtGui import QImage, QPixmap, QFont, QPalette, QColor
from pyzbar import pyzbar
import winsound  # Thư viện phát âm thanh (Windows)
import time
import pyperclip  # Thư viện để sao chép vào clipboard
from openpyxl import Workbook

class BarcodeReaderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ứng dụng Đọc Barcode")
        self.setGeometry(100, 100, 800, 700)
        self.setMinimumSize(800, 600)

        # Thiết lập icon cho ứng dụng
        try:
            self.setWindowIcon(QPixmap("Logo.ico"))
        except:
            pass

        # Tạo widget chính và layout
        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QVBoxLayout(self.central_widget)
        self.main_layout.setSpacing(15)
        self.main_layout.setContentsMargins(20, 20, 20, 20)

        # Tạo header
        self.create_header()

        # Tạo vùng control
        self.create_control_panel()

        # Tạo vùng hiển thị video
        self.create_video_panel()

        # Tạo vùng kết quả
        self.create_result_panel()

        # Khởi tạo webcam
        self.capture = cv2.VideoCapture(0)
        self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)

        # Timer để cập nhật frame
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.timer.start(30)  # Cập nhật mỗi 30ms

        # Timer để reset trạng thái
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.reset_status)
        self.status_timer.setSingleShot(True)

        # Biến để theo dõi barcode cuối cùng được đọc
        self.last_barcode = None
        self.last_beep_time = 0

        # Khởi tạo cơ sở dữ liệu SQLite
        self.init_database()

        # Áp dụng stylesheet
        self.apply_stylesheet()

    def create_header(self):
        """Tạo header của ứng dụng"""
        header_frame = QFrame()
        header_frame.setFrameStyle(QFrame.StyledPanel)
        header_frame.setMaximumHeight(80)
        header_layout = QHBoxLayout(header_frame)

        # Logo và tiêu đề
        title_label = QLabel("📱 Ứng dụng Đọc Barcode")
        title_font = QFont("Arial", 18, QFont.Bold)
        title_label.setFont(title_font)
        title_label.setStyleSheet("color: #2c3e50;")

        # Thông tin phiên bản
        version_label = QLabel("v1.0")
        version_label.setStyleSheet("color: #7f8c8d; font-size: 12px;")

        header_layout.addWidget(title_label)
        header_layout.addStretch()
        header_layout.addWidget(version_label)

        self.main_layout.addWidget(header_frame)

    def create_control_panel(self):
        """Tạo panel điều khiển"""
        control_group = QGroupBox("⚙️ Cài đặt Xuất Excel")
        control_group.setMaximumHeight(120)
        control_layout = QGridLayout(control_group)

        # Từ ngày
        from_label = QLabel("Từ ngày:")
        from_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDisplayFormat("dd/MM/yyyy")
        self.from_date.setDate(QDate.currentDate())
        self.from_date.setStyleSheet(
            """
            QDateEdit {
                padding: 8px;
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                background-color: white;
                font-size: 12px;
            }
            QDateEdit:focus {
                border-color: #3498db;
            }
        """
        )

        # Đến ngày
        to_label = QLabel("Đến ngày:")
        to_label.setStyleSheet("font-weight: bold; color: #2c3e50;")
        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDisplayFormat("dd/MM/yyyy")
        self.to_date.setDate(QDate.currentDate())
        self.to_date.setStyleSheet(
            """
            QDateEdit {
                padding: 8px;
                border: 2px solid #bdc3c7;
                border-radius: 5px;
                background-color: white;
                font-size: 12px;
            }
            QDateEdit:focus {
                border-color: #3498db;
            }
        """
        )

        # Nút xuất Excel
        self.export_btn = QPushButton("📊 Xuất Excel")
        self.export_btn.clicked.connect(self.export_to_excel)
        self.export_btn.setStyleSheet(
            """
            QPushButton {
                background-color: #27ae60;
                color: white;
                border: none;
                padding: 10px 20px;
                border-radius: 5px;
                font-weight: bold;
                font-size: 12px;
            }
            QPushButton:hover {
                background-color: #2ecc71;
            }
            QPushButton:pressed {
                background-color: #229954;
            }
        """
        )

        control_layout.addWidget(from_label, 0, 0)
        control_layout.addWidget(self.from_date, 0, 1)
        control_layout.addWidget(to_label, 0, 2)
        control_layout.addWidget(self.to_date, 0, 3)
        control_layout.addWidget(self.export_btn, 0, 4)

        self.main_layout.addWidget(control_group)

    def create_video_panel(self):
        """Tạo panel hiển thị video"""
        video_group = QGroupBox("📹 Camera")
        video_layout = QVBoxLayout(video_group)

        # Label để hiển thị video từ webcam
        self.video_label = QLabel()
        self.video_label.setAlignment(Qt.AlignCenter)
        self.video_label.setMinimumSize(640, 360)
        self.video_label.setStyleSheet(
            """
            QLabel {
                border: 3px solid #bdc3c7;
                border-radius: 10px;
                background-color: #ecf0f1;
            }
        """
        )

        # Thông báo trạng thái camera
        self.status_label = QLabel(
            "🟢 Camera đang hoạt động - Đưa barcode vào khung hình"
        )
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet(
            """
            QLabel {
                color: #27ae60;
                font-weight: bold;
                font-size: 12px;
                padding: 5px;
            }
        """
        )

        video_layout.addWidget(self.video_label)
        video_layout.addWidget(self.status_label)

        self.main_layout.addWidget(video_group)

    def create_result_panel(self):
        """Tạo panel hiển thị kết quả"""
        result_group = QGroupBox("📋 Kết quả đọc barcode")
        result_layout = QVBoxLayout(result_group)

        # Ô input để hiển thị kết quả barcode
        self.result_input = QLineEdit()
        self.result_input.setReadOnly(True)
        self.result_input.setPlaceholderText("Barcode sẽ hiển thị ở đây...")
        self.result_input.setStyleSheet(
            """
            QLineEdit {
                padding: 15px;
                border: 2px solid #3498db;
                border-radius: 8px;
                background-color: #f8f9fa;
                font-size: 16px;
                font-weight: bold;
                color: #2c3e50;
            }
            QLineEdit:focus {
                border-color: #2980b9;
                background-color: white;
            }
        """
        )

        # Thông tin bổ sung
        info_label = QLabel("💡 Barcode sẽ tự động được sao chép vào clipboard")
        info_label.setAlignment(Qt.AlignCenter)
        info_label.setStyleSheet(
            """
            QLabel {
                color: #7f8c8d;
                font-size: 11px;
                font-style: italic;
            }
        """
        )

        result_layout.addWidget(self.result_input)
        result_layout.addWidget(info_label)

        self.main_layout.addWidget(result_group)

    def apply_stylesheet(self):
        """Áp dụng stylesheet cho toàn bộ ứng dụng"""
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #f5f6fa;
            }
            QGroupBox {
                font-weight: bold;
                font-size: 13px;
                color: #2c3e50;
                border: 2px solid #bdc3c7;
                border-radius: 8px;
                margin-top: 10px;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 5px 0 5px;
            }
            QLabel {
                color: #2c3e50;
            }
        """
        )

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
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("⚠️ Cảnh báo")
            msg.setText("Khoảng ngày không hợp lệ")
            msg.setInformativeText("'Từ ngày' phải nhỏ hơn hoặc bằng 'Đến ngày'.")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
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
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("ℹ️ Thông báo")
            msg.setText("Không có dữ liệu")
            msg.setInformativeText(
                "Không tìm thấy bản ghi nào trong khoảng ngày đã chọn."
            )
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
            return

        # Hộp thoại chọn nơi lưu
        default_filename = f"barcode_scans_{start_date}_to_{end_date}.xlsx"
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "💾 Lưu file Excel",
            os.path.join(os.path.expanduser("~"), "Desktop", default_filename),
            "Excel Files (*.xlsx)",
        )
        if not save_path:
            return

        # Ghi Excel bằng openpyxl với định dạng đẹp hơn
        wb = Workbook()
        ws = wb.active
        ws.title = "Barcode Scans"

        # Định dạng header
        headers = ["STT", "Nội dung Barcode", "Ngày quét", "Giờ quét"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col, value=header)
            cell.font = cell.font.copy(bold=True)
            cell.fill = cell.fill.copy(fill_type="solid", fgColor="CCCCCC")

        # Thêm dữ liệu
        for row_idx, (rid, content, scan_date, scan_time) in enumerate(rows, 2):
            # Định dạng ngày hiển thị dd/MM/yyyy
            try:
                d_disp = datetime.strptime(scan_date, "%Y-%m-%d").strftime("%d/%m/%Y")
            except Exception:
                d_disp = scan_date or ""

            ws.cell(row=row_idx, column=1, value=row_idx - 1)  # STT
            ws.cell(row=row_idx, column=2, value=content)  # Nội dung
            ws.cell(row=row_idx, column=3, value=d_disp)  # Ngày
            ws.cell(row=row_idx, column=4, value=scan_time)  # Giờ

        # Tự động điều chỉnh độ rộng cột
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

        try:
            wb.save(save_path)
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            msg.setWindowTitle("✅ Thành công")
            msg.setText("Đã xuất Excel thành công!")
            msg.setInformativeText(f"File được lưu tại:\n{save_path}")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()
        except Exception as e:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Critical)
            msg.setWindowTitle("❌ Lỗi")
            msg.setText("Không thể lưu file")
            msg.setInformativeText(f"Lỗi: {e}")
            msg.setStandardButtons(QMessageBox.Ok)
            msg.exec_()

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

                        # Cập nhật trạng thái
                        self.status_label.setText("✅ Đã đọc thành công barcode!")
                        self.status_label.setStyleSheet(
                            """
                            QLabel {
                                color: #27ae60;
                                font-weight: bold;
                                font-size: 12px;
                                padding: 5px;
                            }
                        """
                        )
                        # Reset trạng thái sau 3 giây
                        self.status_timer.start(3000)

                # Vẽ khung bao quanh barcode với màu sắc đẹp hơn
                (x, y, w, h) = barcode.rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 3)
                # Thêm text "BARCODE" phía trên khung
                cv2.putText(
                    frame,
                    "BARCODE",
                    (x, y - 10),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.7,
                    (0, 255, 0),
                    2,
                )

            # Chuyển frame sang định dạng RGB
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            # Chuyển thành QImage
            h, w, ch = frame.shape
            bytes_per_line = ch * w
            image = QImage(frame.data, w, h, bytes_per_line, QImage.Format_RGB888)

            # Chuyển thành QPixmap và hiển thị
            pixmap = QPixmap.fromImage(image)
            scaled_pixmap = pixmap.scaled(
                self.video_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation
            )
            self.video_label.setPixmap(scaled_pixmap)
        else:
            # Nếu không đọc được camera
            self.status_label.setText("❌ Không thể kết nối camera")
            self.status_label.setStyleSheet(
                """
                QLabel {
                    color: #e74c3c;
                    font-weight: bold;
                    font-size: 12px;
                    padding: 5px;
                }
            """
            )

    def reset_status(self):
        """Reset trạng thái về mặc định"""
        self.status_label.setText(
            "🟢 Camera đang hoạt động - Đưa barcode vào khung hình"
        )
        self.status_label.setStyleSheet(
            """
            QLabel {
                color: #27ae60;
                font-weight: bold;
                font-size: 12px;
                padding: 5px;
            }
        """
        )

    def play_beep_sound(self):
        """Phát âm thanh bíp khi đọc thành công barcode"""
        try:
            # Phát âm thanh bíp (Windows) - âm thanh vui vẻ hơn
            winsound.Beep(800, 150)  # Tần số 800Hz, thời gian 150ms
            time.sleep(0.05)
            winsound.Beep(1000, 150)  # Tần số 1000Hz, thời gian 150ms
        except:
            pass  # Bỏ qua nếu không thể phát âm thanh

    def closeEvent(self, event):
        """Xử lý khi đóng ứng dụng"""
        # Dừng timers
        if hasattr(self, "timer"):
            self.timer.stop()
        if hasattr(self, "status_timer"):
            self.status_timer.stop()

        # Giải phóng webcam khi đóng ứng dụng
        if hasattr(self, "capture"):
            self.capture.release()

        # Đóng kết nối database
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

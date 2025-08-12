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
    QGraphicsOpacityEffect,
)
from PyQt5.QtCore import (
    Qt,
    QTimer,
    QDate,
    QPropertyAnimation,
    QEasingCurve,
    QThread,
    pyqtSignal,
)
from PyQt5.QtGui import QImage, QPixmap, QFont, QPalette, QColor
from pyzbar import pyzbar
import winsound  # Thư viện phát âm thanh (Windows)
import time
import pyperclip  # Thư viện để sao chép vào clipboard
from openpyxl import Workbook
import threading
from queue import Queue


class DatabaseWorker(QThread):
    """Worker thread để xử lý database operations async"""

    def __init__(self, db_path):
        super().__init__()
        self.db_path = db_path
        self.queue = Queue()
        self.running = True

    def run(self):
        """Chạy thread xử lý database"""
        # Tạo connection riêng cho thread này
        conn = sqlite3.connect(self.db_path)

        while self.running:
            try:
                # Lấy task từ queue với timeout
                task = self.queue.get(timeout=1)
                if task is None:
                    break

                operation, data = task
                if operation == "save_scan":
                    content, scanned_at, scan_date, scan_time = data
                    conn.execute(
                        "INSERT INTO scans(content, scanned_at, scan_date, scan_time) VALUES(?, ?, ?, ?)",
                        (content, scanned_at, scan_date, scan_time),
                    )
                    conn.commit()

            except:
                continue  # Timeout hoặc lỗi, tiếp tục loop

        conn.close()

    def add_task(self, operation, data):
        """Thêm task vào queue"""
        self.queue.put((operation, data))

    def stop(self):
        """Dừng worker"""
        self.running = False
        self.queue.put(None)


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
        # self.create_header()

        # Thêm spacing đẹp hơn
        self.main_layout.addSpacing(10)

        # Tạo vùng control
        self.create_control_panel()
        self.main_layout.addSpacing(5)

        # Tạo vùng hiển thị video
        self.create_video_panel()
        self.main_layout.addSpacing(5)

        # Tạo vùng kết quả
        self.create_result_panel()
        self.main_layout.addSpacing(10)

        # Khởi tạo webcam với tối ưu hiệu suất
        self.capture = cv2.VideoCapture(0)
        # Giảm resolution để tăng tốc độ xử lý
        self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
        self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
        # Giảm buffer size để tránh lag
        self.capture.set(cv2.CAP_PROP_BUFFERSIZE, 1)
        # Tăng FPS nếu có thể
        self.capture.set(cv2.CAP_PROP_FPS, 30)

        # Timer để cập nhật frame - cân bằng giữa hiệu suất và detection
        self.timer = QTimer()
        self.timer.timeout.connect(self.update_frame)
        self.timer.start(40)  # Cập nhật mỗi 40ms (25 FPS) để balance tốt hơn

        # Cache cho tối ưu detection
        self.frame_count = 0
        self.detection_interval = 2  # Giảm interval để detect tốt hơn

        # Timer để reset trạng thái
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.reset_status)
        self.status_timer.setSingleShot(True)

        # Biến để theo dõi barcode cuối cùng được đọc
        self.last_barcode = None
        self.last_beep_time = 0

        # Hiệu ứng animation cho kết quả
        self.success_effect = QGraphicsOpacityEffect()
        self.result_input.setGraphicsEffect(self.success_effect)

        self.success_animation = QPropertyAnimation(self.success_effect, b"opacity")
        self.success_animation.setDuration(800)
        self.success_animation.setEasingCurve(QEasingCurve.OutCubic)

        # Khởi tạo cơ sở dữ liệu SQLite
        self.init_database()

        # Khởi tạo database worker thread
        self.db_worker = DatabaseWorker(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "barcodes.db")
        )
        self.db_worker.start()

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
        from_label = QLabel("📅 Từ ngày:")
        from_label.setStyleSheet("font-weight: 600; color: #495057; font-size: 13px;")
        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDisplayFormat("dd/MM/yyyy")
        self.from_date.setDate(QDate.currentDate())
        self.from_date.setStyleSheet(
            """
            QDateEdit {
                padding: 12px 15px;
                border: 2px solid #e1e8ed;
                border-radius: 8px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #f8f9fa);
                font-size: 13px;
                font-weight: 500;
                color: #495057;
                min-height: 20px;
            }
            QDateEdit:focus {
                border-color: #3498db;
                background: #ffffff;
                box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);
            }
            QDateEdit:hover {
                border-color: #74b9ff;
                background: #ffffff;
            }
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #e1e8ed;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
                background: #f8f9fa;
            }
            QDateEdit::down-arrow {
                image: none;
                border: 2px solid #6c757d;
                border-radius: 2px;
                background: #6c757d;
                width: 8px;
                height: 8px;
            }
        """
        )

        # Đến ngày
        to_label = QLabel("📅 Đến ngày:")
        to_label.setStyleSheet("font-weight: 600; color: #495057; font-size: 13px;")
        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDisplayFormat("dd/MM/yyyy")
        self.to_date.setDate(QDate.currentDate())
        self.to_date.setStyleSheet(
            """
            QDateEdit {
                padding: 12px 15px;
                border: 2px solid #e1e8ed;
                border-radius: 8px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #f8f9fa);
                font-size: 13px;
                font-weight: 500;
                color: #495057;
                min-height: 20px;
            }
            QDateEdit:focus {
                border-color: #3498db;
                background: #ffffff;
                box-shadow: 0 0 0 3px rgba(52, 152, 219, 0.1);
            }
            QDateEdit:hover {
                border-color: #74b9ff;
                background: #ffffff;
            }
            QDateEdit::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #e1e8ed;
                border-top-right-radius: 8px;
                border-bottom-right-radius: 8px;
                background: #f8f9fa;
            }
            QDateEdit::down-arrow {
                image: none;
                border: 2px solid #6c757d;
                border-radius: 2px;
                background: #6c757d;
                width: 8px;
                height: 8px;
            }
        """
        )

        # Nút xuất Excel
        self.export_btn = QPushButton("📊 Xuất Excel")
        self.export_btn.clicked.connect(self.export_to_excel)
        self.export_btn.setStyleSheet(
            """
            QPushButton {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #00b894, stop: 1 #00a085);
                color: white;
                border: none;
                padding: 12px 24px;
                border-radius: 10px;
                font-weight: 600;
                font-size: 13px;
                min-height: 20px;
                box-shadow: 0 4px 15px rgba(0, 184, 148, 0.3);
            }
            QPushButton:hover {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #00d4aa, stop: 1 #00b894);
                transform: translateY(-2px);
                box-shadow: 0 6px 20px rgba(0, 184, 148, 0.4);
            }
            QPushButton:pressed {
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #00a085, stop: 1 #008f75);
                transform: translateY(0px);
                box-shadow: 0 2px 8px rgba(0, 184, 148, 0.2);
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
                border: 2px solid #e1e8ed;
                border-radius: 15px;
                background: qlineargradient(x1: 0, y1: 0, x2: 1, y2: 1,
                    stop: 0 #f8f9fa, stop: 1 #e9ecef);
                padding: 10px;
                box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.1);
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
                padding: 18px 20px;
                border: 2px solid #74b9ff;
                border-radius: 12px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #f1f3f4);
                font-size: 18px;
                font-weight: 600;
                color: #2d3436;
                font-family: 'Consolas', 'Monaco', monospace;
                min-height: 25px;
                box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.05);
            }
            QLineEdit:focus {
                border-color: #0984e3;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #e3f2fd);
                box-shadow: 0 0 0 3px rgba(9, 132, 227, 0.1);
            }
            QLineEdit[text=""]:!focus {
                color: #636e72;
                font-style: italic;
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
        """Áp dụng stylesheet tối ưu cho toàn bộ ứng dụng"""
        # Dùng stylesheet đơn giản hơn để giảm overhead
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #f5f6fa;
                color: #2c3e50;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            
            QGroupBox {
                font-weight: 600;
                font-size: 14px;
                color: #2c3e50;
                border: 1px solid #d0d7de;
                border-radius: 10px;
                margin-top: 12px;
                padding-top: 15px;
                background-color: #ffffff;
            }
            
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 15px;
                padding: 5px 10px;
                background-color: #ffffff;
                border: 1px solid #e1e8ed;
                border-radius: 6px;
                color: #495057;
                font-weight: 600;
            }
            
            QLabel {
                color: #495057;
                font-weight: 500;
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
        """Lưu một dòng quét vào DB với thời gian hệ thống (đồng bộ)."""
        now = datetime.now()
        scanned_at = now.strftime("%Y-%m-%d %H:%M:%S")
        scan_date = now.strftime("%Y-%m-%d")
        scan_time = now.strftime("%H:%M:%S")
        self.conn.execute(
            "INSERT INTO scans(content, scanned_at, scan_date, scan_time) VALUES(?, ?, ?, ?)",
            (content, scanned_at, scan_date, scan_time),
        )
        self.conn.commit()

    def save_scan_async(self, content: str) -> None:
        """Lưu một dòng quét vào DB async để không block UI."""
        now = datetime.now()
        scanned_at = now.strftime("%Y-%m-%d %H:%M:%S")
        scan_date = now.strftime("%Y-%m-%d")
        scan_time = now.strftime("%H:%M:%S")

        # Gửi task vào queue để xử lý async
        self.db_worker.add_task(
            "save_scan", (content, scanned_at, scan_date, scan_time)
        )

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
        """Cập nhật frame với tối ưu hiệu suất"""
        ret, frame = self.capture.read()
        if not ret:
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
            return

        # Tăng frame counter
        self.frame_count += 1

        # Chỉ detect barcode mỗi vài frame để tối ưu CPU
        barcodes = []
        if self.frame_count % self.detection_interval == 0:
            # Tăng kích thước frame detection để đọc barcode tốt hơn
            detection_frame = cv2.resize(frame, (480, 360))  # Tăng từ 320x240
            gray = cv2.cvtColor(detection_frame, cv2.COLOR_BGR2GRAY)

            # Cải thiện chất lượng ảnh cho detection
            # Tăng contrast
            gray = cv2.convertScaleAbs(gray, alpha=1.2, beta=10)

            # Blur nhẹ để giảm noise
            gray = cv2.GaussianBlur(gray, (3, 3), 0)

            # Detection với pyzbar
            barcodes = pyzbar.decode(gray)

            # Nếu không tìm thấy, thử với frame gốc (fallback)
            if not barcodes:
                original_gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
                barcodes = pyzbar.decode(original_gray)
                # Nếu tìm thấy với frame gốc, không cần scale coordinates
                if barcodes:
                    self.use_original_coords = True
                    print("Detected barcode using original frame")
                else:
                    self.use_original_coords = False
            else:
                self.use_original_coords = False
                print("Detected barcode using processed frame")

        # Xử lý barcode nếu tìm thấy
        for barcode in barcodes:
            try:
                barcode_data = barcode.data.decode("utf-8")

                # Chỉ xử lý nếu barcode mới hoặc sau 1 giây
                current_time = time.time()
                if barcode_data != self.last_barcode or (
                    current_time - self.last_beep_time > 1
                ):
                    self.result_input.setText(barcode_data)
                    self.play_beep_sound()
                    pyperclip.copy(barcode_data)
                    self.save_scan_async(barcode_data)  # Lưu DB async
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
                    self.status_timer.start(3000)
                    self.animate_success()

                # Vẽ khung bao quanh barcode
                (x, y, w, h) = barcode.rect

                # Scale coordinates nếu cần thiết
                if (
                    not hasattr(self, "use_original_coords")
                    or not self.use_original_coords
                ):
                    # Scale từ detection frame (480x360) về frame gốc (640x480)
                    scale_x, scale_y = frame.shape[1] / 480, frame.shape[0] / 360
                    x, y, w, h = (
                        int(x * scale_x),
                        int(y * scale_y),
                        int(w * scale_x),
                        int(h * scale_y),
                    )

                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                cv2.putText(
                    frame,
                    "BARCODE",
                    (x, y - 10),
                    cv2.FONT_HERSHEY_SIMPLEX,
                    0.6,
                    (0, 255, 0),
                    2,
                )

            except Exception as e:
                print(f"Lỗi decode barcode: {e}")

        # Tối ưu hiển thị frame
        self.display_frame(frame)

    def display_frame(self, frame):
        """Hiển thị frame với tối ưu memory"""
        # Chuyển sang RGB
        frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)

        # Tạo QImage
        h, w, ch = frame_rgb.shape
        bytes_per_line = ch * w

        # Tối ưu QImage creation
        qt_image = QImage(frame_rgb.data, w, h, bytes_per_line, QImage.Format_RGB888)

        # Scale với tối ưu
        pixmap = QPixmap.fromImage(qt_image)
        label_size = self.video_label.size()

        # Chỉ scale nếu cần thiết
        if pixmap.size() != label_size:
            scaled_pixmap = pixmap.scaled(
                label_size,
                Qt.KeepAspectRatio,
                Qt.FastTransformation,  # Dùng FastTransformation thay vì SmoothTransformation
            )
            self.video_label.setPixmap(scaled_pixmap)
        else:
            self.video_label.setPixmap(pixmap)

    def animate_success(self):
        """Tạo hiệu ứng khi đọc barcode thành công"""
        # Hiệu ứng nhấp nháy
        self.success_animation.setStartValue(0.3)
        self.success_animation.setEndValue(1.0)
        self.success_animation.start()

        # Thay đổi style tạm thời
        self.result_input.setStyleSheet(
            """
            QLineEdit {
                padding: 18px 20px;
                border: 2px solid #00b894;
                border-radius: 12px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #d4f9f0, stop: 1 #a8e6cf);
                font-size: 18px;
                font-weight: 600;
                color: #00695c;
                font-family: 'Consolas', 'Monaco', monospace;
                min-height: 25px;
                box-shadow: 0 0 15px rgba(0, 184, 148, 0.4);
            }
        """
        )

        # Timer để trở về style bình thường
        QTimer.singleShot(1500, self.reset_result_style)

    def reset_result_style(self):
        """Reset style của result input về bình thường"""
        self.result_input.setStyleSheet(
            """
            QLineEdit {
                padding: 18px 20px;
                border: 2px solid #74b9ff;
                border-radius: 12px;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #f1f3f4);
                font-size: 18px;
                font-weight: 600;
                color: #2d3436;
                font-family: 'Consolas', 'Monaco', monospace;
                min-height: 25px;
                box-shadow: inset 0 2px 4px rgba(0, 0, 0, 0.05);
            }
            QLineEdit:focus {
                border-color: #0984e3;
                background: qlineargradient(x1: 0, y1: 0, x2: 0, y2: 1,
                    stop: 0 #ffffff, stop: 1 #e3f2fd);
                box-shadow: 0 0 0 3px rgba(9, 132, 227, 0.1);
            }
            QLineEdit[text=""]:!focus {
                color: #636e72;
                font-style: italic;
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
        """Xử lý khi đóng ứng dụng với cleanup tối ưu"""
        # Dừng timers
        if hasattr(self, "timer"):
            self.timer.stop()
        if hasattr(self, "status_timer"):
            self.status_timer.stop()

        # Dừng animation
        if hasattr(self, "success_animation"):
            self.success_animation.stop()

        # Giải phóng webcam
        if hasattr(self, "capture"):
            self.capture.release()
            cv2.destroyAllWindows()  # Đóng tất cả cửa sổ OpenCV

        # Dừng database worker thread
        if hasattr(self, "db_worker"):
            self.db_worker.stop()
            self.db_worker.wait(3000)  # Đợi tối đa 3 giây

        # Đóng kết nối database chính
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

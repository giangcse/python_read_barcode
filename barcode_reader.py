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
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QGraphicsOpacityEffect,
    QSpacerItem,
    QSizePolicy,
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
from PyQt5.QtGui import QImage, QPixmap, QFont, QColor
from pyzbar import pyzbar
import winsound
import time
import pyperclip
from openpyxl import Workbook
from openpyxl.styles import Font
from queue import Queue

def get_base_path():
    """Lấy đường dẫn cơ sở, hoạt động cho cả script và file .exe từ PyInstaller"""
    if getattr(sys, 'frozen', False):
        # Chạy từ file .exe đã được đóng gói
        base_path = os.path.dirname(sys.executable)
    else:
        # Chạy từ file .py bình thường
        base_path = os.path.dirname(os.path.abspath(__file__))
    return base_path

class DatabaseWorker(QThread):
    # 2. TẠO MỘT TÍN HIỆU (SIGNAL)
    scan_saved = pyqtSignal() 

    def __init__(self, db_path):
        super().__init__()
        self.db_path = db_path
        self.queue = Queue()
        self.running = True

    def run(self):
        conn = sqlite3.connect(self.db_path)
        while self.running:
            try:
                task = self.queue.get(timeout=1)
                if task is None: break
                operation, data = task
                if operation == "save_scan":
                    content, scanned_at, scan_date, scan_time = data
                    conn.execute(
                        "INSERT INTO scans(content, scanned_at, scan_date, scan_time) VALUES(?, ?, ?, ?)",
                        (content, scanned_at, scan_date, scan_time),
                    )
                    conn.commit()
                    # 3. PHÁT TÍN HIỆU SAU KHI LƯU THÀNH CÔNG
                    self.scan_saved.emit() 
            except:
                continue
        conn.close()

    def add_task(self, operation, data):
        self.queue.put((operation, data))

    def stop(self):
        self.running = False
        self.queue.put(None)

class BarcodeReaderApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Ứng dụng Đọc Barcode")
        self.setGeometry(100, 100, 1600, 900)
        self.setMinimumSize(1280, 720)

        try:
            self.setWindowIcon(QPixmap("Logo.ico"))
        except:
            pass

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        
        self.main_layout = QHBoxLayout(self.central_widget)
        self.main_layout.setSpacing(15)
        self.main_layout.setContentsMargins(20, 20, 20, 20)

        left_column_widget = QWidget()
        left_layout = QVBoxLayout(left_column_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(5)

        video_panel = self.create_video_panel()
        result_panel = self.create_result_panel()
        
        left_layout.addWidget(video_panel)
        left_layout.addWidget(result_panel)

        history_panel = self.create_history_and_export_panel()

        self.main_layout.addWidget(left_column_widget)
        self.main_layout.addWidget(history_panel)
        self.main_layout.setStretch(0, 2)
        self.main_layout.setStretch(1, 3)

        self.capture = cv2.VideoCapture(0, cv2.CAP_DSHOW)
        self.capture.set(cv2.CAP_PROP_FRAME_WIDTH, 1280)
        self.capture.set(cv2.CAP_PROP_FRAME_HEIGHT, 720)
        self.capture.set(cv2.CAP_PROP_BUFFERSIZE, 1)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_frame)
        self.timer.start(50)

        self.status_timer = QTimer(self)
        self.status_timer.setSingleShot(True)
        self.status_timer.timeout.connect(self.reset_status)

        self.last_barcode = None
        self.last_beep_time = 0

        self.success_effect = QGraphicsOpacityEffect(self.result_input)
        self.result_input.setGraphicsEffect(self.success_effect)
        self.success_animation = QPropertyAnimation(self.success_effect, b"opacity")

        self.init_database()

        db_path = os.path.join(get_base_path(), "barcodes.db")
        self.db_worker = DatabaseWorker(db_path)
        self.db_worker.start()

        # 4. KẾT NỐI TÍN HIỆU TỪ WORKER VỚI HÀM LÀM MỚI BẢNG
        self.db_worker.scan_saved.connect(self.populate_history_table)

        self.populate_history_table() # Chạy lần đầu để tải dữ liệu cũ
        self.apply_stylesheet()
    
    def create_video_panel(self):
        video_group = QGroupBox("📹 Camera")
        video_layout = QVBoxLayout(video_group)
        self.video_label = QLabel(alignment=Qt.AlignCenter)
        video_group.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.video_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.status_label = QLabel("🟢 Camera đang hoạt động - Đưa barcode vào khung hình", alignment=Qt.AlignCenter)
        video_layout.addWidget(self.video_label)
        video_layout.addWidget(self.status_label)
        return video_group

    def create_result_panel(self):
        result_group = QGroupBox("📋 Kết quả đọc barcode")
        result_group.setFixedHeight(150) 
        result_layout = QVBoxLayout(result_group)
        self.result_input = QLineEdit(readOnly=True, placeholderText="Barcode sẽ hiển thị ở đây...")
        info_label = QLabel("💡 Barcode sẽ tự động được sao chép vào clipboard", alignment=Qt.AlignCenter)
        result_layout.addWidget(self.result_input)
        result_layout.addWidget(info_label)
        return result_group

    def create_history_and_export_panel(self):
        history_group = QGroupBox("📚 Lịch sử quét & Xuất File")
        main_group_layout = QVBoxLayout(history_group)
        controls_layout = QHBoxLayout()
        self.from_date = QDateEdit(QDate.currentDate(), calendarPopup=True, displayFormat="dd/MM/yyyy")
        self.to_date = QDateEdit(QDate.currentDate(), calendarPopup=True, displayFormat="dd/MM/yyyy")
        self.export_btn = QPushButton("📊 Xuất Excel")
        self.export_btn.clicked.connect(self.export_to_excel)
        controls_layout.addWidget(QLabel("📅 Từ ngày:"))
        controls_layout.addWidget(self.from_date)
        controls_layout.addWidget(QLabel("📅 Đến ngày:"))
        controls_layout.addWidget(self.to_date)
        controls_layout.addSpacerItem(QSpacerItem(40, 20, QSizePolicy.Expanding, QSizePolicy.Minimum))
        controls_layout.addWidget(self.export_btn)
        main_group_layout.addLayout(controls_layout)
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(5)
        self.history_table.setHorizontalHeaderLabels(["ID", "Nội dung Barcode", "Ngày quét", "Giờ quét", "Ghi chú"])
        self.history_table.verticalHeader().setVisible(False)
        self.history_table.setAlternatingRowColors(True)
        header = self.history_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.Stretch)
        self.history_table.cellChanged.connect(self.handle_note_change)
        main_group_layout.addWidget(self.history_table)
        return history_group

    def apply_stylesheet(self):
        self.setStyleSheet("""
            QMainWindow { background-color: #f8f9fa; }
            QGroupBox {
                font-weight: 600; font-size: 14px; border: 1px solid #dee2e6;
                border-radius: 8px; margin-top: 1em; padding: 1.5em 1em 1em 1em; background-color: #ffffff;
            }
            QGroupBox::title {
                subcontrol-origin: margin; left: 15px; padding: 2px 8px;
                background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px;
            }
            QTableWidget {
                border: 1px solid #ced4da; border-radius: 5px; gridline-color: #e9ecef;
            }
            QHeaderView::section {
                background-color: #e9ecef; padding: 8px; border: none; font-weight: 600;
            }
            QLineEdit, QDateEdit { padding: 8px 10px; border: 1px solid #ced4da; border-radius: 5px; }
            QLineEdit:focus, QDateEdit:focus { border-color: #80bdff; }
            QPushButton {
                background-color: #007bff; color: white; border: none;
                padding: 8px 16px; border-radius: 5px; font-weight: 600;
            }
            QPushButton:hover { background-color: #0069d9; }
        """)

    def init_database(self):
        db_path = os.path.join(get_base_path(), "barcodes.db")
        self.conn = sqlite3.connect(db_path, check_same_thread=False)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS scans (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                content TEXT NOT NULL,
                scanned_at TEXT NOT NULL,
                scan_date TEXT,
                scan_time TEXT,
                note TEXT
            )
        """)
        cursor = self.conn.cursor()
        cursor.execute("PRAGMA table_info(scans)")
        if "note" not in {row[1] for row in cursor.fetchall()}:
            self.conn.execute("ALTER TABLE scans ADD COLUMN note TEXT")
        self.conn.commit()

    def populate_history_table(self):
        self.history_table.blockSignals(True)
        self.history_table.setRowCount(0)
        cursor = self.conn.cursor()
        cursor.execute("SELECT id, content, scan_date, scan_time, note FROM scans ORDER BY id DESC")
        for row_data in cursor.fetchall():
            row_position = self.history_table.rowCount()
            self.history_table.insertRow(row_position)
            (scan_id, content, scan_date, scan_time, note) = row_data
            for col, item_data in enumerate([str(scan_id), content, scan_date, scan_time]):
                item = QTableWidgetItem(item_data)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                self.history_table.setItem(row_position, col, item)
            self.history_table.setItem(row_position, 4, QTableWidgetItem(note or ""))
        self.history_table.blockSignals(False)

    def handle_note_change(self, row, column):
        if column == 4:
            scan_id = int(self.history_table.item(row, 0).text())
            new_note = self.history_table.item(row, column).text()
            try:
                cursor = self.conn.cursor()
                cursor.execute("UPDATE scans SET note = ? WHERE id = ?", (new_note, scan_id))
                self.conn.commit()
            except sqlite3.Error as e:
                QMessageBox.critical(self, "Lỗi Database", f"Không thể cập nhật ghi chú: {e}")

    def save_scan_async(self, content: str):
        now = datetime.now()
        self.db_worker.add_task(
            "save_scan", (content, now.strftime("%Y-%m-%d %H:%M:%S"), now.strftime("%Y-%m-%d"), now.strftime("%H:%M:%S"))
        )

    def export_to_excel(self):
        start_date = self.from_date.date().toString("yyyy-MM-dd")
        end_date = self.to_date.date().toString("yyyy-MM-dd")
        if self.from_date.date() > self.to_date.date():
            QMessageBox.warning(self, "⚠️ Cảnh báo", "'Từ ngày' phải nhỏ hơn hoặc bằng 'Đến ngày'.")
            return
        cursor = self.conn.cursor()
        cursor.execute(
            "SELECT content, scan_date, scan_time, note FROM scans WHERE scan_date BETWEEN ? AND ? ORDER BY id ASC",
            (start_date, end_date),
        )
        rows = cursor.fetchall()
        if not rows:
            QMessageBox.information(self, "ℹ️ Thông báo", "Không có dữ liệu trong khoảng ngày đã chọn.")
            return
        save_path, _ = QFileDialog.getSaveFileName(self, "💾 Lưu file Excel", f"barcode_scans_{start_date}_to_{end_date}.xlsx", "Excel Files (*.xlsx)")
        if not save_path: return
        wb = Workbook()
        ws = wb.active
        ws.title = "Barcode Scans"
        headers = ["STT", "Nội dung Barcode", "Ngày quét", "Giờ quét", "Ghi chú"]
        ws.append(headers)
        for cell in ws[1]:
            cell.font = Font(bold=True)
        for i, row_data in enumerate(rows, 1):
            (content, scan_date, scan_time, note) = row_data
            display_date = datetime.strptime(scan_date, "%Y-%m-%d").strftime("%d/%m/%Y") if scan_date else ""
            ws.append([i, content, display_date, scan_time, note])
        for col in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 2
        try:
            wb.save(save_path)
            QMessageBox.information(self, "✅ Thành công", f"Đã xuất Excel thành công!\nFile lưu tại: {save_path}")
        except Exception as e:
            QMessageBox.critical(self, "❌ Lỗi", f"Không thể lưu file.\nLỗi: {e}")

    # *** CẬP NHẬT HÀM NÀY ***
    def update_frame(self):
        ret, frame = self.capture.read()
        if not ret:
            self.status_label.setText("❌ Không thể kết nối camera")
            self.status_label.setStyleSheet("color: #dc3545; font-weight: bold;")
            return
        try:
            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            barcodes = pyzbar.decode(gray)
            if barcodes:
                for barcode in barcodes:
                    barcode_data = barcode.data.decode("utf-8")
                    current_time = time.time()
                    if barcode_data != self.last_barcode or (current_time - self.last_beep_time > 1.5):
                        self.last_barcode = barcode_data
                        self.last_beep_time = current_time
                        self.result_input.setText(barcode_data)
                        self.play_beep_sound()
                        pyperclip.copy(barcode_data)
                        
                        # Chỉ cần gửi yêu cầu lưu, không cần làm mới bảng ở đây
                        self.save_scan_async(barcode_data)
                        
                        self.status_label.setText("✅ Đã đọc thành công!")
                        self.status_label.setStyleSheet("color: #28a745; font-weight: bold;")
                        self.status_timer.start(2000)
                        self.animate_success()

                    (x, y, w, h) = barcode.rect
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (40, 167, 69), 3)
            self.display_frame(frame)
        except cv2.error:
            pass

    def display_frame(self, frame):
        try:
            h, w, ch = frame.shape
            frame_rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            qt_image = QImage(frame_rgb.data, w, h, ch * w, QImage.Format_RGB888)
            self.video_label.setPixmap(QPixmap.fromImage(qt_image).scaled(self.video_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation))
        except Exception as e:
            print(f"Lỗi hiển thị frame: {e}")
            pass

    def animate_success(self):
        self.success_animation.setDuration(800)
        self.success_animation.setStartValue(0.3)
        self.success_animation.setEndValue(1.0)
        self.success_animation.setEasingCurve(QEasingCurve.OutCubic)
        self.success_animation.start()
        self.result_input.setStyleSheet("border: 2px solid #28a745; background: #e2f0e6; padding: 10px 12px; border-radius: 5px;")
        QTimer.singleShot(1500, self.reset_result_style)

    def reset_result_style(self):
        self.result_input.setStyleSheet("border: 1px solid #ced4da; padding: 10px 12px; border-radius: 5px;")

    def reset_status(self):
        self.status_label.setText("🟢 Camera đang hoạt động - Đưa barcode vào khung hình")
        self.status_label.setStyleSheet("color: #28a745; font-weight: bold;")

    def play_beep_sound(self):
        try: winsound.Beep(1000, 150)
        except: pass

    def closeEvent(self, event):
        self.timer.stop()
        if self.capture.isOpened(): self.capture.release()
        self.db_worker.stop()
        self.db_worker.wait(2000)
        self.conn.close()
        event.accept()

def main():
    app = QApplication(sys.argv)
    window = BarcodeReaderApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
import cv2
from pyzbar import pyzbar
import winsound
import time
import pyperclip

def main():
    cap = cv2.VideoCapture(0)
    cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
    cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
    
    last_barcode = None
    last_beep_time = 0
    last_read_time = 0  # Thêm biến để theo dõi thời gian đọc cuối cùng

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        current_time = time.time()
        
        # Chỉ đọc barcode nếu đã qua 3 giây kể từ lần đọc cuối
        if current_time - last_read_time >= 3:
            barcodes = pyzbar.decode(frame)
            for barcode in barcodes:
                barcode_data = barcode.data.decode("utf-8")
                
                # Kiểm tra barcode mới hoặc đã qua 1 giây kể từ tiếng bíp cuối
                if barcode_data != last_barcode or (current_time - last_beep_time > 1):
                    print(f"Barcode: {barcode_data}")
                    winsound.Beep(1000, 200)  # Phát tiếng bíp
                    pyperclip.copy(barcode_data)  # Copy vào clipboard
                    last_barcode = barcode_data
                    last_beep_time = current_time
                    last_read_time = current_time  # Cập nhật thời gian đọc cuối cùng

                # Vẽ khung quanh barcode
                (x, y, w, h) = barcode.rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                
                # Thêm text giá trị barcode bên dưới
                text_position = (x, y + h + 30)
                cv2.putText(frame, barcode_data, text_position, 
                          cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2)
        else:
            # Nếu chưa đủ 3 giây, vẫn hiển thị frame nhưng không đọc barcode
            barcodes = pyzbar.decode(frame)
            for barcode in barcodes:
                # Chỉ vẽ khung và text, không xử lý đọc
                (x, y, w, h) = barcode.rect
                cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                barcode_data = barcode.data.decode("utf-8")
                text_position = (x, y + h + 30)
                cv2.putText(frame, barcode_data, text_position, 
                          cv2.FONT_HERSHEY_SIMPLEX, 0.9, (0, 255, 0), 2)

        # Hiển thị frame
        cv2.imshow("Barcode Reader by giangcse", frame)

        # Nhấn Esc hoặc đóng cửa sổ để thoát
        key = cv2.waitKey(1)
        if key == 27:
            break

    cap.release()
    cv2.destroyAllWindows()

if __name__ == "__main__":
    main()
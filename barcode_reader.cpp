#include <opencv2/opencv.hpp>
#include <zbar.h>
#include <windows.h>
#include <chrono>
#include <string>
#include <vector>
#include <iostream>

using namespace cv;
using namespace std;
using namespace zbar;

// Hàm copy text vào clipboard (thay cho pyperclip)
void copyToClipboard(const string& text) {
    if (OpenClipboard(NULL)) {
        EmptyClipboard();
        HGLOBAL hMem = GlobalAlloc(GMEM_MOVEABLE, text.size() + 1);
        if (hMem) {
            char* pMem = (char*)GlobalLock(hMem);
            memcpy(pMem, text.c_str(), text.size() + 1);
            GlobalUnlock(hMem);
            SetClipboardData(CF_TEXT, hMem);
        }
        CloseClipboard();
    }
}

int main() {
    // Khởi tạo webcam
    VideoCapture cap(0);
    if (!cap.isOpened()) {
        cout << "Cannot open webcam" << endl;
        return -1;
    }
    cap.set(CAP_PROP_FRAME_WIDTH, 640);
    cap.set(CAP_PROP_FRAME_HEIGHT, 480);

    // Khởi tạo ZBar scanner
    ImageScanner scanner;
    scanner.set_config(ZBAR_NONE, ZBAR_CFG_ENABLE, 1);

    string last_barcode;
    auto last_beep_time = chrono::steady_clock::now();
    auto last_read_time = chrono::steady_clock::now();

    while (true) {
        Mat frame;
        bool ret = cap.read(frame);
        if (!ret) {
            cout << "Cannot read frame" << endl;
            break;
        }

        // Tính thời gian hiện tại
        auto current_time = chrono::steady_clock::now();
        double time_diff_read = chrono::duration<double>(current_time - last_read_time).count();
        double time_diff_beep = chrono::duration<double>(current_time - last_beep_time).count();

        // Chuyển frame sang grayscale để ZBar đọc
        Mat gray;
        cvtColor(frame, gray, COLOR_BGR2GRAY);

        // Chỉ đọc barcode nếu đã qua 3 giây
        if (time_diff_read >= 3.0) {
            // Chuẩn bị dữ liệu cho ZBar
            Image image(gray.cols, gray.rows, "Y800", gray.data, gray.cols * gray.rows);
            int n = scanner.scan(image);

            // Xử lý barcode nếu tìm thấy
            for (Image::SymbolIterator symbol = image.symbol_begin(); symbol != image.symbol_end(); ++symbol) {
                string barcode_data = symbol->get_data();

                if (barcode_data != last_barcode || time_diff_beep > 1.0) {
                    cout << "Barcode: " << barcode_data << endl;
                    Beep(1000, 200); // Tiếng bíp: 1000Hz, 200ms
                    copyToClipboard(barcode_data);
                    last_barcode = barcode_data;
                    last_beep_time = current_time;
                    last_read_time = current_time;
                }

                // Vẽ khung quanh barcode
                vector<Point> points;
                for (int i = 0; i < symbol->get_location_size(); i++) {
                    points.push_back(Point(symbol->get_location_x(i), symbol->get_location_y(i)));
                }
                Rect rect = boundingRect(points);
                rectangle(frame, rect, Scalar(0, 255, 0), 2);

                // Thêm text bên dưới
                Point text_position(rect.x, rect.y + rect.height + 30);
                putText(frame, barcode_data, text_position, FONT_HERSHEY_SIMPLEX, 0.9, Scalar(0, 255, 0), 2);
            }

            image.set_data(NULL, 0); // Giải phóng dữ liệu ZBar
        }
        else {
            // Nếu chưa đủ 3 giây, vẫn vẽ khung và text nhưng không xử lý
            Image image(gray.cols, gray.rows, "Y800", gray.data, gray.cols * gray.rows);
            scanner.scan(image);

            for (Image::SymbolIterator symbol = image.symbol_begin(); symbol != image.symbol_end(); ++symbol) {
                string barcode_data = symbol->get_data();
                vector<Point> points;
                for (int i = 0; i < symbol->get_location_size(); i++) {
                    points.push_back(Point(symbol->get_location_x(i), symbol->get_location_y(i)));
                }
                Rect rect = boundingRect(points);
                rectangle(frame, rect, Scalar(0, 255, 0), 2);
                Point text_position(rect.x, rect.y + rect.height + 30);
                putText(frame, barcode_data, text_position, FONT_HERSHEY_SIMPLEX, 0.9, Scalar(0, 255, 0), 2);
            }

            image.set_data(NULL, 0);
        }

        // Hiển thị frame
        imshow("Barcode Reader by giangcse", frame);

        // Nhấn Esc để thoát
        int key = waitKey(1);
        if (key == 27) { // Esc
            break;
        }
    }

    cap.release();
    destroyAllWindows();
    return 0;
}
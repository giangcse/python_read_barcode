cmake_minimum_required(VERSION 3.10)
project(BarcodeReader)

find_package(OpenCV REQUIRED)
find_package(ZBar REQUIRED)

add_executable(BarcodeReader main.cpp)
target_link_libraries(BarcodeReader ${OpenCV_LIBS} ${ZBAR_LIBRARIES})
# 📚 RAW CODE trả về translations cho các file PPT/PPTX 

## 🌟 Overview

Đây là phần mô tả về các file trong chương trình, hỗ trợ 2 cách: Legacy (no AI, pure GGTrans) và AI (NLLB hoặc GEMINI).

## 🛠️ Prerequisites & Setup

1.  **Python:** Cài đặt Python (3.8+).

2.  **Môi trường:** Tạo và kích hoạt môi trường ảo.

3.  **Thư viện:** Cài đặt các thư viện cần thiết (`google-genai`, `python-dotenv`, `pywin32`, etc.).

4.  **Tệp `.env`:** Thiết lập khóa API Gemini trong tệp `.env` tại thư mục gốc của dự án:

    ```
    GEMINI_API_KEY="YOUR_API_KEY_HERE"

    ```
    Lưu ý: GEMINI đã mở hỗ trợ free plan cho các model AI từ Gemini 2.5 Pro trở xuống. Chương trình này sẽ sử dụng API Key tự tạo, và phải đăng kí lên Google AI Studio. (also free)

## 1\. ⚙️ Workflow Gián tiếp (Legacy/Template-Based)

Luồng này là quy trình nhiều bước, lý tưởng cho việc kiểm soát chất lượng thủ công hoặc khi dịch thuật bằng AI không khả dụng.

| Bước | Module (Thư mục) | Mô tả |
| ----- | ----- | ----- |
| **1. Trích xuất** | `ConvertPPTToTXT` | Đọc tệp PPT/PPTX gốc và trích xuất tất cả văn bản vào các tệp TXT (dạng `engTXT`). |
| **2. Dịch** | `ConvertEngToVN` | Lấy các tệp `engTXT`, thực hiện dịch sang tiếng Việt để tạo các tệp `VN TXT`. (Bước này ban đầu có thể là thủ công hoặc sử dụng một công cụ dịch thuật đơn giản hơn). |
| **3. Định dạng** | `ConvertTxtToJson` | Chuyển đổi các tệp `VN TXT` đã dịch sang định dạng JSON để dễ dàng tái cấu trúc và chèn vào PPT. |
| **4. Tái cấu trúc** | `ConvertBackToPPTWithExample` | Đọc dữ liệu từ tệp JSON và chèn vào tệp PPT mới, sử dụng một template PowerPoint được định sẵn. |

## 2\. ⚡ Workflow Dịch thuật Trực tiếp (AI-Powered)

Luồng này bỏ qua các bước trung gian (TXT, JSON) và dịch văn bản trực tiếp trong tệp PowerPoint bằng cách sử dụng các mô hình AI tiên tiến, sau đó chèn lại bản dịch vào hình dạng (shape) tương ứng.

Dịch thuật được xử lý trong các module như `.directTrans`, sử dụng hai mô hình khác nhau:

### A. Mô hình NLLB (Meta)

| Đặc điểm | Mô tả |
| ----- | ----- |
| **Perks (Ưu điểm)** | Dịch thuật chất lượng cao, có thể so sánh với các mô hình AI thương mại. **Khả năng chạy Local:** Có thể tải xuống và sử dụng cục bộ mà không cần kết nối internet. |
| **Cons (Nhược điểm)** | **Kích thước Lớn:** Các mô hình như NLLB-200-1.3B có dung lượng rất lớn (khoảng 12GB), gây khó khăn cho việc triển khai và yêu cầu phần cứng mạnh. |

### B. Mô hình Gemini (Google)

| Đặc điểm | Mô tả |
| ----- | ----- |
| **Perks (Ưu điểm)** | **Dễ sử dụng:** Tích hợp API đơn giản, dễ dàng điều chỉnh. **Mô hình Thông minh:** Các mô hình miễn phí (như Gemini 1.5 Flash) cũng rất mạnh mẽ và thông minh. **FREE PLAN:** Cung cấp gói miễn phí với giới hạn lớn (hoặc không giới hạn đối với các mô hình cấp độ Flash/Nano), giúp tiết kiệm chi phí. |
| **Cons (Nhược điểm)** | **Yêu cầu Kết nối:** Hoàn toàn phụ thuộc vào kết nối Internet. Không thể sử dụng AI khi ngoại tuyến. |

## 🚀 Usage

Để dịch một thư mục chứa tệp PPTX, cho các tập lệnh cần dịch theo đường dẫn yêu cầu thư mục `.directTrans` và cung cấp đường dẫn thư mục đầu vào và đầu ra trong khối `if __name__ == "__main__":`.

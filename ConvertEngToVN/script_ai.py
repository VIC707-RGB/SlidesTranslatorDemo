import os
import time
from dotenv import load_dotenv
import google.generativeai as genai

# --- CẤU HÌNH GEMINI CLIENT VÀ CACHE ---
# Tải biến môi trường (ví dụ: GEMINI_API_KEY từ tệp .env)
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
# print(f"Key loaded: {bool(os.getenv('GEMINI_API_KEY'))}")

try:
    # Khởi tạo Gemini Client
    genai.configure(api_key=GEMINI_API_KEY)
    CLIENT = genai.Client()
    MODEL_NAME = "gemini-1.5-flash"
except Exception as e:
    print(f"❌ Lỗi khi khởi tạo Gemini Client: {e}")
    CLIENT = None # Đặt CLIENT thành None nếu thất bại

# --- HÀM DỊCH BẰNG GEMINI ---
def ai_translate_text(text, target_lang="vi", retries=3):
    """
    Dịch một đoạn văn bản bằng Gemini 1.5 Flash.
    """
    if CLIENT is None:
        print("❌ Gemini Client chưa được khởi tạo. Bỏ qua dịch.")
        return text
    
    # Bỏ qua nếu văn bản rỗng
    if not text.strip():
        return text

    for attempt in range(retries):
        try:
            prompt = (
                f"Translate the following text to {target_lang} "
                f"without adding extra text, explanations, or prefixes (like 'Title:' or '- '): \n\n{text}"
            )

            response = CLIENT.models.generate_content(
                model=MODEL_NAME,
                contents=prompt,
                generation_config={"temperature": 0}
            )
            
            translated = response.text.strip()
            
            # Đảm bảo kết quả không rỗng
            if translated:
                return translated
            
        except Exception as e:
            print(f"⚠️ Dịch thất bại (thử {attempt+1}/{retries}) cho '{text[:50]}...': {e}")
            time.sleep(1) # Chờ một chút trước khi thử lại

    print(f"❌ Thất bại hoàn toàn khi dịch, trả về bản gốc: {text}")
    return text

# --- LOGIC DỊCH FILE ---
def translate_file(input_path, output_path, target_lang="vi"):
    """
    Dịch một tệp văn bản dòng theo dòng, giữ lại cấu trúc slide.
    """
    print(f"📄 Đang dịch: {os.path.basename(input_path)}")
    
    with open(input_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    translated_lines = []
    
    for line in lines:
        line = line.rstrip("\n")  # Giữ lại khoảng trắng đầu dòng nhưng loại bỏ xuống dòng
        original_text = line.strip()

        # Giữ nguyên các dòng cấu trúc và dòng trống
        if line.startswith("Slide ") or line.startswith("Contents:") or not original_text:
            translated_lines.append(line)
            continue
        
        # Xử lý các dòng Title
        if line.startswith("Title:"):
            # Lấy văn bản tiêu đề (loại bỏ "Title:")
            title_text = line[len("Title:"):].strip()
            # Dịch
            translated_title = ai_translate_text(title_text, target_lang)
            # Thêm lại prefix "Title: "
            translated_lines.append(f"Title: {translated_title}")
            
        # Xử lý các dòng nội dung có dấu gạch ngang
        elif line.startswith("- "):
            # Lấy nội dung (loại bỏ "- ")
            content_text = line[2:].strip()
            # Dịch
            translated_content = ai_translate_text(content_text, target_lang)
            # Thêm lại prefix "- "
            translated_lines.append(f"- {translated_content}")

        # Xử lý các dòng nội dung khác
        else:
            # Dịch dòng
            translated = ai_translate_text(line.strip(), target_lang)
            translated_lines.append(translated)

    # Lưu file đã dịch
    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        for tline in translated_lines:
            f.write(tline + "\n")

# --- LOGIC DỊCH HÀNG LOẠT ---
def mass_translate(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    txt_files = [f for f in os.listdir(input_folder) if f.lower().endswith(".txt")]
    if not txt_files:
        print("⚠️ Không tìm thấy tệp .txt nào trong thư mục đầu vào.")
        return

    for file in txt_files:
        input_path = os.path.join(input_folder, file)
        output_path = os.path.join(output_folder, file)
        
        translate_file(input_path, output_path)
        print(f"✅ Đã lưu → {output_path}\n")

# --- CHẠY CHÍNH ---
if __name__ == "__main__":
    # Thay đổi các đường dẫn này cho phù hợp với môi trường của bạn
    INPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertPPTXToTXT\AD-txt"
    OUTPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertEngToVN\AD-ppt-vn"

    mass_translate(INPUT_FOLDER, OUTPUT_FOLDER)
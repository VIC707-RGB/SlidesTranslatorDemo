import os
import time
import json
import re
from dotenv import load_dotenv
import google.generativeai as genai
# Lưu ý: Yêu cầu cài đặt thư viện win32com.client (pywin32) trên Windows
try:
    import win32com.client
except ImportError:
    print("Warning: win32com.client not found. This script requires Windows and pywin32.")
    win32com = None

# ---------- CONFIGURATION AND AI SETUP ----------
load_dotenv()
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY") 

# Hướng dẫn chi tiết cho mô hình để giữ nguyên định dạng và ID
SYSTEM_INSTRUCTION = (
    "You are a professional technical document translator. "
    "Your task is to translate the entire list of text chunks from English to Vietnamese. "
    "Each chunk is prefixed with a unique identifier like '[TXT_001]', '[TXT_002]', etc. "
    "You MUST preserve these identifiers (e.g., '[TXT_001]') exactly as they are in the output. "
    "You MUST NOT translate the structural keywords 'Slide', 'Contents', and 'Title' if they appear in the text. "
    "The output MUST only contain the translated text and the preserved identifiers, without extra explanations or remarks."
)

# Khởi tạo model/client (Sử dụng GenerativeModel như bạn đã xác nhận)
client = None
MODEL_NAME = "gemini-2.5-pro"

try:
    if not GEMINI_API_KEY:
        raise ValueError("GEMINI_API_KEY not found in environment. Please check your .env file.")
        
    genai.configure(api_key=GEMINI_API_KEY)
    model = genai.GenerativeModel(MODEL_NAME)
    client = model 
    print("✅ Cấu hình Gemini bằng GenerativeModel thành công.")
except Exception as e:
    print(f"❌ Lỗi cấu hình Gemini: {e}. Vui lòng kiểm tra API Key và thư viện.")
    client = None

# ---------- CACHE SYSTEM (Dùng cho nội dung đã gán ID) ----------
CACHE_FILE = "ppt_translation_cache.json"
if os.path.exists(CACHE_FILE):
    try:
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            CACHE = json.load(f)
    except json.JSONDecodeError:
        print("⚠️ Cache file bị hỏng, tạo cache mới.")
        CACHE = {}
else:
    CACHE = {}

def save_cache():
    """Lưu toàn bộ cache"""
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(CACHE, f, ensure_ascii=False, indent=2)

# ---------- TRANSLATION CORE LOGIC ----------

def translate_chunks_with_gemini(raw_text_with_ids, target_lang="vi", retries=3):
    """
    Gửi tất cả các đoạn văn bản (đã gán ID) đến Gemini để dịch.
    """
    if client is None:
        return {} # Trả về từ điển rỗng

    content_key = raw_text_with_ids.strip()
    if content_key in CACHE:
        # Tải bản dịch thô từ cache
        translated_raw = CACHE[content_key]
    else:
        # Nếu chưa có trong cache, gọi API
        for attempt in range(retries):
            try:
                # Gửi System Instruction và nội dung
                response = client.generate_content(
                    contents=[SYSTEM_INSTRUCTION, raw_text_with_ids],
                    generation_config={"temperature": 0}
                )
                translated_raw = response.text
                
                if not translated_raw:
                    raise Exception("API returned empty response.")
                
                CACHE[content_key] = translated_raw
                save_cache()
                break

            except Exception as e:
                print(f"⚠️ Dịch batch thất bại (thử {attempt+1}/{retries}).")
                print(f"Lỗi: {e}")
                time.sleep(2 * (attempt + 1))
        else:
            print("❌ Dịch batch thất bại sau nhiều lần thử, không thể dịch tệp.")
            return {}
            
    # Phân tích cú pháp phản hồi và trả về từ điển {ID: Bản dịch}
    # Sử dụng regex để tìm tất cả các ID [TXT_XXX] và văn bản theo sau
    
    # Thêm ID giả ở cuối để đảm bảo đoạn cuối cùng được capture
    translated_raw += "\n[TXT_END]" 
    
    # Regex tìm: [TXT_XXX] + (bất kỳ nội dung nào, kể cả xuống dòng) + (cho đến ID tiếp theo hoặc END)
    pattern = re.compile(r'(\[TXT_\d+\])(.*?)(?=\[TXT_\d+\]|\[TXT_END\])', re.DOTALL)
    
    translated_chunks = {}
    for match in pattern.finditer(translated_raw):
        # match.group(1) là ID (e.g., [TXT_001])
        # match.group(2) là nội dung dịch
        chunk_id = match.group(1).strip()
        text = match.group(2).strip()
        translated_chunks[chunk_id] = text

    return translated_chunks


def translate_ppt_text(input_ppt, output_ppt, target_lang="vi"):
    """
    Trích xuất, dịch batch và đưa văn bản dịch vào lại PPT.
    """
    if win32com is None:
        print("Lỗi: win32com.client không khả dụng. Không thể tự động hóa PowerPoint.")
        return
    if client is None:
        print("Lỗi: Gemini Model chưa được cấu hình. Dừng dịch.")
        return

    # 1. Trích xuất văn bản và tạo danh sách mapping
    text_chunks_to_translate = []
    shape_map = [] # Lưu trữ tham chiếu đến Shape để đưa văn bản dịch vào lại
    text_id_counter = 1

    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True 
    
    try:
        presentation = ppt_app.Presentations.Open(input_ppt, WithWindow=True)

        for slide in presentation.Slides:
            for shape in slide.Shapes:
                if shape.HasTextFrame and shape.TextFrame.HasText:
                    original_text = shape.TextFrame.TextRange.Text
                    
                    # Trích xuất văn bản thô, bao gồm khoảng trắng và ngắt dòng
                    if original_text and original_text.strip():
                        unique_id = f"[TXT_{text_id_counter:03d}]"
                        
                        text_chunks_to_translate.append(f"{unique_id} {original_text}")
                        
                        shape_map.append({
                            "id": unique_id,
                            "shape": shape,
                            "original_text": original_text # Giữ nguyên text để tham chiếu
                        })
                        text_id_counter += 1

        if not text_chunks_to_translate:
            print("Không tìm thấy văn bản nào để dịch trong tệp.")
            presentation.Close()
            ppt_app.Quit()
            return
            
        # 2. Dịch batch bằng Gemini
        raw_text_with_ids = "\n\n".join(text_chunks_to_translate)
        print(f"   -> Gửi {len(shape_map)} đoạn văn bản đến Gemini...")
        
        translated_chunks = translate_chunks_with_gemini(raw_text_with_ids, target_lang)

        # 3. Chèn văn bản đã dịch vào lại PPT
        if translated_chunks:
            for item in shape_map:
                translated_text = translated_chunks.get(item["id"])
                
                if translated_text:
                    # Kiểm tra xem có chứa ID không (để tránh lỗi parser)
                    # Nếu có, chỉ lấy phần văn bản sau ID
                    if translated_text.startswith(item["id"]):
                         translated_text = translated_text[len(item["id"]):].lstrip()
                         
                    item["shape"].TextFrame.TextRange.Text = translated_text
                else:
                    print(f"Warning: Không tìm thấy bản dịch cho ID {item['id']}. Giữ nguyên văn bản gốc.")


        # 4. Lưu và đóng PPT
        if output_ppt.lower().endswith(".ppt"):
            file_format = 1 
        elif output_ppt.lower().endswith(".pptx"):
            file_format = 24 # Use 24 (pptx file) instead of 12 for modern files
        else:
            raise ValueError("Output file must end with .ppt or .pptx")

        presentation.SaveAs(output_ppt, FileFormat=file_format)
        presentation.Close()
        
    except Exception as e:
        print(f"Lỗi trong quá trình xử lý PowerPoint: {e}")
    finally:
        ppt_app.Quit()
    
    print(f"Translated presentation saved to: {output_ppt}")


def mass_translate_ppt(input_folder, output_folder, target_lang="vi"):
    os.makedirs(output_folder, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if file_name.lower().endswith((".ppt", ".pptx")):
            input_path = os.path.join(input_folder, file_name)
            output_path = os.path.join(output_folder, file_name)
            print(f"Translating {input_path}...")
            translate_ppt_text(input_path, output_path, target_lang)
            print(f"Saved → {output_path}")


if __name__ == "__main__":
    INPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertPPTXToTXT\AD-ppt"
    OUTPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\TranslatedSlides"
    
    if client is not None:
        mass_translate_ppt(INPUT_FOLDER, OUTPUT_FOLDER, target_lang="vi")
    else:
        print("❌ Dừng lại: Không thể chạy dịch thuật do Gemini Model chưa được cấu hình.")
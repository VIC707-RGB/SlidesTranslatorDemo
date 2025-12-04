import os
import json
import win32com.client
from transformers import AutoTokenizer, AutoModelForSeq2SeqLM

# ---------------- CACHE SYSTEM ----------------
CACHE_FILE = "translation_cache.json"
if os.path.exists(CACHE_FILE):
    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        CACHE = json.load(f)
else:
    CACHE = {}

def save_cache():
    """Save cached translations to JSON."""
    with open(CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(CACHE, f, ensure_ascii=False, indent=2)

# ---------------- NLLB-200 MODEL SETUP ----------------
# Model for offline translation
MODEL_NAME = "facebook/nllb-200-1.3B"
SRC_LANG = "eng_Latn"  # English
TGT_LANG = "vie_Latn"  # Vietnamese

# Load tokenizer and model
tokenizer = AutoTokenizer.from_pretrained(MODEL_NAME)
model = AutoModelForSeq2SeqLM.from_pretrained(MODEL_NAME)

# Manual BOS token ID for Vietnamese (for older transformers without lang_code_to_id)
FORCED_BOS_TOKEN_ID = 250004

# ---------------- AI TRANSLATOR ----------------
def translate_text(text: str) -> str:
    """
    Translate English text to Vietnamese using NLLB-200.
    Uses cache to avoid repeated translation.
    """
    key = text.strip()
    if not key:
        return ""  # skip empty lines

    if key in CACHE:
        return CACHE[key]

    # Tokenize input text
    inputs = tokenizer(key, return_tensors="pt")

    # Generate translation, force Vietnamese output
    translated_tokens = model.generate(
        **inputs,
        forced_bos_token_id=FORCED_BOS_TOKEN_ID
    )

    # Decode translation
    translated_text = tokenizer.batch_decode(translated_tokens, skip_special_tokens=True)[0]

    # Save to cache
    CACHE[key] = translated_text
    save_cache()

    return translated_text

# ---------------- PPT TRANSLATION ----------------
def translate_ppt_text(input_ppt: str, output_ppt: str):
    """
    Translate all text in a PPT/PPTX while keeping images, charts, and layouts.
    """
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True
    presentation = ppt_app.Presentations.Open(input_ppt, WithWindow=True)

    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                original_text = shape.TextFrame.TextRange.Text.strip()
                if original_text:
                    try:
                        translated_text = translate_text(original_text)
                        shape.TextFrame.TextRange.Text = translated_text
                    except Exception as e:
                        print(f"⚠️ Warning: failed to translate '{original_text}': {e}")

    # Save translated presentation
    if output_ppt.lower().endswith(".ppt"):
        file_format = 1  # PPT 97-2003
    elif output_ppt.lower().endswith(".pptx"):
        file_format = 12  # PPTX
    else:
        raise ValueError("Output file must end with .ppt or .pptx")

    presentation.SaveAs(output_ppt, FileFormat=file_format)
    presentation.Close()
    ppt_app.Quit()
    print(f"✅ Translated presentation saved to: {output_ppt}")

# ---------------- MASS TRANSLATION ----------------
def mass_translate_ppt(input_folder: str, output_folder: str):
    """
    Translate all PPT/PPTX files in a folder.
    """
    os.makedirs(output_folder, exist_ok=True)

    for file_name in os.listdir(input_folder):
        if file_name.lower().endswith((".ppt", ".pptx")):
            input_path = os.path.join(input_folder, file_name)
            output_path = os.path.join(output_folder, file_name)
            print(f"📄 Translating {input_path}...")
            translate_ppt_text(input_path, output_path)
            print(f"✅ Saved → {output_path}\n")

# ---------------- MAIN ----------------
if __name__ == "__main__":
    INPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertPPTXToTXT\AD-ppt"
    OUTPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\TranslatedSlides"

    mass_translate_ppt(INPUT_FOLDER, OUTPUT_FOLDER)

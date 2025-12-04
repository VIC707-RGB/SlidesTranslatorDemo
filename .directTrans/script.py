import os
import json
import win32com.client
from deep_translator import GoogleTranslator

def translate_ppt_text(input_ppt, output_ppt, target_lang="vi"):
    """
    Translate all text in a PPT/PPTX presentation while keeping images, charts, and layouts intact.
    """
    # Open PowerPoint
    ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    ppt_app.Visible = True  # Must be visible to avoid SaveAs errors

    presentation = ppt_app.Presentations.Open(input_ppt, WithWindow=True)

    # Iterate through all slides
    for slide in presentation.Slides:
        for shape in slide.Shapes:
            if shape.HasTextFrame and shape.TextFrame.HasText:
                original_text = shape.TextFrame.TextRange.Text.strip()
                if original_text:
                    try:
                        translated_text = GoogleTranslator(source='auto', target=target_lang).translate(original_text)
                        shape.TextFrame.TextRange.Text = translated_text
                    except Exception as e:
                        print(f"Warning: failed to translate '{original_text}': {e}")

    # Save as PPT or PPTX
    if output_ppt.lower().endswith(".ppt"):
        file_format = 1  # PPT 97-2003
    elif output_ppt.lower().endswith(".pptx"):
        file_format = 12  # PPTX
    else:
        raise ValueError("Output file must end with .ppt or .pptx")

    presentation.SaveAs(output_ppt, FileFormat=file_format)
    presentation.Close()
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

    mass_translate_ppt(INPUT_FOLDER, OUTPUT_FOLDER, target_lang="vi")

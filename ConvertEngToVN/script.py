import os
from deep_translator import GoogleTranslator

def translate_file(input_path, output_path, target_lang="vi"):
    # Read original text
    with open(input_path, "r", encoding="utf-8") as f:
        lines = f.readlines()

    translated_lines = []
    for line in lines:
        line = line.rstrip("\n")  # keep indentation but remove trailing newline

        # Keep structural slide lines as is
        if line.startswith("Slide ") or line.startswith("Contents:"):
            translated_lines.append(line)

        elif line.startswith("Title:"):
            # Translate title text but keep "Title:" prefix
            title_text = line[len("Title:"):].strip()
            if title_text:
                try:
                    translated_title = GoogleTranslator(source='auto', target=target_lang).translate(title_text)
                    if translated_title is None:
                        translated_title = title_text
                except Exception as e:
                    print(f"Warning: failed to translate title '{title_text}': {e}")
                    translated_title = title_text
            else:
                translated_title = ""
            translated_lines.append(f"Title: {translated_title}")

        elif line.startswith("- "):
            # Translate content after "- "
            content = line[2:].strip()
            if content:
                try:
                    translated = GoogleTranslator(source='auto', target=target_lang).translate(content)
                    if translated is None:
                        translated = content
                except Exception as e:
                    print(f"Warning: failed to translate line '{content}': {e}")
                    translated = content
            else:
                translated = content
            translated_lines.append(f"- {translated}")

        else:
            # Translate any other non-empty line normally
            if line.strip():
                try:
                    translated = GoogleTranslator(source='auto', target=target_lang).translate(line.strip())
                    if translated is None:
                        translated = line.strip()
                except Exception as e:
                    print(f"Warning: failed to translate line '{line.strip()}': {e}")
                    translated = line.strip()
                translated_lines.append(translated)
            else:
                translated_lines.append("")

    # Save translated file
    with open(output_path, "w", encoding="utf-8") as f:
        for tline in translated_lines:
            f.write(tline + "\n")


def mass_translate(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    for file in os.listdir(input_folder):
        if file.lower().endswith(".txt"):
            input_path = os.path.join(input_folder, file)
            output_path = os.path.join(output_folder, file)
            print(f"Translating: {input_path}")
            translate_file(input_path, output_path)
            print(f"Saved → {output_path}")


if __name__ == "__main__":
    INPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertPPTXToTXT\AD-txt"
    OUTPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertEngToVN\AD-ppt-vn"

    mass_translate(INPUT_FOLDER, OUTPUT_FOLDER)

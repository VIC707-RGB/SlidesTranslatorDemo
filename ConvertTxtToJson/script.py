import os
import json

def convert_txt_to_json(input_path):
    slides_dict = {}
    with open(input_path, "r", encoding="utf-8") as f:
        lines = [line.rstrip() for line in f]

    slide_num = 0
    title = ""
    content_lines = []
    reading_content = False

    for line in lines:
        if line.startswith("Slide "):
            # Save previous slide
            if slide_num > 0:
                slides_dict[f"Slide {slide_num}"] = {
                    "Title": title,
                    "Contents": content_lines
                }

            slide_num += 1
            title = ""
            content_lines = []
            reading_content = False

        elif line.startswith("Title:"):
            title = line.replace("Title:", "").strip()

        elif line.startswith("Contents:"):
            reading_content = True

        elif reading_content:
            if line.startswith("- "):
                content_lines.append(line[2:].strip())
            else:
                content_lines.append(line.strip())

    # Save last slide
    if slide_num > 0:
        slides_dict[f"Slide {slide_num}"] = {
            "Title": title,
            "Contents": content_lines
        }

    return slides_dict

def save_json(data, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def mass_convert_txt_to_json(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)
    for file in os.listdir(input_folder):
        if file.lower().endswith(".txt"):
            input_path = os.path.join(input_folder, file)
            print(f"Processing: {input_path}")
            slides_data = convert_txt_to_json(input_path)
            output_name = os.path.splitext(file)[0] + ".json"
            output_path = os.path.join(output_folder, output_name)
            save_json(slides_data, output_path)
            print(f"Saved → {output_path}")

if __name__ == "__main__":
    INPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertEngToVN\AD-ppt-vn"
    OUTPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertTxtToJson\AD-Json"

    mass_convert_txt_to_json(INPUT_FOLDER, OUTPUT_FOLDER)

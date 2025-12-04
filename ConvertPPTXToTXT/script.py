import os
import win32com.client

def extract_lines_from_ppt(input_path):
    powerpoint = win32com.client.Dispatch("PowerPoint.Application")
    # powerpoint.Visible = 0

    input_path = os.path.abspath(input_path)
    presentation = powerpoint.Presentations.Open(input_path, WithWindow=False)

    slides_data = []

    for idx, slide in enumerate(presentation.Slides, start=1):
        title = None
        content = []

        # Collect all lines in order
        all_lines = []
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                text = shape.TextFrame.TextRange.Text.strip()
                if text:
                    for line in text.split("\n"):
                        line = line.strip()
                        if line:
                            all_lines.append(line)

        if all_lines:
            title = all_lines[0]            # First line becomes the title
            content = all_lines[1:]         # Rest go into contents

        slides_data.append({
            "slide_number": idx,
            "title": title if title else "",
            "content": content
        })

    presentation.Close()
    powerpoint.Quit()
    return slides_data

    # lines = []

    # for slide in presentation.Slides:
    #     for shape in slide.Shapes:
    #         if shape.HasTextFrame:
    #             text = shape.TextFrame.TextRange.Text.strip()
    #             if text:
    #                 for line in text.split("\n"):
    #                     line = line.strip()
    #                     if line:
    #                         lines.append(line)

    # presentation.Close()
    # powerpoint.Quit()

    # return lines

def save_lines_to_txt(lines, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        for line in lines:
            f.write(line + "\n")

def save_slide_txt(slides_data, output_path):
    with open(output_path, "w", encoding="utf-8") as f:
        for slide in slides_data:
            f.write(f"Slide {slide['slide_number']}:\n")
            f.write(f"Title: {slide['title']}\n")
            f.write("Contents:\n")
            for line in slide['content']:
                f.write(f"- {line}\n")
            f.write("\n")  # blank line between slides

def mass_convert(input_folder, output_folder):
    os.makedirs(output_folder, exist_ok=True)

    for file in os.listdir(input_folder):
        if file.lower().endswith((".ppt", ".pptx")):
            input_path = os.path.join(input_folder, file)
            print(f"Reading: {input_path}")

            lines = extract_lines_from_ppt(input_path)

            output_name = os.path.splitext(file)[0] + ".txt"
            output_path = os.path.join(output_folder, output_name)

            # save_lines_to_txt(lines, output_path)
            save_slide_txt(lines, output_path)
            print(f"Saved → {output_path}")

if __name__ == "__main__":
    INPUT = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertPPTXToTXT\AD-ppt"
    OUTPUT = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertPPTXToTXT\AD-txt"

    mass_convert(INPUT, OUTPUT)

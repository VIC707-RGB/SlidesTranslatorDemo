import os
import json
import win32com.client

def append_slides_from_json(template_path, json_path, output_path):
    """
    Reads .ppt template, appends slides according to JSON, 
    and outputs new .ppt with only the JSON slides.
    First and last slides use special layouts, the rest use master layout.
    """
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True  # Make visible to avoid SaveAs errors with .ppt

    # Open template
    presentation = ppt.Presentations.Open(template_path, WithWindow=True)

    # Load JSON
    with open(json_path, "r", encoding="utf-8") as f:
        slides_data = json.load(f)

    slide_keys = list(slides_data.keys())
    total_slides = len(slide_keys)

    # Remember template slides
    first_slide_template = presentation.Slides(1)
    master_slide_template = presentation.Slides(2)
    last_slide_template = presentation.Slides(presentation.Slides.Count)

    # Keep track of the insertion index
    current_index = 1  # start after template slides (or 1 if template will be deleted)

    # Duplicate slides in order
    for idx, slide_key in enumerate(slide_keys):
        slide_info = slides_data[slide_key]
        title_text = slide_info.get("Title", "")
        contents = slide_info.get("Contents", [])

        # Choose template
        if idx == 0:
            base_slide = first_slide_template
        elif idx == total_slides - 1:
            base_slide = last_slide_template
        else:
            base_slide = master_slide_template

        # Duplicate template slide at the end
        dup_slide = base_slide.Duplicate()[0]

        # Move duplicated slide to correct position
        dup_slide.MoveTo(current_index)
        current_index += 1  # next slide goes after this one

        # Fill placeholders: first text box = title, second = content
        title_filled = False
        content_filled = False
        for shape in dup_slide.Shapes:
            if shape.HasTextFrame:
                text_range = shape.TextFrame.TextRange
                if not title_filled:
                    text_range.Text = title_text
                    title_filled = True
                    continue
                if not content_filled:
                    text_range.Text = ""  # clear default content
                    for line in contents:
                        text_range.InsertAfter(line + "\r")
                    content_filled = True

    # Save as .ppt
    if output_path.lower().endswith(".ppt"):
        pp_format = 1  # PPT 97-2003
    elif output_path.lower().endswith(".pptx"):
        pp_format = 12  # OpenXML
    else:
        raise ValueError("Output path must end with .ppt or .pptx")

    presentation.SaveAs(output_path, FileFormat=pp_format)
    presentation.Close()
    ppt.Quit()
    print(f"Created PPT → {output_path}")

    # Delete old template slides AFTER remembering them
    for slide in reversed([first_slide_template, master_slide_template, last_slide_template]): slide.Delete()

def append_slides_from_json2(template_path, json_path, output_path):
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    ppt.Visible = True  # visible avoids some SaveAs .ppt issues

    # Open template
    presentation = ppt.Presentations.Open(template_path, WithWindow=True)

    # Load JSON
    with open(json_path, "r", encoding="utf-8") as f:
        slides_data = json.load(f)

    slide_keys = list(slides_data.keys())
    total_slides = len(slide_keys)

    # Remember template slides
    first_slide_template = presentation.Slides(1)
    master_slide_template = presentation.Slides(2)
    last_slide_template = presentation.Slides(presentation.Slides.Count)

    current_index = 1

    # Duplicate slides in order
    for idx, slide_key in enumerate(slide_keys):
        slide_info = slides_data[slide_key]
        title_text = slide_info.get("Title", "")
        contents = slide_info.get("Contents", [])

        # Select base template
        if idx == 0:
            base_slide = first_slide_template
        elif idx == total_slides - 1:
            base_slide = last_slide_template
        else:
            base_slide = master_slide_template

        # Duplicate and move to current index
        dup_slide = base_slide.Duplicate()[0]
        dup_slide.MoveTo(current_index)
        current_index += 1

        # Fill title and content
        title_filled = False
        content_filled = False
        for shape in dup_slide.Shapes:
            if shape.HasTextFrame:
                text_range = shape.TextFrame.TextRange
                if not title_filled:
                    text_range.Text = title_text
                    title_filled = True
                    continue
                if not content_filled:
                    text_range.Text = ""
                    for line in contents:
                        text_range.InsertAfter(line + "\r")
                    content_filled = True

    # Delete original template slides in reverse order
    for i in sorted([first_slide_template.SlideIndex, 
                     master_slide_template.SlideIndex, 
                     last_slide_template.SlideIndex], reverse=True):
        if i <= presentation.Slides.Count:
            presentation.Slides(i).Delete()

    # Save presentation
    if output_path.lower().endswith(".ppt"):
        pp_format = 1  # PPT 97-2003
    else:
        pp_format = 12  # PPTX

    presentation.SaveAs(output_path, FileFormat=pp_format)
    presentation.Close()
    ppt.Quit()
    print(f"Created PPT → {output_path}")


if __name__ == "__main__":
    TEMPLATE_PPT = r"C:\Users\caoli\PycharmProjects\SlideConverter\Template\base_template.ppt"
    JSON_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertTxtToJson\AD-Json"
    OUTPUT_FOLDER = r"C:\Users\caoli\PycharmProjects\SlideConverter\ConvertBackToPPTWithExample\JSON_to_PPT"

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)

    for json_file in os.listdir(JSON_FOLDER):
        if json_file.lower().endswith(".json"):
            json_path = os.path.join(JSON_FOLDER, json_file)
            output_name = os.path.splitext(json_file)[0] + "_generated.ppt"
            output_path = os.path.join(OUTPUT_FOLDER, output_name)

            append_slides_from_json2(TEMPLATE_PPT, json_path, output_path)

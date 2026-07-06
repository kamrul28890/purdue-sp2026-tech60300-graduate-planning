from pdf2image import convert_from_path
import os

base_path = r"d:\Purdue\Courses\03. Spring 60300 - Graduate Planning\Other reference gantt chart"
output_folder = r"d:\Purdue\Courses\03. Spring 60300 - Graduate Planning\gantt_images"

os.makedirs(output_folder, exist_ok=True)

pdf_files = [
    "Tara Middlebrooks_Gantt_PhD Timeline.pdf",
    "Trisolicha_GanttChart_TECH603_REV1.pdf",
    "_Gantt Chart.pdf"
]

for pdf_file in pdf_files:
    pdf_path = os.path.join(base_path, pdf_file)
    if os.path.exists(pdf_path):
        print(f"Converting: {pdf_file}")
        try:
            name_base = pdf_file.replace('.pdf', '')
            images = convert_from_path(pdf_path, dpi=150)
            for i, image in enumerate(images, 1):
                output_path = os.path.join(output_folder, f"{name_base}_page{i}.png")
                image.save(output_path)
                print(f"  Saved: {output_path}")
        except Exception as e:
            print(f"  Error: {e}")

print(f"\nConversion complete. Images saved to: {output_folder}")

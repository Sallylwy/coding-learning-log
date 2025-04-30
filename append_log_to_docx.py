from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_SECTION
import sys
import os
import re
from datetime import datetime
import pyperclip  # Add this at the top with other imports

def load_log_from_txt(file_path):
    """Open the txt file and return its contents as plain text."""
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()

def format_date(date_str):
    for fmt in ("%b %d", "%B %d"):
        try:
            dt = datetime.strptime(date_str.strip(), fmt)
            return dt.strftime("%B %d, 2025")
        except:
            continue
    return date_str

def write_log_to_docx(log_text, doc, is_replacement=False):
    # 🧾 Start a new section (page break)
    doc.add_page_break()

    reflection_mode = False

    lines = log_text.splitlines()
    for i, line in enumerate(lines):
        stripped = line.strip()

        # 📌 Add title as large heading
        if stripped.startswith("Title:"):
            title_line = stripped.replace("Title:", "").strip()

            # Extract and format date from title
            date_match = re.search(r"\((\w+ \d{1,2})\)", title_line)
            if date_match:
                formatted_date = format_date(date_match.group(1))
                title_line = re.sub(r"\(\w+ \d{1,2}\)", f"({formatted_date})", title_line)

            doc.add_heading(title_line, level=2)

        # ✍️ Bold paragraph for Project Context
        elif stripped.startswith("Project Context:"):
            context = stripped.replace("Project Context:", "").strip()
            p = doc.add_paragraph()
            p.add_run("Project Context: ").bold = True
            p.add_run(context)

        # 🖍️ Section headers like "Key Accomplishments"
        elif stripped in ["Key Accomplishments:", "Next Steps (Optional):"]:
            doc.add_paragraph("")  # spacer line
            doc.add_paragraph(stripped, style='Heading 3')

        elif stripped == "Reflection:":
            doc.add_paragraph("")  # spacer line
            doc.add_paragraph("Reflection:", style='Heading 3')
            reflection_mode = True

        # 🔘 Bullet points or reflection bullets
        elif stripped.startswith("•") or stripped.startswith("-") or reflection_mode:
            bullet = stripped[1:].strip() if stripped.startswith(("•", "-")) else stripped
            para = doc.add_paragraph("• " + bullet)
            para.paragraph_format.left_indent = Pt(18)

        # 🧠 Numbered points: bold the beginning (e.g. '1. Installed ...')
        elif re.match(r"^\d+\.", stripped):
            p = doc.add_paragraph()
            match = re.match(r"^(\d+\.\s*)(.*)", stripped)
            if match:
                p.add_run(match.group(1)).bold = True
                p.add_run(match.group(2))

        # 🔸 Any other line — just a plain paragraph
        elif stripped:
            doc.add_paragraph(stripped)

    # Add a blank line for spacing
    doc.add_paragraph("")

    # Add the reference line, italicized and with the check emoji
    p = doc.add_paragraph()
    run = p.add_run("✅ Saved as a working reference and milestone doc.")
    run.italic = True

    print("✅ Log successfully appended to document.")

def replace_last_log(doc, new_log_text):
    paragraphs = doc.paragraphs
    delete_from_index = None

    # ✅ Find the last heading level 2 — that's our last Title line
    for i in range(len(paragraphs) - 1, -1, -1):
        if paragraphs[i].style.name == 'Heading 2':
            delete_from_index = i
            break

    if delete_from_index is None:
        print("⚠️ Couldn't find a previous log heading to replace.")
        return

    # 🧼 Remove any extra blank paragraph before the last log (e.g. page break)
    if delete_from_index > 0 and paragraphs[delete_from_index - 1].text == "":
        previous_para = paragraphs[delete_from_index - 1]
        previous_para._element.getparent().remove(previous_para._element)
        delete_from_index -= 1

    # 🧽 Remove all paragraphs from that heading to the end
    for _ in range(len(paragraphs) - delete_from_index):
        p = doc.paragraphs[-1]
        p._element.getparent().remove(p._element)

    print("🧽 Old log cleared. Inserting new log...")
    return write_log_to_docx(new_log_text, doc, is_replacement=True)


import pyperclip  # this should be at the top

def main():
    args = sys.argv

    if len(args) < 2:
        print("Usage:")
        print("  python3 append_log_to_docx.py <log_file.txt>")
        print("  python3 append_log_to_docx.py --from-clipboard [--replace-last]")
        return

    from_clipboard = "--from-clipboard" in args
    replace_mode = "--replace-last" in args
    docx_file = "Programming Learning Log.docx"

    if from_clipboard:
        log_text = pyperclip.paste()
        print("📋 Pulled log text from clipboard.")
    else:
        log_file = args[1]
        if not os.path.exists(log_file):
            print("❌ Log file not found.")
            return
        log_text = load_log_from_txt(log_file)

    if not log_text.strip():
        print("⚠️ Log is empty.")
        return

    doc = Document(docx_file)

    if replace_mode:
        replace_last_log(doc, log_text)
    else:
        write_log_to_docx(log_text, doc)

    doc.save(docx_file)
    print(f"✅ Log {'replaced' if replace_mode else 'appended'} successfully to {docx_file}")

if __name__ == "__main__":
    main()

# Functions to replace text in a docx file
from docx import Document

def replace_text_in_paragraph(paragraph, start_pos, end_pos, replace_text):
    full_text = paragraph.text
    n = len(full_text)
    start_pos = max(0, min(start_pos, n))
    end_pos   = max(start_pos, min(end_pos, n)) 

    if start_pos == 0 and end_pos >= n:
        for r in paragraph.runs:
            r.text = ''
        paragraph.add_run(replace_text)
        return True

    text_len = 0
    start_run = end_run = None
    start_off = end_off = 0
    for run in paragraph.runs:
        run_len = len(run.text)
        run_start = text_len
        run_end = text_len + run_len

        if start_run is None and start_pos <= run_end:
            start_run = run
            start_off = max(start_pos - run_start, 0)
        if end_run is None and end_pos <= run_end:
            end_run = run
            end_off = min(end_pos - run_start, run_len)

        text_len += run_len

    if start_run is None or end_run is None:
        prefix = full_text[:start_pos]
        suffix = full_text[end_pos:]
        for r in paragraph.runs:
            r.text = ''
        paragraph.add_run(prefix + replace_text + suffix)
        return True

    if start_run == end_run:
        start_run.text = start_run.text[:start_off] + replace_text + start_run.text[end_off:]
        return True

    start_run.text = start_run.text[:start_off] + replace_text
    clear = False
    for run in paragraph.runs:
        if run is start_run:
            clear = True
            continue
        if run is end_run:
            run.text = run.text[end_off:]
            break
        if clear:
            run.text = ''
    return True

import re
import win32com.client as win32

# Define replacements and exceptions
UK_SPELLINGS = {
    r"\borganize\b": "organise",
    r"\beg\b": "for example"
}

INITIAL_FIX = re.compile(r"(?<=\b)([A-Z]) (?=[A-Z][a-z]+\b)")  # e.g., D Roosevelt -> D. Roosevelt


class NameTracker:
    def __init__(self):
        self.mentioned = {}
        self.ambiguous_last_names = set()

    def check_ambiguity(self):
        last_name_map = {}
        for full_name in self.mentioned:
            parts = full_name.split()
            if len(parts) >= 2:
                last = parts[-1]
                last_name_map.setdefault(last, []).append(full_name)
        for last, names in last_name_map.items():
            if len(names) > 1:
                self.ambiguous_last_names.add(last)

    def replace_names(self, text):
        self.check_ambiguity()
        for full_name in sorted(self.mentioned.keys(), key=len, reverse=True):
            title, *names = full_name.split()
            last = names[-1]
            pattern = re.escape(full_name)
            if self.mentioned[full_name] == 0:
                self.mentioned[full_name] += 1
                continue  # keep full name on first use
            if last in self.ambiguous_last_names:
                continue  # keep full name if ambiguous
            shortened = f"{title} {last}"
            text = re.sub(rf"\b{pattern}\b", shortened, text)
        return text


def is_within_quotes(pos, text):
    quotes = [m.start() for m in re.finditer(r'"', text)]
    return any(start < pos < end for start, end in zip(quotes[::2], quotes[1::2]))


def correct_text(text):
    # Fix initials like D Roosevelt -> D. Roosevelt
    text = INITIAL_FIX.sub(r"\1. ", text)

    # Track names
    name_tracker = NameTracker()
    name_pattern = re.compile(r"\b(Dr|Mr|Ms|Mrs) [A-Z][a-z]+(?: [A-Z][a-z]+)+")
    for match in name_pattern.finditer(text):
        full_name = match.group()
        name_tracker.mentioned.setdefault(full_name, 0)

    # Task 1: UK spellings (outside quotes)
    for pattern, replacement in UK_SPELLINGS.items():
        for match in re.finditer(pattern, text):
            start = match.start()
            if not is_within_quotes(start, text):
                text = text[:start] + re.sub(pattern, replacement, match.group(), count=1) + text[match.end():]
                break  # redo search after each substitution

    # Task 2: Replace names appropriately
    text = name_tracker.replace_names(text)
    return text


def process_docx(input_path, output_path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    doc = word.Documents.Open(input_path)

    full_text = ""
    for para in doc.Paragraphs:
        full_text += para.Range.Text

    corrected = correct_text(full_text.strip())

    # Create a new document
    new_doc = word.Documents.Add()
    new_doc.Content.Text = "Corrected:\n" + corrected
    new_doc.SaveAs(output_path)
    new_doc.Close()
    doc.Close()
    word.Quit()


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python text_corrector.py <input.docx> <output.docx>")
    else:
        process_docx(sys.argv[1], sys.argv[2])

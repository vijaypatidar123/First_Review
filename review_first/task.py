import os
import re
import win32com.client as win32
from win32com.client import constants

def apply_uk_spelling(text):
    # Define US to UK spelling replacements
    replacements = {
        'organize': 'organise',
        'organizing': 'organising',
        'organized': 'organised',
        'organizes': 'organises',
    }

    abbreviation_replacements = {
        r'\beg\b': 'for example',
    }
    quoted_texts = re.findall(r'"([^"]*)"', text)
    for i, quoted in enumerate(quoted_texts):
        text = text.replace(f'"{quoted}"', f'QUOTED_TEXT_{i}')
 
    for pattern, replacement in abbreviation_replacements.items():
        text = re.sub(pattern, replacement, text)
 
    for us_word, uk_word in replacements.items():
        # Skip "organize" in "World Health Organization"
        text = re.sub(r'\b' + us_word + r'\b(?!.{0,20}Organization)', uk_word, text, flags=re.IGNORECASE)
    
    # Restore quoted texts
    for i, quoted in enumerate(quoted_texts):
        text = text.replace(f'QUOTED_TEXT_{i}', f'"{quoted}"')
    
    return text

def add_periods_to_initials(text):
    def repl(match):
        parts = match.group().split()
        return f"{parts[0]} {parts[1][0]}." + (f" {parts[2]}" if len(parts) > 2 else "")

    return re.sub(r'\b(Franklin)\s([A-Z])\b\s(Roosevelt)', repl, text)

def format_names_first_mention(text):
    titled_name_pattern = r'\b(Dr|Mr|Mrs|Ms|Prof)\s+([A-Z][a-z]+)\s+([A-Z][a-z]+)\b'
    matches = list(re.finditer(titled_name_pattern, text))

    name_mentions = {}
    last_name_counts = {}

    for m in matches:
        full_name = m.group(0)
        title = m.group(1)
        last_name = m.group(3)
        name_mentions.setdefault(full_name, []).append(m.start())
        last_name_counts[last_name] = last_name_counts.get(last_name, 0) + 1

    for full_name, positions in name_mentions.items():
        title = full_name.split()[0]
        last_name = full_name.split()[-1]
        if last_name_counts[last_name] == 1:
            for i, pos in enumerate(positions):
                if i > 0:   
                    text = text[:pos] + text[pos:].replace(full_name, f"{title} {last_name}", 1)

    return text

def process_document(input_path, output_path):
    # Initialize Word application
    word = win32.gencache.EnsureDispatch('Word.Application')
    word.Visible = False
    
    try:
        doc = word.Documents.Open(input_path)
    
        content = ""
        for para in doc.Paragraphs:
            content += para.Range.Text
      
        doc.Close(False)
 
        content = apply_uk_spelling(content)
        content = add_periods_to_initials(content)
        content = format_names_first_mention(content)
        
        output_doc = word.Documents.Add()
        
        output_doc.Content.Text = "Corrected:\n" + content
  
        output_doc.SaveAs(output_path)
        output_doc.Close()
        
        print(f"Successfully processed document. Output saved to {output_path}")
    
    except Exception as e:
        print(f"An error occurred: {e}")
    
    finally:
        word.Quit()

if __name__ == "__main__":
    
    current_dir = os.getcwd()

    input_path = os.path.join(current_dir, "input.docx")
    output_path = os.path.join(current_dir, "output.docx")
    
    input_path = os.path.abspath(input_path)
    output_path = os.path.abspath(output_path)
    
    print("Starting text formatting process...")
    process_document(input_path, output_path)
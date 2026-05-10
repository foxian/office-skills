import docx
import sys

def extract_rich_markdown(filepath, overview=False):
    doc = docx.Document(filepath)
    lines = []
    lines.append(f"<!-- DOCUMENT: {filepath} -->\n")
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        style_name = para.style.name
        
        if overview and not style_name.startswith('Heading'):
            continue
            
        lines.append(f"### [p{i}] {text} {{{style_name}}}")
        
    return "\n\n".join(lines)

if __name__ == "__main__":
    filepath = sys.argv[1]
    overview = "--overview" in sys.argv
    print(extract_rich_markdown(filepath, overview))

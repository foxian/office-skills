import docx
import json
import sys

def apply_operations(filepath, ops, outpath=None):
    doc = docx.Document(filepath)
    if outpath is None:
        outpath = filepath
        
    for op in ops:
        if op['op'] == 'rewrite_paragraph':
            idx = int(op['target'].replace('p', ''))
            if idx < len(doc.paragraphs):
                # Simple replacement (clears inline styles for MVP)
                doc.paragraphs[idx].text = op['content']
        elif op['op'] == 'replace_text':
            find_txt = op['find']
            repl_txt = op['replace']
            for p in doc.paragraphs:
                if find_txt in p.text:
                    p.text = p.text.replace(find_txt, repl_txt)
                    
    doc.save(outpath)

if __name__ == "__main__":
    filepath = sys.argv[1]
    with open(sys.argv[2], 'r', encoding='utf-8') as f:
        ops = json.load(f)
    outpath = sys.argv[3] if len(sys.argv) > 3 else filepath
    apply_operations(filepath, ops, outpath)

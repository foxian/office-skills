import docx
import sys
import difflib

def generate_diff(path1, path2):
    d1 = docx.Document(path1)
    d2 = docx.Document(path2)
    
    texts1 = [p.text for p in d1.paragraphs if p.text.strip()]
    texts2 = [p.text for p in d2.paragraphs if p.text.strip()]
    
    diff = difflib.unified_diff(texts1, texts2, lineterm='', n=0)
    return '\n'.join(diff)

if __name__ == "__main__":
    print(generate_diff(sys.argv[1], sys.argv[2]))

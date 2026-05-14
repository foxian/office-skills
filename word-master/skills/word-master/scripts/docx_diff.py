import docx
import re
import sys
import difflib
from dataclasses import dataclass, field
from typing import List, Optional

@dataclass
class ParagraphBlock:
    index: int
    text: str
    style: str
    overrides: List[str] = field(default_factory=list)
    raw: str = ""  # Original markdown line

def parse_markdown_blocks(md: str) -> List[ParagraphBlock]:
    """
    Parse rich markdown into structured ParagraphBlock list.
    Extracts: paragraph index [pN], text, style {StyleName}, overrides.
    """
    blocks = []
    # Match: ### [p0] Text {StyleName} or ### [p0] Text {StyleName | override: ...}
    pattern = r'### \[p(\d+)\] (.+?) \{(.+?)\}'
    for line in md.split('\n'):
        match = re.search(pattern, line)
        if match:
            idx = int(match.group(1))
            text = match.group(2).strip()
            style_part = match.group(3)
            # Split style and overrides
            if '|' in style_part:
                style, override_part = style_part.split('|', 1)
                overrides = [o.strip() for o in override_part.replace('override:', '').split(';')]
                overrides = [o for o in overrides if o]
            else:
                style = style_part
                overrides = []
            blocks.append(ParagraphBlock(
                index=idx,
                text=text,
                style=style.strip(),
                overrides=overrides,
                raw=line
            ))
    return blocks

def generate_diff(path1, path2):
    d1 = docx.Document(path1)
    d2 = docx.Document(path2)
    
    texts1 = [p.text for p in d1.paragraphs if p.text.strip()]
    texts2 = [p.text for p in d2.paragraphs if p.text.strip()]
    
    diff = difflib.unified_diff(texts1, texts2, lineterm='', n=0)
    return '\n'.join(diff)

if __name__ == "__main__":
    print(generate_diff(sys.argv[1], sys.argv[2]))


"""
格式转换：Markdown -&gt; DOCX / HTML / TXT / PDF
"""
import argparse
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from _md_utils import read_utf8


def to_html(content):
    import markdown as md_lib
    css = """&lt;style&gt;
body { font-family: 'Segoe UI', Arial, sans-serif; max-width: 860px; margin: 40px auto; padding: 0 20px; line-height: 1.7; }
h1,h2,h3,h4,h5,h6 { border-bottom: 1px solid #eee; padding-bottom: 4px; }
code { background: #f4f4f4; padding: 2px 6px; border-radius: 3px; font-family: monospace; }
pre { background: #f4f4f4; padding: 16px; border-radius: 4px; overflow-x: auto; }
blockquote { border-left: 4px solid #ccc; margin: 0; padding: 0 16px; color: #666; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ddd; padding: 8px 12px; }
th { background: #f0f0f0; }
&lt;/style&gt;"""
    html_body = md_lib.markdown(content, extensions=["tables", "fenced_code", "toc", "nl2br"])
    return "&lt;!DOCTYPE html&gt;\n&lt;html&gt;\n&lt;head&gt;\n&lt;meta charset='utf-8'&gt;\n" + css + "\n&lt;/head&gt;\n&lt;body&gt;\n" + html_body + "\n&lt;/body&gt;\n&lt;/html&gt;"


def to_txt(content):
    # 移除 frontmatter
    if content.startswith("---"):
        parts = content.split("---", 2)
        if len(parts) &gt;= 3:
            content = parts[2].strip()
    # 移除 Markdown 标记
    text = re.sub(r"^#{1,6}\s+", "", content, flags=re.MULTILINE)
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    text = re.sub(r"\*(.+?)\*", r"\1", text)
    text = re.sub(r"`(.+?)`", r"\1", text)
    text = re.sub(r"```.*?```", "", text, flags=re.DOTALL)
    text = re.sub(r"!\[.*?\]\(.*?\)", "", text)
    text = re.sub(r"\[(.+?)\]\(.*?\)", r"\1", text)
    text = re.sub(r"^[-*+]\s+", "", text, flags=re.MULTILINE)
    text = re.sub(r"^\d+\.\s+", "", text, flags=re.MULTILINE)
    text = re.sub(r"^&gt;+\s?", "", text, flags=re.MULTILINE)
    text = re.sub(r"^---+$", "", text, flags=re.MULTILINE)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def to_docx(content, input_path, template):
    from docx import Document
    from docx.shared import Inches, Pt
    import markdown as md_lib
    from bs4 import BeautifulSoup
    import io

    if template and Path(template).exists():
        doc = Document(template)
        for para in list(doc.paragraphs):
            p = para._element
            p.getparent().remove(p)
    else:
        doc = Document()

    base_dir = Path(input_path).parent
    html = md_lib.markdown(content, extensions=["tables", "fenced_code"])
    soup = BeautifulSoup(html, "html.parser")

    for elem in soup.children:
        tag = getattr(elem, "name", None)
        if tag is None:
            continue
        if tag in ["h1", "h2", "h3", "h4", "h5", "h6"]:
            level = int(tag[1])
            doc.add_heading(elem.get_text(), level=level)
        elif tag == "p":
            img = elem.find("img")
            if img:
                src = img.get("src", "")
                if not src.startswith("http"):
                    img_path = (base_dir / src).resolve()
                    if img_path.exists():
                        doc.add_picture(str(img_path), width=Inches(5.5))
                        if img.get("alt"):
                            doc.add_paragraph(img.get("alt"), style="Caption")
            else:
                doc.add_paragraph(elem.get_text())
        elif tag == "ul":
            for li in elem.find_all("li", recursive=False):
                doc.add_paragraph(li.get_text(), style="List Bullet")
        elif tag == "ol":
            for li in elem.find_all("li", recursive=False):
                doc.add_paragraph(li.get_text(), style="List Number")
        elif tag == "pre":
            code = elem.find("code")
            text = code.get_text() if code else elem.get_text()
            para = doc.add_paragraph()
            run = para.add_run(text)
            run.font.name = "Courier New"
            run.font.size = Pt(9)
        elif tag == "blockquote":
            doc.add_paragraph(elem.get_text(), style="Quote")
        elif tag == "hr":
            doc.add_paragraph("-" * 40)
        elif tag == "table":
            rows = elem.find_all("tr")
            if not rows:
                continue
            cols = 0
            for row in rows:
                cell_count = len(row.find_all(["td", "th"]))
                if cell_count &gt; cols:
                    cols = cell_count
            table = doc.add_table(rows=len(rows), cols=cols)
            table.style = "Table Grid"
            for r_idx, row in enumerate(rows):
                cells = row.find_all(["td", "th"])
                for c_idx, cell in enumerate(cells):
                    table.rows[r_idx].cells[c_idx].text = cell.get_text()

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def to_pdf(content):
    from weasyprint import HTML as WP_HTML
    html = to_html(content)
    return WP_HTML(string=html).write_pdf()


def convert_file(input_path, fmt, output_path, template):
    content = read_utf8(str(input_path))

    if fmt == "html":
        result = to_html(content)
        output_path.write_text(result, encoding="utf-8")
    elif fmt == "txt":
        result = to_txt(content)
        output_path.write_text(result, encoding="utf-8")
    elif fmt == "docx":
        result = to_docx(content, str(input_path), template)
        output_path.write_bytes(result)
    elif fmt == "pdf":
        result = to_pdf(content)
        output_path.write_bytes(result)

    print("✓ " + str(input_path) + " -&gt; " + str(output_path))


def main():
    parser = argparse.ArgumentParser(description="Markdown 格式转换工具")
    parser.add_argument("input", help="输入 Markdown 文件或目录")
    parser.add_argument("--to", required=True, choices=["docx", "html", "txt", "pdf"], dest="fmt", help="目标格式")
    parser.add_argument("-o", "--output", help="输出路径（文件或目录）")
    parser.add_argument("--template", help="DOCX 模板文件路径")

    args = parser.parse_args()
    input_path = Path(args.input)

    if input_path.is_dir():
        md_files = list(input_path.rglob("*.md"))
        out_dir = Path(args.output) if args.output else input_path / (args.fmt + "_output")
        out_dir.mkdir(parents=True, exist_ok=True)
        for md_file in md_files:
            rel = md_file.relative_to(input_path)
            out_file = out_dir / rel.with_suffix("." + args.fmt)
            out_file.parent.mkdir(parents=True, exist_ok=True)
            convert_file(md_file, args.fmt, out_file, args.template)
    else:
        if args.output:
            out_file = Path(args.output)
        else:
            out_file = input_path.with_suffix("." + args.fmt)
        out_file.parent.mkdir(parents=True, exist_ok=True)
        convert_file(input_path, args.fmt, out_file, args.template)


if __name__ == "__main__":
    main()


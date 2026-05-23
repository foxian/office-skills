# 图表题注高精度识别与样式绑定实施计划 (Caption Style Mapping Implementation Plan)

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 实现 Word 样式迁移中“图表题注 (Captions)”的高精度识别（文本正则+物理布局），样式名自动绑定至标准内置 `"Caption"`（支持目标文档缺失时动态补齐及中西文字体映射），并提供彻底的“物理防断页和表格不拆分”的版式保护。

**Architecture:** 
1. 升级 `style_analyzer.py` 的段落角色提取逻辑，支持 `Caption` 的指纹学习；
2. 升级 `style_transfer.py` 中的段落匹配逻辑，在 `Normal` 判定前插入第 1.5 级 `Caption` 检测判定（物理位置 1-2 个空行容错检测 + 正则捕获）；
3. 升级 `docx_editor.py` 的通用函数 `_apply_update_style_definition`，替换其原有的提前快速返回机制，支持缺失样式（如 `"Caption"`）时的动态创建与补齐，并在设置字体时提供精细的中西文字体区分（西文优先从指纹读取或兜底为 Times New Roman，中文映射至 `w:eastAsia` 的“仿宋”）；
4. 在 `style_transfer.py` 的 DSL 处理流后引入版式后处理物理防断页保护（图片段落 `keep_with_next`、表格行除最后一行外 `cantSplit`、表格上一文字段落及最后一行单元格段落 `keep_with_next`、上方题注段落自身 `keep_with_next`），彻底防止图文分裂。

---

## 🛠️ Task List (实施细分任务)

### Task 1: 升级样式检测模块与 Caption 检测算法

**Files:**
- Modify: [style_analyzer.py](file:///d:/melon/Documents/aiwork/office-master/word-master/skills/word-master/scripts/style_analyzer.py:430-512)
- Test: [test_style_analyzer.py](file:///d:/melon/Documents/aiwork/office-master/word-master/tests/test_style_analyzer.py)

- [ ] **Step 1: 编写失败的 Caption 物理判定和降级判定的单元测试**
  在 `tests/test_style_analyzer.py` 中新增 `test_detect_caption_type`，测试以下用例：
  1. 段落样式名为 `"Caption"` 时，能被正确检测。
  2. 文本匹配题注正则且前/后紧邻表格，能被正确检测。
  3. 题注与表格之间存在 1 个空行（非空段落截断），能被正确检测；如果中间夹杂非空段落，则判定为 False。
  4. 题注与图片（含有 `drawing` 的段落）存在空行关联时，能被正确检测。
  5. 物理不关联，但字号较小且居中的题注文本正则强匹配，能降级判定通过。

- [ ] **Step 2: 运行测试并确保失败**
  运行：`pytest tests/test_style_analyzer.py` 并确信上述新测试用例以未定义或属性错误失败。

- [ ] **Step 3: 实现 `_detect_caption_type` 检测函数**
  在 `skills/word-master/scripts/style_analyzer.py` 中引入 `re` 和 `qn`，实现完整的 `_detect_caption_type(paragraph, doc)` 函数。

  ```python
  def _detect_caption_type(paragraph, doc) -> bool:
      import re
      from docx.oxml.ns import qn
      
      style_name = paragraph.style.name.lower()
      if 'caption' in style_name:
          return True
          
      CAPTION_REGEX = re.compile(
          r"^(图|表|附图|附表|图表|Fig|Figure|Table|Chart)\s*(\d+([.-]\d+)*)\s*[:：]?\s*.*", 
          re.IGNORECASE
      )
      text = paragraph.text.strip()
      if not CAPTION_REGEX.match(text):
          return False
          
      # 物理位置检测 (向上看和向下看，最多容错 2 个节点)
      try:
          body_elements = list(doc.element.body)
          para_elem = paragraph._element
          para_idx = body_elements.index(para_elem)
          
          # 向上看最多 2 个元素找 tbl 或 drawing
          for offset in (1, 2):
              if para_idx - offset >= 0:
                  prev_elem = body_elements[para_idx - offset]
                  if prev_elem.tag.endswith('tbl'):
                      return True
                  if prev_elem.tag.endswith('p'):
                      # 若中间的段落有实质内容，说明图文不相邻，中断查找
                      if offset == 2 and body_elements[para_idx - 1].text.strip():
                          break
                      if prev_elem.find(qn('w:drawing')) is not None or prev_elem.find(qn('w:pict')) is not None:
                          return True
                          
          # 向下看最多 2 个元素找 tbl
          for offset in (1, 2):
              if para_idx + offset < len(body_elements):
                  next_elem = body_elements[para_idx + offset]
                  if next_elem.tag.endswith('tbl'):
                      return True
                  if next_elem.tag.endswith('p'):
                      # 若中间夹杂有文字的段落，中断查找
                      if offset == 2 and body_elements[para_idx + 1].text.strip():
                          break
      except Exception:
          pass
          
      # 降级视觉分母判定 (直接调用本地模块内的 _get_effective_value 函数，不使用外部导入)
      pf = paragraph.paragraph_format
      style_pf = paragraph.style.paragraph_format if hasattr(paragraph.style, 'paragraph_format') else None
      align = _get_effective_value(pf.alignment if pf else None, style_pf.alignment if style_pf else None)
      
      from docx.enum.text import WD_ALIGN_PARAGRAPH
      is_center = (align == WD_ALIGN_PARAGRAPH.CENTER)
      
      size = paragraph.style.font.size if paragraph.style.font else None
      is_small = size is not None and size.pt <= 11.0
      
      if is_center or is_small:
          return True
          
      return False
  ```

- [ ] **Step 4: 运行测试确保其通过**
  运行：`pytest tests/test_style_analyzer.py` 验证 `_detect_caption_type` 功能完全正常。

- [ ] **Step 5: 提交代码**
  ```bash
  git add skills/word-master/scripts/style_analyzer.py tests/test_style_analyzer.py
  git commit -m "feat: implement high-precision _detect_caption_type"
  ```

---

### Task 2: 扩展样式自动补齐与中西文字体支持

**Files:**
- Modify: [docx_editor.py](file:///d:/melon/Documents/aiwork/office-master/word-master/skills/word-master/scripts/docx_editor.py:454-502)
- Test: [test_style_transfer.py](file:///d:/melon/Documents/aiwork/office-master/word-master/tests/test_style_transfer.py)

- [ ] **Step 1: 编写样式自动补齐与中西文字体映射测试**
  在 `tests/test_style_transfer.py` 中，编写测试验证：执行 `update_style_definition` 后，即使目标文档原本没有 `"Caption"` 样式，该样式也必须被自动补全，且中文字体应用为“仿宋”，西文应用为指纹设置的值（或 Times New Roman 兜底值），并同时确认该改动未破坏其他普通样式的默认西文字体。

- [ ] **Step 2: 运行测试确保失败**
  运行：`pytest tests/test_style_transfer.py` 验证缺失样式拦截的失败。

- [ ] **Step 3: 重构 `docx_editor.py` 的 `_apply_update_style_definition` 逻辑**
  1. **移除/替换**原有函数的快速返回判断（即删除原有第 463-465 行的 `if style_name not in doc.styles: return` 代码块）。
  2. 实现安全补全和中西文字体映射：
  ```python
  def _apply_update_style_definition(doc, params):
      from docx.shared import Pt
      from docx.enum.text import WD_ALIGN_PARAGRAPH
      style_name = params.get('style')
      fp = params.get('fingerprint', {})
      if not style_name or not fp:
          print("[WARNING] op=update_style_definition: missing style or fingerprint, skipping.")
          return

      # 动态补齐不存在的样式 (解决原有快速返回导致的 Dead Code)
      if style_name not in doc.styles:
          from docx.enum.style import WD_STYLE_TYPE
          try:
              style = doc.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
              style.base_style = doc.styles['Normal']
          except Exception as e:
              print(f"[WARNING] op=update_style_definition: failed to add style '{style_name}': {e}, skipping.")
              return
      else:
          style = doc.styles[style_name]

      # Font size
      size_str = fp.get('size')
      if size_str:
          try:
              style.font.size = Pt(float(size_str.rstrip('pt')))
          except (ValueError, AttributeError):
              pass

      # Bold / italic
      if fp.get('bold') is not None:
          style.font.bold = fp['bold']
      if fp.get('italic') is not None:
          style.font.italic = fp['italic']

      # Font name 精细中西文控制 (不一刀切硬编码全局 Times New Roman)
      font_name = fp.get('font')
      if font_name:
          style.font.name = font_name
          from docx.oxml.ns import qn
          rpr = style._element.get_or_add_rPr()
          rfonts = rpr.find(qn('w:rFonts'))
          if rfonts is None:
              rfonts = rpr.makeelement(qn('w:rFonts'), {})
              rpr.insert(0, rfonts)
          for attr in list(rfonts.attrib):
              localname = attr.split('}')[-1] if '}' in attr else attr
              if localname in ('eastAsiaTheme', 'asciiTheme', 'hAnsiTheme', 'cstheme', 'majorHAnsi', 'majorEastAsia', 'majorBidi', 'minorBidi'):
                  del rfonts.attrib[attr]
          
          # 中文采用指纹设置的字体，西文优先从指纹读取或针对题注兜底使用 Times New Roman
          style_is_caption = (style_name.lower() == 'caption')
          west_font = font_name if not style_is_caption else 'Times New Roman'
          
          rfonts.set(qn('w:eastAsia'), font_name)
          rfonts.set(qn('w:ascii'), west_font)
          rfonts.set(qn('w:hAnsi'), west_font)
  ```

- [ ] **Step 4: 验证测试通过**
  运行：`pytest tests/test_style_transfer.py` 验证补全和字体渲染无误。

- [ ] **Step 5: 提交代码**
  ```bash
  git add skills/word-master/scripts/docx_editor.py tests/test_style_transfer.py
  git commit -m "feat: support dynamic style creation with precise font mapping in docx_editor"
  ```

---

### Task 3: 角色匹配流水线升级与版式后处理防断页机制

**Files:**
- Modify: [style_transfer.py](file:///d:/melon/Documents/aiwork/office-master/word-master/skills/word-master/scripts/style_transfer.py:126-240)
- Modify: [validate_style_profile.py](file:///d:/melon/Documents/aiwork/office-master/word-master/skills/word-master/scripts/validate_style_profile.py)

- [ ] **Step 1: 新增 `validate_style_profile.py` 对 `"Caption"` 角色的支持**
  将 `"Caption"` 添加至语义角色白名单，并同步在相关的格式指纹验证中给予支持。
  
- [ ] **Step 2: 重构 `style_transfer.py` 判定流水线**
  在 `generate_apply_ops` 判定流水线中，在 Level 1 之后且 Level 2 之前插入 【第 1.5 级】题注识别逻辑：
  ```python
  # Level 1.5: Caption detection
  if role is None:
      from style_analyzer import _detect_caption_type
      if _detect_caption_type(para, doc):
          role = "Caption" if "Caption" in available_roles else "Normal"
  ```

- [ ] **Step 3: 在 `style_transfer.py` 中实现防断页保护后处理函数**
  实现 `_apply_pagination_protection(doc_path)` 后处理函数：
  * **A. 图片下方题注**：前 1~2 段为图片宿主段落，设置图片段落 `keep_with_next = True`。
  * **B. 表格下方题注**：在表格上设置防断页物理保护：
    - 遍历该表格的所有行，对于**除最后一行外**的所有行设置 XML 元素 `w:cantSplit = True` 以禁断跨页断行。
    - 对表格的前一个段落（若存在且为普通文字段落）强制设置 `keep_with_next = True`，防止表格与前文断开。
    - 将表格最后一行的单元格内所有段落设置 `keep_with_next = True`。
  * **C. 表格/图表上方题注**：在题注段落自身强制设置 `keep_with_next = True`。
  
  **XML 操作规范（使用 OxmlElement 并维护严密 schema 顺序）：**
  ```python
  def _apply_pagination_protection(doc_path):
      import docx
      from docx.oxml.ns import qn
      from docx.oxml import OxmlElement
      import re

      doc = docx.Document(doc_path)
      paras = doc.paragraphs
      n = len(paras)
      
      caption_regex = re.compile(
          r"^(图|表|附图|附表|图表|Fig|Figure|Table|Chart)\s*(\d+([.-]\d+)*)\s*[:：]?\s*.*", 
          re.IGNORECASE
      )
      
      def is_caption_para(p):
          return p.style.name == "Caption" or bool(caption_regex.match(p.text.strip()))
          
      def set_keep_with_next(paragraph):
          paragraph.paragraph_format.keep_with_next = True
          paragraph.paragraph_format.keep_together = True

      # 1. 物理位置检测及图片关联防断页保护
      for i, para in enumerate(paras):
          if not is_caption_para(para):
              continue
              
          # 图片下方题注：向上寻找 (最多容错 1 个空行)
          for offset in (1, 2):
              if i - offset >= 0:
                  prev_p = paras[i - offset]
                  if offset == 2 and paras[i - 1].text.strip():
                      break
                  if prev_p._element.find(qn('w:drawing')) is not None or prev_p._element.find(qn('w:pict')) is not None:
                      set_keep_with_next(prev_p)
                      break
                      
          # 图片上方题注：向下寻找 (最多容错 1 个空行)
          for offset in (1, 2):
              if i + offset < n:
                  next_p = paras[i + offset]
                  if offset == 2 and paras[i + 1].text.strip():
                      break
                  if next_p._element.find(qn('w:drawing')) is not None or next_p._element.find(qn('w:pict')) is not None:
                      set_keep_with_next(para)
                      break

      # 2. 表格关联防断页与防跨页物理防护
      body_elements = list(doc.element.body)
      for i, elem in enumerate(body_elements):
          # 2.1 表格下方题注检测
          if elem.tag.endswith('tbl'):
              table = next(t for t in doc.tables if t._tbl is elem)
              
              # 向后看 2 个元素检查是否有题注段落
              for offset in (1, 2):
                  if i + offset < len(body_elements):
                      next_elem = body_elements[i + offset]
                      if offset == 2 and body_elements[i + 1].tag.endswith('p') and body_elements[i + 1].text.strip():
                          break
                      if next_elem.tag.endswith('p'):
                          next_p = next(p for p in doc.paragraphs if p._element is next_elem)
                          if is_caption_para(next_p):
                              # 对齐规格书：将表格除最后一行外所有行的 cantSplit 设为 True
                              for r_idx, row in enumerate(table.rows):
                                  if r_idx < len(table.rows) - 1:
                                      trPr = row._tr.get_or_add_trPr()
                                      if trPr.find(qn('w:cantSplit')) is None:
                                          trPr.append(trPr.makeelement(qn('w:cantSplit'), {}))
                              
                              # 补全上一文字段落防断开
                              if i - 1 >= 0 and body_elements[i - 1].tag.endswith('p'):
                                  prev_p = next(p for p in doc.paragraphs if p._element is body_elements[i - 1])
                                  if prev_p.text.strip():
                                      set_keep_with_next(prev_p)
                                      
                              # 表格最后一行的单元格内段落设置 keep_with_next 强制与下方题注绑定
                              if table.rows:
                                  for cell in table.rows[-1].cells:
                                      for cell_p in cell.paragraphs:
                                          set_keep_with_next(cell_p)
                              break
                              
          # 2.2 表格上方题注检测
          elif elem.tag.endswith('p'):
              para = next(p for p in doc.paragraphs if p._element is elem)
              if is_caption_para(para):
                  # 向后看最多 2 个元素看是否有 tbl
                  for offset in (1, 2):
                      if i + offset < len(body_elements):
                          next_elem = body_elements[i + offset]
                          if next_elem.tag.endswith('tbl'):
                              set_keep_with_next(para)
                              break
                              
      doc.save(doc_path)
  ```

  并在 `style_transfer.py` 的 `run` 函数的 `apply_operations` 之后进行调用：
  ```python
  apply_operations(draft_path, ops, output_path)
  _apply_pagination_protection(output_path)  # 后处理防断页
  ```

- [ ] **Step 4: 编写防断页整体集成测试**
  in `tests/test_style_transfer.py` 中编写集成测试，构造带有表格、图片以及相应空行和下方题注的文档，验证迁移后底层 XML 中 `cantSplit`、`keepNext` 的存在及正确性。

- [ ] **Step 5: 运行并验证测试全部通过**
  运行 `pytest` 确认全部测试通过！

- [ ] **Step 6: 提交代码**
  ```bash
  git add skills/word-master/scripts/style_transfer.py skills/word-master/scripts/validate_style_profile.py tests/test_style_transfer.py
  git commit -m "feat: integrate 1.5-level Caption pipeline and pagination protection post-processing"
  ```

---

## 🧪 Verification Plan (验证计划)

### 自动化验证
* 在 `/word-master` 模块下执行单元测试：
  ```powershell
  python -m pytest tests/test_style_analyzer.py tests/test_style_transfer.py -v
  ```
  预期结果：全部测试通过。

### 手动双盲集成校验
1. 使用包含图表与紧邻题注（图表上方和下方各一例，且存在空行空段落）的目标草稿 `draft.docx`。
2. 运行样式迁移脚本：
   ```powershell
   python skills/word-master/scripts/style_transfer.py --profile style_profile.json draft.docx output.docx
   ```
3. 在 Word 中打开 `output.docx`，确认题注字号自动设为了模板字号，中文字体设为仿宋，在页面底部的换行边界处完美防止了图文分离！

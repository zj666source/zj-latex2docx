# -*- coding: utf-8 -*-
"""Build a Word test docx full of LaTeX formulas (stored as plain text runs),
so the formula-conversion tool can be exercised end-to-end."""
import zipfile, os, html

OUT = r"D:\work文件\LaTeX公式测试文件.docx"

# Each item: (paragraph_label, latex_text_or_None)
# label is Chinese context; latex_text is a raw LaTeX snippet stored in a run.
# If latex_text is None, the label itself is a "should NOT convert" plain paragraph.
CASES = [
    ("一、基础行内公式（纯 LaTeX 文本，无定界符）", None),
    ("质能方程：", r"E = mc^2"),
    ("勾股定理：", r"a^2 + b^2 = c^2"),
    ("一般下标：", r"x_i , y_j"),

    ("二、希腊字母与上下标", None),
    ("", r"\alpha + \beta = \gamma"),
    ("", r"a_i^2 + b_j^3 = c"),

    ("三、分数与偏导数", None),
    ("", r"\frac{\partial u}{\partial t} = \alpha \nabla^2 u"),
    ("", r"f(x) = \frac{x}{x+1}"),

    ("四、积分与求和", None),
    ("", r"\int_0^1 x^2 dx = \frac{1}{3}"),
    ("", r"\sum_{k=1}^{n} \frac{1}{k^2} = \frac{\pi^2}{6}"),

    ("五、根号（含 n 次根）", None),
    ("", r"\sqrt{x^2 + y^2}"),
    ("", r"\sqrt[3]{a^3 + b^3}"),

    ("六、极限", None),
    ("", r"\lim_{x \to 0} \frac{\sin x}{x} = 1"),

    ("七、矩阵", None),
    ("", r"\begin{pmatrix} a & b \\ c & d \end{pmatrix}"),

    ("八、中文下标（边界 case）", None),
    ("", r"E_{夏} = \eta \cdot Q_{太阳}"),

    ("九、热力学 / 物性公式", None),
    ("", r"Q = mc\Delta T"),
    ("", r"PV = nRT"),

    ("十、带 $ 定界符格式", None),
    ("", r"$E = mc^2$"),
    ("", r"$\int_0^\infty e^{-x} dx = 1$"),

    ("十一、混合中文正文（下列不应被误转）", None),
    ("室内 CO₂ 浓度 CO₂ > 1000 ppm 时需要通风。", None),
    ("本装置综合效率为 ", r"\eta = 0.85", " 高于传统方案。"),

    ("十二、分段函数", None),
    ("", r"f(x) = \begin{cases} x & x \geq 0 \\ -x & x < 0 \end{cases}"),
]

def esc(t):
    return html.escape(t, quote=False)

body = []
for item in CASES:
    if len(item) == 2:
        label, latex = item
        if latex is None:
            # pure Chinese paragraph, must NOT be converted
            body.append(f'<w:p><w:r><w:t xml:space="preserve">{esc(label)}</w:t></w:r></w:p>')
        else:
            # label (context) + formula run
            runs = ""
            if label:
                runs += f'<w:r><w:t xml:space="preserve">{esc(label)}</w:t></w:r>'
            runs += f'<w:r><w:t xml:space="preserve">{esc(latex)}</w:t></w:r>'
            body.append(f'<w:p>{runs}</w:p>')
    else:
        # (label_before, latex, label_after) inline formula in Chinese paragraph
        before, latex, after = item
        runs = f'<w:r><w:t xml:space="preserve">{esc(before)}</w:t></w:r>'
        runs += f'<w:r><w:t xml:space="preserve">{esc(latex)}</w:t></w:r>'
        runs += f'<w:r><w:t xml:space="preserve">{esc(after)}</w:t></w:r>'
        body.append(f'<w:p>{runs}</w:p>')

document = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    '<w:body>'
    + "".join(body)
    + '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr>'
    + '</w:body></w:document>'
)

content_types = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    '</Types>'
)

rels = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    '</Relationships>'
)

os.makedirs(os.path.dirname(OUT), exist_ok=True)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as z:
    z.writestr("[Content_Types].xml", content_types)
    z.writestr("_rels/.rels", rels)
    z.writestr("word/document.xml", document)

print("WROTE:", OUT)
print("SIZE:", os.path.getsize(OUT), "bytes")

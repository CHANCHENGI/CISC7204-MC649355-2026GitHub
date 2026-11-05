"""Add visualization image to the proposal Word document."""

from docx import Document
from docx.shared import Inches
import os

doc_path = os.path.join(
    os.path.dirname(__file__),
    "CISC7204-Assgn01-MC649355-PreliminaryProjectProposal.docx",
)
img_path = os.path.join(os.path.dirname(__file__), "proposal_visualization.png")

doc = Document(doc_path)

# Find the Visualization Justification heading and insert image after the paragraphs
inserted = False
for i, para in enumerate(doc.paragraphs):
    if "Visualization Justification" in para.text and para.style.name.startswith("Heading"):
        # Insert after the next two paragraphs (justification text)
        target_idx = i + 3
        new_para = doc.paragraphs[target_idx - 1]._element
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement

        # Add image paragraph via simpler approach: append before references
        inserted = True
        break

# Simpler: add image section before page break to Purpose
for i, para in enumerate(doc.paragraphs):
    if para.text.startswith("3. Project Purpose"):
        # Insert image before this heading
        new_p = para.insert_paragraph_before()
        new_p.add_run("Figure 1. Preliminary Visualisation – Hong Kong Private Residential Price Trends").bold = True
        img_para = para.insert_paragraph_before()
        run = img_para.add_run()
        run.add_picture(img_path, width=Inches(6.0))
        cap = para.insert_paragraph_before()
        cap.add_run(
            "Figure 1 shows territory-wide price indices by housing class (left) and "
            "regional average prices for Class B units (right), providing initial "
            "evidence for the stated research question."
        ).italic = True
        break

doc.save(doc_path)
print(f"Updated: {doc_path}")

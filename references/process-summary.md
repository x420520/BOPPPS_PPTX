# Process Summary

Use this reference when explaining the method behind the skill.

## Flow

1. DOCX reading
   - Extract paragraphs and tables with `python-docx`.
   - Read the existing lesson-plan logic before writing new courseware.

2. Markdown courseware
   - Convert the lesson plan into a project-first OBE course design.
   - Use Bloom verbs for learning outcomes.
   - Organize the 80-minute class with BOPPPS.
   - Include code with Chinese comments.

3. PPTX generation
   - Use `pptxgenjs` to build an editable PowerPoint.
   - Split text, tables, and code to protect readable font sizes.
   - Use BOPPPS stage badges on the right side.
   - Validate slide count and font sizes from the PPTX XML.

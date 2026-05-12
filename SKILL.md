---
name: bopps-ppt
description: Create OBE and BOPPPS teaching slide decks from course lesson-plan DOCX files, Markdown courseware, or existing teaching notes. Use when the user asks for "bopps-ppt", BOPPPS PPT, OBE courseware, teaching PPT, 教案转课件, DOCX to Markdown to PPTX, or wants a polished editable PowerPoint for a Chinese Web/programming course with large readable text and Chinese code comments.
---

# BOPPS PPT

Use this skill to turn a course lesson plan into a polished editable PPTX. The usual path is:

1. Read the lesson-plan DOCX.
2. Write or improve a Markdown courseware file.
3. Generate a PPTX from the Markdown.
4. Check the PPTX structure, slide count, and font sizes.

## Output Standard

Use OBE first, then BOPPPS.

- Start from a concrete project scenario.
- Derive learning outcomes from the project.
- Use Bloom verbs in the outcome table.
- Arrange the 80-minute class by BOPPPS.
- Include assessment evidence and Rubrics.
- Keep code examples practical and add Chinese comments.
- Keep PPT text large. Prefer title 30+, body 21+, table 14+, code 14+.
- Split dense content into more slides instead of shrinking text.
- Make the right-side badge show the BOPPPS stage, not generic labels such as 实战、任务、项目、课程.

Right-side badge labels must use these words when possible:

- 导入
- 目标
- 前测
- 参与学习
- 后测
- 总结

## DOCX To Markdown

When the user provides a lesson-plan DOCX:

1. Use `python-docx` to extract paragraphs and tables.
2. Identify the topic, teaching hours, learning objectives, evaluation methods, key points, difficulties, and BOPPPS activity table.
3. Create a Markdown file named after the topic when the user did not give a target file.
4. Build the Markdown as courseware, not only as a lesson-plan table.

Recommended Markdown structure:

- Title page with course name, chapter, duration, OBE and BOPPPS note.
- Course basic information.
- Project-first scenario.
- Project deliverables.
- Learner analysis.
- OBE teaching idea.
- Learning outcomes with Bloom verbs.
- Evaluation methods.
- Key points and difficult points.
- BOPPPS 80-minute table.
- One section for each BOPPPS stage.
- Code pages with Chinese comments.
- Practice, Rubrics, post-assessment, summary, homework, teaching reflection.

## Markdown To PPTX

Prefer `pptxgenjs` when the workspace already uses it. If a suitable local generator exists, adapt it. Otherwise copy or reference:

`scripts/generate_bopps_ppt.js`

Run it from the workspace:

```powershell
node C:\Users\zero\.codex\skills\bopps-ppt\scripts\generate_bopps_ppt.js ".\实现分页.md" ".\实现分页-OBE-BOPPPS-美化版.pptx"
```

If `pptxgenjs` is missing, install it in the project or use the project dependency workflow already present in the repo.

## PPT Design Rules

- Use 16:9 wide slides unless the user asks otherwise.
- Use Microsoft YaHei for Chinese body text and Consolas for code.
- Use a clean teaching style with blue, white, gray, and small safety accent colors.
- Use panels and tables, but do not make dense nested cards.
- Put one main idea on each slide.
- For long code, split by 12 to 16 lines per slide.
- For long tables, split rows across slides.
- Keep the footer simple with course name and page number.
- Use BOPPPS stage badges on the right side.

## BOPPPS Stage Mapping

When the Markdown title does not explicitly contain a BOPPPS stage, infer the badge:

- Course info, project scenario, learner analysis: 导入
- OBE, learning outcomes, Bloom verbs, evaluation method: 目标
- 前测 and diagnostic questions: 前测
- Concepts, SQL, Java, Servlet, JSP, practice, safety analysis: 参与学习
- Rubrics, classroom acceptance, post-assessment: 后测
- Summary, homework, extension, reflection: 总结

## Validation

After generating PPTX:

1. Check that the file exists and is not tiny.
2. Open the PPTX as a zip file and count `ppt/slides/slide*.xml`.
3. Extract font sizes from slide XML and report the minimum and main ranges.
4. If LibreOffice or PowerPoint preview is available, render or open several slides.
5. If preview is not available, say that visual preview was not performed.

## Final Response

Report the final PPTX path, the source Markdown path, generated slide count, and validation result. Keep the answer concise.

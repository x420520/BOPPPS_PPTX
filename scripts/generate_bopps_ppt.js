const fs = require('fs');
const path = require('path');
const { createRequire } = require('module');

function loadPptxGen() {
  try {
    return require('pptxgenjs');
  } catch (error) {
    // When this script is executed from the skill folder, load dependencies from the caller workspace.
    const workspaceRequire = createRequire(path.join(process.cwd(), 'package.json'));
    return workspaceRequire('pptxgenjs');
  }
}

const PptxGenJS = loadPptxGen();

const ShapeType = {
  rect: 'rect',
  line: 'line',
  roundRect: 'roundRect',
};

const W = 13.333;
const H = 7.5;

const palette = {
  ink: '0F172A',
  blue: '2563EB',
  teal: '0F766E',
  amber: 'D97706',
  red: 'DC2626',
  slate: '475569',
  gray: '64748B',
  light: 'F8FAFC',
  panel: 'FFFFFF',
  line: 'CBD5E1',
  blueSoft: 'EAF2FF',
  greenSoft: 'ECFDF5',
  amberSoft: 'FFF7ED',
  redSoft: 'FEF2F2',
  indigo: '4F46E5',
  indigoSoft: 'EEF2FF',
  slateSoft: 'F1F5F9',
  codeBg: '111827',
  codeText: 'E5E7EB',
};

const fonts = {
  title: 'Microsoft YaHei',
  body: 'Microsoft YaHei',
  mono: 'Consolas',
};

const boppsStages = [
  { label: '导入', color: palette.blue, soft: palette.blueSoft },
  { label: '目标', color: palette.teal, soft: palette.greenSoft },
  { label: '前测', color: palette.amber, soft: palette.amberSoft },
  { label: '参与学习', color: palette.indigo, soft: palette.indigoSoft },
  { label: '后测', color: palette.red, soft: palette.redSoft },
  { label: '总结', color: palette.slate, soft: palette.slateSoft },
];

const contentBox = {
  x: 0.72,
  y: 1.72,
  w: 11.86,
  h: 4.96,
};

function usage() {
  console.log('Usage: node generate_bopps_ppt.js <source.md> <target.pptx>');
}

function cleanText(text) {
  return text
    .replace(/\*\*/g, '')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripFrontMatter(markdown) {
  const lines = markdown.replace(/\r\n/g, '\n').split('\n');
  if (lines[0]?.trim() !== '---') return markdown;
  for (let i = 1; i < lines.length; i += 1) {
    if (lines[i].trim() === '---') return lines.slice(i + 1).join('\n');
  }
  return markdown;
}

function parseTable(lines) {
  const rows = [];
  lines.forEach((line) => {
    if (/^\|\s*-+/.test(line)) return;
    rows.push(line.slice(1, -1).split('|').map((cell) => cleanText(cell)));
  });
  return rows;
}

function parseBlocks(text) {
  const lines = text.split('\n');
  const blocks = [];
  let i = 0;

  while (i < lines.length) {
    const trimmed = lines[i].trim();
    if (!trimmed) {
      i += 1;
      continue;
    }

    if (trimmed.startsWith('```')) {
      const language = trimmed.replace('```', '').trim();
      const codeLines = [];
      i += 1;
      while (i < lines.length && !lines[i].trim().startsWith('```')) {
        codeLines.push(lines[i].replace(/\t/g, '    '));
        i += 1;
      }
      blocks.push({ type: 'code', language, lines: codeLines });
      i += 1;
      continue;
    }

    if (trimmed.startsWith('|') && trimmed.endsWith('|')) {
      const tableLines = [];
      while (i < lines.length && lines[i].trim().startsWith('|') && lines[i].trim().endsWith('|')) {
        tableLines.push(lines[i].trim());
        i += 1;
      }
      blocks.push({ type: 'table', rows: parseTable(tableLines) });
      continue;
    }

    if (trimmed.startsWith('- ') || /^\d+\.\s+/.test(trimmed)) {
      blocks.push({ type: 'bullet', text: cleanText(trimmed.replace(/^-\s+/, '').replace(/^\d+\.\s+/, '')) });
      i += 1;
      continue;
    }

    blocks.push({ type: 'text', text: cleanText(trimmed) });
    i += 1;
  }

  return blocks.filter((block) => block.type !== 'table' || block.rows.length);
}

function parseMarkdown(markdown) {
  return stripFrontMatter(markdown)
    .split(/\n---\n/g)
    .map((chunk) => chunk.trim())
    .filter(Boolean)
    .map((chunk) => {
      const lines = chunk.split('\n');
      let title = '';
      const body = [];
      lines.forEach((line) => {
        if (!title && line.startsWith('# ')) title = cleanText(line.slice(2));
        else body.push(line);
      });
      return { title, blocks: parseBlocks(body.join('\n')) };
    })
    .filter((slide) => slide.title);
}

function blockWeight(block) {
  if (block.type === 'table' || block.type === 'code') return 99;
  return block.text.length > 30 ? 1.6 : 1;
}

function chunkTextBlocks(blocks, maxWeight = 9.6) {
  const chunks = [];
  let current = [];
  let weight = 0;
  blocks.forEach((block) => {
    const w = blockWeight(block);
    if (current.length && weight + w > maxWeight) {
      chunks.push(current);
      current = [];
      weight = 0;
    }
    current.push(block);
    weight += w;
  });
  if (current.length) chunks.push(current);
  return chunks;
}

function chunkCodeLines(lines) {
  const cleaned = [...lines];
  while (cleaned.length && !cleaned[cleaned.length - 1].trim()) cleaned.pop();
  const chunks = [];
  for (let i = 0; i < cleaned.length; i += 15) chunks.push(cleaned.slice(i, i + 15));
  return chunks.length ? chunks : [[]];
}

function chunkTableRows(rows) {
  if (rows.length <= 1) return [];
  const header = rows[0];
  const body = rows.slice(1);
  const cols = header.length;
  const rowsPerSlide = cols >= 5 ? 4 : cols >= 4 ? 5 : 7;
  const chunks = [];
  for (let i = 0; i < body.length; i += rowsPerSlide) {
    chunks.push([header, ...body.slice(i, i + rowsPerSlide)]);
  }
  return chunks;
}

function createPlan(slides) {
  const plan = [];
  slides.forEach((slide) => {
    const textBlocks = slide.blocks.filter((b) => b.type === 'text' || b.type === 'bullet');
    const specialBlocks = slide.blocks.filter((b) => b.type === 'table' || b.type === 'code');

    if (!specialBlocks.length) {
      chunkTextBlocks(textBlocks).forEach((chunk, idx) => {
        plan.push({ kind: 'text', title: idx === 0 ? slide.title : `${slide.title} 续`, blocks: chunk });
      });
      return;
    }

    if (textBlocks.length) {
      chunkTextBlocks(textBlocks, 7.9).forEach((chunk, idx) => {
        plan.push({ kind: 'text', title: idx === 0 ? slide.title : `${slide.title} 要点`, blocks: chunk });
      });
    }

    specialBlocks.forEach((block) => {
      if (block.type === 'table') {
        chunkTableRows(block.rows).forEach((rows, idx) => {
          plan.push({ kind: 'table', title: idx === 0 ? slide.title : `${slide.title} 表格续`, rows });
        });
      }
      if (block.type === 'code') {
        chunkCodeLines(block.lines).forEach((lines, idx) => {
          plan.push({ kind: 'code', title: idx === 0 ? slide.title : `${slide.title} 代码续`, language: block.language, lines });
        });
      }
    });
  });
  return plan;
}

function stageByLabel(label) {
  return boppsStages.find((stage) => stage.label === label) || boppsStages[3];
}

function boppsStage(title) {
  if (/Bridge-in|导入|项目先行|项目运行|课程基本|学情分析/.test(title)) return stageByLabel('导入');
  if (/Objective|目标|OBE|学习成果|布鲁姆|评价方式|重点|难点|BOPPPS 时间/.test(title)) return stageByLabel('目标');
  if (/Pre-assessment|前测/.test(title)) return stageByLabel('前测');
  if (/Post-assessment|后测|验收|Rubrics/.test(title)) return stageByLabel('后测');
  if (/Summary|总结|作业|拓展|反思/.test(title)) return stageByLabel('总结');
  return stageByLabel('参与学习');
}

function addBoppsNavigation(slide, currentStage, options = {}) {
  const x = options.x ?? 0.65;
  const y = options.y ?? 1.17;
  const w = options.w ?? 11.5;
  const h = options.h ?? 0.32;
  const gap = options.gap ?? 0.06;
  const segmentW = (w - gap * (boppsStages.length - 1)) / boppsStages.length;
  const centerY = y + h / 2;
  const isCover = options.variant === 'cover';

  if (!isCover) {
    slide.addShape(ShapeType.line, {
      x: x + segmentW / 2, y: centerY, w: w - segmentW, h: 0,
      line: { color: palette.line, pt: 1.1 },
    });
  }

  boppsStages.forEach((navStage, idx) => {
    const active = currentStage?.label === navStage.label;
    const segX = x + idx * (segmentW + gap);
    const colorStage = active || isCover;
    const fillColor = colorStage ? navStage.color : palette.panel;
    const lineColor = colorStage ? navStage.color : palette.line;
    const textColor = colorStage ? 'FFFFFF' : palette.slate;

    slide.addShape(ShapeType.roundRect, {
      x: segX, y, w: segmentW, h, rectRadius: 0.05,
      line: { color: lineColor, pt: active ? 1.1 : 0.7 },
      fill: { color: fillColor },
    });

    if (!colorStage) {
      slide.addShape(ShapeType.rect, {
        x: segX, y, w: segmentW, h: 0.04,
        line: { color: navStage.color, transparency: 100 },
        fill: { color: navStage.color },
      });
    }

    slide.addText(navStage.label, {
      x: segX, y: y + 0.075, w: segmentW, h: 0.16,
      fontFace: fonts.body, fontSize: 10.4, bold: colorStage,
      color: textColor, align: 'center', margin: 0, fit: 'shrink',
    });
  });
}

function addBase(slide, stage, pageNo, total) {
  slide.background = { color: palette.light };
  slide.addShape(ShapeType.rect, { x: 0, y: 0, w: 0.18, h: H, line: { color: stage.color, transparency: 100 }, fill: { color: stage.color } });
  slide.addShape(ShapeType.rect, { x: 0.18, y: 0, w: W - 0.18, h: 0.18, line: { color: stage.color, transparency: 100 }, fill: { color: stage.color } });
  slide.addShape(ShapeType.line, { x: 0.46, y: 6.94, w: 12.24, h: 0, line: { color: palette.line, pt: 1 } });
  slide.addText(`Web程序设计与实践  |  OBE + BOPPPS  |  ${pageNo}/${total}`, {
    x: 0.48, y: 7.05, w: 12.1, h: 0.22, fontFace: fonts.body, fontSize: 10.5, color: palette.gray, align: 'right', margin: 0,
  });
}

function addHeader(slide, title, stage) {
  slide.addText(title, {
    x: 0.65, y: 0.52, w: 9.9, h: 0.56, fontFace: fonts.title, fontSize: 30, bold: true, color: palette.ink, margin: 0, fit: 'shrink',
  });
  slide.addShape(ShapeType.roundRect, {
    x: 10.78, y: 0.53, w: 1.35, h: 0.38, rectRadius: 0.05, line: { color: stage.color, transparency: 100 }, fill: { color: stage.soft },
  });
  slide.addText(stage.label, {
    x: 10.78, y: 0.62, w: 1.35, h: 0.18, fontFace: fonts.body, fontSize: 12, bold: true, color: stage.color, align: 'center', margin: 0,
  });
  addBoppsNavigation(slide, stage);
  slide.addShape(ShapeType.line, { x: 0.65, y: 1.58, w: 11.5, h: 0, line: { color: stage.color, pt: 1.2 } });
}

function addCover(pptx, deckTitle) {
  const slide = pptx.addSlide();
  slide.background = { color: palette.ink };
  slide.addShape(ShapeType.rect, { x: 0, y: 0, w: W, h: H, line: { color: palette.ink, transparency: 100 }, fill: { color: palette.ink } });
  slide.addShape(ShapeType.rect, { x: 0, y: 0, w: W, h: 0.35, line: { color: palette.blue, transparency: 100 }, fill: { color: palette.blue } });
  slide.addShape(ShapeType.rect, { x: 0, y: 6.92, w: W, h: 0.2, line: { color: palette.teal, transparency: 100 }, fill: { color: palette.teal } });
  slide.addText(deckTitle || 'OBE + BOPPPS 教学课件', { x: 0.95, y: 1.38, w: 9.6, h: 0.8, fontFace: fonts.title, fontSize: 44, bold: true, color: 'FFFFFF', margin: 0, fit: 'shrink' });
  slide.addText('OBE + BOPPPS 教学课件', { x: 0.98, y: 2.38, w: 9.8, h: 0.35, fontFace: fonts.body, fontSize: 20, color: 'DBEAFE', margin: 0 });
  addBoppsNavigation(slide, null, { x: 0.98, y: 3.02, w: 9.4, h: 0.34, variant: 'cover' });
}

function addTextSlide(pptx, item, index, total) {
  const stage = boppsStage(item.title);
  const slide = pptx.addSlide();
  addBase(slide, stage, index + 1, total);
  addHeader(slide, item.title, stage);
  slide.addShape(ShapeType.roundRect, { ...contentBox, rectRadius: 0.06, line: { color: palette.line, pt: 1 }, fill: { color: palette.panel } });

  let y = contentBox.y + 0.31;
  item.blocks.forEach((block) => {
    if (block.type === 'bullet') {
      slide.addShape(ShapeType.rect, { x: 1.05, y: y + 0.12, w: 0.13, h: 0.13, line: { color: stage.color, transparency: 100 }, fill: { color: stage.color } });
      slide.addText(block.text, { x: 1.33, y, w: 10.7, h: 0.42, fontFace: fonts.body, fontSize: 21, color: palette.slate, margin: 0, breakLine: false });
      y += block.text.length > 32 ? 0.62 : 0.5;
    } else {
      slide.addText(block.text, { x: 1.03, y, w: 10.9, h: 0.44, fontFace: fonts.body, fontSize: 22, color: palette.ink, bold: item.blocks.length <= 3, margin: 0 });
      y += block.text.length > 28 ? 0.62 : 0.52;
    }
  });
}

function estimateColWidths(rows) {
  const cols = rows[0].length;
  const raw = Array(cols).fill(0);
  rows.forEach((row) => row.forEach((cell, idx) => { raw[idx] = Math.max(raw[idx], Math.min(cell.length, 22)); }));
  const total = raw.reduce((a, b) => a + b, 0) || cols;
  return raw.map((v) => Math.max(1.2, (v / total) * 11.2));
}

function addTableSlide(pptx, item, index, total) {
  const stage = boppsStage(item.title);
  const slide = pptx.addSlide();
  addBase(slide, stage, index + 1, total);
  addHeader(slide, item.title, stage);
  const rows = item.rows.map((row, r) => row.map((cell) => ({
    text: cell,
    options: {
      fontFace: fonts.body, fontSize: r === 0 ? 15 : 14.5, bold: r === 0, color: r === 0 ? 'FFFFFF' : palette.ink,
      valign: 'mid', margin: 0.08, fill: r === 0 ? { color: stage.color } : { color: 'FFFFFF' },
      border: { type: 'solid', color: palette.line, pt: 0.7 }, breakLine: false,
    },
  })));
  const rowH = Math.min(0.68, contentBox.h / rows.length);
  slide.addTable(rows, { x: contentBox.x, y: contentBox.y + 0.03, w: 11.85, h: rowH * rows.length, colW: estimateColWidths(item.rows), rowH: Array(rows.length).fill(rowH), border: { type: 'solid', color: palette.line, pt: 0.7 } });
}

function addCodeSlide(pptx, item, index, total) {
  const stage = boppsStage(item.title);
  const slide = pptx.addSlide();
  addBase(slide, stage, index + 1, total);
  addHeader(slide, item.title, stage);
  const lang = item.language ? item.language.toUpperCase() : 'CODE';
  slide.addShape(ShapeType.roundRect, { ...contentBox, rectRadius: 0.05, line: { color: palette.codeBg, transparency: 100 }, fill: { color: palette.codeBg } });
  slide.addShape(ShapeType.rect, { x: contentBox.x, y: contentBox.y, w: contentBox.w, h: 0.42, line: { color: stage.color, transparency: 100 }, fill: { color: stage.color } });
  slide.addText(lang, { x: 0.98, y: contentBox.y + 0.12, w: 1.7, h: 0.18, fontFace: fonts.body, fontSize: 11.5, bold: true, color: 'FFFFFF', margin: 0 });
  const longest = item.lines.reduce((m, line) => Math.max(m, line.length), 0);
  const fontSize = longest > 92 ? 13.8 : longest > 72 ? 14.2 : 15;
  slide.addText(item.lines.join('\n'), { x: 0.98, y: contentBox.y + 0.58, w: 11.28, h: contentBox.h - 0.84, fontFace: fonts.mono, fontSize, color: palette.codeText, margin: 0.08, breakLine: false, fit: 'shrink', valign: 'top' });
}

async function build() {
  const source = process.argv[2];
  const target = process.argv[3];
  if (!source || !target) {
    usage();
    process.exitCode = 1;
    return;
  }

  const markdown = fs.readFileSync(path.resolve(source), 'utf8');
  const slides = parseMarkdown(markdown);
  const plan = createPlan(slides);
  const pptx = new PptxGenJS();
  pptx.defineLayout({ name: 'WIDE', width: W, height: H });
  pptx.layout = 'WIDE';
  pptx.author = 'Codex';
  pptx.subject = 'OBE BOPPPS teaching PPT';
  pptx.title = path.basename(target, path.extname(target));
  pptx.lang = 'zh-CN';
  pptx.theme = { headFontFace: fonts.title, bodyFontFace: fonts.body, lang: 'zh-CN' };

  addCover(pptx, slides[0]?.title);
  const total = plan.length + 1;
  plan.forEach((item, idx) => {
    if (item.kind === 'text') addTextSlide(pptx, item, idx + 1, total);
    if (item.kind === 'table') addTableSlide(pptx, item, idx + 1, total);
    if (item.kind === 'code') addCodeSlide(pptx, item, idx + 1, total);
  });

  await pptx.writeFile({ fileName: path.resolve(target) });
  console.log(`PPT generated: ${path.resolve(target)}`);
  console.log(`Source slides: ${slides.length}`);
  console.log(`Generated slides: ${total}`);
}

build().catch((err) => {
  console.error(err);
  process.exitCode = 1;
});

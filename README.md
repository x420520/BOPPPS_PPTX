# BOPPPS PPTX

这是一个 Codex 技能，用来把课程教案、Markdown 课件或教学笔记整理成可编辑的 PowerPoint 课件。

它的重点不是简单排版，而是先按 OBE 思路梳理教学目标，再用 BOPPPS 组织课堂过程，最后生成适合上课直接使用的 PPTX。

## 本次课程示例

这个仓库里的命令和示例使用的课程主题是 `实现分页`。

`实现分页` 是本次课程要用的课题，主要面向 Web 程序设计中的列表分页功能。README 里的 `实现分页.md` 和 `实现分页-OBE-BOPPPS.pptx` 都是围绕这个主题写的。

其他课程使用时，把文件名、标题和示例内容换成自己的课程主题即可。

## 适合做什么

- 把 DOCX 教案改成课堂课件
- 把 Markdown 课件生成 PPTX
- 为 Web 程序设计、Java、数据库、前端等课程生成教学 PPT
- 生成带有导入、目标、前测、参与学习、后测、总结标识的 BOPPPS 课件
- 在每页生成贯穿 BOPPPS 全流程的导航栏，并高亮当前教学阶段
- 生成文字较大、代码可读、能继续编辑的 PowerPoint 文件

## 项目结构

```text
BOPPPS_PPTX
├─ SKILL.md
├─ agents
│  └─ openai.yaml
├─ references
│  └─ process-summary.md
└─ scripts
   └─ generate_bopps_ppt.js
```

主要文件说明如下。

- `SKILL.md` 是 Codex 技能说明，定义了使用场景、输出标准和工作流程
- `scripts/generate_bopps_ppt.js` 是 Markdown 转 PPTX 的生成脚本
- `references/process-summary.md` 记录了这个技能的处理思路
- `agents/openai.yaml` 是技能相关的模型配置

## 安装到 Codex

如果你想把它作为 Codex 技能使用，可以把仓库克隆到 Codex 的 skills 目录。

Windows PowerShell 示例。

```powershell
cd $env:USERPROFILE\.codex\skills
git clone https://github.com/x420520/BOPPPS_PPTX.git bopps-ppt
```

如果你已经有 `bopps-ppt` 目录，可以先备份旧目录，再重新克隆。

macOS 或 Linux 示例。

```bash
cd ~/.codex/skills
git clone https://github.com/x420520/BOPPPS_PPTX.git bopps-ppt
```

安装后，可以在 Codex 里直接说类似下面的话。

```text
使用 bopps-ppt 技能，把这个教案 DOCX 做成 OBE+BOPPPS 课件。
```

## 单独运行生成脚本

这个仓库里的生成脚本也可以独立运行。它会读取 Markdown 文件，并生成 PPTX。

先安装 Node.js，然后安装 `pptxgenjs`。

```powershell
cd BOPPPS_PPTX
npm install pptxgenjs
```

运行脚本。

```powershell
node scripts\generate_bopps_ppt.js ".\实现分页.md" ".\实现分页-OBE-BOPPPS.pptx"
```

这里的 `实现分页` 是本仓库的课程示例名称。你可以替换成自己的 Markdown 文件名和 PPTX 输出文件名。

命令格式如下。

```text
node scripts\generate_bopps_ppt.js <输入Markdown文件> <输出PPTX文件>
```

生成后会得到一个 `.pptx` 文件，可以用 PowerPoint 或 WPS 打开继续修改。

## 推荐输入格式

建议先把教案整理成 Markdown。结构可以参考下面的格式。

````markdown
# 实现分页

## 课程基本信息

- 课程名称：Web程序设计与实践
- 教学时长：80分钟
- 教学对象：软件技术专业学生

## 项目情境

学生需要为新闻列表、商品列表或用户列表实现分页显示。

## 学习目标

| 目标 | Bloom层次 | 评价证据 |
| --- | --- | --- |
| 解释分页参数 page 和 size 的作用 | 理解 | 前测问答 |
| 编写 SQL 分页查询语句 | 应用 | 课堂代码 |
| 完成 Servlet 分页控制逻辑 | 应用 | 项目验收 |

## BOPPPS 课堂安排

| 阶段 | 时间 | 教师活动 | 学生活动 | 评价 |
| --- | --- | --- | --- | --- |
| 导入 | 5分钟 | 展示列表数据过多的问题 | 观察页面问题 | 口头回答 |
| 目标 | 5分钟 | 明确本节课学习目标 | 记录目标 | 目标确认 |
| 前测 | 8分钟 | 提问 SQL limit 用法 | 快速作答 | 诊断问题 |
| 参与学习 | 45分钟 | 讲解并演示分页实现 | 编码实践 | 代码检查 |
| 后测 | 12分钟 | 布置分页小任务 | 独立完成 | Rubrics |
| 总结 | 5分钟 | 总结分页实现步骤 | 复盘收获 | 课后任务 |

## 参与学习：分页 SQL

```sql
-- 查询第 2 页数据，每页 10 条
select *
from news
order by id desc
limit 10 offset 10;
```
````

实际写课件时，可以继续加入任务说明、代码示例、课堂练习、Rubrics、课后作业和教学反思。

## 生成 PPT 的思路

这个技能采用三步处理方式。

1. 先读懂教案  
   从 DOCX 或现有文字里提取课程主题、教学目标、重点难点、评价方式和课堂活动。

2. 再改写成 Markdown 课件  
   不直接把教案表格搬进 PPT，而是先按项目情境组织内容。学习目标使用 Bloom 动词，课堂过程使用 BOPPPS。

3. 最后生成可编辑 PPTX  
   生成脚本会把标题、正文、表格和代码分开处理。长文本会拆成多页，代码也会按行数切分，避免字体太小。

## 设计原则

- 先有项目情境，再提出学习目标
- 先说明学生要做成什么，再安排教师讲什么
- 每页只保留一个主要意思
- 代码必须能看清，中文注释要能帮助课堂讲解
- 表格不要过密，内容多时拆成多页
- 每页保留 BOPPPS 导航栏，顺序展示导入、目标、前测、参与学习、后测、总结
- 当前教学阶段在导航栏中高亮显示，帮助课堂过程始终可见
- 右侧标识使用 BOPPPS 阶段名称
- PPT 采用 16:9 宽屏版式
- 中文正文优先使用微软雅黑
- 代码优先使用 Consolas

## BOPPPS 阶段标识

生成 PPT 时，右侧阶段标识尽量使用下面这些名称。
同时，每页顶部会显示完整 BOPPPS 导航栏，按同样顺序展示六个阶段，并自动高亮当前页所属阶段。

- 导入
- 目标
- 前测
- 参与学习
- 后测
- 总结

如果 Markdown 标题里没有直接写阶段名称，脚本会根据标题内容推断阶段。

常见对应关系如下。

- 课程信息、项目情境、学情分析通常归入导入
- OBE、学习目标、评价方式通常归入目标
- 诊断问题和预备知识检查归入前测
- 概念讲解、代码演示、课堂练习通常归入参与学习
- Rubrics、课堂验收和后测题通常归入后测
- 总结、作业、拓展和教学反思通常归入总结

## 生成后检查

建议生成 PPTX 后做一次简单检查。

```powershell
Get-Item ".\实现分页-OBE-BOPPPS.pptx"
```

检查重点如下。

- 文件是否成功生成
- 文件大小是否正常
- PPT 页数是否合理
- 标题和正文是否太小
- 代码是否需要继续拆页
- BOPPPS 阶段标识是否准确

如果需要更严格的检查，可以把 PPTX 当作 zip 文件打开，统计 `ppt/slides/slide*.xml` 的数量，并读取其中的字体大小。

## 常见问题

### 提示找不到 pptxgenjs

在运行脚本的目录安装依赖。

```powershell
npm install pptxgenjs
```

### 生成的 PPT 文字太多

先修改 Markdown，把一个大章节拆成多个小章节。这个脚本会尽量拆页，但输入内容本身太密时，人工拆分效果更好。

### 代码页不够美观

把代码控制在 12 到 16 行左右。每段代码前加一句课堂说明，学生更容易跟上。

### 阶段标识不准确

在 Markdown 标题里直接写阶段名称，例如 `## 参与学习：分页 SQL`。这样脚本就不需要推断。

## 适合的工作流程

推荐按下面顺序使用。

1. 准备原始教案 DOCX 或教学笔记
2. 让 Codex 用 `bopps-ppt` 技能整理成 Markdown 课件
3. 人工检查学习目标、评价方式和代码内容
4. 用脚本生成 PPTX
5. 打开 PPTX 检查视觉效果
6. 根据课堂需要微调文字和案例

## 许可证

本项目采用 GNU General Public License v3.0 许可证。

完整许可证文本见 `LICENSE`。

# mdx-convert

Markdown ↔ Excel 双向转换工具，支持 Windows 右键菜单集成。

适用场景：用 Markdown 作为表格数据的存储和 Git 提交格式，用 Excel 进行筛选、排序等操作，两者按需互转。

## 安装

**依赖（一次性）：**

```bash
pip install pandas openpyxl
```

**注册右键菜单（一次性，无需管理员权限）：**

```bash
python install.py
```

卸载：

```bash
python install.py --uninstall
```

---

## 使用方式

### 方式一：右键菜单（推荐）

| 操作 | 结果 |
|---|---|
| 右键 `.md` 文件 → **To Excel (.xlsx)** | 提取文件中所有表格，每张表生成一个 `.xlsx` |
| 右键 `.xlsx` 文件 → **To Markdown (.md)** | 每个 Sheet 生成一个 `.md` 文件 |

输出文件生成在**原文件所在目录**。

### 方式二：命令行

```bash
python mdx_convert.py to-xl  report.md      # → 每张表一个 .xlsx
python mdx_convert.py to-md  data.xlsx      # → 每个 Sheet 一个 .md
```

---

## 转换规则

### Markdown → Excel

- 文档中的**文字段落、标题、代码块全部忽略**，只提取表格
- 每张表生成一个独立的 `.xlsx` 文件
- 文件名规则：
  - 只有一张表 → `原文件名.xlsx`
  - 多张表 → `原文件名_表格标题.xlsx`（标题取表格前最近的 `#` 标题）
- 自动剥离 Obsidian 高亮（`==text==`）、粗体、斜体等行内标记
- 纯数字列自动识别为数值类型
- 自动冻结首行、调整列宽、表头加蓝色底色

### Excel → Markdown

- 每个 Sheet 生成一个独立的 `.md` 文件
- 文件名规则：
  - 只有一个 Sheet → `原文件名.md`
  - 多个 Sheet → `原文件名_Sheet名.md`
- 每个 `.md` 文件包含 `## Sheet名` 标题 + 标准 GFM 表格

---

## Markdown 表格格式

支持标准 GFM 格式，表格前的标题会作为输出文件名：

```markdown
## 服务器列表

这里可以有说明文字，转换时会被忽略。

| 名称  | IP       | 用途     |
|-------|----------|----------|
| app01 | 10.0.0.1 | 应用服务 |
| app02 | 10.0.0.2 | 应用服务 |

## 数据库清单

| 实例名    | 引擎      | 状态   |
|-----------|-----------|--------|
| lmd-mysql | MySQL 5.7 | 运行中 |
```

以上文件执行 `to-xl` 后会生成：
```
文件名_服务器列表.xlsx
文件名_数据库清单.xlsx
```

---

## Git 工作流建议

```
日常编辑 / git 提交   →  .md 文件
需要筛选 / 排序       →  右键 To Excel → Excel 操作 → 右键 To Markdown → git 提交
```

`.gitignore` 推荐配置：

```gitignore
__pycache__/
*.pyc
*.xlsx
*.md
!README.md
```
# Draco AI 自动化服务 · 交付样例仓库

> 这个仓库不是"作品集展示"，是**可以直接拿走用的成品源码**。
> 每一项都在真实环境跑通过，不含未验证的占位代码。

## 📦 交付样例

### ① `tools/csv-dashboard.html` — 数据看板生成器

单文件 HTML，零依赖，双击即用。

**解决什么**：把 Excel/CSV 报表变成能直接发给老板的看板，不用写代码。

**功能**
- 拖入 `.csv` 文件，或直接粘贴表格文本
- 自动识别列类型：数值列 / 维度列 / 日期列
- 输出 6 个 KPI（记录数、合计、平均值、最大值、最小值、去重数）
- Top-N 排行条形图（按指标列自动排序）
- 识别到日期列时自动生成趋势图
- 明细表（数值自动格式化为「万 / 亿」）
- **导出 HTML**：生成一个独立文件发给客户，双击就能看，不依赖本工具

**隐私**：纯前端，数据全部在浏览器本地解析，不上传任何服务器。

**用法**
```
1. 打开 csv-dashboard.html
2. 拖入 CSV（或点「载入示例数据」看效果）
3. 点「生成看板」
4. 点「导出 HTML」得到可分发的单文件看板
```

**可定制点**：加多表关联 / 定时邮件推送 / 写入指定数据库 / 自定义品牌配色。

---

### ② `tools/batch_tool.py` — 批量文件处理器

Python CLI，把重复的体力活变成一条命令。

**解决什么**：几百个文件重复改名、转换、加水印、提取文字，人工做要一天。

**功能**

| 子命令 | 作用 | 示例 |
|---|---|---|
| `rename` | 批量重命名（前缀 + 日期 + 序号 + 关键词替换） | `rename ./资料 -o ./整理后 --prefix 项目A --date "%Y%m%d"` |
| `convert pdf2img` | PDF → 图片（可指定 DPI） | `convert pdf2img ./合同 -o ./图片 --dpi 200` |
| `convert img2pdf` | 多张图片合并成一个 PDF | `convert img2pdf ./照片 -o ./合集.pdf` |
| `convert csv2xlsx` | CSV → Excel | `convert csv2xlsx ./表.csv -o ./表.xlsx` |
| `convert compress` | 图片批量压缩（质量 / 最大宽度可调） | `convert compress ./图 -o ./小图 --quality 60 --maxw 1200` |
| `watermark` | PDF 批量加水印 | `watermark ./合同 -o ./水印版 --text "仅供内部使用"` |
| `extract` | 批量提取文本，合并成一个 txt | `extract ./全部资料 -o ./汇总.txt` |

**依赖**
```bash
pip install pymupdf python-docx openpyxl pillow
```

**用法**
```bash
python tools/batch_tool.py --help
python tools/batch_tool.py rename --help
```

**已实测**：7 个子命令全部跑通（含真实 PDF / DOCX / XLSX / CSV / JPEG 输入）。

---

## 🧰 附带的通用工具

| 文件 | 说明 |
|---|---|
| `tools/md2pdf.js` | Markdown → PDF（Playwright 渲染，中文友好，支持分页与页眉页脚） |
| `tools/feishu-card.js` | 飞书富文本消息卡片生成器（JSON 输出，直接发到飞书群） |

## 🛠️ 交付标准

每单交付都包含：
1. 完整源码（所有权归客户）
2. `README` 使用说明 + 依赖清单
3. 用**客户真实数据**跑通的验收结果
4. 录屏演示
5. 7 天内免费修改

## 📞 合作

把你的需求发过来，小狐狸 24 小时内回方案和报价。

---

© Draco · AI 自动化交付

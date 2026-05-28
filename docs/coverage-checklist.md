# office-meta-parser 功能覆盖对照清单

> 状态标记:`[x]` 已完成 `[ ]` 未实现 `[~]` 部分实现(类型已定义但未完整解析/序列化)
>
> **范围说明**:本库面向 Web 场景(在线预览/轻量编辑/模板生成),已剔除 3D 模型、SmartArt、自定义几何路径、VBA/宏、OLE 嵌入对象、OMML 数学公式、数据透视表、外部链接、嵌入字体、NotesMaster/HandoutMaster、墨迹注释、拼音注音、自定义放映等复杂特性。

---

## 公共:元数据 (docProps/core.xml)

| 功能 | 解析 | 序列化 | 测试 | 备注 |
|------|------|--------|------|------|
| title | [x] | [x] | [x] | |
| subject | [x] | [x] | [x] | |
| creator | [x] | [x] | [x] | |
| description | [x] | [x] | [x] | |
| keywords | [x] | [x] | [x] | |
| lastModifiedBy | [x] | [x] | [x] | |
| created | [x] | [x] | [x] | |
| modified | [x] | [x] | [x] | |
| revision | [x] | [x] | [x] | |
| category | [x] | [x] | [x] | |
| app.xml 扩展属性 | [x] | [x] | [x] | P2 |
| custom.xml 自定义属性 | [x] | [x] | [x] | P2 |

---

## 公共:核心模块

| 功能 | 状态 | 测试 | 备注 |
|------|------|------|------|
| XML 解析 (parseXml) | [x] | [x] | |
| XML 序列化 (serializeXml) | [x] | [x] | |
| rels 解析 (parseRels) | [x] | [x] | |
| rels 序列化 (serializeRels) | [x] | [x] | round-trip 测试 |
| Content_Types 解析 | [x] | [x] | |
| Content_Types 序列化 | [x] | [x] | round-trip 测试 |
| ZIP 解压 (unzip) | [x] | [x] | |
| ZIP 打包 (zip) | [x] | [x] | |
| 元数据解析 (parseMeta) | [x] | [x] | round-trip 测试;三格式统一使用 |
| 元数据序列化 (serializeMeta) | [x] | [x] | 支持 10 个字段 |

---

## DOCX

### 已完成 - 解析

| 功能 | 状态 | 测试 | 备注 |
|------|------|------|------|
| 段落 (w:p) | [x] | [x] | |
| 文本片段 (w:r) | [x] | [x] | |
| 粗体 (w:b) | [x] | [x] | |
| 斜体 (w:i) | [x] | [x] | |
| 下划线 (w:u) | [x] | [x] | |
| 删除线 (w:strike) | [x] | [x] | |
| 上标 (w:vertAlign superscript) | [x] | [x] | |
| 下标 (w:vertAlign subscript) | [x] | [x] | |
| 字号 (w:sz) | [x] | [x] | 半磅值 |
| 颜色 (w:color) | [x] | [x] | |
| 字体 (w:rFonts) | [x] | [x] | |
| 段落对齐 (w:jc) | [x] | [x] | |
| 段落缩进 (w:ind) | [x] | [x] | |
| 段落间距 (w:spacing) | [x] | [x] | |
| 表格 (w:tbl) | [x] | [x] | |
| 表格行 (w:tr) | [x] | [x] | |
| 表格单元格 (w:tc) | [x] | [x] | |
| 表格边框(表级) | [x] | [x] | |
| 表格合并 | [x] | [x] | vMerge/gridSpan |
| 图片 (w:drawing, inline) | [x] | [x] | wp:inline 中的 a:blip |
| 样式定义 (w:styles) | [x] | [x] | paragraph/character/table 样式 |
| 段落样式引用 (w:pStyle) | [x] | [x] | |
| 批注 (w:comments) | [x] | [x] | |
| 修订记录 (w:trackChanges) | [x] | [x] | w:ins/w:del/rPrChange |
| 页眉 | [x] | [x] | word/header*.xml |
| 页脚 | [x] | [x] | word/footer*.xml |
| 超链接 | [x] | [x] | |
| 列表/编号引用 (w:numPr) | [x] | [x] | numbering.xml 完整定义 |

### 已完成 - 序列化

| 功能 | 状态 | 测试 | 备注 |
|------|------|------|------|
| 段落 / 文本片段 | [x] | [x] | |
| 粗体/斜体/下划线/删除线 | [x] | [x] | |
| 上标/下标/字号/颜色/字体 | [x] | [x] | |
| 段落对齐/缩进/间距 | [x] | [x] | |
| 表格 | [x] | [x] | |
| 图片 | [x] | [x] | w:drawing → wp:inline → pic:pic |
| 样式定义 | [x] | [x] | paragraph/character/table |
| 批注 | [x] | [x] | w:comment |
| 修订记录 | [x] | [x] | w:ins/w:del/w:rPrChange |
| 页眉/页脚 | [x] | [x] | w:hdr/w:ftr |
| 超链接 | [x] | [x] | w:hyperlink |

### 待完成 P0 - Web 渲染必需

| 功能 | OOXML | 解析 | 序列化 | 备注 |
|------|-------|------|--------|------|
| 字符高亮 | w:highlight | [x] | [x] | 标黄等强调 |
| 字符底纹 | w:shd (run) | [x] | [x] | 文字背景色 |
| 大写/小型大写 | w:caps / w:smallCaps | [x] | [x] | 标题排版 |
| 字符间距 | w:spacing / w:kern | [x] | [x] | |
| 双删除线 | w:dstrike | [x] | [x] | |
| 隐藏文本 | w:vanish | [x] | [x] | 决定是否渲染 |
| 大纲级别 | w:outlineLvl | [x] | [x] | 生成目录/导航关键 |
| 段落边框 | w:pBdr | [x] | [x] | 分隔线/引用块 |
| 段落底纹 | w:shd (paragraph) | [x] | [x] | 段落背景色 |
| 制表符 | w:tabs (位置+leader) | [x] | [x] | 目录"标题…页码"格式 |
| 段落分页控制 | w:keepNext / w:keepLines / w:pageBreakBefore | [x] | [x] | 分页预览 |
| 完整列表定义 | numbering.xml (abstractNum/lvl/numFmt/lvlText) | [x] | [x] | 编号样式正确显示 |
| 脚注 | footnotes.xml | [x] | [x] | 报告/论文常见 |
| 书签 | w:bookmarkStart / w:bookmarkEnd | [x] | [x] | 内部跳转锚点 |
| 完整域结构 | 跨 run fldChar (begin/separate/end) + instrText | [x] | [x] | 跨 run instrText 拼接 |
| 节方向 | w:pgSz@orient | [x] | [x] | 横/纵向 |
| 页码格式 | w:pgNumType | [x] | [x] | 起始页码/格式 |
| 首页不同 / 奇偶页不同 | w:titlePg / w:evenAndOddHeaders | [x] | [x] | |
| 节内容垂直对齐 | w:vAlign | [x] | [x] | |

### 待完成 P1 - 常见但可推迟

| 功能 | OOXML | 解析 | 序列化 | 备注 |
|------|-------|------|--------|------|
| 浮动图片 | wp:anchor | [x] | [x] | 真实文档大量使用 |
| 图片旋转 | a:xfrm@rot | [ ] | [ ] | |
| 图片翻转 | flipH / flipV | [ ] | [ ] | |
| 图片裁剪 | a:srcRect | [ ] | [ ] | |
| 文字环绕 | wrapSquare / wrapTight 等 | [x] | [x] | |
| 单元格底纹 | w:shd (cell) | [ ] | [ ] | 斑马纹/表头底色 |
| 单元格独立边框 | w:tcBorders | [x] | [x] | |
| 表格条件格式标志 | w:tblLook | [ ] | [ ] | |
| 单元格不换行 | w:noWrap | [x] | [x] | |
| 字符着重号 | w:em | [x] | [x] | 中日韩 |
| 字符语言 | w:lang | [ ] | [ ] | 拼写检查 |
| 字符边框 | w:bdr | [x] | [x] | |
| 垂直位置偏移 | w:position | [x] | [x] | |

### 待完成 P2 - 低优先级 / Web 场景可不实现

| 功能 | OOXML | 备注 |
|------|-------|------|
| RTL 字符/段落 | w:rtl / w:bidi | 阿拉伯/希伯来 |
| 文本效果 | w:emboss / w:imprint / w:outline / w:shadow | 浮雕等 |
| 不校验 | w:noProof | |
| 取消行号 | w:suppressLineNumbers | |
| 文本框段落 | w:framePr | |
| settings.xml | | 兼容性选项 |
| fontTable.xml | | 字体替换 |
| webSettings.xml | | |
| 尾注 | endnotes.xml | |

---

## XLSX

### 已完成 - 解析

| 功能 | 状态 | 测试 | 备注 |
|------|------|------|------|
| 多工作表 | [x] | [x] | workbook.xml + rels |
| 工作表名称 | [x] | [x] | |
| 共享字符串 (sharedStrings) | [~] | [x] | 仅普通字符串,缺富文本 r/rPr |
| 内联字符串 (inlineStr) | [x] | [x] | |
| 数字 | [x] | [x] | |
| 布尔值 | [x] | [x] | |
| 公式(普通) | [~] | [x] | 缺数组/共享公式 |
| 稀疏行 | [x] | [x] | |
| 样式/格式 (styles.xml) | [x] | [x] | fonts/fills/borders/numFmts/cellXfs |
| 合并单元格 | [x] | [x] | |
| 列宽 | [x] | [x] | |
| 行高 | [x] | [x] | |
| 日期类型 | [x] | [x] | |
| 错误值 | [x] | [x] | t="e" |
| 条件格式 | [x] | [x] | conditionalFormatting/cfRule |
| 数据验证 | [x] | [x] | dataValidations/dataValidation |
| 图片 | [x] | [x] | drawing/xdr:twoCellAnchor |
| 超链接 | [x] | [x] | |
| 筛选/排序 | [x] | [x] | autoFilter/filterColumn |
| 打印区域(基础) | [~] | [x] | 仅 printOptions/pageMargins/fitTo |

### 已完成 - 序列化

| 功能 | 状态 | 测试 | 备注 |
|------|------|------|------|
| 工作簿 / 工作表 | [x] | [x] | |
| 共享字符串 / 内联字符串 | [x] | [x] | |
| 数字 / 布尔值 / 公式 | [x] | [x] | |
| 样式/格式 | [x] | [x] | fonts/fills/borders/numFmts/cellXfs |
| 合并单元格 | [x] | [x] | |
| 列宽/行高 | [x] | [x] | |
| 条件格式 / 数据验证 | [x] | [x] | |
| 筛选/排序 | [x] | [x] | |
| 打印区域 | [x] | [x] | |
| 图片 | [x] | [x] | |

### 待完成 P0 - Web 渲染必需

| 功能 | OOXML | 解析 | 序列化 | 备注 |
|------|-------|------|--------|------|
| 冻结窗格 | sheetView/pane (xSplit/ySplit/state="frozen") | [x] | [x] | Web 表格预览必需 |
| 视图缩放/选区 | sheetView (zoomScale/topLeftCell/selection/activeCell) | [x] | [x] | |
| 工作表隐藏状态 | sheet@state (visible/hidden/veryHidden) | [x] | [x] | |
| 标签颜色 | sheetPr/tabColor | [x] | [x] | 多 sheet 切换 UI |
| 分组折叠 | outlinePr + row/col outlineLevel | [x] | [x] | 分级显示 |
| 完整 pageSetup | orientation/paperSize/scale | [x] | [x] | 当前只有 fitTo |
| 页眉页脚 | headerFooter/oddHeader/oddFooter | [x] | [x] | 打印预览 |
| 默认行高/列宽 | sheetFormatPr | [x] | [x] | 回退值 |
| 命名区域 | workbook/definedNames | [x] | [x] | 公式引用解析 |
| Excel Tables | tableParts + table*.xml | [x] | [x] | 结构化引用 |
| 传统批注 | comments*.xml + vmlDrawing | [x] | [x] | 评论气泡 |
| 现代批注 | threadedComments*.xml | [ ] | [ ] | Office 365 评论 |
| 富文本字符串 | sharedStrings 中 `<r><rPr><t>` | [x] | [x] | 单元格内部分文字格式 |
| 数组/共享公式 | f@t (array/shared) + f@ref/f@si | [x] | [x] | |
| Font.vertAlign | 上下标 | [x] | [x] | m³ 等显示 |
| Alignment.indent | | [x] | [x] | |
| Alignment.textRotation | -90~90 度 | [x] | [x] | 斜表头 |
| Alignment.shrinkToFit | | [x] | [x] | |
| Fill.gradientFill | degree/stops | [x] | [x] | 渐变背景 |
| Border.diagonal | diagonalUp / diagonalDown | [x] | [x] | 斜表头对角线 |
| Font.scheme | major/minor | [x] | [x] | 主题字体引用 |
| theme1.xml | | [x] | [x] | 颜色主题色 + tint |

### 待完成 P1 - 常见但可推迟

| 功能 | OOXML | 解析 | 序列化 | 备注 |
|------|-------|------|--------|------|
| 打印标题 | definedName _xlnm.Print_Titles | [x] | [x] | |
| 工作表保护 | sheetProtection | [x] | [x] | |
| 单元格保护 | Protection.locked/hidden | [x] | [x] | |
| 自动换行已支持 | wrapText | [x] | [x] | (已在 Alignment) |
| readingOrder / justifyLastLine | | [x] | [x] | |

### 待完成 P2

| 功能 | 备注 |
|------|------|
| 工作簿保护 | workbookProtection |
| 计算属性 | calcPr |
| 自定义颜色表 | customColors |

---

## PPTX

### 已完成 - 解析

| 功能 | 状态 | 测试 | 备注 |
|------|------|------|------|
| 多幻灯片 | [x] | [x] | presentation.xml + rels |
| 文本形状 (p:sp) | [x] | [x] | |
| 文本段落 (a:p) | [x] | [x] | algn/marL/indent/lvl/spcBef/spcAft/lnSpc/buChar/buAutoNum |
| 文本片段 (a:r) | [x] | [x] | |
| 图片形状 (p:pic) | [x] | [x] | |
| 占位符 (p:ph) | [x] | [x] | |
| 形状位置 (a:xfrm) | [x] | [x] | 含 flipH/flipV |
| 形状样式 | [x] | [x] | fill/border/shadow/opacity/glow/softEdge/reflection |
| 填充 (a:solidFill) | [x] | [x] | solid/gradient/pattern/noFill/grpFill/blipFill |
| 边框 (a:ln) | [x] | [x] | width/color/dash/cap/cmpd/headEnd/tailEnd |
| 阴影 | [x] | [x] | outer/inner |
| 组合形状 (p:grpSp) | [x] | [x] | 递归 + chOff/chExt |
| SlideMaster | [x] | [x] | 背景 + layouts + txStyles |
| SlideLayout | [x] | [x] | placeholders |
| Theme | [x] | [x] | colorScheme + fontScheme + fmtScheme |
| 幻灯片备注 | [x] | [x] | 从 notesSlide 提取 |
| 切换效果 (p:transition) | [x] | [x] | 类型/方向/advClick/advTm/sound |
| 表格 (a:tbl) | [x] | [x] | tableStyleId + tcPr(fill/border) |
| 动画 | [x] | [x] | p:timing/p:tnLst 等 |
| 超链接 | [x] | [x] | a:hlinkClick |

### 已完成 - 序列化

| 功能 | 状态 | 测试 | 备注 |
|------|------|------|------|
| 演示文稿结构 | [x] | [x] | |
| 幻灯片 | [x] | [x] | |
| 文本形状 | [x] | [x] | |
| 文本段落 | [x] | [x] | |
| 图片形状 | [x] | [x] | |
| 形状位置 | [x] | [x] | |
| 形状样式 | [x] | [x] | |
| 占位符 | [x] | [x] | |
| SlideMaster | [x] | [x] | |
| SlideLayout | [x] | [x] | |
| Theme | [x] | [x] | |
| 切换效果 | [x] | [x] | |
| 动画 | [x] | [x] | |
| 表格 | [x] | [x] | |

### 待完成 P0 - Web 渲染必需

| 功能 | OOXML | 解析 | 序列化 | 备注 |
|------|-------|------|--------|------|
| 预设形状类型 | a:prstGeom@prst | [x] | [x] | 200+ 种;不解析则全部按矩形渲染 |
| 形状水平/垂直翻转 | a:xfrm@flipH / flipV | [x] | [x] | |
| GroupShape 旋转 | grpSpPr/a:xfrm@rot | [x] | [x] | |
| GroupShape 子坐标 | a:chOff / a:chExt | [x] | [x] | 组内变换基准 |
| 文本框垂直锚点 | a:bodyPr@anchor (t/ctr/b) | [x] | [x] | 上/中/下对齐 |
| 文本框内边距 | a:bodyPr lIns/tIns/rIns/bIns | [x] | [x] | |
| 文本框换行 | a:bodyPr@wrap | [x] | [x] | |
| 文本自动缩放 | a:normAutofit / a:spAutoFit | [x] | [x] | |
| 文本方向 | a:bodyPr@vert | [x] | [x] | 竖排 |
| 段落对齐 | a:pPr@algn (l/ctr/r/just) | [x] | [x] | **基础属性** |
| 段落缩进/大纲级别 | a:pPr marL/indent/lvl | [x] | [x] | |
| 段间距 | a:pPr spcBef/spcAft | [x] | [x] | |
| 行距 | a:pPr lnSpc | [x] | [x] | |
| 项目符号 | a:buChar/a:buAutoNum/a:buNone/a:buFont/a:buSzPct/a:buClr | [x] | [x] | 列表渲染关键 |
| 列表样式默认 run | a:lstStyle/a:defRPr | [x] | [x] | |
| 形状图片填充 | a:blipFill (作为 spPr 子元素) | [x] | [x] | 当前只在 picture 中识别 |
| 渐变填充完整结构 | a:gradFill (stops/angle/path linear/radial) | [x] | [x] | 当前只标记 type='gradient' |
| 图案填充 | a:pattFill | [x] | [x] | |
| 无填充/组继承 | a:noFill / a:grpFill | [x] | [x] | 与 solid 区分 |
| 线条箭头端点 | a:headEnd / a:tailEnd | [x] | [x] | 箭头形状必备 |
| 线条复合 | a:ln@cmpd (sng/dbl/thickThin) | [x] | [x] | 双线/三线 |
| 线条笔帽 | a:ln@cap (flat/round/sq) | [x] | [x] | |
| 发光 | a:glow | [x] | [x] | css 可实现 |
| 柔化边缘 | a:softEdge | [x] | [x] | css filter |
| 主题格式表 | a:fmtScheme (fillStyleLst/lnStyleLst/effectStyleLst/bgFillStyleLst) | [x] | [x] | 形状常引用 lnRef/fillRef idx |
| 母版文字样式 | a:txStyles (titleStyle/bodyStyle/otherStyle) | [x] | [x] | 默认文字样式继承 |
| 颜色映射 | p:clrMap / p:clrMapOvr | [x] | [x] | bg1/tx1 占位符颜色 |
| 幻灯片尺寸 | p:sldSz | [x] | [x] | 当前硬编码 16:9 |
| 备注尺寸 | p:notesSz | [x] | [x] | 当前硬编码 |
| Slide 独立背景 | p:cSld/p:bg | [x] | [x] | 当前只 master 有 |
| 母版形状显示 | showMasterSp / showMasterPhAnim | [x] | [x] | |

### 待完成 P1 - 常见但可推迟

| 功能 | OOXML | 解析 | 序列化 | 备注 |
|------|-------|------|--------|------|
| 表格样式引用 | a:tblPr@tableStyleId | [x] | [x] | |
| 表格条件格式标志 | firstRow/firstCol/bandRow/bandCol | [ ] | [ ] | |
| 单元格属性 | a:tcPr (fill/border) | [x] | [x] | 单元格独立背景/边框 |
| 倒影 | a:reflection | [x] | [x] | css 可实现 |
| 内阴影 | a:innerShdw | [x] | [x] | |
| 完整 dash 类型 | a:prstDash 全枚举 | [x] | [x] | |
| 切换详细属性 | advClick/advTm/sound | [x] | [x] | 自动播放计时 |
| 嵌入音视频 | p:videoFile / p:audioFile | [x] | [x] | 可映射为 video/audio |

### 待完成 P2

| 功能 | 备注 |
|------|------|
| a:custGeom | 自定义路径(已在范围外但保留标记) |
| 演示文稿默认文字样式 | p:defaultTextStyle |
| 自定义放映 | p:custShowLst (范围外) |

---

## Round-trip 测试

| 格式 | 解析→序列化→重新解析 | 测试 | 备注 |
|------|----------------------|------|------|
| DOCX | [x] | [x] | body blocks 数量一致;rawXmlParts 保留原始 XML |
| XLSX | [x] | [x] | sheet 数量和名称一致;rawXmlParts 保留原始 XML |
| PPTX | [x] | [x] | slides 和 elements 数量一致;rawXmlParts 保留原始 XML |
| 元数据 (core/app/custom) | [x] | [x] | parseMeta/serializeMeta round-trip |

---

## 统计

### 已完成 vs 待完成

| 格式 | 已完成解析 | 已完成序列化 | P0 待完成 | P1 待完成 | P2 待完成 |
|------|-----------|-------------|-----------|-----------|-----------|
| DOCX | 54 | 37 | 0 | 6 | 9 |
| XLSX | 44 | 36 | 1 | 0 | 3 |
| PPTX | 58 | 52 | 0 | 1 | 3 |
| 公共 | 13 | 13 | 0 | 0 | 0 |

### 部分实现项(需补全)

- **XLSX**:共享字符串(sharedStrings)仅普通字符串,缺富文本 r/rPr
- **XLSX**:公式(普通)缺数组/共享公式
- **XLSX**:打印区域(基础)仅 printOptions/pageMargins/fitTo
- **XLSX**:现代批注(threadedComments)
- **custom.xml**:date/blob 类型已定义但缺测试

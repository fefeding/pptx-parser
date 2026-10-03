# OOXML 合规与已踩过的坑

## 核心原则

**PowerPoint / WPS 对非法 OOXML 是静默忽略，不报错。** 所以"JSON 写对了但真机没效果"时，第一怀疑对象永远是生成端写出的 XML 结构，而不是解析端。

排查顺序固定为：

1. `pptxToFiles(data)` 直接看生成端写了什么（比任何推理都快）
2. 与 ECMA-376 schema 对元素名 / 属性名 / 父子结构 / 枚举值
3. 写断言进自检（`examples/generate-test-pptx.mjs` 的 `assert(...)` 风格），防回归
4. 真机打开确认

## 已修复的真实案例（都是"真机完全无效果"级）

| # | 错误写法 | 正确写法 | 现象 |
|---|---|---|---|
| 1 | `<a:tableCellInsets l="…" r="…"/>` | `<a:tcPr marL="…" marR="…" marT="…" marB="…">` | 内边距是 `a:tcPr` 的**属性**，自造元素被忽略 → 单元格内边距无效 |
| 2 | `<a:lnTlToBr><a:ln w="12700">…</a:ln></a:lnTlToBr>` | `<a:lnTlToBr w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>…</a:solidFill></a:lnTlToBr>` | `lnTlToBr` / `lnBlToTr` **本身**就是 `CT_LineProperties`，多包一层 `a:ln` → 对角线不显示 |
| 3 | `<a:band1Horz>` | `<a:band1H>` | 交替行底纹区域名错误 → 该区域被忽略 |
| 4 | `tableStyleId` 指向 `tableStyles.xml` 里没有的 GUID | 必须在 `ppt/tableStyles.xml` 中有对应 `a:tblStyle` 定义 | 找不到样式 → 退化为"无样式无网格"，**表格连边框都没了** |
| 5 | `vert="wordArtVertical"` | `vert="wordArtVert"`（合法枚举：horz / vert / vert270 / wordArtVert / eaVert / mongolianVert / wordArtVertRtl） | 非法枚举 → 静默回退横排 |
| 6 | `a:gd name="adj1"` 用在 roundRect | roundRect/snip 用 `adj`；箭头/标注/星形用 `adj1`/`adj2` | 名字不匹配 → 忽略整组 avLst |
| 7 | `a:buAutoNum type="…"` 用非 `ST_TextAutonumberScheme` 值 | 如 `ea1ChsPeriod` / `arabicPeriod` / `alphaLcPeriod` / `romanUcPeriod` | WPS 回退成默认中文编号 |
| 8 | 解析端 HTML 渲染：把图案/图片填充塞进同一个 `styleTable`（以图案文本为 key），CSS 用固定 `fill="url(#pattPtrn)"` 这类 id | 每个形状生成独立 SVG `<pattern>`，id 形如 `pattPtrn_<shpId>` / `imgPtrn_<shpId>`，形状用 `fill="url(#…)"` 引用 | 固定 id 时多个形状互相覆盖（后者覆盖前者）；且 CSS background 不会按路径裁剪。**注意：生成端**形状图案填充是标准 `a:pattFill`（prst/fgClr/bgClr），与解析端的 SVG 方案无关 |
| 9 | 多个 duotone 图片共用固定 `filter id="svg_image_duotone"` | id 按形状加后缀 | 后者覆盖前者，所有形状都用最后一组的双色调 |
| 10 | 用同步 `new Image()` 探测图片尺寸 | 直接解析图片二进制头（PNG/JPEG/GIF/BMP/WEBP） | 图片未加载时 `image.width` 恒为 0 → `a:tile` 平铺尺寸算成 0 → 静默退回铺满 |

## 易错点清单（写生成端时逐条对照）

- **属性 vs 元素**：OOXML 里同一个语义有时是属性（如 `a:tcPr@marL`），有时是元素（如 `a:srcRect@l` 也是属性，但 `a:tile` 是元素）。不确定就查 ECMA-376 的 CT_ 类型定义。
- **枚举值**：`ST_*` 简单类型有固定取值集合（`vert`、`buAutoNum@type`、`prstDash`、`anchor`…），拼错即静默失效。
- **元素顺序**：`CT_BlipFillProperties` 里 `a:srcRect` 在 `a:stretch`/`a:tile` 之前；`CT_TableStyle` 的 region 顺序为 `tblBg → wholeTbl → band1H/2H/1V/2V → lastCol/firstCol/lastRow → seCell/swCell → firstRow`。顺序错会被判无效。
- **ID 唯一性**：SVG 的 `pattern` / `filter`、关系 rId、图表编号等在同一文档内必须唯一，否则互相覆盖。
- **千分比**：`a:tile@sx`、`a:srcRect@l`、`a:lumMod@val` 都是千分比/百分比整数（`比例 × 100000`）。
- **EMU 与 px**：1px = 9525 EMU，1pt = 12700 EMU，1 inch = 914400 EMU。
- **data URI 注入 HTML**：生成的内联样式若含未编码的引号，会截断 HTML 属性（`<td style='…url("data:image/svg+xml,…'…")'>`）。SVG 内联时务必把 `'` 编码为 `%27`。

## 解析端（HTML 渲染）常见差异

- 渲染结果缺少样式 → 检查是否注入了 `result.styles.global`。
- 图表不显示 → `pptxToHtml` 只给数据，需用 `examples/chart-lib/chart-renderer.js` + 全局 `echarts` 渲染。
- 形状图片/图案填充看不见 → 先确认素材本身是有效图片（曾经因为示例里放了一段截断的 base64，整块"看起来没解析"）。
- 单元格文字贴边 → 解析端按 ECMA-376 默认值（左右 0.1in / 上下 0.05in）渲染 padding；显式 `a:tcPr@mar*` 优先。

## 自检写法参考

```js
// 生成 → 重新解析 → 断言 XML（examples/generate-test-pptx.mjs 的风格）
const zip = await JSZip.loadAsync(buffer);
const xml = await zip.file('ppt/slides/slide10.xml').async('string');
assert('T10 对角线', /<a:lnTlToBr w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>/.test(xml));
assert('T10 无双层 a:ln', !/<a:lnTlToBr><a:ln/.test(xml));
```

断言要同时写"应该有什么"和"不应该有什么"（否定式断言能防止旧写法回潮）。

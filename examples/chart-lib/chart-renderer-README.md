# Chart Renderer 模块

这是一个独立的图表渲染模块，用于在浏览器中渲染从 PPTX 文件提取的图表。

## 功能特性

- 支持多种图表类型：折线图、柱状图、饼图、3D 饼图、面积图、散点图
- 支持 NVD3/D3 图表库
- 自动处理数据格式转换（特别是饼图）
- 支持图表样式（颜色、边距、图例位置等）
- 提供图表实例管理（创建、更新、销毁）

## 使用方法

### 基本使用

```html
<!-- 引入必要的库 -->
<script type="text/javascript" src="/src/lib/d3.min.js"></script>
<script type="text/javascript" src="/src/lib/nv.d3.min.js"></script>

<!-- 引入 ChartRenderer -->
<script type="module">
import { chartRenderer } from '/src/lib/chart-renderer.js';
window.chartRenderer = chartRenderer;
</script>

<!-- 在 PPTX 解析完成后渲染图表 -->
<script>
// 获取解析结果
const result = await pptxToHtml(fileData, options);

// 渲染所有图表
if (result.charts && result.charts.length > 0) {
    chartRenderer.renderCharts(result.charts);
}
</script>
```

### 创建自定义实例

```javascript
import { ChartRenderer } from '/src/lib/chart-renderer.js';

// 创建新的渲染器实例
const myRenderer = new ChartRenderer();

// 渲染图表
myRenderer.renderCharts(result.charts);
```

## API 参考

### ChartRenderer 类

#### `renderCharts(charts, container?)`

渲染所有图表。

**参数:**
- `charts` (Array): 图表数据数组，从 PPTX 解析器返回
- `container` (HTMLElement, 可选): 容器元素

**示例:**
```javascript
chartRenderer.renderCharts(result.charts);
```

#### `renderChart(chartInfo)`

渲染单个图表。

**参数:**
- `chartInfo` (Object): 图表信息对象
  - `chartId` (string): 图表容器 ID
  - `type` (string): 图表类型
  - `data` (Array): 图表数据
  - `style` (Object): 图表样式

**支持的图表类型:**
- `lineChart`: 折线图
- `barChart`: 柱状图
- `pieChart`: 饼图
- `pie3DChart`: 3D 饼图
- `areaChart`: 面积图
- `scatterChart`: 散点图

#### `updateChart(chartId, newData)`

更新图表数据。

**参数:**
- `chartId` (string): 图表 ID
- `newData` (Array): 新数据

**示例:**
```javascript
chartRenderer.updateChart('chart1', newData);
```

#### `destroyChart(chartId)`

销毁单个图表。

**参数:**
- `chartId` (string): 图表 ID

**示例:**
```javascript
chartRenderer.destroyChart('chart1');
```

#### `destroyAllCharts()`

销毁所有图表。

**示例:**
```javascript
chartRenderer.destroyAllCharts();
```

## 图表数据格式

### 折线图、柱状图、面积图

```javascript
[
    {
        key: "Series 1",
        values: [
            { x: 0, y: 10 },
            { x: 1, y: 20 },
            { x: 2, y: 15 }
        ],
        xlabels: ["Label 1", "Label 2", "Label 3"],
        style: {
            fillColor: "#FF0000",
            gradientFill: {
                color: ["#FF0000", "#00FF00"]
            }
        }
    }
]
```

### 饼图

饼图数据会被自动转换：

```javascript
// 输入格式（从 PPTX 解析器）
[
    {
        key: "Series 1",
        values: [
            { x: 0, y: 10 },
            { x: 1, y: 20 },
            { x: 2, y: 15 }
        ],
        xlabels: ["Label 1", "Label 2", "Label 3"]
    }
]

// 自动转换为
[
    { x: "Label 1", y: 10 },
    { x: "Label 2", y: 20 },
    { x: "Label 3", y: 15 }
]
```

### 散点图

```javascript
[
    {
        key: "Series 1",
        values: [
            { x: 0.5, y: 0.8 },
            { x: 0.6, y: 0.9 },
            { x: 0.7, y: 0.85 }
        ]
    }
]
```

## 图表样式

样式对象包含以下属性：

```javascript
{
    title: {
        color: "#000000",
        fontSize: 18
    },
    chartArea: {
        fillColor: "#FFFFFF",
        borderColor: "#000000",
        borderWidth: 1
    },
    legend: {
        position: "right", // b, t, l, r, tr, tl, br, bl
        fontSize: 12,
        color: "#000000"
    },
    categoryAxis: {
        color: "#000000",
        fontSize: 12,
        lineColor: "#000000"
    },
    valueAxis: {
        color: "#000000",
        fontSize: 12,
        lineColor: "#000000",
        gridlineColor: "#CCCCCC",
        gridlineWidth: 1
    }
}
```

## 依赖项

- D3.js
- NVD3.js

## 注意事项

1. 确保在引入 `chart-renderer.js` 之前先引入 D3 和 NVD3
2. 图表容器元素必须存在（通过 `chartId` 引用）
3. 饼图数据会自动转换格式
4. 图表样式会自动应用，但部分样式可能不被图表库支持

## 迁移到其他图表库

如果需要使用其他图表库（如 ECharts、Chart.js 等），可以：

1. 修改 `createChart()` 方法创建目标库的图表实例
2. 修改 `renderChartElement()` 方法使用目标库的渲染 API
3. 根据需要调整样式应用逻辑

## 示例

完整的示例请参考 `examples/basic-demo.html`。

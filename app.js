// 全局变量
let workbookData = {};
let currentSheet = null;
let currentData = [];
let filteredData = [];
let charts = [];
let currentRCode = { ggplot2: '', base: '', plotly: '' };
let currentModalRCode = { ggplot2: '', base: '', plotly: '' };

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    initializeEventListeners();
});

function initializeEventListeners() {
    const uploadBox = document.getElementById('uploadBox');
    const fileInput = document.getElementById('fileInput');
    const sheetSelector = document.getElementById('sheetSelector');
    const searchInput = document.getElementById('searchInput');
    const generateChartBtn = document.getElementById('generateChartBtn');
    const exportPdfBtn = document.getElementById('exportPdfBtn');
    const exportRCodeBtn = document.getElementById('exportRCodeBtn');

    // 上传事件
    uploadBox.addEventListener('click', () => fileInput.click());
    uploadBox.addEventListener('dragover', handleDragOver);
    uploadBox.addEventListener('dragleave', handleDragLeave);
    uploadBox.addEventListener('drop', handleDrop);
    fileInput.addEventListener('change', handleFileSelect);

    // 工作表选择
    sheetSelector.addEventListener('change', (e) => {
        if (e.target.value) {
            loadSheet(e.target.value);
        }
    });

    // 搜索功能
    searchInput.addEventListener('input', (e) => {
        filterData(e.target.value);
    });

    // 生成图表
    generateChartBtn.addEventListener('click', generateChart);

    // 导出PDF
    exportPdfBtn.addEventListener('click', exportToPdf);

    // 导出R代码
    if (exportRCodeBtn) {
        exportRCodeBtn.addEventListener('click', exportRCodeToFile);
    }

    // R代码标签页切换
    const rCodeTabs = document.querySelectorAll('.r-code-tab');
    rCodeTabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            switchRCodeTab(e.target.dataset.tab);
        });
    });

    // 模态框标签页
    const modalTabs = document.querySelectorAll('.modal-tab');
    modalTabs.forEach(tab => {
        tab.addEventListener('click', (e) => {
            switchModalTab(e.target.textContent.toLowerCase());
        });
    });
}

// 拖拽上传处理
function handleDragOver(e) {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.classList.add('dragover');
}

function handleDragLeave(e) {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.classList.remove('dragover');
}

function handleDrop(e) {
    e.preventDefault();
    e.stopPropagation();
    e.currentTarget.classList.remove('dragover');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

// 文件选择处理
function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

// 处理文件
function processFile(file) {
    const statusEl = document.getElementById('uploadStatus');
    
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        statusEl.textContent = '❌ 仅支持 Excel 文件 (.xlsx, .xls)';
        statusEl.className = 'upload-status error';
        return;
    }

    statusEl.textContent = '📁 正在解析文件...';
    statusEl.className = 'upload-status';
    statusEl.display = 'block';

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            workbookData = {};
            const sheetSelector = document.getElementById('sheetSelector');
            sheetSelector.innerHTML = '<option>请选择工作表</option>';

            // 解析所有工作表
            workbook.SheetNames.forEach(sheetName => {
                const worksheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(worksheet);
                workbookData[sheetName] = jsonData;

                // 添加到下拉菜单
                const option = document.createElement('option');
                option.value = sheetName;
                option.textContent = sheetName;
                sheetSelector.appendChild(option);
            });

            statusEl.textContent = `✅ 成功导入 ${Object.keys(workbookData).length} 个工作表`;
            statusEl.className = 'upload-status success';
            showControls();

        } catch (error) {
            statusEl.textContent = '❌ 文件解析失败：' + error.message;
            statusEl.className = 'upload-status error';
        }
    };

    reader.readAsArrayBuffer(file);
}

// 显示控制面板
function showControls() {
    document.getElementById('controlPanel').style.display = 'block';
}

// 加载工作表
function loadSheet(sheetName) {
    currentSheet = sheetName;
    currentData = JSON.parse(JSON.stringify(workbookData[sheetName]));
    filteredData = [...currentData];

    // 清空搜索框
    document.getElementById('searchInput').value = '';

    // 显示表格
    renderTable();

    // 更新图表下拉菜单
    updateChartSelects();

    // 显示相关部分
    document.getElementById('tableSection').style.display = 'block';
    document.getElementById('chartConfig').style.display = 'block';
    document.getElementById('exportSection').style.display = 'block';
}

// 渲染表格
function renderTable() {
    const table = document.getElementById('dataTable');
    const thead = table.querySelector('thead');
    const tbody = table.querySelector('tbody');

    // 清空表格
    thead.innerHTML = '';
    tbody.innerHTML = '';

    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="100%" style="text-align:center; padding: 40px;">无数据</td></tr>';
        return;
    }

    // 获取列名
    const columns = Object.keys(filteredData[0]);

    // 创建表头
    const headerRow = document.createElement('tr');
    columns.forEach(col => {
        const th = document.createElement('th');
        th.textContent = col;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    // 创建数据行
    filteredData.forEach(row => {
        const tr = document.createElement('tr');
        columns.forEach(col => {
            const td = document.createElement('td');
            td.textContent = row[col] || '';
            tr.appendChild(td);
        });
        tbody.appendChild(tr);
    });
}

// 搜索过滤
function filterData(searchTerm) {
    if (!searchTerm.trim()) {
        filteredData = [...currentData];
    } else {
        const term = searchTerm.toLowerCase();
        filteredData = currentData.filter(row => {
            return Object.values(row).some(val => 
                String(val).toLowerCase().includes(term)
            );
        });
    }
    renderTable();
}

// 更新图表下拉菜单
function updateChartSelects() {
    const columns = Object.keys(currentData[0] || {});
    const xAxis = document.getElementById('xAxisSelect');
    const yAxis = document.getElementById('yAxisSelect');

    [xAxis, yAxis].forEach(select => {
        const value = select.value;
        select.innerHTML = '<option>选择列</option>';
        columns.forEach(col => {
            const option = document.createElement('option');
            option.value = col;
            option.textContent = col;
            select.appendChild(option);
        });
        if (columns.includes(value)) {
            select.value = value;
        }
    });
}

// 生成图表
function generateChart() {
    const xAxis = document.getElementById('xAxisSelect').value;
    const yAxis = document.getElementById('yAxisSelect').value;
    const chartType = document.getElementById('chartTypeSelect').value;

    if (!xAxis || !yAxis) {
        alert('请选择X轴和Y轴数据');
        return;
    }

    // 准备数据
    const chartData = prepareChartData(xAxis, yAxis);

    // 创建图表容器
    const chartsSection = document.getElementById('chartsSection');
    chartsSection.style.display = 'block';
    const container = document.getElementById('chartsContainer');
    const chartId = 'chart_' + Date.now();
    
    const cardHtml = `
        <div class="chart-card">
            <div class="chart-header">
                <h3>${chartType === 'bar' ? '📊' : chartType === 'line' ? '📈' : '🥧'} ${xAxis} vs ${yAxis}</h3>
                <div class="chart-actions">
                    <button class="chart-btn" onclick="exportChartImage('${chartId}')" title="导出为PNG">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="7 10 12 15 17 10"/>
                            <line x1="12" y1="15" x2="12" y2="3"/>
                        </svg>
                    </button>
                    <button class="chart-btn" onclick="showRCode('${chartType}', '${xAxis}', '${yAxis}')" title="查看R代码">
                        <svg viewBox="0 0 24 24" fill="none" stroke="currentColor">
                            <path d="M10 6H6a2 2 0 0 0-2 2v10a2 2 0 0 0 2 2h4m4-16h4a2 2 0 0 1 2 2v10a2 2 0 0 1-2 2h-4m-4-16v16m4-16h4v16h-4"/>
                        </svg>
                    </button>
                </div>
            </div>
            <div class="chart-canvas-container">
                <canvas id="${chartId}"></canvas>
            </div>
        </div>
    `;
    container.innerHTML += cardHtml;

    // 绘制图表
    const ctx = document.getElementById(chartId).getContext('2d');
    let chartObj = null;

    const commonConfig = {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
            legend: {
                display: true,
                labels: {
                    font: { family: 'inherit', size: 12 },
                    color: '#6b7280',
                    padding: 15
                }
            }
        }
    };

    if (chartType === 'pie') {
        chartObj = new Chart(ctx, {
            type: 'doughnut',
            data: {
                labels: chartData.labels,
                datasets: [{
                    data: chartData.values,
                    backgroundColor: generateColors(chartData.values.length),
                    borderColor: 'white',
                    borderWidth: 2
                }]
            },
            options: commonConfig
        });
    } else {
        const isLine = chartType === 'line';
        chartObj = new Chart(ctx, {
            type: isLine ? 'line' : 'bar',
            data: {
                labels: chartData.labels,
                datasets: [{
                    label: yAxis,
                    data: chartData.values,
                    backgroundColor: isLine ? 'rgba(212, 165, 116, 0.3)' : 'rgba(107, 114, 128, 0.7)',
                    borderColor: isLine ? '#d4a574' : '#6b7280',
                    borderWidth: 2,
                    fill: isLine,
                    tension: 0.3,
                    pointRadius: isLine ? 5 : 0,
                    pointBackgroundColor: '#d4a574'
                }]
            },
            options: {
                ...commonConfig,
                scales: {
                    y: {
                        beginAtZero: true,
                        ticks: {
                            font: { family: 'inherit' },
                            color: '#6b7280'
                        },
                        grid: {
                            color: 'rgba(0, 0, 0, 0.05)'
                        }
                    },
                    x: {
                        ticks: {
                            font: { family: 'inherit' },
                            color: '#6b7280'
                        },
                        grid: {
                            display: false
                        }
                    }
                }
            }
        });
    }

    charts.push({
        id: chartId,
        obj: chartObj,
        type: chartType,
        xAxis: xAxis,
        yAxis: yAxis,
        data: chartData
    });

    // 显示R代码板块
    document.getElementById('rCodeSection').style.display = 'block';
    
    // 生成R代码
    generateAllRCodes(chartType, xAxis, yAxis);
}

// 准备图表数据
function prepareChartData(xAxis, yAxis) {
    const labels = [];
    const values = [];
    const seen = new Set();

    currentData.forEach(row => {
        const xVal = String(row[xAxis]);
        if (!seen.has(xVal)) {
            seen.add(xVal);
            labels.push(xVal);
            const yVal = parseFloat(row[yAxis]);
            values.push(isNaN(yVal) ? 0 : yVal);
        }
    });

    return { labels, values };
}

// 生成颜色
function generateColors(count) {
    const colors = [
        '#6b7280', '#d4a574', '#a0937d', '#7c9e8a',
        '#c97169', '#8b7355', '#9e8d7e', '#6b8b7c'
    ];
    const result = [];
    for (let i = 0; i < count; i++) {
        result.push(colors[i % colors.length]);
    }
    return result;
}

// 导出图表为PNG
function exportChartImage(chartId) {
    const canvas = document.getElementById(chartId);
    const link = document.createElement('a');
    link.href = canvas.toDataURL('image/png');
    link.download = `chart_${chartId}_${Date.now()}.png`;
    link.click();
}

// 显示R代码模态框
function showRCode(chartType, xAxis, yAxis) {
    generateModalRCodes(chartType, xAxis, yAxis);
    document.getElementById('rCodeModal').classList.add('show');
    // 默认显示第一个标签
    switchModalTab('ggplot2');
}

// 关闭R代码模态框
function closeRCodeModal() {
    document.getElementById('rCodeModal').classList.remove('show');
}

// 生成所有版本的R代码
function generateAllRCodes(chartType, xAxis, yAxis) {
    currentRCode = {
        ggplot2: generateRCodeGgplot2(chartType, xAxis, yAxis),
        base: generateRCodeBase(chartType, xAxis, yAxis),
        plotly: generateRCodePlotly(chartType, xAxis, yAxis)
    };

    // 更新R代码板块
    document.getElementById('rCodeGgplot2Content').textContent = currentRCode.ggplot2;
    document.getElementById('rCodeBaseContent').textContent = currentRCode.base;
    document.getElementById('rCodePlotlyContent').textContent = currentRCode.plotly;

    // 代码高亮
    if (typeof hljs !== 'undefined') {
        document.querySelectorAll('.r-code-content code').forEach(block => {
            hljs.highlightElement(block);
        });
    }
}

// 生成模态框中的R代码
function generateModalRCodes(chartType, xAxis, yAxis) {
    currentModalRCode = {
        ggplot2: generateRCodeGgplot2(chartType, xAxis, yAxis),
        base: generateRCodeBase(chartType, xAxis, yAxis),
        plotly: generateRCodePlotly(chartType, xAxis, yAxis)
    };

    document.getElementById('rCodeModalContent').textContent = currentModalRCode.ggplot2;
    document.getElementById('rCodeModalBaseContent').textContent = currentModalRCode.base;
    document.getElementById('rCodeModalPlotlyContent').textContent = currentModalRCode.plotly;
}

// ggplot2 版本
function generateRCodeGgplot2(chartType, xAxis, yAxis) {
    let code = `# 加载必要的库
library(dplyr)
library(ggplot2)
library(readxl)

# 数据导入
df <- read_excel("your_file.xlsx")

# 数据处理
data_summary <- df %>%
  group_by(${xAxis}) %>%
  summarise(${yAxis} = sum(${yAxis}, na.rm = TRUE))

# 绘制${chartType === 'bar' ? '柱状图' : chartType === 'line' ? '折线图' : '饼图'}\n`;

    if (chartType === 'bar') {
        code += `ggplot(data_summary, aes(x = reorder(${xAxis}, -${yAxis}), y = ${yAxis})) +
  geom_col(fill = '#6b7280', color = '#6b7280', alpha = 0.8) +
  geom_text(aes(label = round(${yAxis}, 2)), vjust = -0.5, size = 3) +
  theme_minimal() +
  labs(title = '${xAxis} vs ${yAxis}',
       subtitle = '数据分析可视化',
       x = '${xAxis}',
       y = '${yAxis}',
       caption = paste('Generated on', Sys.Date())) +
  theme(plot.title = element_text(size = 14, face = 'bold', hjust = 0.5),
        plot.subtitle = element_text(size = 11, hjust = 0.5),
        axis.text.x = element_text(angle = 45, hjust = 1),
        panel.grid.major.x = element_blank())`;
    } else if (chartType === 'line') {
        code += `ggplot(data_summary, aes(x = ${xAxis}, y = ${yAxis}, group = 1)) +
  geom_line(color = '#d4a574', size = 1) +
  geom_point(color = '#d4a574', size = 3) +
  geom_text(aes(label = round(${yAxis}, 2)), vjust = -0.8, size = 3) +
  theme_minimal() +
  labs(title = '${xAxis} vs ${yAxis} - 趋势分析',
       subtitle = '随时间或分类的变化趋势',
       x = '${xAxis}',
       y = '${yAxis}',
       caption = paste('Generated on', Sys.Date())) +
  theme(plot.title = element_text(size = 14, face = 'bold', hjust = 0.5),
        plot.subtitle = element_text(size = 11, hjust = 0.5),
        axis.text.x = element_text(angle = 45, hjust = 1))`;
    } else {
        code += `# 饼图数据处理
pie_data <- data_summary %>%
  mutate(percentage = round(${yAxis} / sum(${yAxis}) * 100, 2),
         label = paste0(${xAxis}, '\\n', percentage, '%'))

# 绘制饼图
ggplot(pie_data, aes(x = "", y = ${yAxis}, fill = ${xAxis})) +
  geom_col(width = 1) +
  geom_text(aes(label = label), position = position_stack(vjust = 0.5), size = 3) +
  coord_polar(theta = "y") +
  theme_void() +
  labs(title = '${xAxis} 分布占比分析',
       fill = '${xAxis}',
       caption = paste('Generated on', Sys.Date())) +
  theme(plot.title = element_text(size = 14, face = 'bold', hjust = 0.5),
        legend.position = 'right')`;
    }

    return code;
}

// Base R 版本
function generateRCodeBase(chartType, xAxis, yAxis) {
    let code = `# Base R 绘图
# 数据导入与处理
df <- read.csv("your_file.csv")
data_summary <- aggregate(${yAxis} ~ ${xAxis}, data = df, FUN = sum)

# 绘制${chartType === 'bar' ? '柱状图' : chartType === 'line' ? '折线图' : '饼图'}\n`;

    if (chartType === 'bar') {
        code += `par(mar = c(5, 4, 4, 2))
barplot(data_summary$${yAxis}, 
        names.arg = data_summary$${xAxis},
        main = '${xAxis} vs ${yAxis}',
        xlab = '${xAxis}',
        ylab = '${yAxis}',
        col = '#6b7280',
        border = NA,
        las = 2)
grid(NA, NULL)`;
    } else if (chartType === 'line') {
        code += `par(mar = c(5, 4, 4, 2))
plot(data_summary$${xAxis}, 
     data_summary$${yAxis},
     type = 'b',
     main = '${xAxis} vs ${yAxis} - 趋势',
     xlab = '${xAxis}',
     ylab = '${yAxis}',
     col = '#d4a574',
     pch = 19,
     lwd = 2)
grid(NULL, NULL, lty = 2, col = 'gray')`;
    } else {
        code += `pie(data_summary$${yAxis},
    labels = data_summary$${xAxis},
    main = '${xAxis} 分布占比',
    col = c('#6b7280', '#d4a574', '#a0937d', '#7c9e8a', '#c97169'),
    percentage = TRUE)`;
    }

    return code;
}

// Plotly 交互式版本
function generateRCodePlotly(chartType, xAxis, yAxis) {
    let code = `# 交互式绘图使用 Plotly
library(plotly)
library(dplyr)

# 数据处理
data_summary <- df %>%
  group_by(${xAxis}) %>%
  summarise(${yAxis} = sum(${yAxis}, na.rm = TRUE))

# 绘制${chartType === 'bar' ? '柱状图' : chartType === 'line' ? '折线图' : '饼图'}\n`;

    if (chartType === 'bar') {
        code += `plot_ly(data_summary, 
        x = ~${xAxis}, 
        y = ~${yAxis},
        type = 'bar',
        marker = list(color = '#6b7280')) %>%
  layout(title = '${xAxis} vs ${yAxis}',
         xaxis = list(title = '${xAxis}'),
         yaxis = list(title = '${yAxis}'),
         showlegend = FALSE,
         hovermode = 'x unified')`;
    } else if (chartType === 'line') {
        code += `plot_ly(data_summary,
        x = ~${xAxis},
        y = ~${yAxis},
        type = 'scatter',
        mode = 'lines+markers',
        line = list(color = '#d4a574', width = 2),
        marker = list(color = '#d4a574', size = 8)) %>%
  layout(title = '${xAxis} vs ${yAxis} - 趋势分析',
         xaxis = list(title = '${xAxis}'),
         yaxis = list(title = '${yAxis}'),
         hovermode = 'x unified')`;
    } else {
        code += `plot_ly(data_summary,
        labels = ~${xAxis},
        values = ~${yAxis},
        type = 'pie',
        marker = list(colors = c('#6b7280', '#d4a574', '#a0937d', '#7c9e8a'))) %>%
  layout(title = '${xAxis} 分布占比分析')`;
    }

    return code;
}

// R代码板块标签页切换
function switchRCodeTab(tab) {
    // 隐藏所有内容
    document.querySelectorAll('.r-code-content').forEach(el => {
        el.classList.remove('active');
    });
    // 移除所有标签的active类
    document.querySelectorAll('.r-code-tab').forEach(el => {
        el.classList.remove('active');
    });

    // 显示选中的内容和标签
    document.getElementById(`rCode${tab.charAt(0).toUpperCase() + tab.slice(1)}`).classList.add('active');
    event.target.classList.add('active');
}

// 模态框标签页切换
function switchModalTab(tab) {
    // 隐藏所有内容
    document.querySelectorAll('.modal-code-content').forEach(el => {
        el.classList.remove('active');
    });
    // 移除所有标签的active类
    document.querySelectorAll('.modal-tab').forEach(el => {
        el.classList.remove('active');
    });

    // 显示选中的内容和标签
    const contentId = `modal-${tab}`;
    const element = document.getElementById(contentId);
    if (element) {
        element.classList.add('active');
    }
    
    // 标记当前标签为active
    event.target?.classList.add('active');
}

// 复制R代码
function copyRCode(type = 'ggplot2') {
    const code = currentRCode[type];
    navigator.clipboard.writeText(code).then(() => {
        alert('代码已复制到剪贴板！');
    }).catch(() => {
        alert('复制失败，请手动复制');
    });
}

// 复制模态框中的R代码
function copyModalRCode() {
    // 获取当前显示的标签页
    const activeTab = document.querySelector('.modal-tab.active');
    const tab = activeTab?.textContent.toLowerCase() || 'ggplot2';
    copyRCode(tab);
}

// 导出R代码到文件
function exportRCodeToFile() {
    if (Object.values(currentRCode).every(code => code === '')) {
        alert('请先生成图表和R代码');
        return;
    }

    let content = `# 数据分析 R 代码导出
# 生成时间: ${new Date().toLocaleString()}
# 工作表: ${currentSheet}
\n`;

    content += `# ============================================\n`;
    content += `# ggplot2 版本（推荐用于学术论文）\n`;
    content += `# ============================================\n`;
    content += currentRCode.ggplot2 + '\n\n';

    content += `# ============================================\n`;
    content += `# Base R 版本（无需额外包）\n`;
    content += `# ============================================\n`;
    content += currentRCode.base + '\n\n';

    content += `# ============================================\n`;
    content += `# Plotly 版本（交互式）\n`;
    content += `# ============================================\n`;
    content += currentRCode.plotly + '\n';

    // 创建下载链接
    const element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(content));
    element.setAttribute('download', `R_code_${Date.now()}.R`);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
}

// 导出为PDF
async function exportToPdf() {
    const { jsPDF } = window.jspdf;
    const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4'
    });

    const pageHeight = pdf.internal.pageSize.getHeight();
    const pageWidth = pdf.internal.pageSize.getWidth();
    let yPosition = 10;

    try {
        // 添加标题
        pdf.setFontSize(16);
        pdf.text('数据分析报告', 10, yPosition);
        yPosition += 15;

        // 添加表格
        const tableCanvas = await html2canvas(document.querySelector('.table-wrapper'), {
            scale: 2,
            allowTaint: true,
            backgroundColor: '#ffffff'
        });
        const tableImg = tableCanvas.toDataURL('image/png');
        const tableHeight = (tableCanvas.height * pageWidth / tableCanvas.width - 20) / pageHeight * pageHeight;
        
        if (yPosition + tableHeight > pageHeight) {
            pdf.addPage();
            yPosition = 10;
        }
        pdf.addImage(tableImg, 'PNG', 10, yPosition, pageWidth - 20, tableHeight);
        yPosition += tableHeight + 10;

        // 添加图表
        const chartCards = document.querySelectorAll('.chart-card');
        for (let i = 0; i < chartCards.length; i++) {
            if (yPosition > pageHeight - 60) {
                pdf.addPage();
                yPosition = 10;
            }

            const chartCanvas = await html2canvas(chartCards[i], {
                scale: 2,
                allowTaint: true,
                backgroundColor: '#ffffff'
            });
            const chartImg = chartCanvas.toDataURL('image/png');
            const chartHeight = (chartCanvas.height * pageWidth / chartCanvas.width - 20) / pageHeight * pageHeight;
            
            pdf.addImage(chartImg, 'PNG', 10, yPosition, pageWidth - 20, Math.min(chartHeight, 100));
            yPosition += Math.min(chartHeight, 100) + 10;
        }

        // 添加页码
        const pageCount = pdf.internal.pages.length - 1;
        for (let i = 1; i <= pageCount; i++) {
            pdf.setPage(i);
            pdf.setFontSize(10);
            pdf.text(`第 ${i} 页`, pageWidth - 20, pageHeight - 10);
        }

        pdf.save(`report_${Date.now()}.pdf`);
    } catch (error) {
        alert('PDF导出失败: ' + error.message);
    }
}

// 关闭模态框点击外部
window.onclick = function(event) {
    const modal = document.getElementById('rCodeModal');
    if (event.target === modal) {
        modal.classList.remove('show');
    }
}

document.getElementById('labelForm').addEventListener('submit', function(e) {
    e.preventDefault();
    generateLabels();
});

// 添加自定义样本类型功能
let customSampleCounter = 0;

// 等待DOM加载完成后添加事件监听器
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('addCustomSample').addEventListener('click', function() {
        addCustomSampleType();
    });
});

function addCustomSampleType() {
    customSampleCounter++;
    const customSamplesDiv = document.getElementById('customSamples');
    
    const customSampleGroup = document.createElement('div');
    customSampleGroup.className = 'custom-sample-group';
    customSampleGroup.innerHTML = `
        <div class="form-group">
            <label for="customName${customSampleCounter}">样本类型名称:</label>
            <input type="text" id="customName${customSampleCounter}" placeholder="例如: 尿液" required>
        </div>
        <div class="form-group">
            <label for="customCount${customSampleCounter}">管数:</label>
            <input type="number" id="customCount${customSampleCounter}" min="0" value="1">
        </div>
        <button type="button" class="remove-btn" onclick="removeCustomSample(this)">删除</button>
    `;
    
    customSamplesDiv.appendChild(customSampleGroup);
}

function removeCustomSample(button) {
    button.parentElement.remove();
}

function generateLabels() {
    // 获取表单数据
    const location = document.getElementById('location').value.trim();
    const startNum = parseInt(document.getElementById('startNum').value);
    const endNum = parseInt(document.getElementById('endNum').value);
    const namesText = document.getElementById('names').value.trim();
    const sampleTime = document.getElementById('sampleTime').value;
    const packTime = document.getElementById('packTime').value;
    
    // 样本类型数量
    const wholeBlood = parseInt(document.getElementById('wholeBlood').value) || 0;
    const bloodCells = parseInt(document.getElementById('bloodCells').value) || 0;
    const plasma = parseInt(document.getElementById('plasma').value) || 0;
    const serum = parseInt(document.getElementById('serum').value) || 0;
    const other = parseInt(document.getElementById('other').value) || 0;
    
    // 获取自定义样本类型
    const customSamples = [];
    const customSampleGroups = document.querySelectorAll('.custom-sample-group');
    customSampleGroups.forEach((group) => {
        const nameInput = group.querySelector('input[type="text"]');
        const countInput = group.querySelector('input[type="number"]');
        const name = nameInput.value.trim();
        const count = parseInt(countInput.value) || 0;
        
        if (name && count > 0) {
            customSamples.push({ name, count });
        }
    });

    // 验证输入
    if (!location || !namesText || !sampleTime || !packTime) {
        alert('请填写所有必填字段！');
        return;
    }

    if (startNum > endNum) {
        alert('起始序号不能大于结束序号！');
        return;
    }

    // 处理姓名列表
    const names = namesText.split('\n').map(name => name.trim()).filter(name => name);
    const personCount = endNum - startNum + 1;

    if (names.length !== personCount) {
        alert(`姓名数量(${names.length})与人员序号范围(${personCount})不匹配！`);
        return;
    }

    // 格式化时间
    const formattedSampleTime = formatDateTime(sampleTime);
    const formattedPackTime = formatDateTime(packTime);

    // 生成标签
    const labels = [];
    
    for (let i = 0; i < personCount; i++) {
        const personNum = String(startNum + i).padStart(6, '0'); // 6位数字，前面补0
        const name = names[i];
        
        // 为每种样本类型生成标签
        const sampleTypes = [
            { name: '全血', count: wholeBlood },
            { name: '血细胞', count: bloodCells },
            { name: '血浆', count: plasma },
            { name: '血清', count: serum },
            { name: '其它', count: other },
            ...customSamples  // 添加自定义样本类型
        ];

        sampleTypes.forEach(sampleType => {
            for (let j = 1; j <= sampleType.count; j++) {
                const label = `${location}${personNum}_${name}_${sampleType.name}-${j}_${formattedSampleTime}_${formattedPackTime}`;
                labels.push(label);
            }
        });
    }

    if (labels.length === 0) {
        alert('没有生成任何标签，请检查样本数量设置！');
        return;
    }

    // 显示预览
    showPreview(labels);
    
    // 准备生成信息
    const generationInfo = {
        location,
        startNum,
        endNum,
        names,
        sampleTime,
        packTime,
        formattedSampleTime,
        formattedPackTime,
        sampleTypes: [
            { name: '全血', count: wholeBlood },
            { name: '血细胞', count: bloodCells },
            { name: '血浆', count: plasma },
            { name: '血清', count: serum },
            { name: '其它', count: other },
            ...customSamples
        ].filter(type => type.count > 0),
        totalLabels: labels.length,
        generatedAt: new Date()
    };
    
    // 生成Excel文件
    generateExcel(labels, generationInfo);
}

function formatDateTime(dateTimeString) {
    const date = new Date(dateTimeString);
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    
    return `${year}${month}${day}-${hours}${minutes}`;
}

function showPreview(labels) {
    const previewSection = document.getElementById('preview');
    const previewContent = document.getElementById('previewContent');
    
    const previewLabels = labels.slice(0, 10); // 只显示前10条
    previewContent.innerHTML = previewLabels.map(label => 
        `<div class="preview-item">${label}</div>`
    ).join('');
    
    if (labels.length > 10) {
        previewContent.innerHTML += `<div class="preview-item" style="color: #666; font-style: italic;">... 还有 ${labels.length - 10} 条标签</div>`;
    }
    
    previewSection.style.display = 'block';
}

function generateExcel(labels, generationInfo) {
    // 创建工作簿
    const wb = XLSX.utils.book_new();
    
    // 创建第一个工作表 - 标签数据
    const wsData = [['_'], ...labels.map(label => [label])];
    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = [{ wch: 60 }];
    XLSX.utils.book_append_sheet(wb, ws, '样本标签');
    
    // 创建第二个工作表 - 生成信息详细记录
    const infoData = [
        ['样本标签生成详细信息记录'],
        [''],
        ['=== 基本参数 ==='],
        ['地点缩写', generationInfo.location],
        ['人员序号起始', generationInfo.startNum],
        ['人员序号结束', generationInfo.endNum],
        ['人员总数', generationInfo.names.length],
        ['原始采样时间', generationInfo.sampleTime],
        ['原始分装时间', generationInfo.packTime],
        ['格式化采样时间', generationInfo.formattedSampleTime],
        ['格式化分装时间', generationInfo.formattedPackTime],
        [''],
        ['=== 人员名单详情 ==='],
        ['序号', '格式化序号', '姓名'],
        ...generationInfo.names.map((name, index) => [
            generationInfo.startNum + index,
            String(generationInfo.startNum + index).padStart(6, '0'), 
            name
        ]),
        [''],
        ['=== 样本类型配置 ==='],
        ['样本类型', '每人管数', '人员总数', '该类型总标签数'],
        ...generationInfo.sampleTypes.map(type => [
            type.name, 
            type.count, 
            generationInfo.names.length,
            type.count * generationInfo.names.length
        ]),
        [''],
        ['=== 生成统计 ==='],
        ['总标签数量', generationInfo.totalLabels],
        ['生成时间', generationInfo.generatedAt.toLocaleString('zh-CN', {
            year: 'numeric',
            month: '2-digit', 
            day: '2-digit',
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit'
        })],
        [''],
        ['=== 标签格式说明 ==='],
        ['标签格式', '地点缩写 + 6位序号 + _ + 姓名 + _ + 样本类型-管号 + _ + 采样时间 + _ + 分装时间'],
        ['示例', `${generationInfo.location}${String(generationInfo.startNum).padStart(6, '0')}_${generationInfo.names[0]}_${generationInfo.sampleTypes[0]?.name || '样本'}-1_${generationInfo.formattedSampleTime}_${generationInfo.formattedPackTime}`]
    ];
    
    const wsInfo = XLSX.utils.aoa_to_sheet(infoData);
    
    // 设置信息表格式，让列宽更合适
    wsInfo['!cols'] = [
        { wch: 25 },  // 第一列 - 标题列
        { wch: 20 },  // 第二列 - 数据列
        { wch: 15 },  // 第三列
        { wch: 20 }   // 第四列
    ];
    
    // 添加信息工作表
    XLSX.utils.book_append_sheet(wb, wsInfo, '生成信息');
    
    // 生成文件名
    const now = new Date();
    const timestamp = now.getFullYear() + 
                     String(now.getMonth() + 1).padStart(2, '0') + 
                     String(now.getDate()).padStart(2, '0') + '_' +
                     String(now.getHours()).padStart(2, '0') + 
                     String(now.getMinutes()).padStart(2, '0');
    
    const filename = `样本标签_${generationInfo.location}_${timestamp}.xlsx`;
    
    // 下载文件
    XLSX.writeFile(wb, filename);
    
    alert(`成功生成 ${labels.length} 个标签！文件已下载：${filename}`);
}

// 页面加载时设置默认时间
window.addEventListener('load', function() {
    const now = new Date();
    const dateTimeString = now.getFullYear() + '-' + 
                          String(now.getMonth() + 1).padStart(2, '0') + '-' + 
                          String(now.getDate()).padStart(2, '0') + 'T' + 
                          String(now.getHours()).padStart(2, '0') + ':' + 
                          String(now.getMinutes()).padStart(2, '0');
    
    document.getElementById('sampleTime').value = dateTimeString;
    document.getElementById('packTime').value = dateTimeString;
});
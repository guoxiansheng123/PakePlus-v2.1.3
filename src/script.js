// 存储解析后的户籍数据
let householdData = [];

// 处理拖拽经过事件
function handleDragOver(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.style.backgroundColor = '#f0f8ff';
    event.currentTarget.style.borderColor = '#2980b9';
}

// 处理拖拽离开事件
function handleDragLeave(event) {
    event.preventDefault();
    event.stopPropagation();
    event.currentTarget.style.backgroundColor = '';
    event.currentTarget.style.borderColor = '#3498db';
}

// 处理文件放下事件
function handleDrop(event) {
    event.preventDefault();
    event.stopPropagation();
    
    // 恢复样式
    event.currentTarget.style.backgroundColor = '';
    event.currentTarget.style.borderColor = '#3498db';
    
    // 获取拖拽的文件
    const files = event.dataTransfer.files;
    if (files.length > 0) {
        const file = files[0];
        handleFile(file);
    }
}

// 处理文件上传
function handleFileUpload(event) {
    const file = event.target.files[0];
    handleFile(file);
}

// 处理文件的核心函数
function handleFile(file) {
    if (!file) return;
    
    // 检查文件类型
    if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
        alert('请上传Excel文件（.xlsx或.xls格式）');
        return;
    }
    
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        try {
            // 使用SheetJS解析Excel文件
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // 将工作表转换为JSON格式
            const jsonData = XLSX.utils.sheet_to_json(worksheet);
            
            // 处理数据，按户分组
            processHouseholdData(jsonData);
        } catch (error) {
            console.error('解析Excel文件失败:', error);
            alert('解析Excel文件失败，请检查文件格式是否正确');
        }
    };
    
    reader.readAsArrayBuffer(file);
}

// 处理户籍数据，按户分组
function processHouseholdData(rawData) {
    // 按户号分组
    const householdMap = new Map();
    
    rawData.forEach(item => {
        // 确保户号存在
        const householdId = item['户号'] || item['householdId'] || '未知户号';
        
        if (!householdMap.has(householdId)) {
            householdMap.set(householdId, {
                householdId: householdId,
                address: item['地址'] || item['address'] || '',
                members: []
            });
        }
        
        // 添加家庭成员
        householdMap.get(householdId).members.push({
            name: item['姓名'] || item['name'] || '',
            relationship: item['与户主关系'] || item['relationship'] || '',
            gender: item['性别'] || item['gender'] || '',
            birthDate: item['出生日期'] || item['birthDate'] || '',
            idCard: item['身份证号'] || item['idCard'] || '',
            nationality: item['民族'] || item['nationality'] || '',
            education: item['文化程度'] || item['education'] || '',
            occupation: item['职业'] || item['occupation'] || ''
        });
    });
    
    // 转换为数组并排序
    householdData = Array.from(householdMap.values());
    householdData.sort((a, b) => a.householdId.localeCompare(b.householdId));
    
    // 显示户籍信息
    displayHouseholdData();
}

// 显示户籍信息
function displayHouseholdData() {
    const container = document.getElementById('householdContainer');
    
    if (householdData.length === 0) {
        container.innerHTML = '<div class="no-data">未找到户籍数据</div>';
        return;
    }
    
    let html = '';
    
    householdData.forEach((household, index) => {
        html += `
            <div class="household-item" id="household-${index}">
                <div class="household-header">
                    <h3>户号：${household.householdId} ${household.address ? `| 地址：${household.address}` : ''}</h3>
                    <button class="print-btn" onclick="printHousehold(${index})">打印此户</button>
                </div>
                <div class="member-list">
        `;
        
        household.members.forEach(member => {
            html += `
                <div class="member-item">
                    <div><span class="label">姓名：</span><span class="value">${member.name}</span></div>
                    <div><span class="label">与户主关系：</span><span class="value">${member.relationship}</span></div>
                    <div><span class="label">性别：</span><span class="value">${member.gender}</span></div>
                    <div><span class="label">出生日期：</span><span class="value">${member.birthDate}</span></div>
                    <div><span class="label">身份证号：</span><span class="value">${member.idCard}</span></div>
                    <div><span class="label">民族：</span><span class="value">${member.nationality}</span></div>
                    <div><span class="label">文化程度：</span><span class="value">${member.education}</span></div>
                    <div><span class="label">职业：</span><span class="value">${member.occupation}</span></div>
                </div>
            `;
        });
        
        html += `
                </div>
            </div>
        `;
    });
    
    container.innerHTML = html;
}

// 打印指定户的信息
function printHousehold(index) {
    const householdElement = document.getElementById(`household-${index}`);
    if (!householdElement) return;
    
    // 创建打印区域
    const printContainer = document.createElement('div');
    printContainer.innerHTML = householdElement.outerHTML;
    printContainer.style.width = '100%';
    printContainer.style.padding = '20px';
    
    // 移除打印按钮
    const printBtn = printContainer.querySelector('.print-btn');
    if (printBtn) printBtn.remove();
    
    // 创建打印窗口
    const printWindow = window.open('', '_blank');
    printWindow.document.write(`
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="UTF-8">
            <title>户籍信息打印</title>
            <style>
                body {
                    font-family: Arial, sans-serif;
                    margin: 0;
                    padding: 20px;
                }
                .household-item {
                    border: 1px solid #e0e0e0;
                    border-radius: 8px;
                    padding: 20px;
                    background-color: #fafafa;
                }
                .household-header {
                    margin-bottom: 15px;
                    padding-bottom: 10px;
                    border-bottom: 1px solid #e0e0e0;
                }
                .household-header h3 {
                    color: #2c3e50;
                }
                .member-list {
                    margin-top: 15px;
                }
                .member-item {
                    display: grid;
                    grid-template-columns: 1fr 1fr 1fr 1fr;
                    gap: 10px;
                    padding: 10px;
                    background-color: white;
                    border-radius: 4px;
                    margin-bottom: 8px;
                    border-left: 4px solid #3498db;
                }
                .member-item .label {
                    font-weight: bold;
                    color: #555;
                }
                .member-item .value {
                    color: #333;
                }
            </style>
        </head>
        <body>
            ${printContainer.outerHTML}
        </body>
        </html>
    `);
    
    printWindow.document.close();
    printWindow.print();
    printWindow.close();
}

// 初始化页面
function init() {
    console.log('户籍套打工具已初始化');
}

// 页面加载完成后初始化
window.addEventListener('DOMContentLoaded', init);
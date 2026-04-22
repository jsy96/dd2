// Vercel serverless function for processing manifest files
const path = require('path');
const fs = require('fs').promises;
const stream = require('stream');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');

// 解析舱单 Excel 文件
function parseManifestExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const jsonData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: ''
  });

  const getCellValue = (row, col) => {
    if (row < 0 || row >= jsonData.length) return '';
    const rowData = jsonData[row];
    if (!rowData || col < 0 || col >= rowData.length) return '';
    return String(rowData[col] || '').trim();
  };

  const data = {
    船名: getCellValue(3, 1),
    航次: getCellValue(3, 4),
    目的港: getCellValue(3, 7),
    提单号: getCellValue(4, 1),
    箱号: getCellValue(12, 0),
    封号: getCellValue(12, 1),
    箱型: getCellValue(12, 2),
    英文品名: getCellValue(20, 4),
    件数: getCellValue(20, 6),
    包装单位: getCellValue(20, 7),
    毛重: getCellValue(20, 8),
    体积: getCellValue(20, 9),
    唛头: getCellValue(20, 10),
  };

  // 调试日志
  console.log('DEBUG parseManifestExcel: 英文品名原始值:', JSON.stringify(data.英文品名));
  console.log('DEBUG parseManifestExcel: 解析后商品列表长度:', data.英文品名 ? data.英文品名.split(',').map(s => s.trim()).filter(item => item !== '').length : 0);

  return data;
}

// 安全转换为整数：如果是数字或数字字符串则转换为整数，否则返回空字符串
function safeToInt(value) {
  if (value === '' || value === null || value === undefined) {
    return '';
  }
  // 尝试转换为数字
  const num = Number(value);
  // 检查是否为有效数字且不是NaN
  if (!isNaN(num) && isFinite(num)) {
    // 转换为整数
    return Math.floor(num).toString();
  }
  // 如果不是有效数字，返回空字符串
  return '';
}

// 更新求和公式（数据行范围：15-21行，求和行：第22行）
function updateSumFormulasAfterRowDeletion(worksheet, deletedRows) {
  // 计算在数据行范围（15-21）内删除了多少行
  const dataRowStart = 15;
  const dataRowEnd = 21;
  const originalSumRow = 22; // 原始求和公式所在行

  // 统计在数据行范围内删除了多少行
  let deletedInDataRange = 0;
  // 统计在求和行之前删除了多少行（行号小于originalSumRow）
  let deletedBeforeSum = 0;

  deletedRows.forEach(rowNum => {
    if (rowNum >= dataRowStart && rowNum <= dataRowEnd) {
      deletedInDataRange++;
    }
    if (rowNum < originalSumRow) {
      deletedBeforeSum++;
    }
  });

  if (deletedInDataRange === 0) {
    // 没有在数据行范围内删除行，无需更新公式
    return;
  }

  // 计算新的结束行
  const newDataRowEnd = dataRowEnd - deletedInDataRange;
  // 计算新的求和行号（原始行号减去之前删除的行数）
  const newSumRow = originalSumRow - deletedBeforeSum;

  // 更新求和公式
  // 公式列：C列、E列、G列
  const formulaColumns = ['C', 'E', 'G'];

  formulaColumns.forEach(col => {
    const originalCellAddress = `${col}${originalSumRow}`;
    const newCellAddress = `${col}${newSumRow}`;
    const cell = worksheet.getCell(newCellAddress);

    // 如果新单元格没有公式，尝试查找包含公式的单元格
    if (!cell.formula) {
      // 可能公式在其他行，尝试搜索
      console.log(`单元格${newCellAddress}没有公式，尝试查找公式...`);
      // 简单起见，我们假设公式就在新行
      return;
    }

    // 解析原始公式，更新引用范围
    const originalFormula = cell.formula;
    // 预期公式格式：SUM(C15:C21)
    const regex = new RegExp(`SUM\\(${col}(\\d+):${col}(\\d+)\\)`, 'i');
    const match = originalFormula.match(regex);

    if (match) {
      const startRow = parseInt(match[1], 10);
      const endRow = parseInt(match[2], 10);

      // 检查公式是否符合预期格式
      if (startRow === dataRowStart && endRow === dataRowEnd) {
        // 更新公式
        const newFormula = `SUM(${col}${dataRowStart}:${col}${newDataRowEnd})`;
        cell.value = {
          formula: newFormula,
          result: null // 清除缓存结果，让Excel重新计算
        };
        console.log(`更新公式：${newCellAddress} = ${newFormula}（原公式：${originalFormula}，删除了${deletedInDataRange}行，求和行从${originalSumRow}移动到${newSumRow}）`);
      } else {
        // 公式格式不匹配，但可能已经被调整过，尝试更新为新的结束行
        // 保持起始行不变，更新结束行
        const newFormula = `SUM(${col}${startRow}:${col}${newDataRowEnd})`;
        cell.value = {
          formula: newFormula,
          result: null
        };
        console.log(`更新公式（调整）：${newCellAddress} = ${newFormula}（原公式：${originalFormula}，删除了${deletedInDataRange}行）`);
      }
    } else {
      console.log(`无法解析公式：${newCellAddress} = ${originalFormula}`);
    }
  });
}

// 生成 Word 文档
async function generateWordDocument(data) {
  const templatePath = path.join(__dirname, '../templates/提单确认件的格式.docx');
  const templateBuffer = await fs.readFile(templatePath);

  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // 商品列表 - 严格按舱单文件中的英文品名数量处理
  const englishNames = data.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const goodsData = {};
  for (let i = 1; i <= 22; i++) {
    // 只使用舱单文件中存在的商品，不存在则设置为空字符串
    goodsData[`商品${i}`] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  // 调试日志
  console.log('DEBUG Word生成: 英文品名原始值:', JSON.stringify(data.英文品名));
  console.log('DEBUG Word生成: 解析后商品列表:', JSON.stringify(goodsList));
  console.log('DEBUG Word生成: 商品数据:', JSON.stringify(goodsData));

  doc.setData({
    船名: data.船名,
    航次: data.航次,
    目的港: data.目的港,
    提单号: data.提单号,
    箱号: data.箱号,
    封号: data.封号,
    箱型: data.箱型,
    件数: safeToInt(data.件数),
    毛重: safeToInt(data.毛重),
    体积: safeToInt(data.体积),
    ...goodsData,
  });

  doc.render();
  return doc.getZip().generate({ type: 'nodebuffer' });
}

// 生成 Excel 文档
async function generateExcelDocument(data) {
  const templatePath = path.join(__dirname, '../templates/装箱单发票的格式.xlsx');
  const templateBuffer = await fs.readFile(templatePath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  if (workbook.worksheets.length === 0) {
    throw new Error('无法加载 Excel 模板');
  }

  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

  const replacePlaceholder = (cell, placeholder, replacement) => {
    const text = getCellText(cell);
    if (text.includes(placeholder)) {
      // 检查替换值是否为数字（整数）
      if (replacement !== '' && !isNaN(Number(replacement)) && replacement !== null && replacement !== undefined) {
        // 设置为数字类型
        cell.value = Number(replacement);
        // 设置数字格式为整数（无小数）
        cell.numFmt = '0';
        // 设置单元格类型为数字
        cell.type = 'n';
      } else {
        // 非数字值，保持原有格式
        let font = {};
        if (cell.value?.richText && cell.value.richText.length > 0) {
          font = cell.value.richText[0].font || {};
        }
        cell.value = {
          richText: [{
            font: font,
            text: replacement || ''
          }]
        };
      }
      return true;
    }
    return false;
  };

  // 准备替换数据 - 严格按舱单文件中的英文品名数量处理
  const englishNames = data.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const replacementData = {
    '{发票日期}': formattedDate
  };

  // 添加商品占位符替换数据 - 只使用舱单文件中存在的商品
  for (let i = 1; i <= 22; i++) {
    const placeholder = `{商品${i}}`;
    replacementData[placeholder] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  console.log('替换数据:', {
    英文品名: data.英文品名,
    商品列表长度: goodsList.length,
    商品列表内容: goodsList,
    占位符数量: Object.keys(replacementData).length,
    占位符列表: Object.keys(replacementData)
  });

  // 处理所有 sheet
  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    let replacedCount = 0;
    // 遍历所有行和单元格
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        // 尝试替换所有可能的占位符
        for (const [placeholder, replacement] of Object.entries(replacementData)) {
          if (replacePlaceholder(cell, placeholder, replacement)) {
            replacedCount++;
          }
        }
      });
    });
    console.log(`Sheet ${sheetIndex + 1} "${worksheet.name}" 替换了 ${replacedCount} 个占位符`);
  });

  // 专门处理 PACKING LIST 工作表的 B11-B32 单元格
  const packingListSheet = workbook.getWorksheet('PACKING LIST') || workbook.worksheets[1];
  if (packingListSheet) {
    console.log(`专门处理 PACKING LIST 工作表: ${packingListSheet.name}`);
    let specificReplaced = 0;
    for (let row = 11; row <= 32; row++) {
      const cell = packingListSheet.getCell(`B${row}`);
      const placeholderIndex = row - 10; // B11 对应商品1, B12 对应商品2...
      const placeholder = `{商品${placeholderIndex}}`;
      const replacement = replacementData[placeholder] || '';

      if (replacePlaceholder(cell, placeholder, replacement)) {
        specificReplaced++;
      } else {
        // 如果没找到占位符，直接设置单元格值
        const cellText = getCellText(cell);
        if (cellText.includes('{商品')) {
          // 单元格包含其他商品占位符，尝试替换所有可能的占位符
          for (const [ph, repl] of Object.entries(replacementData)) {
            if (replacePlaceholder(cell, ph, repl)) {
              specificReplaced++;
              break;
            }
          }
        }
      }
    }
    console.log(`PACKING LIST 工作表 B11-B32 替换了 ${specificReplaced} 个单元格`);
  }

  return workbook.xlsx.writeBuffer();
}

// 生成并单保函 Word 文档
async function generateCombinedLetter(firstData, allCargoData) {
  const templatePath = path.join(__dirname, '../templates/并单保函的格式.docx');
  const templateBuffer = await fs.readFile(templatePath);

  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: function() {
      return '';
    }
  });

  // 商品列表 - 严格按舱单文件中的英文品名数量处理（使用第一个文件的数据）
  const englishNames = firstData.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const goodsData = {};
  for (let i = 1; i <= 22; i++) {
    // 只使用舱单文件中存在的商品，不存在则设置为空字符串
    goodsData[`商品${i}`] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  // 生成所有舱单字段映射：提单号1, 箱号1, 箱型1, 封号1, 件数1, 毛重1, 体积1, 并单号1, ...
  const containerData = {};
  const maxContainers = 20; // 假设模板最多支持20个舱单

  for (let i = 0; i < maxContainers; i++) {
    const suffix = i + 1;
    if (i < allCargoData.length) {
      const cargo = allCargoData[i];
      const billNumber = cargo.提单号 || '';
      // 如果提单号为空，设置字段为空字符串（清空整行）
      if (billNumber === '') {
        containerData[`提单号${suffix}`] = '';
        containerData[`箱号${suffix}`] = '';
        containerData[`箱型${suffix}`] = '';
        containerData[`封号${suffix}`] = '';
        containerData[`件数${suffix}`] = '';
        containerData[`毛重${suffix}`] = '';
        containerData[`体积${suffix}`] = '';
        containerData[`并单号${suffix}`] = ''; // 清空并单号字段
      } else {
        containerData[`提单号${suffix}`] = billNumber;
        containerData[`箱号${suffix}`] = cargo.箱号 || '';
        containerData[`箱型${suffix}`] = cargo.箱型 || '';
        containerData[`封号${suffix}`] = cargo.封号 || '';
        containerData[`件数${suffix}`] = safeToInt(cargo.件数);
        containerData[`毛重${suffix}`] = safeToInt(cargo.毛重);
        containerData[`体积${suffix}`] = safeToInt(cargo.体积);
        // 如果当前舱单的提单号不为空，并单号等于第一个舱单的提单号
        containerData[`并单号${suffix}`] = firstData.提单号 || '';
      }
    } else {
      // 没有更多舱单数据
      if (suffix >= 2) {
        // 对于提单号2及以上，设置字段为空字符串（清空整行）
        containerData[`提单号${suffix}`] = '';
        containerData[`箱号${suffix}`] = '';
        containerData[`箱型${suffix}`] = '';
        containerData[`封号${suffix}`] = '';
        containerData[`件数${suffix}`] = '';
        containerData[`毛重${suffix}`] = '';
        containerData[`体积${suffix}`] = '';
        containerData[`并单号${suffix}`] = ''; // 清空并单号字段
      } else {
        // 提单号1，设置空字符串
        containerData[`提单号${suffix}`] = '';
        containerData[`箱号${suffix}`] = '';
        containerData[`箱型${suffix}`] = '';
        containerData[`封号${suffix}`] = '';
        containerData[`件数${suffix}`] = '';
        containerData[`毛重${suffix}`] = '';
        containerData[`体积${suffix}`] = '';
        containerData[`并单号${suffix}`] = ''; // 清空并单号字段
      }
    }
  }

  console.log('并单保函舱单数据（前3个）:', {
    提单号1: containerData['提单号1'],
    箱号1: containerData['箱号1'],
    箱型1: containerData['箱型1'],
    封号1: containerData['封号1'],
    件数1: containerData['件数1'],
    毛重1: containerData['毛重1'],
    体积1: containerData['体积1'],
  });

  doc.setData({
    船名: firstData.船名,
    航次: firstData.航次,
    目的港: firstData.目的港,
    提单号: firstData.提单号,
    箱号: firstData.箱号,
    封号: firstData.封号,
    箱型: firstData.箱型,
    件数: safeToInt(firstData.件数),
    毛重: safeToInt(firstData.毛重),
    体积: safeToInt(firstData.体积),
    并单号: firstData.提单号, // 新增并单号占位符
    ...goodsData,
    ...containerData,
  });

  doc.render();
  return doc.getZip().generate({ type: 'nodebuffer' });
}

// 生成总提单OK件（带HS）Excel 文档
async function generateOKBillWithHS(firstData, allCargoData, hsCodeMap = null) {
  const templatePath = path.join(__dirname, '../templates/总提单OK件的格式(带HS的.xlsx');
  const templateBuffer = await fs.readFile(templatePath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  if (workbook.worksheets.length === 0) {
    throw new Error('无法加载 Excel 模板');
  }

  // 使用与现有 Excel 生成相同的替换逻辑
  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

  const replacePlaceholder = (cell, placeholder, replacement) => {
    const text = getCellText(cell);
    if (text.includes(placeholder)) {
      // 检查替换值是否为数字（整数）
      if (replacement !== '' && !isNaN(Number(replacement)) && replacement !== null && replacement !== undefined) {
        // 设置为数字类型
        cell.value = Number(replacement);
        // 设置数字格式为整数（无小数）
        cell.numFmt = '0';
        // 设置单元格类型为数字
        cell.type = 'n';
      } else {
        // 非数字值，保持原有格式
        let font = {};
        if (cell.value?.richText && cell.value.richText.length > 0) {
          font = cell.value.richText[0].font || {};
        }
        cell.value = {
          richText: [{
            font: font,
            text: replacement || ''
          }]
        };
      }
      return true;
    }
    return false;
  };

  // 准备替换数据 - 严格按舱单文件中的英文品名数量处理
  const englishNames = firstData.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const replacementData = {
    '{发票日期}': formattedDate,
    '{船名}': firstData.船名 || '',
    '{航次}': firstData.航次 || '',
    '{目的港}': firstData.目的港 || '',
    '{提单号}': firstData.提单号 || '',
    '{箱号}': firstData.箱号 || '',
    '{封号}': firstData.封号 || '',
    '{箱型}': firstData.箱型 || '',
    '{件数}': safeToInt(firstData.件数),
    '{毛重}': safeToInt(firstData.毛重),
    '{体积}': safeToInt(firstData.体积),
    '{并单号}': firstData.提单号 || '', // 新增并单号占位符
  };

  // 添加商品占位符替换数据 - 只使用舱单文件中存在的商品
  for (let i = 1; i <= 22; i++) {
    const placeholder = `{商品${i}}`;
    replacementData[placeholder] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  // 添加所有舱单字段映射：提单号1, 箱号1, 箱型1, 封号1, 件数1, 毛重1, 体积1, 并单号1, ...
  const maxContainers = 20; // 假设模板最多支持20个舱单
  for (let i = 0; i < maxContainers; i++) {
    const suffix = i + 1;
    if (i < allCargoData.length) {
      const cargo = allCargoData[i];
      replacementData[`{提单号${suffix}}`] = cargo.提单号 || '';
      replacementData[`{箱号${suffix}}`] = cargo.箱号 || '';
      replacementData[`{箱型${suffix}}`] = cargo.箱型 || '';
      replacementData[`{封号${suffix}}`] = cargo.封号 || '';
      replacementData[`{件数${suffix}}`] = safeToInt(cargo.件数);
      replacementData[`{毛重${suffix}}`] = safeToInt(cargo.毛重);
      replacementData[`{体积${suffix}}`] = safeToInt(cargo.体积);
      // 如果当前舱单的提单号不为空，并单号等于第一个舱单的提单号；否则为空
      replacementData[`{并单号${suffix}}`] = cargo.提单号 ? (firstData.提单号 || '') : '';
    } else {
      // 填充空的占位符
      replacementData[`{提单号${suffix}}`] = '';
      replacementData[`{箱号${suffix}}`] = '';
      replacementData[`{箱型${suffix}}`] = '';
      replacementData[`{封号${suffix}}`] = '';
      replacementData[`{件数${suffix}}`] = '';
      replacementData[`{毛重${suffix}}`] = '';
      replacementData[`{体积${suffix}}`] = '';
      replacementData[`{并单号${suffix}}`] = ''; // 清空并单号字段
    }
  }

  // 为每个舱单生成带HS的商品列表
  const cargoListsWithHS = [];
  for (let i = 0; i < allCargoData.length; i++) {
    const cargo = allCargoData[i];
    const englishNames = cargo.英文品名 || '';
    const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');

    // 使用HS编码映射表或默认值
    const goodsWithHS = goodsList.map((goods) => {
      const hsCode = hsCodeMap && hsCodeMap[goods] ? hsCodeMap[goods] : '88886666';
      return `${goods} ${hsCode}`;
    });

    // 用逗号连接成字符串
    const cargoListString = goodsWithHS.join(', ');
    cargoListsWithHS.push(cargoListString);
    // 添加到替换数据
    replacementData[`{带HS的商品列表${i + 1}}`] = cargoListString;
  }

  console.log('总提单OK件（带HS）替换数据:', {
    提单号: firstData.提单号,
    商品列表长度: goodsList.length,
    商品列表内容: goodsList,
    提单号总数: allCargoData.length,
    所有提单号: allCargoData.map(d => d.提单号),
    带HS的商品列表数量: cargoListsWithHS.length,
    带HS的商品列表: cargoListsWithHS,
  });

  // 处理所有 sheet
  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    let replacedCount = 0;
    const rowsToDelete = new Set();

    // 第一遍：扫描并标记需要删除的行
    worksheet.eachRow((row, rowNumber) => {
      let hasEmptyBillNumber = false;
      let hasFormula = false;

      row.eachCell((cell) => {
        // 检查单元格是否包含公式
        if (cell.formula) {
          hasFormula = true;
        }

        const cellText = getCellText(cell);
        // 检查是否包含提单号占位符，且不是提单号1
        const billMatch = cellText.match(/\{提单号(\d+)\}/);
        if (billMatch) {
          const billNum = parseInt(billMatch[1], 10);
          if (billNum >= 2) {
            const placeholder = billMatch[0];
            const replacement = replacementData[placeholder];
            // 如果替换值为空，标记该行删除
            if (replacement === '') {
              hasEmptyBillNumber = true;
            }
          }
        }
      });

      // 如果行包含公式或行号是22，不标记删除
      if (hasEmptyBillNumber && !hasFormula && rowNumber !== 22) {
        rowsToDelete.add(rowNumber);
      }
    });

    // 第二遍：替换占位符（跳过D13单元格，它有特殊的富文本处理）
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        // 跳过D13单元格，它后面有特殊处理
        if (rowNumber === 13 && cell.address === 'D13') {
          return;
        }
        for (const [placeholder, replacement] of Object.entries(replacementData)) {
          if (replacePlaceholder(cell, placeholder, replacement)) {
            replacedCount++;
          }
        }
      });
    });

    // 清空标记的行（不删除行，只将整行单元格设置为空）
    rowsToDelete.forEach(rowNumber => {
      const row = worksheet.getRow(rowNumber);
      row.eachCell(cell => {
        cell.value = null;
        cell.formula = null;
        cell.numFmt = null;
        cell.type = null;
      });
      console.log(`总提单OK件（带HS） Sheet ${sheetIndex + 1}: 清空第 ${rowNumber} 行（提单号为空）`);
    });

    // 替换D13单元格中的商品列表占位符，保留原始格式
    const goodsListCell = worksheet.getCell('D13');
    if (goodsListCell.value && goodsListCell.value.richText) {
      const originalRichText = goodsListCell.value.richText;

      console.log(`总提单OK件（带HS）D13片段数: ${originalRichText.length}`);
      console.log(`总提单OK件（带HS）商品列表数量: ${cargoListsWithHS.length}`);

      // 创建新的富文本，按照指定格式：第1,3,5,7个片段红色，第2,4,6个片段黑色，字体Times New Roman
      const newRichText = [];
      let hasContent = false; // 标记前面是否已经有内容

      for (let i = 0; i < 7; i++) {
        const cargoList = cargoListsWithHS[i];

        // 跳过空的商品列表
        if (!cargoList || cargoList.trim() === '') {
          console.log(`总提单OK件（带HS）片段${i + 1}: "(空，跳过)"`);
          continue;
        }

        // 确定颜色：偶数索引(0,2,4,6)红色，奇数索引(1,3,5)黑色
        const isRed = i % 2 === 0; // 0-based: 0,2,4,6 是红色（对应第1,3,5,7个）
        const color = isRed ? 'FF0000' : '000000'; // 红色: FF0000, 黑色: 000000

        // 字体配置
        const font = {
          name: 'Times New Roman',
          color: { argb: color }
        };

        // 如果前面已经有内容，先添加逗号分隔符
        if (hasContent) {
          // 逗号用黑色，Times New Roman字体
          newRichText.push({
            font: {
              name: 'Times New Roman',
              color: { argb: '000000' } // 黑色逗号
            },
            text: ', '
          });
        }

        // 添加商品列表片段
        newRichText.push({
          font: font,
          text: cargoList
        });

        hasContent = true;
        console.log(`总提单OK件（带HS）片段${i + 1}: "${cargoList}" (颜色: ${isRed ? '红色' : '黑色'})`);
      }

      goodsListCell.value = { richText: newRichText };
      console.log(`总提单OK件（带HS）D13单元格已设置 ${newRichText.length} 个片段，${hasContent ? '有内容' : '无内容'}`);
    }

    // 更新第22行的求和公式（数据行范围：15-21行）
    // 注意：现在不删除行，只清空行内容，因此不需要更新公式
    // updateSumFormulasAfterRowDeletion(worksheet, rowsToDelete);

    console.log(`总提单OK件（带HS） Sheet ${sheetIndex + 1} "${worksheet.name}" 替换了 ${replacedCount} 个占位符，清空了 ${rowsToDelete.size} 行，设置了 ${cargoListsWithHS.length} 个商品列表`);
  });

  return workbook.xlsx.writeBuffer();
}

// 生成总提单OK件（无HS）Excel 文档
async function generateOKBillWithoutHS(firstData, allCargoData) {
  const templatePath = path.join(__dirname, '../templates/总提单OK件的格式(无HS的.xlsx');
  const templateBuffer = await fs.readFile(templatePath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  if (workbook.worksheets.length === 0) {
    throw new Error('无法加载 Excel 模板');
  }

  // 使用与现有 Excel 生成相同的替换逻辑
  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

  const replacePlaceholder = (cell, placeholder, replacement) => {
    const text = getCellText(cell);
    if (text.includes(placeholder)) {
      // 检查替换值是否为数字（整数）
      if (replacement !== '' && !isNaN(Number(replacement)) && replacement !== null && replacement !== undefined) {
        // 设置为数字类型
        cell.value = Number(replacement);
        // 设置数字格式为整数（无小数）
        cell.numFmt = '0';
        // 设置单元格类型为数字
        cell.type = 'n';
      } else {
        // 非数字值，保持原有格式
        let font = {};
        if (cell.value?.richText && cell.value.richText.length > 0) {
          font = cell.value.richText[0].font || {};
        }
        cell.value = {
          richText: [{
            font: font,
            text: replacement || ''
          }]
        };
      }
      return true;
    }
    return false;
  };

  // 准备替换数据 - 严格按舱单文件中的英文品名数量处理
  const englishNames = firstData.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const replacementData = {
    '{发票日期}': formattedDate,
    '{船名}': firstData.船名 || '',
    '{航次}': firstData.航次 || '',
    '{目的港}': firstData.目的港 || '',
    '{提单号}': firstData.提单号 || '',
    '{箱号}': firstData.箱号 || '',
    '{封号}': firstData.封号 || '',
    '{箱型}': firstData.箱型 || '',
    '{件数}': safeToInt(firstData.件数),
    '{毛重}': safeToInt(firstData.毛重),
    '{体积}': safeToInt(firstData.体积),
    '{并单号}': firstData.提单号 || '', // 新增并单号占位符
  };

  // 添加商品占位符替换数据 - 只使用舱单文件中存在的商品
  for (let i = 1; i <= 22; i++) {
    const placeholder = `{商品${i}}`;
    replacementData[placeholder] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  // 添加所有舱单字段映射：提单号1, 箱号1, 箱型1, 封号1, 件数1, 毛重1, 体积1, 并单号1, ...
  const maxContainers = 20; // 假设模板最多支持20个舱单
  for (let i = 0; i < maxContainers; i++) {
    const suffix = i + 1;
    if (i < allCargoData.length) {
      const cargo = allCargoData[i];
      replacementData[`{提单号${suffix}}`] = cargo.提单号 || '';
      replacementData[`{箱号${suffix}}`] = cargo.箱号 || '';
      replacementData[`{箱型${suffix}}`] = cargo.箱型 || '';
      replacementData[`{封号${suffix}}`] = cargo.封号 || '';
      replacementData[`{件数${suffix}}`] = safeToInt(cargo.件数);
      replacementData[`{毛重${suffix}}`] = safeToInt(cargo.毛重);
      replacementData[`{体积${suffix}}`] = safeToInt(cargo.体积);
      // 如果当前舱单的提单号不为空，并单号等于第一个舱单的提单号；否则为空
      replacementData[`{并单号${suffix}}`] = cargo.提单号 ? (firstData.提单号 || '') : '';
    } else {
      // 填充空的占位符
      replacementData[`{提单号${suffix}}`] = '';
      replacementData[`{箱号${suffix}}`] = '';
      replacementData[`{箱型${suffix}}`] = '';
      replacementData[`{封号${suffix}}`] = '';
      replacementData[`{件数${suffix}}`] = '';
      replacementData[`{毛重${suffix}}`] = '';
      replacementData[`{体积${suffix}}`] = '';
      replacementData[`{并单号${suffix}}`] = ''; // 清空并单号字段
    }
  }

  // 为每个舱单生成无HS的商品列表
  const cargoListsWithoutHS = [];
  for (let i = 0; i < allCargoData.length; i++) {
    const cargo = allCargoData[i];
    const englishNames = cargo.英文品名 || '';
    const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
    // 用逗号连接成字符串
    const cargoListString = goodsList.join(', ');
    cargoListsWithoutHS.push(cargoListString);
    // 添加到替换数据
    replacementData[`{无HS的商品列表${i + 1}}`] = cargoListString;
  }

  console.log('总提单OK件（无HS）替换数据:', {
    提单号: firstData.提单号,
    商品列表长度: goodsList.length,
    商品列表内容: goodsList,
    提单号总数: allCargoData.length,
    所有提单号: allCargoData.map(d => d.提单号),
    无HS的商品列表数量: cargoListsWithoutHS.length,
    无HS的商品列表: cargoListsWithoutHS,
  });

  // 处理所有 sheet
  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    let replacedCount = 0;
    const rowsToDelete = new Set();

    // 第一遍：扫描并标记需要删除的行
    worksheet.eachRow((row, rowNumber) => {
      let hasEmptyBillNumber = false;
      let hasFormula = false;

      row.eachCell((cell) => {
        // 检查单元格是否包含公式
        if (cell.formula) {
          hasFormula = true;
        }

        const cellText = getCellText(cell);
        // 检查是否包含提单号占位符，且不是提单号1
        const billMatch = cellText.match(/\{提单号(\d+)\}/);
        if (billMatch) {
          const billNum = parseInt(billMatch[1], 10);
          if (billNum >= 2) {
            const placeholder = billMatch[0];
            const replacement = replacementData[placeholder];
            // 如果替换值为空，标记该行删除
            if (replacement === '') {
              hasEmptyBillNumber = true;
            }
          }
        }
      });

      // 如果行包含公式或行号是22，不标记删除
      if (hasEmptyBillNumber && !hasFormula && rowNumber !== 22) {
        rowsToDelete.add(rowNumber);
      }
    });

    // 第二遍：替换占位符（跳过D13单元格，它有特殊的富文本处理）
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        // 跳过D13单元格，它后面有特殊处理
        if (rowNumber === 13 && cell.address === 'D13') {
          return;
        }
        for (const [placeholder, replacement] of Object.entries(replacementData)) {
          if (replacePlaceholder(cell, placeholder, replacement)) {
            replacedCount++;
          }
        }
      });
    });

    // 清空标记的行（不删除行，只将整行单元格设置为空）
    rowsToDelete.forEach(rowNumber => {
      const row = worksheet.getRow(rowNumber);
      row.eachCell(cell => {
        cell.value = null;
        cell.formula = null;
        cell.numFmt = null;
        cell.type = null;
      });
      console.log(`总提单OK件（无HS） Sheet ${sheetIndex + 1}: 清空第 ${rowNumber} 行（提单号为空）`);
    });

    // 替换D13单元格中的商品列表占位符，保留原始格式
    const goodsListCell = worksheet.getCell('D13');
    if (goodsListCell.value && goodsListCell.value.richText) {
      const originalRichText = goodsListCell.value.richText;

      console.log(`总提单OK件（无HS）D13片段数: ${originalRichText.length}`);
      console.log(`总提单OK件（无HS）商品列表数量: ${cargoListsWithoutHS.length}`);

      // 创建新的富文本，按照指定格式：第1,3,5,7个片段红色，第2,4,6个片段黑色，字体Times New Roman
      const newRichText = [];
      let hasContent = false; // 标记前面是否已经有内容

      for (let i = 0; i < 7; i++) {
        const cargoList = cargoListsWithoutHS[i];

        // 跳过空的商品列表
        if (!cargoList || cargoList.trim() === '') {
          console.log(`总提单OK件（无HS）片段${i + 1}: "(空，跳过)"`);
          continue;
        }

        // 确定颜色：偶数索引(0,2,4,6)红色，奇数索引(1,3,5)黑色
        const isRed = i % 2 === 0; // 0-based: 0,2,4,6 是红色（对应第1,3,5,7个）
        const color = isRed ? 'FF0000' : '000000'; // 红色: FF0000, 黑色: 000000

        // 字体配置
        const font = {
          name: 'Times New Roman',
          color: { argb: color }
        };

        // 如果前面已经有内容，先添加逗号分隔符
        if (hasContent) {
          // 逗号用黑色，Times New Roman字体
          newRichText.push({
            font: {
              name: 'Times New Roman',
              color: { argb: '000000' } // 黑色逗号
            },
            text: ', '
          });
        }

        // 添加商品列表片段
        newRichText.push({
          font: font,
          text: cargoList
        });

        hasContent = true;
        console.log(`总提单OK件（无HS）片段${i + 1}: "${cargoList}" (颜色: ${isRed ? '红色' : '黑色'})`);
      }

      goodsListCell.value = { richText: newRichText };
      console.log(`总提单OK件（无HS）D13单元格已设置 ${newRichText.length} 个片段，${hasContent ? '有内容' : '无内容'}`);
    }

    // 更新第22行的求和公式（数据行范围：15-21行）
    // 注意：现在不删除行，只清空行内容，因此不需要更新公式
    // updateSumFormulasAfterRowDeletion(worksheet, rowsToDelete);

    console.log(`总提单OK件（无HS） Sheet ${sheetIndex + 1} "${worksheet.name}" 替换了 ${replacedCount} 个占位符，清空了 ${rowsToDelete.size} 行，设置了 ${cargoListsWithoutHS.length} 个商品列表`);
  });

  return workbook.xlsx.writeBuffer();
}

// 生成带HS的汇总单 Word 文档
async function generateSummaryWithHS(firstData, allCargoData, hsCodeMap = null) {
  const templatePath = path.join(__dirname, '../templates/带HS的汇总单.docx');
  const templateBuffer = await fs.readFile(templatePath);

  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    nullGetter: function() {
      return '';
    }
  });

  // 为每个舱单生成带HS的商品列表（商品之间用换行符分隔）
  const cargoListsWithHS = [];
  for (let i = 0; i < allCargoData.length; i++) {
    const cargo = allCargoData[i];
    const englishNames = cargo.英文品名 || '';
    const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');

    // 使用HS编码映射表或默认值
    const goodsWithHS = goodsList.map((goods) => {
      const hsCode = hsCodeMap && hsCodeMap[goods] ? hsCodeMap[goods] : '88886666';
      return `${goods} ${hsCode}`;
    });

    // 各个商品之间用换行符分隔
    const cargoListString = goodsWithHS.join('\n');
    cargoListsWithHS.push(cargoListString);
  }

  // 商品列表数据（最多7个）
  const goodsListData = {};
  for (let i = 1; i <= 7; i++) {
    goodsListData[`带HS的商品列表${i}`] = i <= cargoListsWithHS.length ? cargoListsWithHS[i - 1] : '';
  }

  // 计算总件数、总毛重、总体积
  let totalPieces = 0;
  let totalWeight = 0;
  let totalVolume = 0;
  allCargoData.forEach(cargo => {
    totalPieces += Number(cargo.件数) || 0;
    totalWeight += Number(cargo.毛重) || 0;
    totalVolume += Number(cargo.体积) || 0;
  });

  // 舱单字段映射（最多7个）
  const containerData = {};
  for (let i = 0; i < 7; i++) {
    const suffix = i + 1;
    if (i < allCargoData.length) {
      const cargo = allCargoData[i];
      containerData[`箱号${suffix}`] = cargo.箱号 || '';
      containerData[`封号${suffix}`] = cargo.封号 || '';
      containerData[`箱型${suffix}`] = cargo.箱型 || '';
      containerData[`件数${suffix}`] = safeToInt(cargo.件数);
      containerData[`毛重${suffix}`] = safeToInt(cargo.毛重);
      containerData[`体积${suffix}`] = safeToInt(cargo.体积);
      containerData[`提单号${suffix}`] = cargo.提单号 || '';
      // 提单号为空时，单位也设为空
      if (!cargo.提单号) {
        containerData[`数量单位${suffix}`] = '';
        containerData[`重量单位${suffix}`] = '';
        containerData[`体积单位${suffix}`] = '';
      } else {
        containerData[`数量单位${suffix}`] = 'CTNS';
        containerData[`重量单位${suffix}`] = 'KGS';
        containerData[`体积单位${suffix}`] = 'CBM/';
      }
    } else {
      containerData[`箱号${suffix}`] = '';
      containerData[`封号${suffix}`] = '';
      containerData[`箱型${suffix}`] = '';
      containerData[`件数${suffix}`] = '';
      containerData[`毛重${suffix}`] = '';
      containerData[`体积${suffix}`] = '';
      containerData[`提单号${suffix}`] = '';
      containerData[`数量单位${suffix}`] = '';
      containerData[`重量单位${suffix}`] = '';
      containerData[`体积单位${suffix}`] = '';
    }
  }

  doc.setData({
    船名: firstData.船名,
    航次: firstData.航次,
    目的港: firstData.目的港,
    提单号: firstData.提单号,
    总件数: totalPieces.toString(),
    总毛重: totalWeight.toString(),
    总体积: totalVolume.toString(),
    ...goodsListData,
    ...containerData,
  });

  doc.render();
  return doc.getZip().generate({ type: 'nodebuffer' });
}

// Buffer to base64
function bufferToBase64(buffer) {
  return buffer.toString('base64');
}

module.exports = async (req, res) => {
  // Handle CORS preflight
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  try {
    const formData = await parseFormData(req);

    const manifestFiles = formData.files.manifest;
    if (!manifestFiles || (Array.isArray(manifestFiles) && manifestFiles.length === 0) || manifestFiles.length === 0) {
      return res.status(400).json({ success: false, message: '请上传舱单文件' });
    }

    // 确保是数组
    const files = Array.isArray(manifestFiles) ? manifestFiles : [manifestFiles];

    // 解析HS编码映射表
    let hsCodeMap = null;
    const hsCodeMapField = formData.fields.hsCodeMap;
    if (hsCodeMapField) {
      try {
        hsCodeMap = JSON.parse(hsCodeMapField);
        console.log(`收到HS编码映射表，共 ${Object.keys(hsCodeMap).length} 条记录`);
      } catch (e) {
        console.error('解析HS编码映射表失败:', e.message);
      }
    }

    console.log(`开始批量处理 ${files.length} 个文件`);

    // 收集所有舱单数据用于生成汇总文件
    const allCargoData = [];
    let firstCargoData = null;

    // 创建 ZIP 归档
    const archive = archiver('zip', { zlib: { level: 9 } });
    const chunks = [];
    archive.on('data', (chunk) => chunks.push(chunk));

    const zipPromise = new Promise((resolve, reject) => {
      archive.on('end', () => {
        const zipBuffer = Buffer.concat(chunks);
        resolve(zipBuffer);
      });
      archive.on('error', reject);
    });

    // 处理每个文件
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      console.log(`处理第 ${i + 1} 个文件: ${file.filename || '未命名文件'}`);

      try {
        const cargoData = parseManifestExcel(file.buffer);

        // 收集数据用于汇总文件
        allCargoData.push(cargoData);
        if (i === 0) {
          firstCargoData = cargoData;
        }

        const wordBuffer = await generateWordDocument(cargoData);
        const excelBuffer = await generateExcelDocument(cargoData);

        // 生成安全文件名
        const safeBillName = cargoData.提单号 ? cargoData.提单号.replace(/[^a-zA-Z0-9]/g, '_') : `bill_${i + 1}`;
        const safeContainerName = cargoData.箱号 ? cargoData.箱号.replace(/[^a-zA-Z0-9]/g, '_') : `container_${i + 1}`;

        // 添加到 ZIP，按文件夹结构组织
        archive.append(wordBuffer, { name: `A/B/${safeBillName}.docx` });
        archive.append(excelBuffer, { name: `A/C/${safeContainerName}.xlsx` });

        console.log(`文件 ${i + 1} 处理完成: 提单号=${cargoData.提单号}, 箱号=${cargoData.箱号}`);
        console.log(`  文件名: B/${safeBillName}.docx, C/${safeContainerName}.xlsx`);
      } catch (fileError) {
        console.error(`处理文件 ${i + 1} 失败:`, fileError);
        // 跳过失败的文件，继续处理其他文件
        continue;
      }
    }

    // 生成三个汇总文件（使用第一个文件的数据）
    if (firstCargoData) {
      try {
        const safeBillNumber = firstCargoData.提单号 ? firstCargoData.提单号.replace(/[^a-zA-Z0-9]/g, '_') : '汇总';

        // 生成并单保函
        const combinedLetterBuffer = await generateCombinedLetter(firstCargoData, allCargoData);
        archive.append(combinedLetterBuffer, { name: `A/${safeBillNumber}并单保函的格式.docx` });
        console.log(`生成汇总文件: A/${safeBillNumber}并单保函的格式.docx`);

        // 生成总提单OK件（带HS）
        const okWithHSBuffer = await generateOKBillWithHS(firstCargoData, allCargoData, hsCodeMap);
        archive.append(okWithHSBuffer, { name: `A/${safeBillNumber}总提单OK件的格式(带HS的.xlsx` });
        console.log(`生成汇总文件: A/${safeBillNumber}总提单OK件的格式(带HS的.xlsx`);

        // 生成总提单OK件（无HS）
        const okWithoutHSBuffer = await generateOKBillWithoutHS(firstCargoData, allCargoData);
        archive.append(okWithoutHSBuffer, { name: `A/${safeBillNumber}总提单OK件的格式(无HS的.xlsx` });
        console.log(`生成汇总文件: A/${safeBillNumber}总提单OK件的格式(无HS的.xlsx`);

        // 生成带HS的汇总单
        const summaryWithHSBuffer = await generateSummaryWithHS(firstCargoData, allCargoData, hsCodeMap);
        archive.append(summaryWithHSBuffer, { name: `A/${safeBillNumber}带HS的汇总单.docx` });
        console.log(`生成汇总文件: A/${safeBillNumber}带HS的汇总单.docx`);
      } catch (summaryError) {
        console.error('生成汇总文件失败，跳过:', summaryError);
      }
    }

    // 完成 ZIP 归档
    await archive.finalize();
    const zipBuffer = await zipPromise;
    const zipBase64 = bufferToBase64(zipBuffer);

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.json({
      success: true,
      message: `批量处理完成，共处理 ${files.length} 个文件`,
      zipFileBase64: zipBase64,
      fileCount: files.length,
    });
  } catch (error) {
    console.error('处理文件失败:', error);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.status(500).json({ success: false, message: '处理文件失败：' + error.message });
  }
};

// Helper to parse form data
async function parseFormData(req) {
  const Busboy = require('busboy');
  return new Promise((resolve, reject) => {
    const busboy = Busboy({ headers: req.headers });
    const files = {};
    const fields = {};

    busboy.on('file', (fieldname, file, info) => {
      const chunks = [];
      file.on('data', (chunk) => chunks.push(chunk));
      file.on('end', () => {
        const fileData = { buffer: Buffer.concat(chunks), ...info };
        if (!files[fieldname]) {
          files[fieldname] = [fileData];
        } else if (Array.isArray(files[fieldname])) {
          files[fieldname].push(fileData);
        } else {
          // 如果已存在但不是数组，转换为数组
          files[fieldname] = [files[fieldname], fileData];
        }
      });
    });

    busboy.on('field', (fieldname, value) => {
      fields[fieldname] = value;
    });

    busboy.on('finish', () => resolve({ files, fields }));
    busboy.on('error', reject);

    req.pipe(busboy);
  });
}

module.exports.parseManifestExcel = parseManifestExcel;
module.exports.generateWordDocument = generateWordDocument;
module.exports.generateExcelDocument = generateExcelDocument;
module.exports.generateCombinedLetter = generateCombinedLetter;
module.exports.generateOKBillWithHS = generateOKBillWithHS;
module.exports.generateOKBillWithoutHS = generateOKBillWithoutHS;
module.exports.generateSummaryWithHS = generateSummaryWithHS;

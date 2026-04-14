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
    发货人名称: getCellValue(27, 2),
    发货人地址: getCellValue(28, 2),
    发货人电话: getCellValue(30, 2),
    收货人名称: getCellValue(34, 2),
    收货人地址: getCellValue(35, 2),
    收货人电话: getCellValue(37, 2),
    收货人联系人: getCellValue(39, 2),
    通知人名称: getCellValue(43, 2),
    通知人地址: getCellValue(44, 2),
    通知人电话: getCellValue(46, 2),
  };

  data.发货人 = [
    data.发货人名称,
    data.发货人地址,
    `TEL: ${data.发货人电话}`
  ].filter(Boolean).join('\n');

  data.收货人 = [
    data.收货人名称,
    data.收货人地址,
    `TEL: ${data.收货人电话}`,
    data.收货人联系人 ? `Contact: ${data.收货人联系人}` : ''
  ].filter(Boolean).join('\n');

  data.通知人 = [
    data.通知人名称,
    data.通知人地址,
    `TEL: ${data.通知人电话}`
  ].filter(Boolean).join('\n');

  return data;
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

  const goodsList = data.英文品名.split(',').map(s => s.trim()).filter(Boolean);
  const goodsData = {};
  for (let i = 1; i <= 22; i++) {
    goodsData[`商品${i}`] = goodsList[i - 1] || '';
  }

  doc.setData({
    船名: data.船名,
    航次: data.航次,
    目的港: data.目的港,
    提单号: data.提单号,
    箱号: data.箱号,
    封号: data.封号,
    箱型: data.箱型,
    件数: data.件数,
    毛重: data.毛重,
    体积: data.体积,
    公司名: data.发货人名称,
    公司地址: data.发货人地址,
    电话: data.发货人电话,
    传真: '',
    电子邮箱: '',
    许可证号: '',
    收货地址: data.收货人地址,
    邮编: '',
    手机号: '',
    电话号码: data.收货人电话,
    姓名: data.通知人名称,
    地址: data.通知人地址,
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
      let font = {};
      if (cell.value?.richText && cell.value.richText.length > 0) {
        font = cell.value.richText[0].font || {};
      }
      cell.value = {
        richText: [{
          font: font,
          text: replacement
        }]
      };
      return true;
    }
    return false;
  };

  // 准备替换数据
  const goodsList = data.英文品名.split(',').map(s => s.trim()).filter(Boolean);
  const replacementData = {
    '{发票日期}': formattedDate
  };

  // 添加商品占位符替换数据
  for (let i = 1; i <= 22; i++) {
    const placeholder = `{商品${i}}`;
    replacementData[placeholder] = goodsList[i - 1] || '';
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

    console.log(`开始批量处理 ${files.length} 个文件`);

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
        throw new Error(`第 ${i + 1} 个文件处理失败: ${fileError.message}`);
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

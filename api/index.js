// Vercel serverless function entry point
const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const PORT = process.env.PORT || 5000;

// 中间件
app.use(express.json());

// Vercel serverless 环境使用 /tmp 目录
const tmpDir = '/tmp/uploads';
const outputDir = '/tmp/output';

// 确保临时目录存在
fs.mkdir(tmpDir, { recursive: true }).catch(() => {});
fs.mkdir(outputDir, { recursive: true }).catch(() => {});

// 配置文件上传（使用内存存储，适合 serverless）
const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB 限制
});

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
    // 行4: 船名,航次,目的港
    船名: getCellValue(3, 1),
    航次: getCellValue(3, 4),
    目的港: getCellValue(3, 7),

    // 行5: 总提单号
    提单号: getCellValue(4, 1),

    // 行13: 箱号,封号,箱型
    箱号: getCellValue(12, 0),
    封号: getCellValue(12, 1),
    箱型: getCellValue(12, 2),

    // 行21: 英文品名,件数,毛重,体积
    英文品名: getCellValue(20, 4),
    件数: getCellValue(20, 6),
    包装单位: getCellValue(20, 7),
    毛重: getCellValue(20, 8),
    体积: getCellValue(20, 9),
    唛头: getCellValue(20, 10),

    // 发货人信息
    发货人名称: getCellValue(27, 2),
    发货人地址: getCellValue(28, 2),
    发货人电话: getCellValue(30, 2),

    // 收货人信息
    收货人名称: getCellValue(34, 2),
    收货人地址: getCellValue(35, 2),
    收货人电话: getCellValue(37, 2),
    收货人联系人: getCellValue(39, 2),

    // 通知人信息
    通知人名称: getCellValue(43, 2),
    通知人地址: getCellValue(44, 2),
    通知人电话: getCellValue(46, 2),
  };

  // 组合完整信息
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
  // Vercel serverless 中读取模板文件
  const templatePath = path.join(__dirname, '../templates/提单确认件的格式.docx');
  const templateBuffer = await fs.readFile(templatePath);

  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // 商品列表
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
    发货人: data.发货人,
    收货人: data.收货人,
    通知人: data.通知人,
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
  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error('无法加载 Excel 模板');
  }

  // 生成日期
  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  // 获取单元格文本
  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

  // 替换占位符
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

  // 填充发票日期
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      replacePlaceholder(cell, '{发票日期}', formattedDate);
    });
  });

  // 填充商品列表
  const goodsList = data.英文品名.split(',').map(s => s.trim()).filter(Boolean);
  for (let i = 0; i < 22; i++) {
    const rowNum = 12 + i;
    const row = worksheet.getRow(rowNum);
    const cell = row.getCell(5);
    const placeholder = `{商品${i + 1}}`;

    if (i < goodsList.length) {
      replacePlaceholder(cell, placeholder, goodsList[i]);
    } else {
      replacePlaceholder(cell, placeholder, '');
    }
  }

  return workbook.xlsx.writeBuffer();
}

// API: 处理舱单文件
app.post('/api/process', upload.single('manifest'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: '请上传舱单文件' });
    }

    // 读取上传的文件（内存中）
    const manifestBuffer = req.file.buffer;

    // 解析舱单数据
    const cargoData = parseManifestExcel(manifestBuffer);

    // 生成文件
    const wordBuffer = await generateWordDocument(cargoData);
    const excelBuffer = await generateExcelDocument(cargoData);

    // 保存文件到 /tmp
    const timestamp = Date.now();
    const wordFileName = `提单确认件_${timestamp}.doc`;
    const excelFileName = `装箱单发票_${timestamp}.xls`;

    const wordFilePath = path.join(outputDir, wordFileName);
    const excelFilePath = path.join(outputDir, excelFileName);

    await fs.writeFile(wordFilePath, wordBuffer);
    await fs.writeFile(excelFilePath, excelBuffer);

    res.json({
      success: true,
      message: '文件处理成功',
      data: cargoData,
      wordFileUrl: `/api/download?file=${encodeURIComponent(wordFileName)}`,
      excelFileUrl: `/api/download?file=${encodeURIComponent(excelFileName)}`,
    });
  } catch (error) {
    console.error('处理文件失败:', error);
    res.status(500).json({ success: false, message: '处理文件失败，请检查文件格式' });
  }
});

// API: 重新生成文件
app.post('/api/regenerate', async (req, res) => {
  try {
    const cargoData = req.body.data;

    if (!cargoData) {
      return res.status(400).json({ success: false, message: '缺少数据' });
    }

    // 生成文件
    const wordBuffer = await generateWordDocument(cargoData);
    const excelBuffer = await generateExcelDocument(cargoData);

    // 保存文件
    const timestamp = Date.now();
    const wordFileName = `提单确认件_${timestamp}.doc`;
    const excelFileName = `装箱单发票_${timestamp}.xls`;

    const wordFilePath = path.join(outputDir, wordFileName);
    const excelFilePath = path.join(outputDir, excelFileName);

    await fs.writeFile(wordFilePath, wordBuffer);
    await fs.writeFile(excelFilePath, excelBuffer);

    res.json({
      success: true,
      message: '文件重新生成成功',
      wordFileUrl: `/api/download?file=${encodeURIComponent(wordFileName)}`,
      excelFileUrl: `/api/download?file=${encodeURIComponent(excelFileName)}`,
    });
  } catch (error) {
    console.error('重新生成文件失败:', error);
    res.status(500).json({ success: false, message: '重新生成文件失败' });
  }
});

// API: 下载文件
app.get('/api/download', async (req, res) => {
  try {
    const filename = req.query.file;
    if (!filename) {
      return res.status(400).json({ success: false, message: '缺少文件名' });
    }

    const filePath = path.join(outputDir, filename);

    // 检查文件是否存在
    try {
      await fs.access(filePath);
    } catch {
      console.error('文件不存在:', filePath);
      return res.status(404).json({ success: false, message: '文件不存在' });
    }

    // 读取文件
    const fileBuffer = await fs.readFile(filePath);

    // 根据文件扩展名设置 MIME 类型
    const ext = path.extname(filename).toLowerCase();
    let contentType = 'application/octet-stream';
    if (ext === '.doc' || ext === '.docx') {
      contentType = 'application/msword';
    } else if (ext === '.xls' || ext === '.xlsx') {
      contentType = 'application/vnd.ms-excel';
    }

    // 设置响应头（兼容手机浏览器）
    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Length', fileBuffer.length);

    // Content-Disposition 支持 UTF-8 文件名 (RFC 5987)
    const encodedFileName = encodeURIComponent(filename);
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"; filename*=UTF-8''${encodedFileName}`);

    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Accept-Ranges', 'bytes');

    res.send(fileBuffer);
  } catch (error) {
    console.error('下载文件失败:', error);
    res.status(500).json({ success: false, message: '下载文件失败' });
  }
});

// 导出为 Vercel serverless 函数
module.exports = app;

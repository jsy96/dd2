// Vercel serverless function for downloading files
const path = require('path');
const fs = require('fs').promises;

const outputDir = '/tmp/output';

module.exports = async (req, res) => {
  // Handle CORS preflight
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'GET, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }

  if (req.method !== 'GET') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

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

    // 设置响应头
    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Length', fileBuffer.length);
    res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(filename)}"`);
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Access-Control-Allow-Origin', '*');

    res.send(fileBuffer);
  } catch (error) {
    console.error('下载文件失败:', error);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.status(500).json({ success: false, message: '下载文件失败' });
  }
};

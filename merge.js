import ExcelJS from 'exceljs';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const OUTPUT_DIR = path.join(__dirname, 'output');
const COLUMNS_PATH = path.join(__dirname, 'columns_186.json');
const COLUMNS = JSON.parse(fs.readFileSync(COLUMNS_PATH, 'utf8'));

async function mergeExcelFiles() {
  if (!fs.existsSync(OUTPUT_DIR)) {
    console.log('[merge] Folder output belum ada.');
    return;
  }

  const files = fs.readdirSync(OUTPUT_DIR)
    .filter(f => f.startsWith('tweets_') && f.endsWith('.xlsx') && !f.startsWith('merged_'));

  if (files.length === 0) {
    console.log('[merge] Tidak ada file tweets_*.xlsx yang ditemukan di folder output.');
    return;
  }

  console.log(`[merge] Memproses ${files.length} file di folder output...`);

  const uniqueTweets = new Map();
  let totalRowsScanned = 0;

  for (const file of files) {
    const filePath = path.join(OUTPUT_DIR, file);
    try {
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.readFile(filePath);
      const worksheet = workbook.getWorksheet(1);
      if (!worksheet) continue;

      let fileRowCount = 0;
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return;

        const rowValues = row.values;
        const obj = {};
        for (let i = 0; i < COLUMNS.length; i++) {
          obj[COLUMNS[i]] = rowValues[i + 1] ?? null;
        }

        const tweetId = String(obj.id || '').trim();
        if (tweetId) {
          totalRowsScanned++;
          fileRowCount++;
          if (!uniqueTweets.has(tweetId)) {
            uniqueTweets.set(tweetId, obj);
          }
        }
      });
      console.log(`- ${file} (${fileRowCount} tweet)`);
    } catch {}
  }

  const uniqueArray = Array.from(uniqueTweets.values());
  const duplicateCount = totalRowsScanned - uniqueArray.length;

  if (uniqueArray.length === 0) {
    console.log('[merge] Selesai: 0 tweet ditemukan.');
    return;
  }

  uniqueArray.sort((a, b) => {
    const timeA = a.createdAt ? new Date(a.createdAt).getTime() : 0;
    const timeB = b.createdAt ? new Date(b.createdAt).getTime() : 0;
    return timeB - timeA;
  });

  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const outFilename = `merged_unique_${uniqueArray.length}_${timestamp}.xlsx`;
  const outPath = path.join(OUTPUT_DIR, outFilename);

  const outWorkbook = new ExcelJS.Workbook();
  const outWorksheet = outWorkbook.addWorksheet('UniqueTweets');
  outWorksheet.columns = COLUMNS.map(c => ({ header: c, key: c }));

  for (const tweet of uniqueArray) {
    outWorksheet.addRow(tweet);
  }

  await outWorkbook.xlsx.writeFile(outPath);
  console.log(`\n[merge] Selesai: ${uniqueArray.length} tweet unik tersimpan -> output/${outFilename} (${duplicateCount} duplikat dibuang)\n`);
}

mergeExcelFiles();

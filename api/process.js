import { createClient } from '@supabase/supabase-js';
import JSZip from 'jszip';

// ── 집계 로직 ──────────────────────────────────────────────
const ROW_FILTERS = [
  { pol: 'VNHPH',             pod: 'KR',                       polList: ['VNHPH'],                                                         podList: null },
  { pol: 'CNSHK',             pod: 'KR',                       polList: ['CNSHK'],                                                         podList: null },
  { pol: 'CNXMN',             pod: 'KR',                       polList: ['CNXMN'],                                                         podList: null },
  { pol: 'VNSGN',             pod: 'KR',                       polList: ['VNSGN'],                                                         podList: null },
  { pol: 'TH',                pod: 'KR',                       polList: ['TH'],                                                            podList: null },
  { pol: 'JP',                pod: 'KR',                       polList: ['JPYOK','JPTYO','JPNGO','JPSMZ','JPOSA','JPUKB','JPHKT','JPMOJ'], podList: null },
  { pol: 'CNSHA',             pod: 'KRINC/KRKUV',              polList: ['CNSHA'],                                                         podList: ['KRINC','KRKUV'] },
  { pol: 'CNNGB',             pod: 'KRINC/KRKUV',              polList: ['CNNGB'],                                                         podList: ['KRINC','KRKUV'] },
  { pol: 'CNSHA',             pod: 'KRPUS/KRKAN/KRUSN',        polList: ['CNSHA'],                                                         podList: ['KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNNGB',             pod: 'KRPUS/KRKAN/KRUSN',        polList: ['CNNGB'],                                                         podList: ['KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNTAO/CNLYG',       pod: 'KR',                       polList: ['CNTAO','CNLYG'],                                                 podList: null },
  { pol: 'CNTXG',             pod: 'KRINC/KRPUS/KRKAN/KRUSN',  polList: ['CNTXG'],                                                        podList: ['KRINC','KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNTXG',             pod: 'KRPTK',                    polList: ['CNTXG'],                                                         podList: ['KRPTK'] },
  { pol: 'CNDLC',             pod: 'KR',                       polList: ['CNDLC'],                                                         podList: null },
  { pol: 'CNYNT',             pod: 'KR',                       polList: ['CNYNT'],                                                         podList: null },
  { pol: 'CNNKG',             pod: 'KR',                       polList: ['CNNKG'],                                                         podList: null },
  { pol: 'CNZJG',             pod: 'KRPUS/KRKAN/KRUSN',        polList: ['CNZJG'],                                                         podList: ['KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNNTG/CNTAG/CNZJG', pod: 'KRINC/KRPTK',             polList: ['CNNTG','CNTAG','CNZJG'],                                         podList: ['KRINC','KRPTK'] },
  { pol: 'RU',                pod: 'KR',                       polList: ['RUVVO','RUVFP'],                                                  podList: null },
];

const ROW_EXCEL = {
  'VNHPH|KR':                        [5,6],
  'CNSHK|KR':                        [9,10],
  'CNXMN|KR':                        [13,14],
  'VNSGN|KR':                        [17,18],
  'TH|KR':                           [21,22],
  'JP|KR':                           [25,26],
  'CNSHA|KRINC/KRKUV':               [29,30],
  'CNNGB|KRINC/KRKUV':               [33,34],
  'CNSHA|KRPUS/KRKAN/KRUSN':         [37,38],
  'CNNGB|KRPUS/KRKAN/KRUSN':         [41,42],
  'CNTAO/CNLYG|KR':                  [45,46],
  'CNTXG|KRINC/KRPUS/KRKAN/KRUSN':  [49,50],
  'CNTXG|KRPTK':                     [53,54],
  'CNDLC|KR':                        [57,58],
  'CNYNT|KR':                        [61,62],
  'CNNKG|KR':                        [65,66],
  'CNZJG|KRPUS/KRKAN/KRUSN':        [69,70],
  'CNNTG/CNTAG/CNZJG|KRINC/KRPTK': [73,74],
  'RU|KR':                           [77,78],
};

// ── XML 직접 수정 (서식 완벽 보존) ─────────────────────────
function setCellValue(xml, cellAddr, value) {
  const val = value !== null && value !== undefined ? String(value) : '0';

  // 수식 셀은 건드리지 않음
  const formulaRe = new RegExp(`<c r="${cellAddr}"[^>]*>[^<]*<f`);
  if (formulaRe.test(xml)) return xml;

  // 빈 셀: <c r="AK5" s="25"/>
  const emptyRe = new RegExp(`(<c r="${cellAddr}"(?:\\s+[^/]*)?)(\\/\\s*>)`);
  const emptyMatch = xml.match(emptyRe);
  if (emptyMatch) {
    const newCell = `${emptyMatch[1]}><v>${val}</v></c>`;
    return xml.slice(0, emptyMatch.index) + newCell + xml.slice(emptyMatch.index + emptyMatch[0].length);
  }

  // 값 있는 셀: <c r="AK5" s="25"><v>123</v></c>
  const valueRe = new RegExp(`(<c r="${cellAddr}"[^>]*>)(<v>[^<]*<\\/v>)(<\\/c>)`);
  const valueMatch = xml.match(valueRe);
  if (valueMatch) {
    const newCell = `${valueMatch[1]}<v>${val}</v></c>`;
    return xml.slice(0, valueMatch.index) + newCell + xml.slice(valueMatch.index + valueMatch[0].length);
  }

  return xml;
}

// ── RAW DATA 파싱 ───────────────────────────────────────────
function parseRawData(buffer) {
  const XLSX = require('xlsx');
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const headers = rows[3]; // 4번째 행 = 헤더

  const findIdx = (name, def) => {
    const i = headers.findIndex(h => String(h||'').trim() === name);
    return i >= 0 ? i : def;
  };
  const iPol = findIdx('POL', 13), iPod = findIdx('POD', 14);
  const iTeu = findIdx('TEU', 20), iPc  = findIdx('P/C', 21);
  const iSttl = findIdx('S.TTL1', 29);

  const data = [];
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;
    const pol  = String(row[iPol]  || '').trim();
    const pod  = String(row[iPod]  || '').trim();
    const teu  = parseFloat(row[iTeu]  || 0);
    const pc   = String(row[iPc]   || '').trim().toUpperCase();
    const sttl = parseFloat(row[iSttl] || 0);
    if (pol && pol !== 'nan' && pol !== 'POL' && teu > 0) {
      data.push({ pol, pod, teu, sttl, pc: pc === 'P' ? 'PP' : 'CLT' });
    }
  }
  return data;
}

function getTeuRpb(data, polList, podList, pc) {
  let sub = data.filter(r => polList.includes(r.pol));
  if (podList) sub = sub.filter(r => podList.includes(r.pod));
  if (pc === 'PP')  sub = sub.filter(r => r.pc === 'PP');
  if (pc === 'CLT') sub = sub.filter(r => r.pc === 'CLT');
  const teu = sub.reduce((s, r) => s + r.teu, 0);
  const inc = sub.reduce((s, r) => s + r.sttl, 0);
  return { teu: Math.round(teu), rpb: teu > 0 ? inc / teu : 0 };
}

// ── Vercel Handler ──────────────────────────────────────────
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { fileData, label } = req.body;
    if (!fileData) return res.status(400).json({ error: 'fileData 없음' });

    // RAW DATA 파싱
    const rawBuffer = Buffer.from(fileData, 'base64');
    const data = parseRawData(rawBuffer);
    const totalTeu = data.reduce((s, r) => s + r.teu, 0);

    // 템플릿 로드 (Supabase)
    const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);
    const fileList = await supabase.storage.from('Templates').list('', { sortBy: { column: 'created_at', order: 'desc' } });
    if (!fileList.data || fileList.data.length === 0) throw new Error('Templates 버킷에 파일이 없습니다.');
    const latest = fileList.data[0].name;
    const { data: tplBlob } = await supabase.storage.from('Templates').download(latest);
    const tplBuffer = Buffer.from(await tplBlob.arrayBuffer());

    // ZIP으로 열어서 sheet XML 직접 수정 (서식 100% 보존)
    const zip = await JSZip.loadAsync(tplBuffer);
    let sheetXml = await zip.file('xl/worksheets/sheet1.xml').async('string');

    // 데이터 입력
    for (const f of ROW_FILTERS) {
      const [ppRow, cltRow] = ROW_EXCEL[`${f.pol}|${f.pod}`];
      for (const [rowNum, pc] of [[ppRow, 'PP'], [cltRow, 'CLT']]) {
        const { teu, rpb } = getTeuRpb(data, f.polList, f.podList, pc);
        sheetXml = setCellValue(sheetXml, `AK${rowNum}`, teu);
        sheetXml = setCellValue(sheetXml, `AL${rowNum}`, parseFloat(rpb.toFixed(10)));
      }
    }

    // 결과 파일 생성
    zip.file('xl/worksheets/sheet1.xml', sheetXml);
    const wbXml = await zip.file('xl/workbook.xml').async('string');
    const wbXmlFixed = wbXml.replace(/<calcPr([^/]*?)(\/?>)/, '<calcPr$1 fullCalcOnLoad="1"$2');
    zip.file('xl/workbook.xml', wbXmlFixed);
    const outBuffer = Buffer.from(await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }));

    // Supabase 저장
    const today = new Date().toISOString().slice(0, 10);
    const ts = Date.now();
    const filename = `RPB_${label || 'result'}_${today}.xlsx`;
    await supabase.storage.from('Results').upload(filename, outBuffer, {
      contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    });
    const { data: urlData } = supabase.storage.from('Results').getPublicUrl(filename);

    await supabase.from('rpb_history').insert({
      filename, label: label || '', url: urlData.publicUrl,
      total_teu: Math.round(totalTeu), created_at: new Date().toISOString(),
    });

    res.status(200).json({
      success: true, filename, url: urlData.publicUrl,
      fileBase64: outBuffer.toString('base64'),
      totalTeu: Math.round(totalTeu),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

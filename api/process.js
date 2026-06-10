import { createClient } from '@supabase/supabase-js';
import * as XLSX from 'xlsx';

// ?Ä?Ä ÏßëÍ≥Ñ Î°úÏßÅ ?ïÏùò ?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä
const ROW_FILTERS = [
  { pol: 'VNHPH',             pod: 'KR',                      polList: ['VNHPH'],                                                          podList: null },
  { pol: 'CNSHK',             pod: 'KR',                      polList: ['CNSHK'],                                                          podList: null },
  { pol: 'CNXMN',             pod: 'KR',                      polList: ['CNXMN'],                                                          podList: null },
  { pol: 'VNSGN',             pod: 'KR',                      polList: ['VNSGN'],                                                          podList: null },
  { pol: 'TH',                pod: 'KR',                      polList: ['TH'],                                                             podList: null },
  { pol: 'JP',                pod: 'KR',                      polList: ['JPYOK','JPTYO','JPNGO','JPSMZ','JPOSA','JPUKB','JPHKT','JPMOJ'],  podList: null },
  { pol: 'CNSHA',             pod: 'KRINC/KRKUV',             polList: ['CNSHA'],                                                          podList: ['KRINC','KRKUV'] },
  { pol: 'CNNGB',             pod: 'KRINC/KRKUV',             polList: ['CNNGB'],                                                          podList: ['KRINC','KRKUV'] },
  { pol: 'CNSHA',             pod: 'KRPUS/KRKAN/KRUSN',       polList: ['CNSHA'],                                                          podList: ['KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNNGB',             pod: 'KRPUS/KRKAN/KRUSN',       polList: ['CNNGB'],                                                          podList: ['KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNTAO/CNLYG',       pod: 'KR',                      polList: ['CNTAO','CNLYG'],                                                  podList: null },
  { pol: 'CNTXG',             pod: 'KRINC/KRPUS/KRKAN/KRUSN', polList: ['CNTXG'],                                                         podList: ['KRINC','KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNTXG',             pod: 'KRPTK',                   polList: ['CNTXG'],                                                          podList: ['KRPTK'] },
  { pol: 'CNDLC',             pod: 'KR',                      polList: ['CNDLC'],                                                          podList: null },
  { pol: 'CNYNT',             pod: 'KR',                      polList: ['CNYNT'],                                                          podList: null },
  { pol: 'CNNKG',             pod: 'KR',                      polList: ['CNNKG'],                                                          podList: null },
  { pol: 'CNZJG',             pod: 'KRPUS/KRKAN/KRUSN',       polList: ['CNZJG'],                                                          podList: ['KRPUS','KRKAN','KRUSN'] },
  { pol: 'CNNTG/CNTAG/CNZJG', pod: 'KRINC/KRPTK',            polList: ['CNNTG','CNTAG','CNZJG'],                                          podList: ['KRINC','KRPTK'] },
  { pol: 'RU',                pod: 'KR',                      polList: ['RUVVO','RUVFP'],                                                  podList: null },
];

// ?ëÏ? ??Î≤àÌò∏: [PP?? CLT?? (1-indexed)
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

function parseRawData(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  const headers = rows[3]; // 4Î≤àÏß∏ ??= ?§Îçî
  const data = [];
  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;
    const obj = {};
    headers.forEach((h, idx) => { if (h) obj[h] = row[idx]; });
    const teu   = parseFloat(obj['TEU']);
    const sttl1 = parseFloat(obj['S.TTL1']);
    const pol   = String(obj['POL'] || '').trim();
    const pod   = String(obj['POD'] || '').trim();
    const pc    = String(obj['P/C'] || '').trim().toUpperCase();
    if (!isNaN(teu) && teu > 0 && pol && pol !== 'undefined' && pol !== 'POL') {
      data.push({ pol, pod, teu, sttl1: isNaN(sttl1) ? 0 : sttl1, pc: pc === 'P' ? 'PP' : 'CLT' });
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
  const inc = sub.reduce((s, r) => s + r.sttl1, 0);
  return { teu: Math.round(teu), rpb: teu > 0 ? inc / teu : 0 };
}

// ?Ä?Ä Vercel Handler ?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä?Ä
export default async function handler(req, res) {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  try {
    const { fileData, label } = req.body;
    if (!fileData) return res.status(400).json({ error: 'fileDataÍ∞Ä ?ÜÏäµ?àÎã§.' });

    // RAW DATA ?åÏã±
    const rawBuffer = Buffer.from(fileData, 'base64');
    const data = parseRawData(rawBuffer);

    // ?úÌîåÎ¶?Î°úÎìú ??templates Î≤ÑÌÇ∑?êÏÑú Í∞Ä??ÏµúÍ∑º ?åÏùº ?êÎèô ?†ÌÉù
    const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_SERVICE_KEY);
    const { data: fileList, error: listErr } = await supabase.storage
      .from('Templates').list('', { sortBy: { column: 'created_at', order: 'desc' } });
    if (listErr || !fileList || fileList.length === 0)
      throw new Error('templates Î≤ÑÌÇ∑???åÏùº???ÜÏäµ?àÎã§.');
    const latestFile = fileList[0].name;
    const { data: tplBlob, error: tplErr } = await supabase.storage
      .from('Templates').download(latestFile);
    if (tplErr) throw new Error(`?úÌîåÎ¶?Î°úÎìú ?§Ìå® (${latestFile}): ` + tplErr.message);

    const templateBuffer = Buffer.from(await tplBlob.arrayBuffer());
    const wb = XLSX.read(templateBuffer, { type: 'buffer' });
    const ws = wb.Sheets['RPB'];

    // INBOUND TOTAL PP/CLT ??AK/AL ?®Ïàò (?ÜÏùÑ Í≤ΩÏö∞ ?ΩÏûÖ)
    const encCell = (r, c) => XLSX.utils.encode_cell({ r: r - 1, c: c - 1 });
    const sumPP  = [5,9,13,17,21,25,29,33,37,41,45,49,53,57,61,65,69,73,77,81].map(r=>`AK${r}`).join('+');
    const sumCLT = [6,10,14,18,22,26,30,34,38,42,46,50,54,58,62,66,70,74,78,82].map(r=>`AK${r}`).join('+');
    if (!ws[encCell(85,37)]) ws[encCell(85,37)] = { t:'f', f: sumPP };
    if (!ws[encCell(86,37)]) ws[encCell(86,37)] = { t:'f', f: sumCLT };
    if (!ws[encCell(85,38)]) ws[encCell(85,38)] = { t:'f', f: 'IFERROR(AW85/AK85,"0")' };
    if (!ws[encCell(86,38)]) ws[encCell(86,38)] = { t:'f', f: 'IFERROR(AW86/AK86,"0")' };

    // ?∞Ïù¥???ÖÎ†• (AK=col37, AL=col38)
    for (const { pol, pod, polList, podList } of ROW_FILTERS) {
      const [ppRow, cltRow] = ROW_EXCEL[`${pol}|${pod}`];
      for (const [rowNum, pc] of [[ppRow,'PP'],[cltRow,'CLT']]) {
        const { teu, rpb } = getTeuRpb(data, polList, podList, pc);
        ws[encCell(rowNum, 37)] = { t: 'n', v: teu };
        ws[encCell(rowNum, 38)] = { t: 'n', v: parseFloat(rpb.toFixed(10)) };
      }
    }

    // Í≤∞Í≥º ?åÏùº ?ùÏÑ±
    const outBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const filename = `RPB_${label || 'result'}_${new Date().toISOString().slice(0,10)}_${Date.now()}.xlsx`;

    // Supabase???ÖÎ°ú??(?àÏä§?†Î¶¨ Î≥¥Í?)
    const { error: uploadErr } = await supabase.storage
      .from('Results').upload(filename, outBuffer, {
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
    if (uploadErr) throw new Error('Í≤∞Í≥º ?Ä???§Ìå®: ' + uploadErr.message);

    const { data: urlData } = supabase.storage.from('Results').getPublicUrl(filename);

    // DB???àÏä§?†Î¶¨ Í∏∞Î°ù
    await supabase.from('rpb_history').insert({
      filename,
      label: label || '',
      url: urlData.publicUrl,
      created_at: new Date().toISOString(),
      total_teu: data.reduce((s,r) => s+r.teu, 0),
    });

    res.status(200).json({
      success: true,
      filename,
      url: urlData.publicUrl,
      fileBase64: outBuffer.toString('base64'),
      totalTeu: data.reduce((s,r) => s+r.teu, 0),
    });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
}

export const config = { api: { bodyParser: { sizeLimit: '10mb' } } };

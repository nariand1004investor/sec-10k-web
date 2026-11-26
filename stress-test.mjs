// stress-test.mjs
// SEC 10-K 자동 스트레스 테스트 + SEC 원본 교차 대조(검증 8) + 번역 품질 검수 덤프 스크립트
// 주의: app/api/analyze/route.js 및 verify.mjs 본체는 일절 수정하지 않고 독립 실행함
import fs from 'node:fs';
import path from 'node:path';
import module from 'node:module';
import { fileURLToPath } from 'node:url';
import ExcelJS from 'exceljs';

// Next.js API Route('next/server')를 별도 dev 서버 없이 Node ESM에서 직접 로드하기 위한 해석 훅
module.registerHooks({
  resolve(specifier, context, nextResolve) {
    if (specifier === 'next/server') {
      return nextResolve('next/server.js', context);
    }
    return nextResolve(specifier, context);
  },
});

const origFetchForStress = globalThis.fetch.bind(globalThis);
let cachedCompanyTickersResponse = null;
let activeCapture = null;

function createEmptyCapture() {
  return {
    cik: null,
    companyFacts: null,
    submissions: null,
    filingSummaries: new Map(), // accNoDashes -> { IS, BS, CF }
    rHtmlEntries: [], // { accNoDashes, htmlFile, html }
    geminiMappings: [], // { stmt, tag, engLabel, aiKor }
  };
}

function buildCleanHeaders(headers) {
  const hdrs = {};
  headers.forEach((v, k) => {
    if (!['content-encoding', 'content-length', 'transfer-encoding'].includes(k.toLowerCase())) {
      hdrs[k] = v;
    }
  });
  return hdrs;
}

function parseFilingSummaryStatements(summaryXml) {
  const reportBlocks = summaryXml.match(/<Report[\s\S]*?<\/Report>/gi) || [];
  const statementFiles = { IS: null, BS: null, CF: null };

  for (const block of reportBlocks) {
    const catMatch = block.match(/<MenuCategory>(.*?)<\/MenuCategory>/i);
    const shortMatch = block.match(/<ShortName>(.*?)<\/ShortName>/i);
    const htmlMatch = block.match(/<HtmlFileName>(.*?)<\/HtmlFileName>/i);

    if (!shortMatch || !htmlMatch) continue;
    const category = (catMatch ? catMatch[1] : '').trim();
    const shortName = shortMatch[1].trim();
    const htmlFile = htmlMatch[1].trim();

    if (category && category !== 'Statements' && category !== 'Uncategorized') continue;
    if (
      /Parenthetical|Equity|Stockholders|Shareholders|Supplemental|Schedule|Details|Tables|Notes/i.test(
        shortName
      )
    ) {
      continue;
    }
    if (
      /Comprehensive/i.test(shortName) &&
      !/Operations|Statements?\s+of\s+(?:Income|Loss|Earnings)(?:\s*\((?:Loss|Income)\))?\s+and/i.test(
        shortName
      )
    ) {
      continue;
    }

    if (!statementFiles.IS && /Operations|Income|Earnings/i.test(shortName)) {
      statementFiles.IS = htmlFile;
    } else if (
      !statementFiles.BS &&
      /Balance Sheet|Financial Position|Financial Condition|Statements?\s+of\s+Condition/i.test(
        shortName
      )
    ) {
      statementFiles.BS = htmlFile;
    } else if (!statementFiles.CF && /Cash Flow/i.test(shortName)) {
      statementFiles.CF = htmlFile;
    }
  }

  return statementFiles;
}

function tryCaptureGeminiMappings(reqBodyStr, resText) {
  if (!activeCapture || !reqBodyStr || !resText) return;
  try {
    const reqJson = JSON.parse(reqBodyStr);
    const promptText = reqJson?.contents?.[0]?.parts?.[0]?.text || '';
    const markerIdx = promptText.indexOf('[번역할 계정 목록]');
    if (markerIdx === -1) return;
    const jsonStart = promptText.indexOf('{', markerIdx);
    if (jsonStart === -1) return;
    const inputMap = JSON.parse(promptText.slice(jsonStart).trim());

    const resJson = JSON.parse(resText);
    const rawOut = resJson?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
    let outMap = null;
    try {
      outMap = JSON.parse(rawOut);
    } catch {
      const m = rawOut.match(/\{[\s\S]*\}/);
      if (m) outMap = JSON.parse(m[0]);
    }
    if (!outMap || typeof outMap !== 'object') return;

    for (const [k, item] of Object.entries(inputMap)) {
      const val = outMap[k];
      const aiKor = typeof val === 'string' ? val : val?.kor;
      if (aiKor && typeof aiKor === 'string' && aiKor.trim()) {
        activeCapture.geminiMappings.push({
          stmt: item?.stmt || '',
          tag: item?.tag || '',
          engLabel: item?.label || '',
          aiKor: aiKor.trim(),
        });
      }
    }
  } catch {
    // 캡처 파싱 실패 시 본체 흐름에 영향 없음
  }
}

globalThis.fetch = async (input, init) => {
  const url = typeof input === 'string' ? input : input?.url || '';

  if (url.includes('sec.gov/files/company_tickers.json')) {
    if (cachedCompanyTickersResponse) {
      const c = cachedCompanyTickersResponse;
      return new Response(c.buf.slice(0), {
        status: c.status,
        statusText: c.statusText,
        headers: c.headers,
      });
    }
    const res = await origFetchForStress(input, init);
    if (res.ok) {
      const buf = await res.arrayBuffer();
      const hdrs = buildCleanHeaders(res.headers);
      cachedCompanyTickersResponse = {
        status: res.status,
        statusText: res.statusText,
        headers: hdrs,
        buf,
      };
      return new Response(buf.slice(0), {
        status: res.status,
        statusText: res.statusText,
        headers: hdrs,
      });
    }
    return res;
  }

  const res = await origFetchForStress(input, init);
  if (!activeCapture || !res.ok) {
    return res;
  }

  const isCompanyFacts = /data\.sec\.gov\/api\/xbrl\/companyfacts\/CIK\d+\.json/i.test(url);
  const isMainSubmissions = /data\.sec\.gov\/submissions\/CIK\d+\.json/i.test(url);
  const filingSummaryMatch = url.match(/Archives\/edgar\/data\/\d+\/(\d+)\/FilingSummary\.xml/i);
  const rHtmlMatch = url.match(/Archives\/edgar\/data\/\d+\/(\d+)\/(R\d+\.htm)/i);
  const isGeminiGenerate =
    url.includes('generativelanguage.googleapis.com') && url.includes(':generateContent');

  if (
    !isCompanyFacts &&
    !isMainSubmissions &&
    !filingSummaryMatch &&
    !rHtmlMatch &&
    !isGeminiGenerate
  ) {
    return res;
  }

  const buf = await res.arrayBuffer();
  const hdrs = buildCleanHeaders(res.headers);
  const text = Buffer.from(buf).toString('utf8');

  try {
    if (isCompanyFacts) {
      const parsedFacts = JSON.parse(text);
      if (!activeCapture.companyFacts) {
        activeCapture.companyFacts = parsedFacts;
      } else if (parsedFacts?.facts) {
        activeCapture.companyFacts.facts = activeCapture.companyFacts.facts || {};
        for (const [tax, concepts] of Object.entries(parsedFacts.facts)) {
          if (!concepts || typeof concepts !== 'object') continue;
          if (!activeCapture.companyFacts.facts[tax]) {
            activeCapture.companyFacts.facts[tax] = concepts;
            continue;
          }
          const targetTax = activeCapture.companyFacts.facts[tax];
          for (const [tag, conceptObj] of Object.entries(concepts)) {
            if (!targetTax[tag]) {
              targetTax[tag] = conceptObj;
            } else if (conceptObj?.units) {
              targetTax[tag].units = targetTax[tag].units || {};
              for (const [unit, recs] of Object.entries(conceptObj.units)) {
                if (!Array.isArray(recs)) continue;
                if (!Array.isArray(targetTax[tag].units[unit])) {
                  targetTax[tag].units[unit] = recs;
                } else {
                  targetTax[tag].units[unit] = [...targetTax[tag].units[unit], ...recs];
                }
              }
            }
          }
        }
      }
    } else if (isMainSubmissions) {
      if (!activeCapture.submissions) {
        activeCapture.submissions = JSON.parse(text);
      }
    } else if (filingSummaryMatch) {
      const accNoDashes = filingSummaryMatch[1];
      activeCapture.filingSummaries.set(accNoDashes, parseFilingSummaryStatements(text));
    } else if (rHtmlMatch) {
      activeCapture.rHtmlEntries.push({
        accNoDashes: rHtmlMatch[1],
        htmlFile: rHtmlMatch[2],
        html: text,
      });
    } else if (isGeminiGenerate && typeof init?.body === 'string') {
      tryCaptureGeminiMappings(init.body, text);
    }
  } catch {
    // 캡처 파싱 오류 무시
  }

  return new Response(buf.slice(0), {
    status: res.status,
    statusText: res.statusText,
    headers: hdrs,
  });
};

const { POST } = await import('./app/api/analyze/route.js');

const DEFAULT_STRESS_TICKERS = [
  'MCD',  // 자본잠식(음의 자본)
  'NEE',  // 유틸리티(규제자산/부채·유틸리티 매출)
  'XOM',  // 에너지(Oil & Gas)
  'UNH',  // 헬스케어(보험료·의료서비스 복합)
  'DUOL', // 10년 미만 단기 상장
  'SNOW', // 적자 기업 + 1월 결산
  'CAT',  // S&P500 산업재·중장비 (금융자회사 포함)
  'DIS',  // S&P500 커뮤니케이션/미디어 (9월 결산)
  'PG',   // S&P500 필수소비재 (6월 결산)
  'SHW',  // S&P500 소재/특수화학
  'TSM',  // 해외기업 20-F 예외 처리 확인
];

const TICKER_ROLE_LABELS = {
  MCD: '자본잠식',
  NEE: '유틸리티',
  XOM: '에너지',
  UNH: '헬스케어',
  DUOL: '단기상장(<10Y)',
  SNOW: '적자기업/1월결산',
  CAT: 'S&P500(산업재)',
  DIS: 'S&P500(미디어)',
  PG: 'S&P500(필수소비재)',
  SHW: 'S&P500(소재)',
  TSM: '20-F 예외처리',
};

const SHEET_NAMES = ['손익계산서', '재무상태표', '현금흐름표'];
const BS_IDENTITY_TOLERANCE = 3;
const SCALE_JUMP_THRESHOLD = 1000;
const DEFAULT_DELAY_MS = 6000; // SEC EDGAR rate-limit 방지용 기본 6초 간격

const ALLOWLISTED_ACRONYMS = [
  'CAPEX',
  'CP',
  'IB',
  'OTTI',
  'ROU',
  'EPS',
  'AOCI',
  'LIFO',
  'FIFO',
  'VIE',
  'M&A',
  'R&D',
  'SG&A',
  'EBIT',
  'EBITDA',
  'IT',
  'FX',
  'US',
  'U.S.',
  'RSU',
  'PSU',
  'RSS',
  'GAAP',
  'IFRS',
  'SEC',
  'IRS',
  'ASC',
  'ASU',
  'USD',
  'EUR',
  'GBP',
  'JPY',
  'BM&F',
  'NEX',
  'MSA',
  'CME',
  'RLI',
  'SFM',
  'AAPL',
  'BRK',
  'S&P',
  'OSTTRA',
  'BNSF',
  'BHE',
  'GEICO',
  'LLC',
  'INC',
  'CORP',
  'LTD',
  'LLP',
  'LP',
  'PLC',
  'Class A',
  'Class B',
];

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

function loadDotEnvIfPresent() {
  const scriptDir = path.dirname(fileURLToPath(import.meta.url));
  const searchDirs = [...new Set([scriptDir, process.cwd()])];
  const loadedFiles = [];

  for (const dir of searchDirs) {
    for (const filename of ['.env.local', '.env']) {
      const filePath = path.join(dir, filename);
      if (!fs.existsSync(filePath)) continue;
      try {
        const content = fs.readFileSync(filePath, 'utf8').replace(/^\uFEFF/, '');
        let fileHasKey = false;
        for (const rawLine of content.split(/\r?\n/)) {
          const line = rawLine.trim();
          if (!line || line.startsWith('#')) continue;
          const cleanLine = line.replace(/^export\s+/, '');
          const eqIdx = cleanLine.indexOf('=');
          if (eqIdx <= 0) continue;
          const key = cleanLine.slice(0, eqIdx).trim();
          let val = cleanLine.slice(eqIdx + 1).trim();
          if (
            (val.startsWith('"') && val.endsWith('"')) ||
            (val.startsWith("'") && val.endsWith("'"))
          ) {
            val = val.slice(1, -1);
          }
          if (!(key in process.env) || !String(process.env[key]).trim()) {
            process.env[key] = val;
            if ((key === 'GEMINI_API_KEY' || key === 'GOOGLE_API_KEY') && val) {
              fileHasKey = true;
            }
          }
        }
        if (fileHasKey && !loadedFiles.includes(filename)) {
          loadedFiles.push(filename);
        }
      } catch {
        // 환경변수 파일 파싱 실패 시 무시
      }
    }
  }
  return loadedFiles[0] || null;
}

function displayWidth(str) {
  let width = 0;
  for (const ch of String(str)) {
    const code = ch.codePointAt(0);
    if (
      (code >= 0x1100 && code <= 0x11ff) ||
      (code >= 0x3000 && code <= 0x9fff) ||
      (code >= 0xac00 && code <= 0xd7af) ||
      (code >= 0xff01 && code <= 0xff60)
    ) {
      width += 2;
    } else {
      width += 1;
    }
  }
  return width;
}

function padEndDisplay(str, targetWidth) {
  const s = String(str);
  const diff = targetWidth - displayWidth(s);
  return diff > 0 ? s + ' '.repeat(diff) : s;
}

function padStartDisplay(str, targetWidth) {
  const s = String(str);
  const diff = targetWidth - displayWidth(s);
  return diff > 0 ? ' '.repeat(diff) + s : s;
}

function formatNum(v) {
  if (v === null || v === undefined) return 'null';
  return Number(v).toLocaleString('en-US');
}

function extractCellScalar(rawVal) {
  if (rawVal === null || rawVal === undefined) return null;
  if (typeof rawVal === 'object') {
    if ('result' in rawVal) return rawVal.result;
    if (Array.isArray(rawVal.richText)) {
      return rawVal.richText.map((t) => t.text || '').join('');
    }
    if ('text' in rawVal) return rawVal.text;
  }
  return rawVal;
}

function parseWorkbookSheets(workbook) {
  const sheets = {};
  for (const sheetName of SHEET_NAMES) {
    const ws = workbook.getWorksheet(sheetName);
    if (!ws) {
      sheets[sheetName] = null;
      continue;
    }

    const headerRow = ws.getRow(1);
    const years = [];
    const maxCol = ws.columnCount || 0;
    for (let col = 2; col <= maxCol; col++) {
      const cellVal = extractCellScalar(headerRow.getCell(col).value);
      const yrText = cellVal !== null && cellVal !== undefined ? String(cellVal).trim() : '';
      if (yrText) {
        years.push(yrText);
      }
    }

    const rows = [];
    const maxRow = ws.rowCount || 0;
    for (let r = 2; r <= maxRow; r++) {
      const row = ws.getRow(r);
      const rawLabel = extractCellScalar(row.getCell(1).value);
      const label = rawLabel !== null && rawLabel !== undefined ? String(rawLabel).trim() : '';
      if (!label) continue;

      const isBold = Boolean(row.font?.bold || row.getCell(1).font?.bold);
      const values = years.map((_, idx) => {
        const v = extractCellScalar(row.getCell(idx + 2).value);
        return typeof v === 'number' && Number.isFinite(v) ? v : null;
      });

      rows.push({
        rowIndex: r,
        label,
        isBold,
        values,
      });
    }

    sheets[sheetName] = { years, rows };
  }
  return sheets;
}

function makeFailure(ticker, checkId, checkName, sheet, year, account, actual, expected) {
  return {
    ticker,
    checkId,
    checkName,
    sheet,
    year: String(year),
    account,
    actual,
    expected,
  };
}

// =========================================================================================
// [검증 1~7] verify.mjs 회계 감사 규칙 그대로 유지
// =========================================================================================

// [검증 1] 재무상태표 항등식: 자산총계 = 부채총계 + 자본총계 (±3M)
function auditBalanceSheetIdentity(ticker, sheets) {
  const failures = [];
  const bs = sheets['재무상태표'];
  if (!bs || bs.rows.length === 0) {
    failures.push(
      makeFailure(
        ticker,
        'BS_IDENTITY',
        '1. 재무상태표 항등식',
        '재무상태표',
        'ALL',
        '(시트 누락)',
        '재무상태표 시트 없음 또는 빈 시트',
        '자산총계 = 부채총계 + 자본총계 존재'
      )
    );
    return { status: 'FAIL', failures, checkedYears: 0 };
  }

  const { years, rows } = bs;

  const findRow = (exactLabel, fallbackRegex, excludeRegex = null, pickLast = false) => {
    const exactMatches = rows.filter((r) => r.label === exactLabel);
    if (exactMatches.length > 0) {
      return pickLast ? exactMatches[exactMatches.length - 1] : exactMatches[0];
    }
    const regexMatches = rows.filter(
      (r) => fallbackRegex.test(r.label) && (!excludeRegex || !excludeRegex.test(r.label))
    );
    if (regexMatches.length > 0) {
      return pickLast ? regexMatches[regexMatches.length - 1] : regexMatches[0];
    }
    return null;
  };

  const assetsRow = findRow('[[자산총계]]', /^\[*자산총계\]*$|^총자산$/i, /부채/);
  let liabRow = findRow('[[부채총계]]', /^\[*부채총계\]*$|^총부채$/i, /자본/);
  const equityRow = findRow('[[자본총계]]', /^\[*자본총계\]*$|^총자본$/i, /부채/, true);
  const liabAndEqRow = findRow('[[부채와자본총계]]', /부채(?:와|및)자본총계/i);
  const mezzanineRows = rows.filter((r) =>
    /상환가능비지배지분|상환가능지분|상환가능주식|상환가능우선주|메자닌|redeemable\s+(?:noncontrolling|stock|preferred|shares?|units?)|temporary\s+equity/i.test(
      r.label
    )
  );

  if (assetsRow && equityRow && !liabRow) {
    const clIdx = rows.findIndex(
      (r) =>
        r.label === '[유동부채 합계]' ||
        (/유동부채\s*합계|^total current liabilities$/i.test(r.label) && !/비유동/.test(r.label))
    );
    if (clIdx !== -1) {
      const nonCurrLiabRows = [];
      for (let i = clIdx + 1; i < rows.length; i++) {
        const lbl = rows[i].label;
        if (
          /상환가능|메자닌|자본금|자본잉여금|이익잉여금|기타포괄|자기주식|우선주|보통주|지배기업|비지배|\[\[자본총계\]\]/i.test(
            lbl
          ) &&
          !/주식수|shares/i.test(lbl)
        ) {
          break;
        }
        if (/주식수|shares/i.test(lbl)) continue;
        if (/^\[.*합계\]$/i.test(lbl)) continue;
        if (rows[i].isBold && nonCurrLiabRows.length > 0) {
          let isIntermediateSub = false;
          for (let start = 0; start < nonCurrLiabRows.length; start++) {
            const slice = nonCurrLiabRows.slice(start);
            if (slice.length < 2 && !/^total\b|합계$/i.test(lbl)) continue;
            const tol = Math.max(3, Math.ceil((slice.length + 1) * 0.5));
            let checked = 0;
            let matched = 0;
            for (let yi = 0; yi < years.length; yi++) {
              const candVal = rows[i].values[yi];
              if (candVal === null) continue;
              const sliceSum = slice.reduce((acc, r) => acc + (r.values[yi] ?? 0), 0);
              checked++;
              if (Math.abs(sliceSum - candVal) <= tol) matched++;
            }
            if (checked > 0 && matched === checked) {
              isIntermediateSub = true;
              break;
            }
          }
          if (isIntermediateSub) continue;
        }
        nonCurrLiabRows.push(rows[i]);
      }
      if (nonCurrLiabRows.length > 0) {
        liabRow = {
          label: '부채합계(유동+비유동 산출)',
          values: years.map((_, yi) => {
            const clVal = rows[clIdx].values[yi];
            if (clVal === null) return null;
            return nonCurrLiabRows.reduce((acc, r) => acc + (r.values[yi] ?? 0), clVal);
          }),
        };
      }
    }
  }

  const missing = [];
  if (!assetsRow) missing.push('자산총계([[자산총계]])');
  if (!liabRow) missing.push('부채총계([[부채총계]])');
  if (!equityRow) missing.push('자본총계([[자본총계]])');

  if (missing.length > 0) {
    failures.push(
      makeFailure(
        ticker,
        'BS_IDENTITY',
        '1. 재무상태표 항등식',
        '재무상태표',
        'ALL',
        missing.join(', '),
        `필수 총계 행 누락 (${missing.join(', ')})`,
        `자산총계·부채총계·자본총계 행 모두 존재`
      )
    );
    return { status: 'FAIL', failures, checkedYears: 0 };
  }

  let checkedYears = 0;
  for (let yi = 0; yi < years.length; yi++) {
    const yr = years[yi];
    const assets = assetsRow.values[yi];
    const liab = liabRow.values[yi];
    const equity = equityRow.values[yi];
    const liabAndEq = liabAndEqRow ? liabAndEqRow.values[yi] : null;

    if (assets === null) continue;
    checkedYears++;

    if (liab === null || equity === null) {
      failures.push(
        makeFailure(
          ticker,
          'BS_IDENTITY',
          '1. 재무상태표 항등식',
          '재무상태표',
          yr,
          `${assetsRow.label} vs ${liabRow.label}+${equityRow.label}`,
          `부채=${formatNum(liab)}, 자본=${formatNum(equity)} (결측)`,
          `자산총계=${formatNum(assets)}`
        )
      );
      continue;
    }

    const mezz = mezzanineRows.reduce((acc, mr) => acc + (mr.values[yi] ?? 0), 0);
    const sumLE = liab + mezz + equity;
    const diff = sumLE - assets;
    if (Math.abs(diff) > BS_IDENTITY_TOLERANCE) {
      const extra = liabAndEq !== null ? `, 부채와자본총계=${formatNum(liabAndEq)}` : '';
      const mezzLabel = mezz !== 0 ? ' + 메자닌자본' : '';
      const mezzValStr = mezz !== 0 ? ` + 메자닌(${formatNum(mezz)})` : '';
      failures.push(
        makeFailure(
          ticker,
          'BS_IDENTITY',
          '1. 재무상태표 항등식',
          '재무상태표',
          yr,
          `${assetsRow.label} = ${liabRow.label}${mezzLabel} + ${equityRow.label}`,
          `부채(${formatNum(liab)})${mezzValStr} + 자본(${formatNum(equity)}) = ${formatNum(sumLE)}${extra}`,
          `자산총계 = ${formatNum(assets)} (오차: ${diff > 0 ? '+' : ''}${formatNum(diff)}, 허용: ±${BS_IDENTITY_TOLERANCE})`
        )
      );
    }

    if (liabAndEq !== null) {
      const leDiff = liabAndEq - assets;
      if (Math.abs(leDiff) > BS_IDENTITY_TOLERANCE) {
        failures.push(
          makeFailure(
            ticker,
            'BS_IDENTITY',
            '1. 재무상태표 항등식',
            '재무상태표',
            yr,
            `${assetsRow.label} = ${liabAndEqRow.label}`,
            `부채와자본총계 = ${formatNum(liabAndEq)}`,
            `자산총계 = ${formatNum(assets)} (오차: ${leDiff > 0 ? '+' : ''}${formatNum(leDiff)}, 허용: ±${BS_IDENTITY_TOLERANCE})`
          )
        );
      }
    }
  }

  if (checkedYears === 0) {
    failures.push(
      makeFailure(
        ticker,
        'BS_IDENTITY',
        '1. 재무상태표 항등식',
        '재무상태표',
        'ALL',
        assetsRow.label,
        '검증 가능한 연도별 자산총계 수치 없음 (0개 연도)',
        '최소 1개 이상 연도 수치 존재'
      )
    );
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
    checkedYears,
  };
}

// [검증 2] 주요 소계 합산 일치: 유동자산, 비유동자산, 유동부채, 비유동부채, 영업CF, 투자CF, 재무CF
function auditMajorSubtotals(ticker, sheets) {
  const failures = [];
  const breakdown = {
    유동자산: 'FAIL',
    비유동자산: 'N/A',
    유동부채: 'FAIL',
    비유동부채: 'N/A',
    영업현금흐름: 'FAIL',
    투자현금흐름: 'FAIL',
    재무현금흐름: 'FAIL',
  };

  const bs = sheets['재무상태표'];
  const cf = sheets['현금흐름표'];

  const findIdx = (rows, exactLabel, regex, excludeRegex = null) => {
    const exact = rows.findIndex((r) => r.label === exactLabel);
    if (exact !== -1) return exact;
    return rows.findIndex(
      (r) => regex.test(r.label) && (!excludeRegex || !excludeRegex.test(r.label))
    );
  };

  const evaluateSubtotal = ({ sheetName, rows, years, subtotalIdx, componentRows, keyName }) => {
    const subRow = rows[subtotalIdx];
    let evaluatedYears = 0;
    let itemFailed = false;

    for (let yi = 0; yi < years.length; yi++) {
      const yr = years[yi];
      const subVal = subRow.values[yi];
      const compVals = componentRows
        .map((r) => r.values[yi])
        .filter((v) => v !== null && v !== undefined);

      if (subVal === null && compVals.length === 0) continue;
      evaluatedYears++;

      if (subVal === null && compVals.length > 0) {
        itemFailed = true;
        failures.push(
          makeFailure(
            ticker,
            'SUBTOTALS',
            '2. 주요 소계 합산',
            sheetName,
            yr,
            subRow.label,
            `소계값=null (세부 ${compVals.length}개 항목 합계=${formatNum(
              compVals.reduce((a, b) => a + b, 0)
            )})`,
            '유효한 소계 수치 존재'
          )
        );
        continue;
      }

      if (subVal !== null && compVals.length === 0) {
        itemFailed = true;
        failures.push(
          makeFailure(
            ticker,
            'SUBTOTALS',
            '2. 주요 소계 합산',
            sheetName,
            yr,
            subRow.label,
            '세부항목 누락 (0개 항목)',
            `소계=${formatNum(subVal)}에 대응하는 세부항목 존재`
          )
        );
        continue;
      }

      const sum = compVals.reduce((a, b) => a + b, 0);
      const tol = Math.max(3, Math.ceil((compVals.length + 1) * 0.5));
      const diff = sum - subVal;
      if (Math.abs(diff) > tol) {
        itemFailed = true;
        failures.push(
          makeFailure(
            ticker,
            'SUBTOTALS',
            '2. 주요 소계 합산',
            sheetName,
            yr,
            subRow.label,
            `세부합계=${formatNum(sum)} (${compVals.length}개 항목)`,
            `소계=${formatNum(subVal)} (차이: ${diff > 0 ? '+' : ''}${formatNum(diff)}, 허용: ±${tol})`
          )
        );
      }
    }

    if (evaluatedYears === 0) {
      itemFailed = true;
      failures.push(
        makeFailure(
          ticker,
          'SUBTOTALS',
          '2. 주요 소계 합산',
          sheetName,
          'ALL',
          subRow.label,
          '검증 가능한 연도 데이터 없음',
          '최소 1개 이상 연도의 소계 및 세부항목 존재'
        )
      );
    }

    breakdown[keyName] = itemFailed ? 'FAIL' : 'PASS';
  };

  if (!bs || bs.rows.length === 0) {
    for (const k of ['유동자산', '유동부채']) {
      breakdown[k] = 'FAIL';
      failures.push(
        makeFailure(
          ticker,
          'SUBTOTALS',
          '2. 주요 소계 합산',
          '재무상태표',
          'ALL',
          k,
          '재무상태표 시트 없음',
          `${k} 소계 행 및 세부항목 존재`
        )
      );
    }
  } else {
    const bsRows = bs.rows;
    const bsYears = bs.years;
    const hasAssetsRow = bsRows.some(
      (r) =>
        r.label === '[[자산총계]]' ||
        (/^\[*자산총계\]*$|^총자산$/i.test(r.label) && !/부채/.test(r.label))
    );
    const hasAnyClassifiedSubtotal = bsRows.some((r) =>
      /^(?:\[)?(?:비)?유동(?:자산|부채)\s*합계(?:\])?$|^total (?:non-?current )?(?:assets|liabilities)$/i.test(
        r.label
      )
    );

    if (hasAssetsRow && !hasAnyClassifiedSubtotal) {
      breakdown['유동자산'] = 'N/A(미구분BS)';
      breakdown['비유동자산'] = 'N/A(미구분BS)';
      breakdown['유동부채'] = 'N/A(미구분BS)';
      breakdown['비유동부채'] = 'N/A(미구분BS)';
    } else {
      const caIdx = findIdx(
        bsRows,
        '[유동자산 합계]',
        /유동자산\s*합계|^total current assets$/i,
        /비유동/
      );
      if (caIdx === -1) {
        breakdown['유동자산'] = 'FAIL';
        failures.push(
          makeFailure(
            ticker,
            'SUBTOTALS',
            '2. 주요 소계 합산',
            '재무상태표',
            'ALL',
            '[유동자산 합계]',
            '소계 행 누락',
            '[유동자산 합계] 행 존재'
          )
        );
      } else {
        const caCompRows = [];
        for (let i = 0; i < caIdx; i++) {
          const r = bsRows[i];
          if (/^\[.*합계\]$/.test(r.label)) continue;
          const isIntermediateCandidate =
            caCompRows.length >= 2 &&
            (r.isBold ||
              /^(?:현금및단기금융자산|현금및현금성자산\s*합계|총현금및단기투자)$/i.test(r.label));
          if (isIntermediateCandidate) {
            let matchedYears = 0;
            let mismatchedYears = 0;
            for (let yi = 0; yi < bsYears.length; yi++) {
              const v = r.values[yi];
              if (v === null) continue;
              const prevSum = caCompRows.reduce((acc, pr) => acc + (pr.values[yi] ?? 0), 0);
              if (Math.abs(prevSum - v) <= 3) matchedYears++;
              else mismatchedYears++;
            }
            if (matchedYears >= 1 && mismatchedYears === 0) continue;
          }
          caCompRows.push(r);
        }
        evaluateSubtotal({
          sheetName: '재무상태표',
          rows: bsRows,
          years: bsYears,
          subtotalIdx: caIdx,
          componentRows: caCompRows,
          keyName: '유동자산',
        });
      }

      const clIdx = findIdx(
        bsRows,
        '[유동부채 합계]',
        /유동부채\s*합계|^total current liabilities$/i,
        /비유동/
      );
      if (clIdx === -1) {
        breakdown['유동부채'] = 'FAIL';
        failures.push(
          makeFailure(
            ticker,
            'SUBTOTALS',
            '2. 주요 소계 합산',
            '재무상태표',
            'ALL',
            '[유동부채 합계]',
            '소계 행 누락',
            '[유동부채 합계] 행 존재'
          )
        );
      } else {
        let prevTotalIdx = -1;
        for (let i = clIdx - 1; i >= 0; i--) {
          const lbl = bsRows[i].label;
          if (
            lbl === '[[자산총계]]' ||
            lbl === '[비유동자산 합계]' ||
            lbl === '[유동자산 합계]' ||
            (/^\[*(?:자산총계|비유동자산\s*합계|유동자산\s*합계)\]*$/i.test(lbl) &&
              !/부채/.test(lbl))
          ) {
            prevTotalIdx = i;
            break;
          }
        }
        const clCompRows = bsRows
          .slice(prevTotalIdx + 1, clIdx)
          .filter((r) => !/^\[.*합계\]$/.test(r.label));
        evaluateSubtotal({
          sheetName: '재무상태표',
          rows: bsRows,
          years: bsYears,
          subtotalIdx: clIdx,
          componentRows: clCompRows,
          keyName: '유동부채',
        });
      }

      const ncaIdx = findIdx(
        bsRows,
        '[비유동자산 합계]',
        /^\[?비유동자산\s*합계\]?$|^total non-?current assets$/i
      );
      if (ncaIdx !== -1 && caIdx !== -1 && ncaIdx > caIdx) {
        let ncaStart = caIdx + 1;
        const assetsIdx = bsRows.findIndex(
          (r) =>
            r.label === '[[자산총계]]' ||
            (/^\[*자산총계\]*$|^총자산$/i.test(r.label) && !/부채/.test(r.label))
        );
        if (assetsIdx !== -1 && ncaIdx < assetsIdx) {
          let directOk = 0;
          let directFail = 0;
          for (let yi = 0; yi < bsYears.length; yi++) {
            const caV = bsRows[caIdx].values[yi];
            const ncaV = bsRows[ncaIdx].values[yi];
            const aV = bsRows[assetsIdx].values[yi];
            if (caV === null || ncaV === null || aV === null) continue;
            if (Math.abs(caV + ncaV - aV) <= 3) directOk++;
            else directFail++;
          }
          if (directFail > 0 && directOk === 0) {
            for (let k = 1; k <= 3 && caIdx + 1 + k < ncaIdx; k++) {
              const lead = bsRows.slice(caIdx + 1, caIdx + 1 + k);
              if (lead.some((r) => /^\[.*합계\]$/.test(r.label))) break;
              let kOk = 0;
              let kFail = 0;
              for (let yi = 0; yi < bsYears.length; yi++) {
                const caV = bsRows[caIdx].values[yi];
                const ncaV = bsRows[ncaIdx].values[yi];
                const aV = bsRows[assetsIdx].values[yi];
                if (caV === null || ncaV === null || aV === null) continue;
                const leadSum = lead.reduce((acc, lr) => acc + (lr.values[yi] ?? 0), 0);
                if (Math.abs(caV + leadSum + ncaV - aV) <= 3) kOk++;
                else kFail++;
              }
              if (kOk > 0 && kFail === 0) {
                ncaStart = caIdx + 1 + k;
                break;
              }
            }
          }
        }
        const ncaCompRows = bsRows
          .slice(ncaStart, ncaIdx)
          .filter((r) => !/^\[.*합계\]$/.test(r.label));
        evaluateSubtotal({
          sheetName: '재무상태표',
          rows: bsRows,
          years: bsYears,
          subtotalIdx: ncaIdx,
          componentRows: ncaCompRows,
          keyName: '비유동자산',
        });
      }

      const nclIdx = findIdx(
        bsRows,
        '[비유동부채 합계]',
        /^\[?비유동부채\s*합계\]?$|^total (?:non-?current|long-?term) liabilities$/i
      );
      if (nclIdx !== -1 && clIdx !== -1 && nclIdx > clIdx) {
        const nclCompRows = bsRows
          .slice(clIdx + 1, nclIdx)
          .filter(
            (r) =>
              !/^\[.*\]$/.test(r.label) &&
              !/상환가능|메자닌|화해대상|redeemable|temporary\s+equity|liabilities\s+subject\s+to\s+compromise/i.test(
                r.label
              )
          );
        if (nclCompRows.length > 0) {
          evaluateSubtotal({
            sheetName: '재무상태표',
            rows: bsRows,
            years: bsYears,
            subtotalIdx: nclIdx,
            componentRows: nclCompRows,
            keyName: '비유동부채',
          });
        }
      }
    }
  }

  if (!cf || cf.rows.length === 0) {
    for (const k of ['영업현금흐름', '투자현금흐름', '재무현금흐름']) {
      breakdown[k] = 'FAIL';
      failures.push(
        makeFailure(
          ticker,
          'SUBTOTALS',
          '2. 주요 소계 합산',
          '현금흐름표',
          'ALL',
          k,
          '현금흐름표 시트 없음',
          `${k} 소계 행 및 세부항목 존재`
        )
      );
    }
  } else {
    const cfRows = cf.rows;
    const cfYears = cf.years;

    const opIdx = findIdx(
      cfRows,
      '[영업활동현금흐름]',
      /영업활동\s*현금흐름|net cash.*operating activities/i
    );
    const invIdx = findIdx(
      cfRows,
      '[투자활동현금흐름]',
      /투자활동\s*현금흐름|net cash.*investing activities/i
    );
    const finIdx = findIdx(
      cfRows,
      '[재무활동현금흐름]',
      /재무활동\s*현금흐름|net cash.*financing activities/i
    );
    const sortedCfSections = [opIdx, invIdx, finIdx].filter((i) => i !== -1).sort((a, b) => a - b);
    const prevSectionOf = (idx) => {
      let prev = -1;
      for (const s of sortedCfSections) {
        if (s < idx) prev = s;
      }
      return prev;
    };

    if (opIdx === -1) {
      breakdown['영업현금흐름'] = 'FAIL';
      failures.push(
        makeFailure(
          ticker,
          'SUBTOTALS',
          '2. 주요 소계 합산',
          '현금흐름표',
          'ALL',
          '[영업활동현금흐름]',
          '소계 행 누락',
          '[영업활동현금흐름] 행 존재'
        )
      );
    } else {
      const opStartBound = prevSectionOf(opIdx) + 1;
      const netIncIdx = cfRows.findIndex(
        (r, idx) =>
          idx >= opStartBound &&
          idx < opIdx &&
          /^(?:연결)?당기순이익$|^총?순이익$|^net\s+(?:income|earnings)/i.test(r.label)
      );
      const preContOpsIdx =
        netIncIdx !== -1
          ? cfRows.findIndex(
              (r, idx) =>
                idx >= opStartBound &&
                idx < netIncIdx &&
                /^계속영업(?:순|당기)?손?이익$|계속영업.*(?:당기순이익|순손실|순이익)|income\s*\(?loss\)?\s*from\s+continuing\s+operations/i.test(
                  r.label
                )
            )
          : -1;
      const contOpsIdx = cfRows.findIndex(
        (r, idx) =>
          idx >= opStartBound &&
          idx < opIdx &&
          (idx > netIncIdx || netIncIdx === -1) &&
          /^계속영업(?:순|당기)?손?이익$|계속영업.*(?:당기순이익|순손실|순이익)|income\s*\(?loss\)?\s*from\s+continuing\s+operations/i.test(
            r.label
          )
      );
      let effectiveOpStart = -1;
      if (netIncIdx !== -1) {
        effectiveOpStart = preContOpsIdx !== -1 ? preContOpsIdx : netIncIdx;
        if (contOpsIdx > netIncIdx && contOpsIdx <= Math.min(opIdx - 1, netIncIdx + 5)) {
          const leadSlice = cfRows.slice(netIncIdx, contOpsIdx);
          let verifiedWaterfall = false;
          for (let yi = 0; yi < cfYears.length; yi++) {
            const cVal = cfRows[contOpsIdx].values[yi];
            const nVal = cfRows[netIncIdx].values[yi];
            if (cVal === null || nVal === null) continue;
            const directSum = leadSlice.reduce((a, r) => a + (r.values[yi] ?? 0), 0);
            if (Math.abs(directSum - cVal) <= 3) {
              verifiedWaterfall = true;
              break;
            }
            if (leadSlice.length === 2) {
              const bVal = leadSlice[1].values[yi] ?? 0;
              if (bVal !== 0 && Math.abs(nVal - bVal - cVal) <= 3) {
                verifiedWaterfall = true;
                break;
              }
            } else if (leadSlice.length === 3) {
              const b1 = leadSlice[1].values[yi] ?? 0;
              const b2 = leadSlice[2].values[yi] ?? 0;
              if (
                Math.abs(nVal - b1 - b2 - cVal) <= 3 ||
                Math.abs(nVal - b1 + b2 - cVal) <= 3 ||
                Math.abs(nVal + b1 - b2 - cVal) <= 3
              ) {
                verifiedWaterfall = true;
                break;
              }
            } else if (leadSlice.length >= 4 && leadSlice.length <= 6) {
              let lastNetIdxInLead = -1;
              for (let li = leadSlice.length - 1; li >= 0; li--) {
                if (/^(?:연결)?당기순이익$|^총?순이익$/i.test(leadSlice[li].label || '')) {
                  lastNetIdxInLead = li;
                  break;
                }
              }
              if (lastNetIdxInLead > 0 && lastNetIdxInLead < leadSlice.length - 1) {
                const subNetVal = leadSlice[lastNetIdxInLead].values[yi];
                const tailSlice = leadSlice.slice(lastNetIdxInLead + 1);
                if (subNetVal !== null && tailSlice.length <= 2) {
                  if (tailSlice.length === 1) {
                    const b1 = tailSlice[0].values[yi] ?? 0;
                    if (Math.abs(subNetVal + b1 - cVal) <= 3 || Math.abs(subNetVal - b1 - cVal) <= 3) {
                      verifiedWaterfall = true;
                      break;
                    }
                  } else if (tailSlice.length === 2) {
                    const b1 = tailSlice[0].values[yi] ?? 0;
                    const b2 = tailSlice[1].values[yi] ?? 0;
                    if (
                      Math.abs(subNetVal + b1 + b2 - cVal) <= 3 ||
                      Math.abs(subNetVal - b1 - b2 - cVal) <= 3 ||
                      Math.abs(subNetVal - b1 + b2 - cVal) <= 3 ||
                      Math.abs(subNetVal + b1 - b2 - cVal) <= 3
                    ) {
                      verifiedWaterfall = true;
                      break;
                    }
                  }
                }
              }
            }
          }
          if (verifiedWaterfall) effectiveOpStart = contOpsIdx;
        }
      } else if (contOpsIdx !== -1) {
        effectiveOpStart = contOpsIdx;
      }
      const isExcludedOpRow = (lbl) =>
        /^\[.*\]$/.test(lbl) ||
        /^(?:\[)?조정항목\s*합계(?:\])?$|^total\s+adjustments(?:\s+to\s+reconcile.*)?$/i.test(lbl);
      let opCompRows;
      if (effectiveOpStart !== -1) {
        const rawSlice = cfRows.slice(effectiveOpStart, opIdx);
        const skippedInterm = new Set();
        for (let i = 1; i < rawSlice.length; i++) {
          if (isExcludedOpRow(rawSlice[i].label)) continue;
          const prevItems = [];
          for (let j = 1; j < i; j++) {
            if (!isExcludedOpRow(rawSlice[j].label) && !skippedInterm.has(j)) {
              prevItems.push(rawSlice[j]);
            }
          }
          if (prevItems.length < 2) continue;
          let okCnt = 0;
          let totCnt = 0;
          const tol = Math.max(3, Math.ceil((prevItems.length + 1) * 0.5));
          for (let yi = 0; yi < cfYears.length; yi++) {
            const tv = rawSlice[i].values[yi];
            if (tv === null || tv === 0 || !prevItems.some((p) => p.values[yi] !== null)) continue;
            totCnt++;
            const s = prevItems.reduce((acc, p) => acc + (p.values[yi] ?? 0), 0);
            if (Math.abs(s - tv) <= tol) okCnt++;
          }
          if (totCnt >= 2 && okCnt / totCnt >= 0.8) skippedInterm.add(i);
        }
        opCompRows = rawSlice.filter((r, idx) => !isExcludedOpRow(r.label) && !skippedInterm.has(idx));
      } else {
        opCompRows = cfRows
          .slice(opStartBound, opIdx)
          .filter(
            (r) =>
              !isExcludedOpRow(r.label) &&
              !/기초\s*현금|beginning of period/i.test(r.label)
          );
      }
      evaluateSubtotal({
        sheetName: '현금흐름표',
        rows: cfRows,
        years: cfYears,
        subtotalIdx: opIdx,
        componentRows: opCompRows,
        keyName: '영업현금흐름',
      });
    }

    if (invIdx === -1) {
      breakdown['투자현금흐름'] = 'FAIL';
      failures.push(
        makeFailure(
          ticker,
          'SUBTOTALS',
          '2. 주요 소계 합산',
          '현금흐름표',
          'ALL',
          '[투자활동현금흐름]',
          '소계 행 누락',
          '[투자활동현금흐름] 행 존재'
        )
      );
    } else {
      const invStart = prevSectionOf(invIdx) + 1;
      const invCompRows = cfRows.slice(invStart, invIdx).filter((r) => !/^\[.*\]$/.test(r.label));
      evaluateSubtotal({
        sheetName: '현금흐름표',
        rows: cfRows,
        years: cfYears,
        subtotalIdx: invIdx,
        componentRows: invCompRows,
        keyName: '투자현금흐름',
      });
    }

    if (finIdx === -1) {
      breakdown['재무현금흐름'] = 'FAIL';
      failures.push(
        makeFailure(
          ticker,
          'SUBTOTALS',
          '2. 주요 소계 합산',
          '현금흐름표',
          'ALL',
          '[재무활동현금흐름]',
          '소계 행 누락',
          '[재무활동현금흐름] 행 존재'
        )
      );
    } else {
      const finStart = prevSectionOf(finIdx) + 1;
      const finCompRows = cfRows.slice(finStart, finIdx).filter((r) => !/^\[.*\]$/.test(r.label));
      evaluateSubtotal({
        sheetName: '현금흐름표',
        rows: cfRows,
        years: cfYears,
        subtotalIdx: finIdx,
        componentRows: finCompRows,
        keyName: '재무현금흐름',
      });
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
    breakdown,
  };
}

// [검증 3] 단위 급변 탐지: 직전 연도 대비 1,000배 이상 급변 이상치 탐지
function auditScaleJumps(ticker, sheets) {
  const failures = [];

  for (const sheetName of SHEET_NAMES) {
    const sheet = sheets[sheetName];
    if (!sheet) continue;
    const { years, rows } = sheet;

    let maxScanRowIdx = rows.length;
    let invStart = -1;
    let invIdx = -1;
    const invReconciles = years.map(() => false);
    if (sheetName === '현금흐름표') {
      const actIndices = rows
        .map((r, idx) => (/^\[(?:영업|투자|재무)활동현금흐름\]$/.test(r.label) ? idx : -1))
        .filter((x) => x >= 0);
      if (actIndices.length > 0) maxScanRowIdx = Math.max(...actIndices);
      invIdx = rows.findIndex((r) => r.label === '[투자활동현금흐름]');
      if (invIdx !== -1) {
        for (let k = invIdx - 1; k >= 0; k--) {
          if (/^\[(?:영업|투자|재무)활동현금흐름\]$/.test(rows[k].label)) {
            invStart = k;
            break;
          }
        }
        const invComps = rows.slice(invStart + 1, invIdx).filter((r) => !/^\[.*\]$/.test(r.label));
        const invTol = Math.max(3, Math.ceil((invComps.length + 1) * 0.5));
        for (let yi = 0; yi < years.length; yi++) {
          const tv = rows[invIdx].values[yi];
          if (tv === null || !invComps.some((c) => c.values[yi] !== null)) continue;
          const sum = invComps.reduce((acc, c) => acc + (c.values[yi] ?? 0), 0);
          if (Math.abs(sum - tv) <= invTol) invReconciles[yi] = true;
        }
      }
    }

    for (let rIdx = 0; rIdx < maxScanRowIdx; rIdx++) {
      const row = rows[rIdx];
      const isInvMaRow =
        sheetName === '현금흐름표' &&
        invIdx !== -1 &&
        rIdx > invStart &&
        rIdx < invIdx &&
        /기업인수|사업인수|매각|처분|인수.*현금|acquisitions?|divestitures?|business\s+combination|sale\s+of\s+business/i.test(
          row.label
        );
      for (let i = 1; i < years.length; i++) {
        const prev = row.values[i - 1];
        const curr = row.values[i];
        if (prev === null || curr === null || prev === 0 || curr === 0) continue;
        if (isInvMaRow && invReconciles[i - 1] && invReconciles[i]) continue;

        const maxAbs = Math.max(Math.abs(prev), Math.abs(curr));
        const minAbs = Math.min(Math.abs(prev), Math.abs(curr));
        const ratio = maxAbs / minAbs;

        if (ratio >= SCALE_JUMP_THRESHOLD && (minAbs >= 5 || maxAbs >= 100000)) {
          failures.push(
            makeFailure(
              ticker,
              'SCALE_JUMP',
              '3. 단위 급변 탐지',
              sheetName,
              `${years[i - 1]}→${years[i]}`,
              row.label,
              `${years[i]}년 ${formatNum(curr)} (${Math.round(ratio).toLocaleString('en-US')}배 변동)`,
              `직전연도(${years[i - 1]}년 ${formatNum(prev)}) 대비 1,000배 미만`
            )
          );
        }
      }
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
  };
}

// [검증 4] 현금유출 부호 오류 탐지
function auditCashOutflowSigns(ticker, sheets) {
  const failures = [];
  const cf = sheets['현금흐름표'];
  if (!cf) {
    failures.push(
      makeFailure(
        ticker,
        'OUTFLOW_SIGN',
        '4. 현금유출 부호',
        '현금흐름표',
        'ALL',
        '(시트 누락)',
        '현금흐름표 시트 없음',
        '현금흐름표 존재'
      )
    );
    return { status: 'FAIL', failures };
  }

  const exclusionRegex =
    /순증감|순액|증감|유입|유출입|순발행|순상환|순차입|순조달|발행\(순\)|발행\s*대금|발행\s*금액|투자금\s*상환|(?:우선주|지분)\s*투자.*?상환|(?:발행|차입)\s*(?:및|·|\()\s*상환|처분|매각및|사업매각|만기도래|원금회수|만기회수|조기상환|만기상환|회수|조정|수익|상각|손익|손실|이익|미지급|^\[.*\]$|proceeds|net\s+change|increase|decrease|issuance|amortization|extinguishment\s+of\s+debt|sale\s+of|maturities|collections|accrued|payable|beginning|end\s+of/i;
  const outflowRegex =
    /(?:CAPEX|자본적\s*지출|유형자산의?\s*(?:취득|매입)|생산설비(?:자산)?의?\s*취득|무형자산의?\s*취득|배당금의?\s*지급|현금배당(?:금)?\s*지급|자사주\s*매입|자기주식의?\s*(?:취득|매입)|(?:채무증권|지분증권|유가증권|단기금융상품|투자자산|대출채권|리스채권|비지배지분).*?(?:취득|매입)|기업인수로\s*인한\s*현금유출(?!입)|관련\s*현금유출(?!입)|(?:차입금|사채|부채|리스부채|우선주).*?상환|상환에\s*따른\s*현금유출(?!입)|원천징수세액\s*납부|이연금융원가의?\s*지급|^(?:purchases?\s+of|additions?\s+to|capital\s+expenditures?|payments?\s+(?:to\s+acquire|for\s+repurchase|of\s+dividends|for\s+dividends|on\s+long-?term\s+debt|on\s+.*lease)|repayments?\s+of|repurchases?\s+of|common\s+stock\s+repurchased|treasury\s+stock\s+purchases?|acquisitions?,?\s*net\s+of\s+cash|business\s+acquisitions?|cash\s+dividends\s+paid|dividends\s+paid|redemption\s+of|retirement\s+of)|\bcapital\s+expenditures?\b|\bdividends?\s+paid\b|\brepurchases?\s+of\s+(?:common|treasury|capital)\s+stock\b)/i;

  const { years, rows } = cf;
  const findCfIdx = (exact, re) => {
    const e = rows.findIndex((r) => r.label === exact);
    return e !== -1 ? e : rows.findIndex((r) => re.test(r.label));
  };
  const opIdx = findCfIdx(
    '[영업활동현금흐름]',
    /영업활동\s*현금흐름|net cash.*operating activities/i
  );
  const invIdx = findCfIdx(
    '[투자활동현금흐름]',
    /투자활동\s*현금흐름|net cash.*investing activities/i
  );
  const finIdx = findCfIdx(
    '[재무활동현금흐름]',
    /재무활동\s*현금흐름|net cash.*financing activities/i
  );
  const sortedCfSections = [opIdx, invIdx, finIdx].filter((i) => i !== -1).sort((a, b) => a - b);
  let opStart = 0;
  if (opIdx !== -1) {
    for (const s of sortedCfSections) {
      if (s < opIdx) opStart = s + 1;
    }
  }
  let invStart = 0;
  if (invIdx !== -1) {
    for (const s of sortedCfSections) {
      if (s < invIdx) invStart = s + 1;
    }
  }
  const lastActivityIdx = Math.max(opIdx, invIdx, finIdx);
  const activityRows = lastActivityIdx !== -1 ? rows.slice(0, lastActivityIdx) : rows;
  for (let ri = 0; ri < activityRows.length; ri++) {
    const row = activityRows[ri];
    if (/^\[.*\]$/.test(row.label)) continue;
    if (exclusionRegex.test(row.label)) continue;
    if (
      opIdx !== -1 &&
      ri >= opStart &&
      ri < opIdx &&
      /(?:차입금|사채|부채|리스부채|우선주).*?상환/i.test(row.label)
    ) {
      continue;
    }
    if (
      invIdx !== -1 &&
      ri >= invStart &&
      ri < invIdx &&
      /(?:투자|우선주).*?상환/i.test(row.label)
    ) {
      continue;
    }
    if (!outflowRegex.test(row.label)) continue;

    for (let yi = 0; yi < years.length; yi++) {
      const val = row.values[yi];
      if (val !== null && val > 0) {
        if (
          /^기업인수로\s*인한\s*현금유출$/i.test(row.label) &&
          row.values.some((v) => v !== null && v < 0) &&
          invIdx !== -1
        ) {
          const invCompRows = rows.slice(invStart, invIdx).filter((r) => !/^\[.*\]$/.test(r.label));
          const invSubVal = rows[invIdx].values[yi];
          const invCompVals = invCompRows
            .map((r) => r.values[yi])
            .filter((v) => v !== null && v !== undefined);
          if (invSubVal !== null && invCompVals.length > 0) {
            const invSum = invCompVals.reduce((a, b) => a + b, 0);
            const invTol = Math.max(3, Math.ceil((invCompVals.length + 1) * 0.5));
            if (
              Math.abs(invSum - invSubVal) <= invTol &&
              Math.abs(invSum - 2 * val - invSubVal) > invTol
            ) {
              continue;
            }
          }
        }
        failures.push(
          makeFailure(
            ticker,
            'OUTFLOW_SIGN',
            '4. 현금유출 부호',
            '현금흐름표',
            years[yi],
            row.label,
            `+${formatNum(val)} (양수)`,
            '<= 0 (음수 현금유출)'
          )
        );
      }
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
  };
}

// [검증 5] 미번역 영문 잔존 검사
function auditUntranslatedLabels(ticker, sheets) {
  const failures = [];
  const escapedAcronyms = ALLOWLISTED_ACRONYMS.map((a) =>
    a.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  )
    .sort((a, b) => b.length - a.length)
    .join('|');
  const allowlistedParenRegex = new RegExp(
    `\\((?:${escapedAcronyms})(?:[\\s,/&]+(?:${escapedAcronyms}))*\\)`,
    'g'
  );
  const allowlistedStandaloneRegex = new RegExp(
    `(?<![A-Za-z])(?:${escapedAcronyms})(?![A-Za-z])`,
    'g'
  );
  const xbrlPrefixRegex = /(?:us-gaap|dei|defref|srt|xbrl)[_:]/i;
  const camelCaseTagRegex = /(?:^|[\s(\[])[A-Z][a-z]+[A-Z][A-Za-z0-9]*/;

  for (const sheetName of SHEET_NAMES) {
    const sheet = sheets[sheetName];
    if (!sheet) continue;

    for (const row of sheet.rows) {
      const label = row.label;
      const hasKorean = /[가-힣]/.test(label);
      const hasXbrlPrefix = xbrlPrefixRegex.test(label);
      const hasCamelTag = camelCaseTagRegex.test(label);

      const cleaned = hasKorean
        ? label.replace(allowlistedParenRegex, '').replace(allowlistedStandaloneRegex, '')
        : label;
      const latinTokens = cleaned.match(/[A-Za-z]+/g) || [];

      if (hasXbrlPrefix || hasCamelTag || latinTokens.length > 0) {
        const reasons = [];
        if (hasXbrlPrefix) reasons.push('XBRL 접두사 포함');
        if (hasCamelTag) reasons.push('CamelCase XBRL 태그 의심');
        if (latinTokens.length > 0) {
          reasons.push(`영문 잔존: [${latinTokens.slice(0, 6).join(', ')}]`);
        }

        failures.push(
          makeFailure(
            ticker,
            'UNTRANSLATED',
            '5. 미번역 영문',
            sheetName,
            `ALL(A열:${row.rowIndex}행)`,
            label,
            `미번역 영문/태그 잔존 (${reasons.join('; ')})`,
            '순수 한글 계정명'
          )
        );
      }
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
  };
}

// [검증 6] BRK-A 전용 구조 검사
function auditBrkAStructure(ticker, sheets) {
  if (ticker !== 'BRK-A') {
    return { status: 'N/A', failures: [] };
  }
  const is = sheets['손익계산서'];
  const bs = sheets['재무상태표'];
  const cf = sheets['현금흐름표'];
  const failures = [];
  if (!is || !bs || !cf) {
    failures.push(makeFailure(ticker, 'BRK_STRUCTURE', '6. BRK-A 구조', '-', 'ALL', '시트', '누락', '존재'));
  }
  return { status: failures.length === 0 ? 'PASS' : 'FAIL', failures };
}

// [검증 7] 손익계산서 정합성·부호·연간매출 급변 감사 (Tax Benefit 정상 환급 허용)
function auditIncomeStatementIntegrity(ticker, sheets, capture = null) {
  const failures = [];
  const is = sheets['손익계산서'];
  if (!is || is.rows.length === 0) {
    failures.push(
      makeFailure(
        ticker,
        'IS_INTEGRITY',
        '7. IS 정합성',
        '손익계산서',
        'ALL',
        '(시트 누락)',
        '손익계산서 시트 없음 또는 빈 시트',
        '손익계산서 존재'
      )
    );
    return { status: 'FAIL', failures };
  }

  const { years, rows } = is;
  const firstMetaIdx = rows.findIndex((r) =>
    /주당순이익|발행주식수|주당포괄손익|주당현금배당|per\s+share/i.test(r.label)
  );
  const mainRows = firstMetaIdx !== -1 ? rows.slice(0, firstMetaIdx) : rows;
  const isSegRow = (r) => /^\[.*\]/.test(r.label);
  const findMain = (exactLabels, regex = null, startAt = 0) => {
    const list = Array.isArray(exactLabels) ? exactLabels : [exactLabels];
    for (const lbl of list) {
      const idx = mainRows.findIndex((r, i) => i >= startAt && !isSegRow(r) && r.label === lbl);
      if (idx !== -1) return idx;
    }
    if (regex) {
      return mainRows.findIndex((r, i) => i >= startAt && !isSegRow(r) && regex.test(r.label));
    }
    return -1;
  };

  const pushFail = (year, account, actual, expected) => {
    failures.push(
      makeFailure(ticker, 'IS_INTEGRITY', '7. IS 정합성', '손익계산서', year, account, actual, expected)
    );
  };

  const inferSignVectors = (componentRows, baseRow, targetRow, _startIdx, signFn) => {
    const k = componentRows.length;
    if (k === 0 || k > 5) return null;

    const completeYearIndices = [];
    for (let yi = years.length - 1; yi >= 0 && completeYearIndices.length < 3; yi--) {
      if (targetRow.values[yi] === null) continue;
      if (baseRow && baseRow.values[yi] === null) continue;
      const anyNonZero = componentRows.some((cr) => cr.values[yi] !== null && cr.values[yi] !== 0);
      if (anyNonZero) completeYearIndices.push(yi);
    }
    if (completeYearIndices.length === 0) return null;

    const masks = [];
    masks.push(new Array(k).fill(1));
    masks.push(new Array(k).fill(-1));
    for (let m = 0; m < 1 << k; m++) {
      const vec = [];
      for (let b = 0; b < k; b++) vec.push((m >> b) & 1 ? -1 : 1);
      if (vec.every((v) => v === 1) || vec.every((v) => v === -1)) continue;
      masks.push(vec);
    }

    const tol = Math.max(3, Math.ceil((k + 2) * 0.5));
    let bestVec = null;
    let bestMaxErr = Infinity;
    for (const vec of masks) {
      let maxErr = 0;
      for (const yi of completeYearIndices) {
        const err = Math.abs(signFn(yi, vec) - targetRow.values[yi]);
        if (err > maxErr) maxErr = err;
      }
      if (maxErr <= tol && maxErr < bestMaxErr) {
        bestMaxErr = maxErr;
        bestVec = vec;
        if (maxErr === 0) break;
      }
    }
    return bestVec ? { vec: bestVec, tol } : null;
  };

  const revIdx = findMain(['매출액', '영업수익']);
  const cogsIdx = findMain(['매출원가']);
  const gpIdx = findMain(['매출총이익']);
  const opExpIdx = findMain(['영업비용합계', '총비용합계', '총원가및영업비용', '영업비용']);
  const opIdx = findMain(['영업이익']);

  const intIncIdx = findMain(['이자수익']);
  const intExpIdx = findMain(['이자비용']);
  const netIntIdx = findMain(['순이자손익']);
  const nonIntIncIdx = findMain(['비이자수익합계']);
  const netRevIdx = findMain(['순영업수익(이자비용차감후총수익)']);
  const provIdx = findMain(['신용손실충당금전입액(대손상각비)']);
  const nonIntExpIdx = findMain(['비이자비용합계']);
  const preTaxIdx = findMain(['법인세비용차감전순이익']);

  if (
    netRevIdx !== -1 &&
    intIncIdx !== -1 &&
    intExpIdx !== -1 &&
    netIntIdx !== -1 &&
    intIncIdx < intExpIdx &&
    intExpIdx < netIntIdx
  ) {
    const intFit = inferSignVectors(
      [mainRows[intExpIdx]],
      mainRows[intIncIdx],
      mainRows[netIntIdx],
      0,
      (yi, v) => mainRows[intIncIdx].values[yi] + v[0] * (mainRows[intExpIdx].values[yi] ?? 0)
    );
    if (intFit) {
      years.forEach((yr, yi) => {
        const a = mainRows[intIncIdx].values[yi];
        const b = mainRows[intExpIdx].values[yi];
        const t = mainRows[netIntIdx].values[yi];
        if (a !== null && b !== null && t !== null) {
          const calc = a + intFit.vec[0] * b;
          if (Math.abs(calc - t) > intFit.tol) {
            pushFail(
              yr,
              '이자수익→순이자손익',
              `계산=${formatNum(calc)}`,
              `순이자손익=${formatNum(t)} (차이: ${formatNum(calc - t)})`
            );
          }
        }
      });
    }
    if (nonIntIncIdx !== -1 && netIntIdx < nonIntIncIdx && nonIntIncIdx < netRevIdx) {
      years.forEach((yr, yi) => {
        const a = mainRows[netIntIdx].values[yi];
        const b = mainRows[nonIntIncIdx].values[yi];
        const t = mainRows[netRevIdx].values[yi];
        if (a !== null && b !== null && t !== null && Math.abs(a + b - t) > 3) {
          pushFail(
            yr,
            '순이자손익+비이자수익=순영업수익',
            `합계=${formatNum(a + b)}`,
            `순영업수익=${formatNum(t)}`
          );
        }
      });
    }
    if (nonIntExpIdx !== -1 && preTaxIdx !== -1 && netRevIdx < nonIntExpIdx && nonIntExpIdx < preTaxIdx) {
      const comps =
        provIdx !== -1 && provIdx > netRevIdx && provIdx < nonIntExpIdx
          ? [mainRows[provIdx], mainRows[nonIntExpIdx]]
          : [mainRows[nonIntExpIdx]];
      const fit = inferSignVectors(comps, mainRows[netRevIdx], mainRows[preTaxIdx], 0, (yi, v) =>
        comps.reduce((acc, cr, ci) => acc + v[ci] * (cr.values[yi] ?? 0), mainRows[netRevIdx].values[yi])
      );
      if (fit) {
        years.forEach((yr, yi) => {
          const base = mainRows[netRevIdx].values[yi];
          const t = mainRows[preTaxIdx].values[yi];
          if (base === null || t === null || comps.some((c) => c.values[yi] === null)) return;
          const calc = comps.reduce((acc, cr, ci) => acc + fit.vec[ci] * cr.values[yi], base);
          if (Math.abs(calc - t) > fit.tol) {
            pushFail(
              yr,
              '순영업수익→법인세차감전순이익',
              `계산=${formatNum(calc)}`,
              `세전이익=${formatNum(t)} (차이: ${formatNum(calc - t)})`
            );
          }
        });
      }
    }
  } else {
    if (revIdx !== -1 && cogsIdx !== -1 && gpIdx !== -1 && revIdx < cogsIdx && cogsIdx < gpIdx) {
      const gpFit = inferSignVectors([mainRows[cogsIdx]], mainRows[revIdx], mainRows[gpIdx], 0, (yi, v) =>
        mainRows[revIdx].values[yi] + v[0] * (mainRows[cogsIdx].values[yi] ?? 0)
      );
      if (gpFit) {
        years.forEach((yr, yi) => {
          const r = mainRows[revIdx].values[yi];
          const c = mainRows[cogsIdx].values[yi];
          const g = mainRows[gpIdx].values[yi];
          if (r !== null && c !== null && g !== null) {
            const calc = r + gpFit.vec[0] * c;
            if (Math.abs(calc - g) > gpFit.tol) {
              pushFail(
                yr,
                '매출액-매출원가=매출총이익',
                `계산=${formatNum(calc)}`,
                `매출총이익=${formatNum(g)} (차이: ${formatNum(calc - g)})`
              );
            }
          }
        });
      }
    }

    if (opIdx !== -1) {
      const baseForOpIdx = gpIdx !== -1 && gpIdx < opIdx ? gpIdx : revIdx;
      if (baseForOpIdx !== -1 && baseForOpIdx < opIdx) {
        let expCompRows = [];
        let postOpExpRows = [];
        if (opExpIdx !== -1 && opExpIdx > baseForOpIdx && opExpIdx < opIdx) {
          postOpExpRows = mainRows.slice(opExpIdx + 1, opIdx).filter((r) => !isSegRow(r));
          const recentStartYi = Math.max(0, years.length - 3);
          const recentPostRows = postOpExpRows.filter((r) =>
            r.values.slice(recentStartYi).some((v) => v !== null && v !== 0)
          );
          if (recentPostRows.length > 0 && recentPostRows.length <= 4) {
            expCompRows = [mainRows[opExpIdx], ...recentPostRows];
          } else {
            expCompRows = [mainRows[opExpIdx]];
          }
        } else {
          expCompRows = mainRows.slice(baseForOpIdx + 1, opIdx).filter((r) => !isSegRow(r));
        }
        const opFit = inferSignVectors(expCompRows, mainRows[baseForOpIdx], mainRows[opIdx], 0, (yi, v) =>
          expCompRows.reduce(
            (acc, cr, ci) => acc + v[ci] * (cr.values[yi] ?? 0),
            mainRows[baseForOpIdx].values[yi]
          )
        );
        if (opFit) {
          const recentStartYi = Math.max(0, years.length - 3);
          const anchorExpIndices = expCompRows
            .map((cr, ci) => (cr.values.slice(recentStartYi).some((v) => v !== null) ? ci : -1))
            .filter((ci) => ci !== -1);
          const expSet = new Set(expCompRows);
          const histOnlyPostRows = postOpExpRows.filter((r) => !expSet.has(r));
          years.forEach((yr, yi) => {
            const b = mainRows[baseForOpIdx].values[yi];
            const op = mainRows[opIdx].values[yi];
            if (b === null || op === null) return;
            const anyVal = expCompRows.some((cr) => cr.values[yi] !== null);
            if (!anyVal) return;
            const calc = expCompRows.reduce((acc, cr, ci) => acc + opFit.vec[ci] * (cr.values[yi] ?? 0), b);
            if (Math.abs(calc - op) <= opFit.tol) return;
            if (anchorExpIndices.length > 0 && anchorExpIndices.length < expCompRows.length) {
              const calcAnchor = anchorExpIndices.reduce(
                (acc, ci) => acc + opFit.vec[ci] * (expCompRows[ci].values[yi] ?? 0),
                b
              );
              if (Math.abs(calcAnchor - op) <= opFit.tol) return;
            }
            if (histOnlyPostRows.some((r) => r.values[yi] !== null && r.values[yi] !== 0)) {
              const histSum = histOnlyPostRows.reduce((acc, r) => acc + (r.values[yi] ?? 0), 0);
              const histTol = Math.max(
                opFit.tol,
                Math.ceil((expCompRows.length + histOnlyPostRows.length + 2) * 0.5)
              );
              if (Math.abs(calc + histSum - op) <= histTol || Math.abs(calc - histSum - op) <= histTol) {
                return;
              }
            }
            pushFail(
              yr,
              `${mainRows[baseForOpIdx].label}→영업이익`,
              `계산=${formatNum(calc)}`,
              `영업이익=${formatNum(op)} (차이: ${calc - op > 0 ? '+' : ''}${formatNum(calc - op)})`
            );
          });
        }
      }
    }
  }

  if (opIdx !== -1 && preTaxIdx !== -1 && opIdx < preTaxIdx) {
    const nonOpTotIdx = mainRows.findIndex(
      (r, i) => i > opIdx && i < preTaxIdx && !isSegRow(r) && r.label === '영업외손익합계'
    );
    const comps =
      nonOpTotIdx !== -1
        ? [mainRows[nonOpTotIdx]]
        : mainRows.slice(opIdx + 1, preTaxIdx).filter((r) => !isSegRow(r));
    if (nonOpTotIdx !== -1) {
      const subComps = mainRows.slice(opIdx + 1, nonOpTotIdx).filter((r) => !isSegRow(r));
      const subFit = inferSignVectors(subComps, null, mainRows[nonOpTotIdx], 0, (yi, v) =>
        subComps.reduce((acc, cr, ci) => acc + v[ci] * (cr.values[yi] ?? 0), 0)
      );
      if (subFit) {
        years.forEach((yr, yi) => {
          const t = mainRows[nonOpTotIdx].values[yi];
          if (t === null || !subComps.some((c) => c.values[yi] !== null)) return;
          const calc = subComps.reduce((acc, cr, ci) => acc + subFit.vec[ci] * (cr.values[yi] ?? 0), 0);
          if (Math.abs(calc - t) > subFit.tol) {
            pushFail(yr, '영업외손익합계', `세부합계=${formatNum(calc)}`, `영업외손익합계=${formatNum(t)}`);
          }
        });
      }
    }
    const preTaxFit = inferSignVectors(comps, mainRows[opIdx], mainRows[preTaxIdx], 0, (yi, v) =>
      comps.reduce((acc, cr, ci) => acc + v[ci] * (cr.values[yi] ?? 0), mainRows[opIdx].values[yi])
    );
    if (preTaxFit) {
      const recentStartYi = Math.max(0, years.length - 3);
      const anchorCompIndices = comps
        .map((cr, ci) => (cr.values.slice(recentStartYi).some((v) => v !== null) ? ci : -1))
        .filter((ci) => ci !== -1);
      years.forEach((yr, yi) => {
        const op = mainRows[opIdx].values[yi];
        const pt = mainRows[preTaxIdx].values[yi];
        if (op === null || pt === null) return;
        const calc = comps.reduce((acc, cr, ci) => acc + preTaxFit.vec[ci] * (cr.values[yi] ?? 0), op);
        if (Math.abs(calc - pt) <= preTaxFit.tol) return;
        if (anchorCompIndices.length > 0 && anchorCompIndices.length < comps.length) {
          const calcAnchor = anchorCompIndices.reduce(
            (acc, ci) => acc + preTaxFit.vec[ci] * (comps[ci].values[yi] ?? 0),
            op
          );
          if (Math.abs(calcAnchor - pt) <= preTaxFit.tol) return;
        }
        pushFail(
          yr,
          '영업이익→법인세비용차감전순이익',
          `계산=${formatNum(calc)}`,
          `법인세비용차감전순이익=${formatNum(pt)} (차이: ${calc - pt > 0 ? '+' : ''}${formatNum(calc - pt)})`
        );
      });
    }
  }

  const taxIdx = preTaxIdx !== -1 ? findMain(['법인세비용'], null, preTaxIdx + 1) : -1;
  if (preTaxIdx !== -1 && taxIdx !== -1) {
    const netCandidatesBase = mainRows.filter(
      (r, i) =>
        i > taxIdx &&
        !isSegRow(r) &&
        (r.label === '계속영업순이익' || r.label === '연결당기순이익' || r.label === '당기순이익')
    );
    const hasDiscAfterMeta =
      firstMetaIdx !== -1 &&
      rows.some((r, i) => i >= firstMetaIdx && /중단영업|계속영업/.test(r.label));
    const contAfterMeta =
      hasDiscAfterMeta && !netCandidatesBase.some((r) => r.label === '계속영업순이익')
        ? rows.find((r, i) => i >= firstMetaIdx && !isSegRow(r) && r.label === '계속영업순이익')
        : null;
    const netCandidates = contAfterMeta ? [contAfterMeta, ...netCandidatesBase] : netCandidatesBase;
    if (netCandidates.length > 0) {
      const firstNetIdx = mainRows.findIndex(
        (r, i) =>
          i > taxIdx &&
          !isSegRow(r) &&
          (r.label === '계속영업순이익' || r.label === '연결당기순이익' || r.label === '당기순이익')
      );
      const betweenTaxAndNet =
        firstNetIdx > taxIdx + 1
          ? mainRows.slice(taxIdx + 1, firstNetIdx).filter((r) => !isSegRow(r))
          : [];
      const getNetAt = (yi) => {
        for (const nc of netCandidates) {
          if (nc.values[yi] !== null) return nc.values[yi];
        }
        return null;
      };
      if (betweenTaxAndNet.length === 0) {
        let minusOk = 0;
        let plusOk = 0;
        for (let yi = years.length - 1; yi >= Math.max(0, years.length - 4); yi--) {
          const pt = mainRows[preTaxIdx].values[yi];
          const tx = mainRows[taxIdx].values[yi];
          const net = getNetAt(yi);
          if (pt === null || tx === null || tx === 0 || net === null) continue;
          if (Math.abs(pt - tx - net) <= 3) minusOk++;
          if (Math.abs(pt + tx - net) <= 3) plusOk++;
        }
        if ((minusOk > 0) !== (plusOk > 0)) {
          const taxOp = minusOk > plusOk ? -1 : 1;
          years.forEach((yr, yi) => {
            const pt = mainRows[preTaxIdx].values[yi];
            const tx = mainRows[taxIdx].values[yi];
            const net = getNetAt(yi);
            if (pt === null || tx === null || net === null) return;
            const calc = pt + taxOp * tx;
            if (Math.abs(calc - net) > 3) {
              pushFail(
                yr,
                '법인세비용차감전순이익→순이익 (법인세 산식)',
                `계산=${formatNum(calc)} (세전 ${formatNum(pt)} ${taxOp === -1 ? '-' : '+'} 법인세 ${formatNum(tx)})`,
                `순이익=${formatNum(net)} (차이: ${calc - net > 0 ? '+' : ''}${formatNum(calc - net)})`
              );
            }
          });
        }
      } else {
        const combos = [
          { taxOp: -1, btwOp: 1, ok: 0, fail: 0 },
          { taxOp: 1, btwOp: 1, ok: 0, fail: 0 },
          { taxOp: -1, btwOp: -1, ok: 0, fail: 0 },
          { taxOp: 1, btwOp: -1, ok: 0, fail: 0 },
        ];
        const btwSumAt = (yi) => betweenTaxAndNet.reduce((acc, r) => acc + (r.values[yi] ?? 0), 0);
        for (let yi = years.length - 1; yi >= Math.max(0, years.length - 4); yi--) {
          const pt = mainRows[preTaxIdx].values[yi];
          const tx = mainRows[taxIdx].values[yi];
          const net = getNetAt(yi);
          if (pt === null || tx === null || tx === 0 || net === null) continue;
          const bSum = btwSumAt(yi);
          for (const c of combos) {
            if (Math.abs(pt + c.taxOp * tx + c.btwOp * bSum - net) <= 3) c.ok++;
            else c.fail++;
          }
        }
        const matchedCombos = combos.filter((c) => c.ok > 0 && c.fail === 0);
        if (matchedCombos.length === 1) {
          const { taxOp, btwOp } = matchedCombos[0];
          years.forEach((yr, yi) => {
            const pt = mainRows[preTaxIdx].values[yi];
            const tx = mainRows[taxIdx].values[yi];
            const net = getNetAt(yi);
            if (pt === null || tx === null || net === null) return;
            const bSum = btwSumAt(yi);
            const calc = pt + taxOp * tx + btwOp * bSum;
            if (Math.abs(calc - net) > 3) {
              pushFail(
                yr,
                '법인세비용차감전순이익→순이익 (법인세 산식)',
                `계산=${formatNum(calc)} (세전 ${formatNum(pt)} ${taxOp === -1 ? '-' : '+'} 법인세 ${formatNum(tx)} + 세후항목 ${formatNum(btwOp * bSum)})`,
                `순이익=${formatNum(net)} (차이: ${calc - net > 0 ? '+' : ''}${formatNum(calc - net)})`
              );
            }
          });
        }
      }
    }
  }

  const primaryRevRowIdx = revIdx !== -1 ? revIdx : netRevIdx;
  if (primaryRevRowIdx !== -1) {
    const rRow = mainRows[primaryRevRowIdx];
    const factsObj = capture?.companyFacts?.facts;
    const fyeMonth = getFiscalYearEndMonth(capture?.submissions);
    const revSecTags = [
      'Revenues',
      'RevenueFromContractWithCustomerExcludingAssessedTax',
      'RevenueFromContractWithCustomerIncludingAssessedTax',
      'SalesRevenueNet',
      'RevenuesNetOfInterestExpense',
      'RegulatedAndUnregulatedOperatingRevenue',
      'ElectricUtilityRevenue',
      'OilAndGasRevenue',
      'SalesRevenueGoodsNet',
    ];
    const matchesSecAnnualRev = (yr, excelVal) => {
      if (!factsObj?.['us-gaap']) return false;
      for (const tag of revSecTags) {
        const fact = getSecAnnualFactValueForTag(factsObj, tag, yr, false, fyeMonth);
        if (!fact || !(fact.valM > 0)) continue;
        if (Math.abs(excelVal - fact.valM) <= Math.max(2, Math.abs(fact.valM) * 0.01)) {
          return true;
        }
      }
      return false;
    };
    for (let i = 1; i < years.length; i++) {
      const prev = rRow.values[i - 1];
      const curr = rRow.values[i];
      if (prev === null || curr === null || prev <= 0 || curr <= 0) continue;
      const ratio = Math.max(prev, curr) / Math.min(prev, curr);
      if (ratio >= 3.0) {
        if (matchesSecAnnualRev(years[i - 1], prev) && matchesSecAnnualRev(years[i], curr)) {
          continue;
        }
        pushFail(
          `${years[i - 1]}→${years[i]}`,
          `${rRow.label} (매출 급변)`,
          `${years[i - 1]}년=${formatNum(prev)}, ${years[i]}년=${formatNum(curr)} (${ratio.toFixed(1)}배 변동)`,
          '연간 매출 연속성 유지 (< 3.0배)'
        );
      }
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
  };
}

// =========================================================================================
// [검증 8] SEC 원본 수치 교차 대조 (Ground-Truth Check via CompanyFacts 10-K)
// =========================================================================================

function getFiscalYearEndMonth(submissions) {
  const fye = String(submissions?.fiscalYearEnd || '').trim();
  if (/^\d{4}$/.test(fye)) {
    const m = parseInt(fye.slice(0, 2), 10);
    if (m >= 1 && m <= 12) return m;
  }
  return null;
}

function circularMonthDiff(m1, m2) {
  const d = Math.abs(m1 - m2);
  return Math.min(d, 12 - d);
}

function getSecAnnualFactValueForTag(factsObj, tag, targetYear, isInstant, fyeMonth) {
  const usGaap = factsObj?.['us-gaap'] || {};
  const concept = usGaap[tag];
  const records = concept?.units?.USD;
  if (!Array.isArray(records) || records.length === 0) return null;

  const targetYrNum = parseInt(String(targetYear), 10);

  let matches = records.filter((r) => {
    if (r.form !== '10-K' && r.form !== '10-K/A') return false;
    if (typeof r.val !== 'number' || !Number.isFinite(r.val)) return false;

    if (isInstant) {
      if (r.start || !r.end) return false;
    } else {
      if (!r.start || !r.end) return false;
      const days = (new Date(r.end) - new Date(r.start)) / 86400000;
      if (!Number.isFinite(days) || days < 300 || days > 400) return false;
    }

    let recYr = parseInt(r.end.slice(0, 4), 10);
    if (/^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(r.end)) recYr -= 1;
    return recYr === targetYrNum;
  });

  if (matches.length === 0) return null;

  if (fyeMonth) {
    const nearFye = matches.filter((r) => {
      const endM = parseInt(r.end.slice(5, 7), 10);
      return circularMonthDiff(endM, fyeMonth) <= 1;
    });
    if (nearFye.length > 0) matches = nearFye;
  }

  // 동일 회계연도(fyMatch) 내에서 복수의 10-K/A 수정보고서나 비교연도가 존재할 경우
  // 가장 최근(최신 10-K) 보고서의 filed 일자 기준 값을 우선 채택하고, 전 수치는 allValuesM / allRecords에 보존한다.
  matches.sort((a, b) => {
    const filedCmp = (a.filed || '').localeCompare(b.filed || '');
    if (filedCmp !== 0) return filedCmp;
    return (a.end || '').localeCompare(b.end || '');
  });

  const latest =
    [...matches].reverse().find((m) => m.form === '10-K') || matches[matches.length - 1];
  return {
    tag,
    valRaw: latest.val,
    valM: latest.val / 1e6,
    allValuesM: matches.map((m) => m.val / 1e6),
    allRecords: matches.map((m) => ({ valM: m.val / 1e6, filed: m.filed || '', form: m.form || '' })),
    end: latest.end,
    start: latest.start || null,
    filed: latest.filed || '',
    form: latest.form || '',
  };
}

function getSecAnnualFactBest(factsObj, tagCandidates, targetYear, isInstant, fyeMonth, preferMaxOverride = false) {
  const foundList = [];
  for (const tag of tagCandidates) {
    const res = getSecAnnualFactValueForTag(factsObj, tag, targetYear, isInstant, fyeMonth);
    if (res) foundList.push(res);
  }
  if (foundList.length === 0) return null;

  const primary = foundList[0];
  if (preferMaxOverride && foundList.length > 1) {
    const revTagSet = new Set([
      'Revenues',
      'RegulatedAndUnregulatedOperatingRevenue',
      'RevenueFromContractWithCustomerExcludingAssessedTax',
      'RevenueFromContractWithCustomerIncludingAssessedTax',
      'SalesRevenueNet',
    ]);
    const revItems = foundList.filter((c) => revTagSet.has(c.tag));
    // 10-K 본문보다 180일 이상 후속 공시된 신규 합병/재작성 10-K 수치가 0.5x~2.0x 범위라면 역합병/회계기준 변경 후속 10-K 값을 채택한다.
    const bestNonZeroRev = [...revItems]
      .filter((c) => Math.abs(c.valM) > 0)
      .sort((a, b) => (b.filed || '').localeCompare(a.filed || ''))[0];
    if (revItems.length > 1 && primary.filed) {
      const primaryMs = new Date(primary.filed).getTime();
      const newerRestated = revItems
        .filter((c) => {
          if (!c.filed || Math.abs(c.valM) <= 0 || Math.abs(primary.valM) <= 0) return false;
          const diffDays = (new Date(c.filed).getTime() - primaryMs) / 86400000;
          const ratio = Math.abs(c.valM) / Math.abs(primary.valM);
          return (
            diffDays >= 180 &&
            (c.form === '10-K' || !primary.form || primary.form !== '10-K') &&
            ratio >= 0.5 &&
            ratio <= 2.0
          );
        })
        .sort((a, b) => (b.filed || '').localeCompare(a.filed || ''));
      if (newerRestated.length > 0 && newerRestated[0] !== primary) {
        return {
          ...newerRestated[0],
          allCandidateValuesM: foundList.flatMap((c) => c.allValuesM),
          allCandidateRecords: foundList.flatMap((c) => c.allRecords || []),
        };
      }
    }
    if (bestNonZeroRev && Math.abs(primary.valM) === 0) {
      return {
        ...bestNonZeroRev,
        allCandidateValuesM: foundList.flatMap((c) => c.allValuesM),
        allCandidateRecords: foundList.flatMap((c) => c.allRecords || []),
      };
    }
    const maxItem = foundList.reduce((best, cur) =>
      Math.abs(cur.valM) > Math.abs(best.valM) ? cur : best
    );
    if (
      revTagSet.has(maxItem.tag) &&
      Math.abs(primary.valM) > 0 &&
      Math.abs(maxItem.valM) > Math.abs(primary.valM) * 1.5
    ) {
      return {
        ...maxItem,
        allCandidateValuesM: foundList.flatMap((c) => c.allValuesM),
        allCandidateRecords: foundList.flatMap((c) => c.allRecords || []),
      };
    }
  }
  return primary;
}

// [신규 감사 A·B] 동일 시트 다개년 동일 수치 중복 행 탐지 & BS 주식수(sharesItemType) 혼입 탐지
function auditStructuralAnomalies(ticker, sheets, translations = null) {
  const failures = [];
  const segOf = (lbl) => {
    const m = String(lbl || '').match(/^(\[[^\]]+\])\s+/);
    return m ? m[1] : '';
  };
  const isSubtotalRow = (r) =>
    Boolean(
      r.isBold ||
        /^\[.*\]$/.test(r.label) ||
        /(?:손익|비용|수익|자산|부채|자본|투자|현금흐름)\s*합계$/.test(r.label)
    );
  const isEpsOrShares = (lbl) => /주당순이익|발행주식수|주식수/i.test(lbl);
  const isIsCoreSummary = (lbl) =>
    /^(?:매출액|영업수익|매출총이익|영업이익|법인세비용차감전순이익|계속영업순이익|연결당기순이익|당기순이익|보통주귀속당기순이익|총포괄손익|포괄손익)$|보통주.*(?:기본|희석).*순이익|희석\s*(?:당기)?순이익|포괄손익/i.test(
      lbl
    );
  const isSoleActiveComponentOfSubtotal = (rows, idxA, idxB, matchedYearIndices, yearsLength) => {
    const lo = Math.min(idxA, idxB);
    const hi = Math.max(idxA, idxB);
    if (!isSubtotalRow(rows[hi]) || isSubtotalRow(rows[lo])) return false;
    let prevSub = -1;
    for (let k = hi - 1; k >= 0; k--) {
      if (isSubtotalRow(rows[k])) {
        prevSub = k;
        break;
      }
    }
    if (lo <= prevSub) return false;
    for (let startIdx = prevSub + 1; startIdx <= lo; startIdx++) {
      const compRows = [];
      for (let k = startIdx; k < hi; k++) {
        if (!isSubtotalRow(rows[k])) compRows.push({ idx: k, row: rows[k] });
      }
      if (compRows.length === 0) continue;
      let othersOk = true;
      for (const yi of matchedYearIndices) {
        let othersSum = 0;
        for (const c of compRows) {
          if (c.idx === lo) continue;
          othersSum += c.row.values[yi] ?? 0;
        }
        if (Math.abs(othersSum) > 2) {
          othersOk = false;
          break;
        }
      }
      if (!othersOk) continue;
      const tol = Math.max(3, Math.ceil((compRows.length + 1) * 0.5));
      let validYears = 0;
      let okYears = 0;
      let recentFail = false;
      for (let yi = 0; yi < yearsLength; yi++) {
        const subVal = rows[hi].values[yi];
        const compVal = rows[lo].values[yi];
        if (subVal === null || compVal === null) continue;
        validYears++;
        let fullSum = 0;
        for (const c of compRows) fullSum += c.row.values[yi] ?? 0;
        if (Math.abs(fullSum - subVal) <= tol) {
          okYears++;
        } else if (yi >= yearsLength - 3) {
          recentFail = true;
        }
      }
      if (validYears > 0 && !recentFail && okYears / validYears >= 0.8) return true;
    }
    return false;
  };

  // Rule B: BS 내 주식수(sharesItemType) 라벨·태그 혼입 탐지
  const flaggedBsRows = new Set();
  const bs = sheets['재무상태표'];
  if (bs) {
    for (const row of bs.rows) {
      if (
        /주식수|주수|\bauthorized\s+shares\b|\bissued\s+shares\b|\boutstanding\s+shares\b|\btreasury.*shares\b/i.test(
          row.label
        )
      ) {
        flaggedBsRows.add(row.rowIndex);
        failures.push(
          makeFailure(
            ticker,
            'STRUCTURAL_ANOMALY',
            '구조/중복 이상',
            '재무상태표',
            `ALL(${row.rowIndex}행)`,
            row.label,
            '재무상태표(BS) 내 주식수(비금액) 행 혼입',
            'BS에는 금액(USD) 행만 포함 (sharesItemType 배제)'
          )
        );
      }
    }
  }
  if (translations?.entries) {
    for (const entry of translations.entries) {
      if (entry.sheetName !== '재무상태표') continue;
      if (flaggedBsRows.has(entry.rowIndex)) continue;
      if (/(?:SharesAuthorized|SharesIssued|SharesOutstanding|TreasuryStock(?:Common|Preferred)?Shares)$/i.test(entry.xbrlTag || '')) {
        flaggedBsRows.add(entry.rowIndex);
        failures.push(
          makeFailure(
            ticker,
            'STRUCTURAL_ANOMALY',
            '구조/중복 이상',
            '재무상태표',
            `ALL(${entry.rowIndex}행)`,
            `${entry.korName} (${entry.xbrlTag})`,
            '재무상태표(BS) 내 주식수(sharesItemType) 태그 행 혼입',
            'BS에는 금액(USD) 태그 행만 포함'
          )
        );
      }
    }
  }

  // Rule A: 동일 시트 내 5개년 이상 0이 아닌 수치가 100% 동일한 중복 행 탐지
  for (const sheetName of SHEET_NAMES) {
    const sheet = sheets[sheetName];
    if (!sheet) continue;
    const { rows, years } = sheet;
    const bsAssetsTotalIdx =
      sheetName === '재무상태표'
        ? rows.findIndex(
            (r) =>
              r.label === '[[자산총계]]' ||
              (/^\[*자산총계\]*$|^총자산$/i.test(r.label) && !/부채/.test(r.label))
          )
        : -1;
    const cfNetChangeIdx =
      sheetName === '현금흐름표'
        ? rows.findIndex((r) => /현금및현금성자산의\s*순증감|현금.*순증감/.test(r.label))
        : -1;
    for (let i = 0; i < rows.length; i++) {
      for (let j = i + 1; j < rows.length; j++) {
        const ra = rows[i];
        const rb = rows[j];
        if (segOf(ra.label) !== segOf(rb.label)) continue;
        if (ra.isBold && rb.isBold) continue;
        if (isEpsOrShares(ra.label) && isEpsOrShares(rb.label)) continue;
        const isDivPerShare = (lbl) =>
          /주당.*배당|배당.*주당|dividends?.*per\s+share/i.test(String(lbl || ''));
        if (sheetName === '손익계산서' && isDivPerShare(ra.label) && isDivPerShare(rb.label)) continue;
        if (sheetName === '손익계산서' && isIsCoreSummary(ra.label) && isIsCoreSummary(rb.label)) continue;
        if (sheetName === '현금흐름표' && (/기초|기말/.test(ra.label) || /기초|기말/.test(rb.label))) continue;
        const normLbl = (lbl) => String(lbl || '').replace(/\s+/g, '');
        if (
          sheetName === '현금흐름표' &&
          cfNetChangeIdx !== -1 &&
          ((i > cfNetChangeIdx && j > cfNetChangeIdx) ||
            (i < cfNetChangeIdx && j > cfNetChangeIdx && normLbl(ra.label) !== normLbl(rb.label)))
        ) {
          continue;
        }
        if (sheetName === '재무상태표' && bsAssetsTotalIdx !== -1 && i < bsAssetsTotalIdx && j > bsAssetsTotalIdx) {
          continue;
        }
        if (
          sheetName === '재무상태표' &&
          ((/지배기업주주지분|지배주주지분/.test(ra.label) && /자본총계/.test(rb.label)) ||
            (/지배기업주주지분|지배주주지분/.test(rb.label) && /자본총계/.test(ra.label)))
        ) {
          continue;
        }

        let matchedCount = 0;
        let bothNonNullCount = 0;
        let maxAbsMatched = 0;
        let hasDecimalVal = false;
        const matchedYearIndices = [];
        const distinctVals = new Set();
        for (let yi = 0; yi < years.length; yi++) {
          const va = ra.values[yi];
          const vb = rb.values[yi];
          if (va !== null && vb !== null) bothNonNullCount++;
          if (va !== null && vb !== null && va !== 0 && vb !== 0 && va === vb) {
            matchedCount++;
            matchedYearIndices.push(yi);
            distinctVals.add(va);
            if (Math.abs(va) > maxAbsMatched) maxAbsMatched = Math.abs(va);
            if (va % 1 !== 0) hasDecimalVal = true;
          }
        }
        if (
          matchedCount >= 5 &&
          bothNonNullCount > 0 &&
          matchedCount / bothNonNullCount >= 0.85 &&
          distinctVals.size >= 2 &&
          (maxAbsMatched >= 5 || hasDecimalVal) &&
          !isSoleActiveComponentOfSubtotal(rows, i, j, matchedYearIndices, years.length)
        ) {
          failures.push(
            makeFailure(
              ticker,
              'STRUCTURAL_ANOMALY',
              '구조/중복 이상',
              sheetName,
              `ALL(${ra.rowIndex}·${rb.rowIndex}행)`,
              `${ra.label} ↔ ${rb.label}`,
              `${matchedCount}개년 0 이외 수치 100% 동일 (예: ${Array.from(distinctVals)
                .slice(0, 3)
                .map(formatNum)
                .join(', ')})`,
              '독립 계정 간 5개년 이상 수치 중복 백필 없음'
            )
          );
        }
      }
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
  };
}

function auditSecGroundTruth(ticker, sheets, capture) {
  const failures = [];
  const comparisons = [];
  const factsObj = capture?.companyFacts?.facts;
  if (!factsObj || !factsObj['us-gaap']) {
    failures.push(
      makeFailure(
        ticker,
        'SEC_GROUND_TRUTH',
        '8. SEC 원본 대조',
        '-',
        'ALL',
        'CompanyFacts US-GAAP',
        'US-GAAP CompanyFacts 데이터 없음',
        '10-K US-GAAP 공식 수치 존재'
      )
    );
    return { status: 'FAIL', failures, comparisons, sampledYears: [] };
  }

  const fyeMonth = getFiscalYearEndMonth(capture?.submissions);
  const is = sheets['손익계산서'];
  const bs = sheets['재무상태표'];
  const cf = sheets['현금흐름표'];

  const findRow = (sheet, exactList, fallbackRegex, excludeRegex = null, pickLast = false) => {
    if (!sheet || !sheet.rows) return null;
    const nonSeg = sheet.rows.filter((r) => !/^\[.*\]/.test(r.label));
    const pool = exactList.some((lbl) => /^\[/.test(lbl)) ? sheet.rows : nonSeg;
    for (const lbl of exactList) {
      const m = pool.filter((r) => r.label === lbl);
      if (m.length > 0) return pickLast ? m[m.length - 1] : m[0];
    }
    if (fallbackRegex) {
      const m = pool.filter(
        (r) => fallbackRegex.test(r.label) && (!excludeRegex || !excludeRegex.test(r.label))
      );
      if (m.length > 0) return pickLast ? m[m.length - 1] : m[0];
    }
    return null;
  };

  const revRow = findRow(
    is,
    ['매출액', '영업수익', '순영업수익(이자비용차감후총수익)', '총수익'],
    /^(?:총)?매출액$|^영업수익$|^총수익$/i,
    /원가|비용|차감전/
  );
  let virtualBankRevRow = null;
  if (!revRow && is?.rows) {
    const netIntIncRow = findRow(is, ['순이자손익']);
    const nonIntIncRow = findRow(is, ['비이자수익합계']);
    if (netIntIncRow && nonIntIncRow) {
      virtualBankRevRow = {
        label: '순영업수익(순이자+비이자 산출)',
        values: is.years.map((_, yi) => {
          const ni = netIntIncRow.values[yi];
          const non = nonIntIncRow.values[yi];
          return ni !== null && non !== null ? ni + non : null;
        }),
      };
    }
  }
  const effectiveRevRow = revRow || virtualBankRevRow;
  const isVirtualBankRevenue = !revRow && Boolean(virtualBankRevRow);
  const findBestPopulatedRow = (sheet, exactList, fallbackRegex, excludeRegex = null) => {
    if (!sheet || !sheet.rows) return null;
    const nonSeg = sheet.rows.filter((r) => !/^\[.*\]/.test(r.label));
    const pool = exactList.some((lbl) => /^\[/.test(lbl)) ? sheet.rows : nonSeg;
    const candidates = [];
    for (const lbl of exactList) {
      candidates.push(...pool.filter((r) => r.label === lbl));
    }
    if (candidates.length === 0 && fallbackRegex) {
      candidates.push(
        ...pool.filter((r) => fallbackRegex.test(r.label) && (!excludeRegex || !excludeRegex.test(r.label)))
      );
    }
    if (candidates.length === 0) return null;
    return candidates.reduce((best, cur) => {
      const bestCnt = best.values.filter((v) => v !== null).length;
      const curCnt = cur.values.filter((v) => v !== null).length;
      return curCnt > bestCnt ? cur : best;
    });
  };
  const netIncRow = findBestPopulatedRow(
    is,
    ['당기순이익', '연결당기순이익', '보통주귀속당기순이익'],
    /^(?:연결)?당기순이익$|^순이익$/i,
    /주당|비지배|차감전|계속영업/
  );
  const assetsRow = findRow(bs, ['[[자산총계]]'], /^\[*자산총계\]*$|^총자산$/i, /부채/);
  const equityRow = findRow(bs, ['[[자본총계]]'], /^\[*자본총계\]*$|^총자본$/i, /부채/, true);
  const ocfRow = findRow(
    cf,
    ['[영업활동현금흐름]'],
    /영업활동\s*현금흐름/i,
    /계속영업|중단영업/
  ) || findRow(cf, ['[계속영업 영업활동현금흐름]'], /영업활동\s*현금흐름/i);

  const hasNciSplitInBs = Boolean(
    bs?.rows?.some((r) => /지배기업주주지분|지배주주지분|비지배지분/i.test(r.label))
  );
  const liabRow = findRow(bs, ['[[부채총계]]'], /^\[*부채총계\]*$|^총부채$/i, /자본/);
  const curLiabRow = findRow(bs, ['[유동부채 합계]'], /^\[?유동부채\s*합계\]?$/i);
  const nonCurLiabRow = findRow(bs, ['[비유동부채 합계]'], /^\[?비유동부채\s*합계\]?$/i, /기타비유동부채/);
  const getLiabAt = (yi) => {
    if (yi === null || yi === undefined || yi < 0) return null;
    if (liabRow && liabRow.values[yi] !== null && liabRow.values[yi] !== undefined) {
      return liabRow.values[yi];
    }
    const cl = curLiabRow?.values?.[yi];
    const ncl = nonCurLiabRow?.values?.[yi];
    if (cl !== null && cl !== undefined && ncl !== null && ncl !== undefined) return cl + ncl;
    if (assetsRow?.values?.[yi] !== null && equityRow?.values?.[yi] !== null) {
      return assetsRow.values[yi] - equityRow.values[yi];
    }
    return null;
  };

  const refYears = is?.years || bs?.years || cf?.years || [];
  const coreRowsForYrScan = [effectiveRevRow, netIncRow, assetsRow, equityRow, ocfRow];
  const populatedIndices = [];
  for (let yi = 0; yi < refYears.length; yi++) {
    const presentCnt = coreRowsForYrScan.filter((r) => r && r.values[yi] !== null).length;
    if (presentCnt >= 2) populatedIndices.push(yi);
  }

  let sampleIndices = [];
  if (populatedIndices.length <= 4) {
    sampleIndices = populatedIndices;
  } else {
    const n = populatedIndices.length;
    sampleIndices = [
      ...new Set([
        populatedIndices[0],
        populatedIndices[Math.floor((n - 1) / 2)],
        populatedIndices[n - 1],
      ]),
    ].sort((a, b) => a - b);
  }

  const sampledYears = sampleIndices.map((yi) => refYears[yi]);

  const isBankRevenue = effectiveRevRow?.label === '순영업수익(이자비용차감후총수익)';
  const metrics = [
    {
      metric: '매출액',
      sheetName: '손익계산서',
      sheetObj: is,
      excelRow: effectiveRevRow,
      isInstant: false,
      preferMaxOverride: !isBankRevenue && !isVirtualBankRevenue,
      tagCandidates: isBankRevenue
        ? [
            'RevenuesNetOfInterestExpense',
            'Revenues',
            'RevenueFromContractWithCustomerExcludingAssessedTax',
            'SalesRevenueNet',
          ]
        : isVirtualBankRevenue
        ? ['RevenuesNetOfInterestExpense', 'Revenues']
        : [
            'Revenues',
            'RegulatedAndUnregulatedOperatingRevenue',
            'RevenueFromContractWithCustomerExcludingAssessedTax',
            'RevenueFromContractWithCustomerIncludingAssessedTax',
            'SalesRevenueNet',
            'RevenuesNetOfInterestExpense',
            'ElectricUtilityRevenue',
            'OilAndGasRevenue',
            'SalesRevenueGoodsNet',
          ],
    },
    {
      metric: '당기순이익',
      sheetName: '손익계산서',
      sheetObj: is,
      excelRow: netIncRow,
      isInstant: false,
      preferMaxOverride: false,
      tagCandidates:
        netIncRow?.label === '연결당기순이익'
          ? ['ProfitLoss', 'NetIncomeLoss']
          : netIncRow?.label === '보통주귀속당기순이익'
          ? ['NetIncomeLossAvailableToCommonStockholdersBasic', 'NetIncomeLoss', 'ProfitLoss']
          : ['NetIncomeLoss', 'ProfitLoss', 'NetIncomeLossAvailableToCommonStockholdersBasic'],
    },
    {
      metric: '자산총계',
      sheetName: '재무상태표',
      sheetObj: bs,
      excelRow: assetsRow,
      isInstant: true,
      preferMaxOverride: false,
      tagCandidates: ['Assets'],
    },
    {
      metric: '자본총계',
      sheetName: '재무상태표',
      sheetObj: bs,
      excelRow: equityRow,
      isInstant: true,
      preferMaxOverride: false,
      resolveTagCandidates: (yr, yi) => {
        const allEqTags = [
          'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
          'StockholdersEquity',
          'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest',
          'PartnersCapital',
        ];
        // 자산 - 부채로 실제 총자본 기준(부모주주 vs 비지배지분 포함, 일반 vs 파트너십) 확인
        const aVal = assetsRow?.values?.[yi];
        const lVal = getLiabAt(yi);
        if (aVal !== null && aVal !== undefined && lVal !== null && lVal !== undefined) {
          const impliedTotalEq = aVal - lVal;
          const scored = allEqTags
            .map((t) => ({ tag: t, fact: getSecAnnualFactValueForTag(factsObj, t, yr, true, fyeMonth) }))
            .filter((x) => x.fact)
            .sort((a, b) => Math.abs(a.fact.valM - impliedTotalEq) - Math.abs(b.fact.valM - impliedTotalEq));
          if (scored.length > 0) {
            const ordered = scored.map((x) => x.tag);
            for (const t of allEqTags) {
              if (!ordered.includes(t)) ordered.push(t);
            }
            return ordered;
          }
        }
        return hasNciSplitInBs
          ? allEqTags
          : [
              'StockholdersEquity',
              'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
              'PartnersCapital',
              'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest',
            ];
      },
    },
    {
      metric: '영업활동현금흐름',
      sheetName: '현금흐름표',
      sheetObj: cf,
      excelRow: ocfRow,
      isInstant: false,
      preferMaxOverride: false,
      tagCandidates: [
        'NetCashProvidedByUsedInOperatingActivities',
        'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
      ],
    },
  ];

  for (const m of metrics) {
    if (!m.excelRow) {
      const anyYear = sampledYears[sampledYears.length - 1];
      const candidates = m.resolveTagCandidates
        ? m.resolveTagCandidates(anyYear, sampleIndices[sampleIndices.length - 1])
        : m.tagCandidates;
      const probeFact = anyYear
        ? getSecAnnualFactBest(factsObj, candidates, anyYear, m.isInstant, fyeMonth, m.preferMaxOverride)
        : null;
      if (probeFact) {
        failures.push(
          makeFailure(
            ticker,
            'SEC_GROUND_TRUTH',
            '8. SEC 원본 대조',
            m.sheetName,
            String(anyYear),
            m.metric,
            `엑셀 행 누락 (추출 실패)`,
            `SEC=${formatNum(Number(probeFact.valM.toFixed(1)))}M (us-gaap:${probeFact.tag})`
          )
        );
        comparisons.push({
          metric: m.metric,
          year: String(anyYear),
          excelLabel: '(누락)',
          excelVal: null,
          secTag: probeFact.tag,
          secValM: Number(probeFact.valM.toFixed(1)),
          diff: null,
          status: 'FAIL',
        });
      }
      continue;
    }

    for (const yi of sampleIndices) {
      const yr = refYears[yi];
      const sheetYi = m.sheetObj?.years ? m.sheetObj.years.indexOf(yr) : yi;
      const excelVal = sheetYi >= 0 ? m.excelRow.values[sheetYi] : null;
      const candidates = m.resolveTagCandidates
        ? m.resolveTagCandidates(yr, sheetYi)
        : m.tagCandidates;

      let secFact = null;
      if (m.metric === '매출액' && isVirtualBankRevenue) {
        const niiFact = getSecAnnualFactBest(
          factsObj,
          ['InterestIncomeExpenseNet'],
          yr,
          false,
          fyeMonth,
          false
        );
        const nonIntFact = getSecAnnualFactBest(
          factsObj,
          ['NoninterestIncome'],
          yr,
          false,
          fyeMonth,
          false
        );
        if (niiFact && nonIntFact) {
          secFact = {
            tag: 'InterestIncomeExpenseNet+NoninterestIncome',
            valM: niiFact.valM + nonIntFact.valM,
            filed: niiFact.filed || nonIntFact.filed,
          };
        }
      }
      if (!secFact) {
        secFact = getSecAnnualFactBest(
          factsObj,
          candidates,
          yr,
          m.isInstant,
          fyeMonth,
          m.preferMaxOverride
        );
      }

      if (!secFact) {
        comparisons.push({
          metric: m.metric,
          year: String(yr),
          excelLabel: m.excelRow.label,
          excelVal,
          secTag: '(SEC N/A)',
          secValM: null,
          diff: null,
          status: 'N/A',
        });
        continue;
      }

      let effectiveSecValM = secFact.valM;
      let effectiveSecTagNote = '';
      if (m.metric === '자본총계' && Array.isArray(secFact.allRecords) && secFact.filed) {
        const secAssetsFactForYr = getSecAnnualFactBest(factsObj, ['Assets'], yr, true, fyeMonth);
        if (secAssetsFactForYr?.filed && secFact.filed > secAssetsFactForYr.filed) {
          const bsAligned = secFact.allRecords.filter(
            (r) => r.filed && r.filed <= secAssetsFactForYr.filed
          );
          if (bsAligned.length > 0) {
            effectiveSecValM = bsAligned[bsAligned.length - 1].valM;
            effectiveSecTagNote = ' (BS동일공시기준)';
          }
        }
      }
      const secValM = effectiveSecValM;
      const secValRounded = Number(secValM.toFixed(1));
      const tol = Math.max(2, Math.abs(secValM) * 0.005);

      if (excelVal === null) {
        if (m.metric === '자본총계') {
          const excelAssetsVal =
            sheetYi >= 0 && assetsRow ? assetsRow.values[sheetYi] : null;
          const secAssetsFact = getSecAnnualFactBest(
            factsObj,
            ['Assets'],
            yr,
            true,
            fyeMonth
          );
          if (excelAssetsVal === null && !secAssetsFact) {
            comparisons.push({
              metric: m.metric,
              year: String(yr),
              excelLabel: m.excelRow.label,
              excelVal: null,
              secTag: '(BS 미공시 N/A)',
              secValM: null,
              diff: null,
              status: 'N/A',
            });
            continue;
          }
        }
        failures.push(
          makeFailure(
            ticker,
            'SEC_GROUND_TRUTH',
            '8. SEC 원본 대조',
            m.sheetName,
            String(yr),
            m.metric,
            `엑셀=null (${m.excelRow.label})`,
            `SEC=${formatNum(secValRounded)}M (us-gaap:${secFact.tag}, end=${secFact.end})`
          )
        );
        comparisons.push({
          metric: m.metric,
          year: String(yr),
          excelLabel: m.excelRow.label,
          excelVal: null,
          secTag: secFact.tag,
          secValM: secValRounded,
          diff: null,
          status: 'FAIL',
        });
        continue;
      }

      const diff = excelVal - secValM;
      const absDiff = Math.abs(diff);
      const matchesPrior10K =
        absDiff > tol &&
        ((secValM !== 0 &&
          Math.abs(excelVal - secValM) / Math.abs(secValM) <= 0.02 &&
          Array.isArray(secFact.allValuesM) &&
          secFact.allValuesM.some((altM) => Math.abs(excelVal - altM) <= 2)) ||
          (Array.isArray(secFact.allRecords) &&
            secFact.allRecords.some(
              (r) => r.form === '10-K' && Math.abs(excelVal - r.valM) <= Math.max(2, Math.abs(r.valM) * 0.005)
            )));
      const isOk = absDiff <= tol || matchesPrior10K;
      if (!isOk) {
        const pct = secValM !== 0 ? ((diff / Math.abs(secValM)) * 100).toFixed(2) : 'inf';
        failures.push(
          makeFailure(
            ticker,
            'SEC_GROUND_TRUTH',
            '8. SEC 원본 대조',
            m.sheetName,
            String(yr),
            m.metric,
            `엑셀=${formatNum(excelVal)}M (${m.excelRow.label})`,
            `SEC=${formatNum(secValRounded)}M (us-gaap:${secFact.tag}, 차이=${diff >= 0 ? '+' : ''}${diff.toFixed(1)}M / ${pct}%, 허용=±${tol.toFixed(1)}M)`
          )
        );
      }

      comparisons.push({
        metric: m.metric,
        year: String(yr),
        excelLabel: m.excelRow.label,
        excelVal,
        secTag: matchesPrior10K
          ? `${secFact.tag} (원공시일치)`
          : `${secFact.tag}${effectiveSecTagNote}`,
        secValM: secValRounded,
        diff: Number(diff.toFixed(1)),
        status: isOk ? 'PASS' : 'FAIL',
      });
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
    comparisons,
    sampledYears,
  };
}

// =========================================================================================
// [번역 검수 덤프] 캡처된 R*.htm + Gemini 요청/응답으로부터 [시트 | 태그 | 원문 → 한글명] 역추적
// =========================================================================================

function decodeHtmlEntities(str) {
  return String(str || '')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;|&apos;/gi, "'")
    .replace(/&#8212;/gi, '—')
    .replace(/&#8211;/gi, '–')
    .replace(/&#160;|&nbsp;/gi, ' ')
    .replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code)));
}

function normalizeEngForMatch(str) {
  return String(str || '')
    .toLowerCase()
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[^a-z0-9]/g, '');
}

function parseCapturedRHtmlRows(capture) {
  const factsObj = capture?.companyFacts?.facts || {};
  const usGaap = factsObj['us-gaap'] || {};
  const rowsByStmt = { IS: [], BS: [], CF: [] };

  // accNoDashes -> stmtType lookup per file
  for (const entry of capture?.rHtmlEntries || []) {
    const summary = capture.filingSummaries.get(entry.accNoDashes);
    if (!summary) continue;
    let stmtType = null;
    if (summary.IS === entry.htmlFile) stmtType = 'IS';
    else if (summary.BS === entry.htmlFile) stmtType = 'BS';
    else if (summary.CF === entry.htmlFile) stmtType = 'CF';
    if (!stmtType) continue;

    const html = entry.html;
    const headerHtmlChunk = html.split(/<td[^>]*class=["']nump?["']/i)[0] || html.slice(0, 4500);
    const headerMetaText = headerHtmlChunk
      .replace(/<[^>]+>/g, ' ')
      .replace(/&#160;|&nbsp;/gi, ' ')
      .replace(/\s+/g, ' ');
    const dollarMeta = headerMetaText.replace(/shares\s+in\s+(?:thousands|millions|billions)/gi, ' ');
    let defaultScale = 1;
    if (/\bin\s+thousands\b/i.test(dollarMeta)) defaultScale = 0.001;

    let shareScale = defaultScale;
    if (/shares\s+in\s+thousands/i.test(headerMetaText)) shareScale = 0.001;
    else if (/shares\s+in\s+millions/i.test(headerMetaText)) shareScale = 1;
    else if (/except\s+(?:share|number\s+of\s+shares)/i.test(headerMetaText)) shareScale = 0.000001;

    const trBlocks = html.match(/<tr[\s\S]*?<\/tr>/gi) || [];
    const headerRows = [];
    let colDurations = [];

    for (const tr of trBlocks) {
      if (/class=["']nump?["']/i.test(tr)) break;
      if (!/<th[\s>]/i.test(tr)) continue;
      const thCells = tr.match(/<th[\s\S]*?<\/th>/gi) || [];
      const extractedCells = [];
      const expandedDurations = [];
      let hasDurationCell = false;

      for (const th of thCells) {
        if (/class=["']tl["']/i.test(th)) continue;
        const colspanMatch = th.match(/colspan=["']?(\d+)["']?/i);
        const colspan = colspanMatch ? Math.max(1, parseInt(colspanMatch[1], 10)) : 1;
        const text = decodeHtmlEntities(th.replace(/<[^>]+>/g, ' '))
          .replace(/\s+/g, ' ')
          .trim();
        if (/^\d+\s*(?:months?|weeks?)\s*ended$/i.test(text)) {
          hasDurationCell = true;
          for (let k = 0; k < colspan; k++) expandedDurations.push(text);
          continue;
        }
        for (let k = 0; k < colspan; k++) expandedDurations.push(text);
        extractedCells.push(text);
      }
      if (hasDurationCell && extractedCells.every((c) => !c)) {
        colDurations = expandedDurations;
      } else if (extractedCells.length > 0) {
        headerRows.push(extractedCells);
      }
    }

    const isAnnualDur = (d) => /^(?:12\s*months?|5[23]\s*weeks?)\s*ended$/i.test(d);
    const colEntries = [];
    const yearRowIdx = headerRows.findLastIndex((r) => r.some((c) => /\b(20\d{2})\b/.test(c)));
    if (yearRowIdx !== -1) {
      const yearRow = headerRows[yearRowIdx];
      const prevRow =
        yearRowIdx > 0 && headerRows[yearRowIdx - 1].length === yearRow.length
          ? headerRows[yearRowIdx - 1]
          : null;
      const hasMixedDurations =
        stmtType !== 'BS' &&
        colDurations.length === yearRow.length &&
        colDurations.some(isAnnualDur) &&
        colDurations.some((d) => !isAnnualDur(d));

      for (let c = 0; c < yearRow.length; c++) {
        if (hasMixedDurations && !isAnnualDur(colDurations[c])) continue;
        const fullDateText = prevRow ? `${prevRow[c]} ${yearRow[c]}` : yearRow[c];
        const yrMatch = fullDateText.match(/\b(20\d{2})\b/);
        if (yrMatch) {
          let yr = parseInt(yrMatch[1], 10);
          if (/\bJan(?:uary)?\.?\s*(?:0?[1-9]|1[0-5])\b/i.test(fullDateText)) yr -= 1;
          colEntries.push({ colIdx: c, yr });
        }
      }
    }
    for (let i = 1; i < colEntries.length; i++) {
      if (colEntries[i].yr === colEntries[i - 1].yr) {
        colEntries[i].yr = colEntries[i - 1].yr - 1;
      }
    }
    const maxYr = colEntries.length > 0 ? Math.max(...colEntries.map((c) => c.yr)) : 0;

    // table-level actual scale check against CompanyFacts
    const tableRawRows = [];
    const votes = { 1: 0, 0.001: 0, 0.000001: 0 };
    let currentSection = '';
    let currentDim = '';

    for (const tr of trBlocks) {
      if (/<th[\s>]/i.test(tr)) continue;
      let rawLabel = '';
      const linkMatch = tr.match(/<a[^>]*onclick=["'][^"']*showAR[^"']*["'][^>]*>([\s\S]*?)<\/a>/i);
      if (linkMatch) {
        rawLabel = linkMatch[1];
      } else {
        const plMatch = tr.match(/<td[^>]*class=["']pl[^"']*["'][^>]*>([\s\S]*?)<\/td>/i);
        if (plMatch) rawLabel = plMatch[1];
      }
      rawLabel = decodeHtmlEntities(rawLabel.replace(/<[^>]+>/g, ''))
        .replace(/\s+/g, ' ')
        .trim();

      if (!/class=["']nump?["']/i.test(tr)) {
        if (/class=["']rh["']/i.test(tr)) {
          currentDim = rawLabel.replace(/\[member\]/gi, '').trim();
          currentSection = '';
        } else if (rawLabel) {
          currentSection = rawLabel;
        }
        continue;
      }

      const fullDefMatch = tr.match(/(defref_([a-zA-Z0-9-]+?)_([a-zA-Z0-9]+))/);
      const prefix = fullDefMatch ? fullDefMatch[2] : 'custom';
      const tag = fullDefMatch ? fullDefMatch[3] : '';

      const valueCells =
        tr.match(/<td[^>]*class=["'](?:nump|num|text)["'][^>]*>[\s\S]*?<\/td>/gi) || [];
      const rawNumbersByYear = {};
      let hasAny = false;
      for (const { colIdx, yr } of colEntries) {
        if (colIdx >= valueCells.length) continue;
        const rawText = decodeHtmlEntities(valueCells[colIdx].replace(/<[^>]+>/g, ''))
          .replace(/\$|,/g, '')
          .trim();
        if (!rawText || rawText === '-' || rawText === '—' || rawText === '–') continue;
        const isNeg =
          rawText.includes('(') ||
          rawText.startsWith('-') ||
          rawText.startsWith('−') ||
          rawText.startsWith('–');
        const num = parseFloat(rawText.replace(/[()\-−–]/g, ''));
        if (!isNaN(num)) {
          rawNumbersByYear[yr] = isNeg ? -Math.abs(num) : Math.abs(num);
          hasAny = true;
        }
      }
      if (!hasAny) continue;

      const isEpsRow =
        /PerShare|EarningsPerShare|DividendsPerShare/i.test(tag) ||
        /per\s+share|in\s+dollars\s+per/i.test(`${rawLabel} ${currentSection}`);
      const isSharesRow =
        (/Shares|WeightedAverageNumber/i.test(tag) && !/PerShare/i.test(tag)) ||
        (/\bshares?\b/i.test(`${rawLabel} ${currentSection}`) &&
          !/share-?based|stock-?based|repurchase|dividend|treasury|per\s+share/i.test(
            `${rawLabel} ${currentSection}`
          ));

      if (!isEpsRow && !isSharesRow && tag) {
        const concept = factsObj?.[prefix]?.[tag] || usGaap[tag];
        const usdRecs = concept?.units?.USD;
        if (Array.isArray(usdRecs)) {
          for (const { yr } of colEntries) {
            const rv = rawNumbersByYear[yr];
            if (!rv || Math.abs(rv) < 1) continue;
            const rec = usdRecs.find((r) => {
              if (r.form !== '10-K' && r.form !== '10-K/A') return false;
              let endYr = r.end ? parseInt(r.end.slice(0, 4), 10) : r.fy;
              if (r.end && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(r.end)) endYr -= 1;
              return endYr === yr;
            });
            if (!rec?.val) continue;
            const tMil = Math.abs(rec.val) / 1e6;
            if (tMil < 0.5) continue;
            const ratio = Math.abs(rv) / tMil;
            if (ratio > 0.4 && ratio < 2.5) votes[1]++;
            else if (ratio > 400 && ratio < 2500) votes[0.001]++;
            else if (ratio > 400000 && ratio < 2500000) votes[0.000001]++;
          }
        }
      }

      tableRawRows.push({
        stmtType,
        prefix,
        tag,
        engLabel: rawLabel.replace(/:$/, '').trim(),
        section: currentSection,
        dimLabel: currentDim,
        isEpsRow,
        isSharesRow,
        maxYr,
        rawNumbersByYear,
      });
    }

    const maxVote = Math.max(votes[1], votes[0.001], votes[0.000001]);
    const verifiedScale =
      maxVote === 0
        ? defaultScale
        : votes[1] === maxVote
        ? 1
        : votes[0.001] === maxVote
        ? 0.001
        : 0.000001;

    for (const r of tableRawRows) {
      const scaledByYear = {};
      for (const [yr, num] of Object.entries(r.rawNumbersByYear)) {
        if (r.isEpsRow) {
          scaledByYear[yr] = Number(num.toFixed(2));
        } else if (r.isSharesRow) {
          let s = num * shareScale;
          if (Math.abs(s) >= 40000 && Math.abs(s) < 50000000) s *= 0.001;
          else if (Math.abs(s) >= 50000000) s *= 0.000001;
          scaledByYear[yr] = Math.round(s);
        } else {
          scaledByYear[yr] = Math.round(num * verifiedScale);
        }
      }
      rowsByStmt[stmtType].push({
        ...r,
        scaledByYear,
      });
    }
  }

  // 최신 10-K 우선 정렬 후 동일 (prefix, tag, normalizedEng) 병합
  const consolidated = { IS: [], BS: [], CF: [] };
  for (const stmtType of ['IS', 'BS', 'CF']) {
    const sorted = rowsByStmt[stmtType].sort((a, b) => b.maxYr - a.maxYr);
    for (const r of sorted) {
      const normE = normalizeEngForMatch(r.engLabel);
      const existing = consolidated[stmtType].find(
        (c) => c.tag === r.tag && c.normEng === normE && c.dimLabel === r.dimLabel
      );
      if (existing) {
        for (const [yr, v] of Object.entries(r.scaledByYear)) {
          if (existing.scaledByYear[yr] === undefined) existing.scaledByYear[yr] = v;
        }
      } else {
        consolidated[stmtType].push({
          ...r,
          normEng: normE,
          scaledByYear: { ...r.scaledByYear },
        });
      }
    }
  }
  return consolidated;
}

const KOR_LABEL_TAG_HINTS = [
  [/^\[\[자산총계\]\]$/, /Assets$/],
  [/^\[\[부채총계\]\]$/, /Liabilities(?:IncludingDeferredIncomeTaxes)?$/],
  [/^\[\[부채와자본총계\]\]$/, /LiabilitiesAndStockholdersEquity/],
  [/^\[\[자본총계\]\]$/, /StockholdersEquity/],
  [/^\[유동자산 합계\]$/, /AssetsCurrent/],
  [/^\[비유동자산 합계\]$/, /AssetsNoncurrent/],
  [/^\[유동부채 합계\]$/, /LiabilitiesCurrent/],
  [/^\[비유동부채 합계\]$/, /LiabilitiesNoncurrent/],
  [/^\[영업활동현금흐름\]$/, /NetCashProvidedByUsedInOperatingActivities$/],
  [/^\[계속영업 영업활동현금흐름\]$/, /NetCashProvidedByUsedInOperatingActivitiesContinuingOperations/],
  [/^\[투자활동현금흐름\]$/, /NetCashProvidedByUsedInInvestingActivities$/],
  [/^\[계속영업 투자활동현금흐름\]$/, /NetCashProvidedByUsedInInvestingActivitiesContinuingOperations/],
  [/^\[재무활동현금흐름\]$/, /NetCashProvidedByUsedInFinancingActivities$/],
  [/^\[계속영업 재무활동현금흐름\]$/, /NetCashProvidedByUsedInFinancingActivitiesContinuingOperations/],
  [/^매출액$|^영업수익$/, /Revenue|SalesRevenue/i],
  [/^매출원가$/, /CostOf/i],
  [/^매출총이익$/, /GrossProfit/],
  [/^연구개발비$/, /ResearchAndDevelopment/],
  [/^판매비와관리비$/, /SellingGeneralAndAdministrative/],
  [/^영업이익$/, /OperatingIncomeLoss/],
  [/^법인세비용차감전순이익$/, /BeforeIncomeTaxes/],
  [/^법인세비용$/, /IncomeTaxExpenseBenefit/],
  [/^연결당기순이익$/, /^ProfitLoss$/],
  [/^당기순이익$/, /^NetIncomeLoss$/],
  [/^보통주귀속당기순이익$/, /NetIncomeLossAvailableToCommonStockholders/],
  [/^기본 주당순이익/, /EarningsPerShareBasic|PerBasicShare/],
  [/^희석 주당순이익/, /EarningsPerShareDiluted|PerDilutedShare/],
  [/^기본 가중평균/, /SharesOutstandingBasic|BasicShares/],
  [/^희석 가중평균/, /DilutedSharesOutstanding/],
  [/^기초 현금및현금성자산$/, /Cash.*CarryingValue|RestrictedCash|beginning/i],
  [/^기말 현금및현금성자산$/, /Cash.*CarryingValue|RestrictedCash|end/i],
  [/^현금및현금성자산$/, /^CashAndCashEquivalents/],
  [/^매출채권/, /AccountsReceivable/i],
  [/^재고자산/, /Inventory/i],
  [/^유형자산/, /PropertyPlantAndEquipment/i],
  [/^영업권$/, /^Goodwill$/],
  [/^무형자산/, /IntangibleAssets/i],
  [/^매입채무/, /AccountsPayable/i],
  [/^이익잉여금|^결손금/, /RetainedEarnings/i],
  [/^자기주식/, /TreasuryStock/i],
  [/^감가상각비/, /Depreciation/i],
  [/^주식보상비용/, /ShareBasedCompensation/i],
];

function buildTranslationMappings(sheets, capture) {
  const consolidatedByStmt = parseCapturedRHtmlRows(capture);
  const geminiList = capture?.geminiMappings || [];
  const sheetToStmt = { 손익계산서: 'IS', 재무상태표: 'BS', 현금흐름표: 'CF' };

  const entries = [];
  const duplicatesBySheet = [];
  let matchedTagCount = 0;
  let totalRows = 0;

  for (const sheetName of SHEET_NAMES) {
    const sheet = sheets[sheetName];
    if (!sheet) continue;
    const stmtType = sheetToStmt[sheetName];
    const candidates = consolidatedByStmt[stmtType] || [];
    const usedCandIndices = new Set();

    // 동일 시트 내 중복 라벨 탐지용 카운터
    const labelRowsMap = new Map();

    for (const row of sheet.rows) {
      totalRows++;
      const korName = row.label;
      const nonSegKor = korName.replace(/^\[[^\]]+\]\s*/, '').trim();

      let matchedCand = null;
      let matchedScore = -Infinity;
      let matchedCandIdx = -1;

      // 1) Gemini 번역 매핑에서 후보 태그 탐색
      const aiHits = geminiList.filter(
        (g) =>
          g.stmt === stmtType &&
          (g.aiKor === korName || g.aiKor === nonSegKor || nonSegKor.includes(g.aiKor))
      );
      const aiTagSet = new Set(aiHits.map((h) => h.tag).filter(Boolean));
      const aiEngSet = new Set(aiHits.map((h) => normalizeEngForMatch(h.engLabel)).filter(Boolean));

      for (let ci = 0; ci < candidates.length; ci++) {
        const cand = candidates[ci];
        let score = 0;

        if (aiTagSet.has(cand.tag)) score += 25;
        if (aiEngSet.has(cand.normEng)) score += 20;

        let exactNumMatches = 0;
        let absNumMatches = 0;
        let numMismatches = 0;
        let overlapCount = 0;

        for (let yi = 0; yi < sheet.years.length; yi++) {
          const v = row.values[yi];
          const cv = cand.scaledByYear[sheet.years[yi]];
          if (v === null || cv === undefined || cv === null) continue;
          overlapCount++;
          const tol = cand.isEpsRow ? 0.05 : 3;
          if (Math.abs(v - cv) <= tol) exactNumMatches++;
          else if (Math.abs(Math.abs(v) - Math.abs(cv)) <= tol) absNumMatches++;
          else numMismatches++;
        }

        if (overlapCount === 0) continue;
        score += exactNumMatches * 6 + absNumMatches * 3 - numMismatches * 14;

        for (const [korRe, tagRe] of KOR_LABEL_TAG_HINTS) {
          if (korRe.test(korName) && tagRe.test(cand.tag)) {
            score += 18;
          }
        }

        if (/기초 현금/.test(korName) && /beginning/i.test(cand.engLabel)) score += 15;
        if (/기말 현금/.test(korName) && /end/i.test(cand.engLabel)) score += 15;
        if (!usedCandIndices.has(ci)) score += 2;

        if (score > matchedScore && exactNumMatches + absNumMatches > 0 && numMismatches === 0) {
          matchedScore = score;
          matchedCand = cand;
          matchedCandIdx = ci;
        }
      }

      // 혹시 소급/부호 반전 등으로 100% 무오차 후보가 없으면 AI 매핑 또는 최다 수치 일치로 fallback
      if (!matchedCand && aiHits.length > 0) {
        const hit = aiHits[0];
        const fallbackCand = candidates.find((c) => c.tag === hit.tag);
        matchedCand = fallbackCand || {
          prefix: 'us-gaap',
          tag: hit.tag,
          engLabel: hit.engLabel,
        };
      } else if (matchedCandIdx !== -1) {
        usedCandIndices.add(matchedCandIdx);
      }

      const xbrlTag = matchedCand
        ? `${matchedCand.prefix || 'us-gaap'}:${matchedCand.tag || 'UNKNOWN'}`
        : aiHits[0]?.tag
        ? `us-gaap:${aiHits[0].tag}`
        : 'UNMATCHED';
      const engLabel = matchedCand?.engLabel || aiHits[0]?.engLabel || '(영문 원문 미매칭)';

      if (xbrlTag !== 'UNMATCHED') matchedTagCount++;

      const item = {
        sheetName,
        rowIndex: row.rowIndex,
        xbrlTag,
        engLabel,
        korName,
        source: aiHits.length > 0 ? 'AI' : 'DICT',
      };
      entries.push(item);

      if (!labelRowsMap.has(korName)) labelRowsMap.set(korName, []);
      labelRowsMap.get(korName).push(item);
    }

    for (const [label, items] of labelRowsMap.entries()) {
      if (items.length >= 2) {
        duplicatesBySheet.push({
          sheetName,
          label,
          count: items.length,
          occurrences: items,
        });
      }
    }
  }

  return { entries, duplicatesBySheet, matchedTagCount, totalRows };
}

// =========================================================================================
// [20-F 판별 및 실행기]
// =========================================================================================

function detect20FIssuer(capture) {
  const recentForms = capture?.submissions?.filings?.recent?.form || [];
  if (!Array.isArray(recentForms) || recentForms.length === 0) return false;
  const has10K = recentForms.some((f) => /^10-K/i.test(f));
  const has20F = recentForms.some((f) => /^(?:20-F|40-F)/i.test(f));
  return !has10K && has20F;
}

function detect1940ActFundIssuer(capture) {
  const recentForms = capture?.submissions?.filings?.recent?.form || [];
  if (!Array.isArray(recentForms) || recentForms.length === 0) return false;
  const has10K = recentForms.some((f) => /^10-K/i.test(f));
  const hasFundForm = recentForms.some((f) => /^(?:N-CSR|N-CSRS|N-CEN|NPORT-P|N-Q)(?:\/A)?$/i.test(f));
  return !has10K && hasFundForm;
}

function detectPre10KIpoIssuer(capture) {
  const recentForms = capture?.submissions?.filings?.recent?.form || [];
  if (!Array.isArray(recentForms) || recentForms.length === 0) return false;
  const has10K = recentForms.some((f) => /^10-K/i.test(f));
  const has20F = recentForms.some((f) => /^(?:20-F|40-F)/i.test(f));
  const hasFundForm = recentForms.some((f) => /^(?:N-CSR|N-CSRS|N-CEN|NPORT-P|N-Q)(?:\/A)?$/i.test(f));
  const hasIpoOrQuarterly = recentForms.some((f) =>
    /^(?:S-1|S-1\/A|S-4|S-4\/A|424B[1-4]|10-Q|10-Q\/A|8-A12B|F-1|F-1\/A)$/i.test(f)
  );
  return !has10K && !has20F && !hasFundForm && hasIpoOrQuarterly;
}

async function analyzeAndStressAuditTicker(ticker, apiKey, userEmail) {
  activeCapture = createEmptyCapture();
  const startedAt = Date.now();

  try {
    const req = new Request('http://localhost/api/analyze', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ticker, apiKey, userEmail }),
    });

    const res = await POST(req);
    const elapsedSec = ((Date.now() - startedAt) / 1000).toFixed(1);
    const capture = activeCapture;

    if (!res.ok) {
      let errMsg = `HTTP ${res.status}`;
      let parsedError = '';
      try {
        const text = await res.text();
        try {
          const parsed = JSON.parse(text);
          parsedError = parsed.error || text;
          errMsg = `${errMsg}: ${parsedError}`;
        } catch {
          parsedError = text;
          errMsg = `${errMsg}: ${text}`;
        }
      } catch {
        // 응답 본문 읽기 실패 시 상태코드 유지
      }

      const isForeign20F = detect20FIssuer(capture);
      const isNon10KFund = !isForeign20F && detect1940ActFundIssuer(capture);
      const isPre10KIpo = !isForeign20F && !isNon10KFund && detectPre10KIpoIssuer(capture);
      const isGracefulKoreanError =
        Boolean(parsedError) &&
        /[가-힣]/.test(parsedError) &&
        /20-F/i.test(parsedError) &&
        res.status === 422;
      const isGracefulFundError =
        Boolean(parsedError) &&
        /[가-힣]/.test(parsedError) &&
        /(?:N-CSR|펀드|투자회사)/i.test(parsedError) &&
        res.status === 422;
      const isGracefulPre10KError =
        Boolean(parsedError) &&
        /[가-힣]/.test(parsedError) &&
        /(?:IPO|신규\s*상장|10-K)/i.test(parsedError) &&
        res.status === 422;

      if (isForeign20F) {
        return {
          ticker,
          role: TICKER_ROLE_LABELS[ticker] || '20-F/해외기업',
          elapsedSec,
          yearRange: '-',
          rowCounts: '-',
          isForeign20F: true,
          isNon10KFund: false,
          httpStatus: res.status,
          apiError: errMsg,
          checks: {
            BS_IDENTITY: { status: 'N/A', failures: [] },
            SUBTOTALS: { status: 'N/A', failures: [], breakdown: {} },
            SCALE_JUMP: { status: 'N/A', failures: [] },
            OUTFLOW_SIGN: { status: 'N/A', failures: [] },
            UNTRANSLATED: { status: 'N/A', failures: [] },
            BRK_STRUCTURE: { status: 'N/A', failures: [] },
            IS_INTEGRITY: { status: 'N/A', failures: [] },
            SEC_GROUND_TRUTH: { status: 'N/A', failures: [], comparisons: [], sampledYears: [] },
            STRUCTURAL_ANOMALY: { status: 'N/A', failures: [] },
          },
          translations: { entries: [], duplicatesBySheet: [], matchedTagCount: 0, totalRows: 0 },
          allFailures: isGracefulKoreanError
            ? []
            : [
                makeFailure(
                  ticker,
                  'FOREIGN_20F',
                  '20-F 예외처리',
                  '-',
                  'ALL',
                  '20-F 안내',
                  errMsg,
                  'HTTP 422 및 20-F 한글 안내 메시지 JSON 반환'
                ),
              ],
          overall: isGracefulKoreanError ? '20-F(PASS)' : '20-F(FAIL)',
        };
      }

      if (isNon10KFund) {
        return {
          ticker,
          role: TICKER_ROLE_LABELS[ticker] || '펀드/투자회사(N-CSR)',
          elapsedSec,
          yearRange: '-',
          rowCounts: '-',
          isForeign20F: false,
          isNon10KFund: true,
          httpStatus: res.status,
          apiError: errMsg,
          checks: {
            BS_IDENTITY: { status: 'N/A', failures: [] },
            SUBTOTALS: { status: 'N/A', failures: [], breakdown: {} },
            SCALE_JUMP: { status: 'N/A', failures: [] },
            OUTFLOW_SIGN: { status: 'N/A', failures: [] },
            UNTRANSLATED: { status: 'N/A', failures: [] },
            BRK_STRUCTURE: { status: 'N/A', failures: [] },
            IS_INTEGRITY: { status: 'N/A', failures: [] },
            SEC_GROUND_TRUTH: { status: 'N/A', failures: [], comparisons: [], sampledYears: [] },
            STRUCTURAL_ANOMALY: { status: 'N/A', failures: [] },
          },
          translations: { entries: [], duplicatesBySheet: [], matchedTagCount: 0, totalRows: 0 },
          allFailures: isGracefulFundError
            ? []
            : [
                makeFailure(
                  ticker,
                  'NON_10K_FUND',
                  '펀드(N-CSR) 예외처리',
                  '-',
                  'ALL',
                  '펀드 공시 안내',
                  errMsg,
                  'HTTP 422 및 N-CSR/펀드 한글 안내 메시지 JSON 반환'
                ),
              ],
          overall: isGracefulFundError ? 'FUND(PASS)' : 'FUND(FAIL)',
        };
      }

      if (isPre10KIpo) {
        return {
          ticker,
          role: TICKER_ROLE_LABELS[ticker] || '신규상장(10-K미공시)',
          elapsedSec,
          yearRange: '-',
          rowCounts: '-',
          isForeign20F: false,
          isNon10KFund: false,
          isPre10KIpo: true,
          httpStatus: res.status,
          apiError: errMsg,
          checks: {
            BS_IDENTITY: { status: 'N/A', failures: [] },
            SUBTOTALS: { status: 'N/A', failures: [], breakdown: {} },
            SCALE_JUMP: { status: 'N/A', failures: [] },
            OUTFLOW_SIGN: { status: 'N/A', failures: [] },
            UNTRANSLATED: { status: 'N/A', failures: [] },
            BRK_STRUCTURE: { status: 'N/A', failures: [] },
            IS_INTEGRITY: { status: 'N/A', failures: [] },
            SEC_GROUND_TRUTH: { status: 'N/A', failures: [], comparisons: [], sampledYears: [] },
            STRUCTURAL_ANOMALY: { status: 'N/A', failures: [] },
          },
          translations: { entries: [], duplicatesBySheet: [], matchedTagCount: 0, totalRows: 0 },
          allFailures: isGracefulPre10KError
            ? []
            : [
                makeFailure(
                  ticker,
                  'PRE_10K_IPO',
                  '신규상장(10-K미공시) 예외처리',
                  '-',
                  'ALL',
                  '10-K 미공시 안내',
                  errMsg,
                  'HTTP 422 및 신규상장 10-K 미공시 한글 안내 메시지 JSON 반환'
                ),
              ],
          overall: isGracefulPre10KError ? 'PRE-10K(PASS)' : 'PRE-10K(FAIL)',
        };
      }

      const apiFail = makeFailure(
        ticker,
        'API_ERROR',
        'API 호출',
        '-',
        'ALL',
        '(API 호출 실패)',
        errMsg,
        'HTTP 200 Excel 응답'
      );

      return {
        ticker,
        role: TICKER_ROLE_LABELS[ticker] || '무작위/미분류',
        elapsedSec,
        yearRange: '-',
        rowCounts: '-',
        isForeign20F: false,
        httpStatus: res.status,
        apiError: errMsg,
        checks: {
          BS_IDENTITY: { status: 'FAIL', failures: [apiFail] },
          SUBTOTALS: { status: 'FAIL', failures: [], breakdown: {} },
          SCALE_JUMP: { status: 'FAIL', failures: [] },
          OUTFLOW_SIGN: { status: 'FAIL', failures: [] },
          UNTRANSLATED: { status: 'FAIL', failures: [] },
          BRK_STRUCTURE: { status: 'N/A', failures: [] },
          IS_INTEGRITY: { status: 'FAIL', failures: [] },
          SEC_GROUND_TRUTH: { status: 'FAIL', failures: [], comparisons: [], sampledYears: [] },
          STRUCTURAL_ANOMALY: { status: 'FAIL', failures: [] },
        },
        translations: { entries: [], duplicatesBySheet: [], matchedTagCount: 0, totalRows: 0 },
        allFailures: [apiFail],
        overall: 'FAIL',
      };
    }

    const buffer = Buffer.from(await res.arrayBuffer());
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const sheets = parseWorkbookSheets(workbook);
    const refYears =
      sheets['손익계산서']?.years ||
      sheets['재무상태표']?.years ||
      sheets['현금흐름표']?.years ||
      [];
    const yearRange =
      refYears.length > 0 ? `${refYears[0]}~${refYears[refYears.length - 1]}` : '-';
    const rowCounts = `${sheets['손익계산서']?.rows.length ?? 0}/${
      sheets['재무상태표']?.rows.length ?? 0
    }/${sheets['현금흐름표']?.rows.length ?? 0}`;

    const c1 = auditBalanceSheetIdentity(ticker, sheets);
    const c2 = auditMajorSubtotals(ticker, sheets);
    const c3 = auditScaleJumps(ticker, sheets);
    const c4 = auditCashOutflowSigns(ticker, sheets);
    const c5 = auditUntranslatedLabels(ticker, sheets);
    const c6 = auditBrkAStructure(ticker, sheets);
    const c7 = auditIncomeStatementIntegrity(ticker, sheets, capture);
    const c8 = auditSecGroundTruth(ticker, sheets, capture);
    const translations = buildTranslationMappings(sheets, capture);
    const c9 = auditStructuralAnomalies(ticker, sheets, translations);

    const allFailures = [
      ...c1.failures,
      ...c2.failures,
      ...c3.failures,
      ...c4.failures,
      ...c5.failures,
      ...c6.failures,
      ...c7.failures,
      ...c8.failures,
      ...c9.failures,
    ];

    const overall =
      c1.status === 'PASS' &&
      c2.status === 'PASS' &&
      c3.status === 'PASS' &&
      c4.status === 'PASS' &&
      c5.status === 'PASS' &&
      c6.status !== 'FAIL' &&
      c7.status === 'PASS' &&
      c8.status === 'PASS' &&
      c9.status === 'PASS'
        ? 'PASS'
        : 'FAIL';

    return {
      ticker,
      role: TICKER_ROLE_LABELS[ticker] || '무작위/신규',
      elapsedSec,
      yearRange,
      rowCounts,
      isForeign20F: false,
      httpStatus: 200,
      apiError: null,
      checks: {
        BS_IDENTITY: c1,
        SUBTOTALS: c2,
        SCALE_JUMP: c3,
        OUTFLOW_SIGN: c4,
        UNTRANSLATED: c5,
        BRK_STRUCTURE: c6,
        IS_INTEGRITY: c7,
        SEC_GROUND_TRUTH: c8,
        STRUCTURAL_ANOMALY: c9,
      },
      translations,
      allFailures,
      overall,
    };
  } finally {
    activeCapture = null;
  }
}

// =========================================================================================
// [CLI 파싱 및 무작위 상장사 N개 추출]
// =========================================================================================

function parseCliArgs(argv) {
  const opts = {
    tickers: [],
    randomCount: null,
    dryRun: false,
    delayMs: DEFAULT_DELAY_MS,
    reportPath: 'audit-report.md',
    translationsPath: 'audit-translations.txt',
  };

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === '--random' || arg === '-r') {
      const next = argv[i + 1];
      if (!next || isNaN(Number(next))) {
        throw new Error('--random 뒤에는 양의 정수를 입력해 주세요. 예: --random 10');
      }
      opts.randomCount = Math.max(1, parseInt(next, 10));
      i++;
    } else if (arg.startsWith('--random=')) {
      opts.randomCount = Math.max(1, parseInt(arg.split('=')[1], 10));
    } else if (arg === '--dry-run') {
      opts.dryRun = true;
    } else if (arg === '--delay' && argv[i + 1]) {
      opts.delayMs = Math.max(1000, parseInt(argv[++i], 10));
    } else if (arg === '--out' && argv[i + 1]) {
      opts.reportPath = argv[++i];
    } else if (arg === '--translations' && argv[i + 1]) {
      opts.translationsPath = argv[++i];
    } else if (!arg.startsWith('-')) {
      opts.tickers.push(arg.trim().toUpperCase());
    }
  }
  return opts;
}

async function sampleRandomTickers(count, userEmail) {
  const safeEmail = userEmail && userEmail.includes('@') ? userEmail.trim() : 'sec_user@gmail.com';
  const res = await fetch('https://www.sec.gov/files/company_tickers.json', {
    headers: {
      'User-Agent': `SecExcelWebApp ${safeEmail}`,
      'Accept-Encoding': 'gzip, deflate',
      Accept: 'application/json',
    },
  });
  if (!res.ok) {
    throw new Error(`SEC company_tickers.json 조회 실패 (HTTP ${res.status})`);
  }
  const data = await res.json();
  const entries = Object.values(data);
  const topPool = entries.slice(0, Math.min(2500, entries.length));
  const cleanPool = [];
  const seen = new Set();

  for (const item of topPool) {
    const t = String(item.ticker || '').trim().toUpperCase();
    // 워런트(.WS), 유닛(.U), 우선주(-P*) 등 비일반 주식 제외
    if (!/^[A-Z]{1,4}(?:-[AB])?$/.test(t)) continue;
    if (seen.has(t)) continue;
    seen.add(t);
    cleanPool.push({ ticker: t, title: item.title || t });
  }

  for (let i = cleanPool.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [cleanPool[i], cleanPool[j]] = [cleanPool[j], cleanPool[i]];
  }
  return cleanPool.slice(0, count);
}

// =========================================================================================
// [출력 포매팅: audit-translations.txt & audit-report.md]
// =========================================================================================

function writeTranslationsDumpFile(outPath, results) {
  const lines = [];
  lines.push('====================================================================================================');
  lines.push(` SEC 10-K 번역 품질 정밀 검수 매핑 덤프 (${new Date().toISOString()})`);
  lines.push(' 형식: [시트 | 원본 XBRL 태그 | SEC 영문 원문 → 최종 한글 계정명]');
  lines.push('====================================================================================================');
  lines.push('');

  for (const r of results) {
    lines.push(`####################################################################################################`);
    lines.push(
      `### ${r.ticker} (${r.role}) | 연도: ${r.yearRange} | 행수(IS/BS/CF): ${r.rowCounts} | 매칭률: ${r.translations.matchedTagCount}/${r.translations.totalRows} | 종합: ${r.overall}`
    );
    lines.push(`####################################################################################################`);

    if (r.apiError) {
      lines.push(`(API/20-F 응답: HTTP ${r.httpStatus} - ${r.apiError})`);
      lines.push('');
      continue;
    }

    for (const item of r.translations.entries) {
      lines.push(`[${item.sheetName} | ${item.xbrlTag} | ${item.engLabel} → ${item.korName}]`);
    }

    if (r.translations.duplicatesBySheet.length > 0) {
      lines.push('');
      lines.push(`--- [${r.ticker} 동일 시트 내 중복 한글 라벨 자동 탐지: ${r.translations.duplicatesBySheet.length}건] ---`);
      for (const dup of r.translations.duplicatesBySheet) {
        lines.push(`  * ${dup.sheetName} | "${dup.label}" (${dup.count}회):`);
        for (const occ of dup.occurrences) {
          lines.push(`    - ${occ.rowIndex}행: ${occ.xbrlTag} | "${occ.engLabel}"`);
        }
      }
    }
    lines.push('');
  }

  fs.writeFileSync(outPath, lines.join('\n'), 'utf8');
}

function formatCheckCell(checkRes) {
  if (!checkRes || checkRes.status === 'N/A') return '-';
  if (checkRes.status === 'PASS') return 'PASS';
  return `FAIL(${checkRes.failures.length})`;
}

function writeMarkdownAuditReport(outPath, results, meta) {
  const lines = [];
  lines.push('# SEC 10-K 자동 스트레스 테스트 감사 리포트');
  lines.push('');
  lines.push(`- 실행 시각: \`${new Date().toISOString()}\``);
  lines.push(`- 점검 종목 (${results.length}개): \`${results.map((r) => r.ticker).join(', ')}\``);
  lines.push(`- 호출 간격: \`${meta.delayMs}ms\``);
  lines.push(`- Gemini 모드: ${meta.geminiMode}`);
  lines.push('');

  lines.push('## 1. 종목별 종합 감사 결과 (검증 1~8 + 구조/중복)');
  lines.push('');
  lines.push(
    '| 종목 | 구분 | 소요(s) | 연도 | 행수(IS/BS/CF) | 1.BS항등식 | 2.주요소계 | 3.단위급변 | 4.유출부호 | 5.미번역 | 7.IS정합성 | 8.SEC원본대조 | 구조/중복 | 최종 |'
  );
  lines.push('|---|---|---:|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|');
  for (const r of results) {
    lines.push(
      `| ${r.ticker} | ${r.role} | ${r.elapsedSec} | ${r.yearRange} | ${r.rowCounts} | ${formatCheckCell(
        r.checks.BS_IDENTITY
      )} | ${formatCheckCell(r.checks.SUBTOTALS)} | ${formatCheckCell(
        r.checks.SCALE_JUMP
      )} | ${formatCheckCell(r.checks.OUTFLOW_SIGN)} | ${formatCheckCell(
        r.checks.UNTRANSLATED
      )} | ${formatCheckCell(r.checks.IS_INTEGRITY)} | ${formatCheckCell(
        r.checks.SEC_GROUND_TRUTH
      )} | ${formatCheckCell(r.checks.STRUCTURAL_ANOMALY)} | **${r.overall}** |`
    );
  }
  lines.push('');

  lines.push('## 2. 검증 2 (주요 소계 합산) 세부 상태');
  lines.push('');
  lines.push('| 종목 | 유동자산 | 비유동자산 | 유동부채 | 비유동부채 | 영업CF | 투자CF | 재무CF |');
  lines.push('|---|---|---|---|---|---|---|---|');
  for (const r of results) {
    const b = r.checks.SUBTOTALS?.breakdown || {};
    lines.push(
      `| ${r.ticker} | ${b['유동자산'] || '-'} | ${b['비유동자산'] || '-'} | ${b['유동부채'] || '-'} | ${
        b['비유동부채'] || '-'
      } | ${b['영업현금흐름'] || '-'} | ${b['투자현금흐름'] || '-'} | ${b['재무현금흐름'] || '-'} |`
    );
  }
  lines.push('');

  lines.push('## 3. 검증 8 — SEC 원본(CompanyFacts 10-K) 교차 대조 상세');
  lines.push('');
  for (const r of results) {
    if (r.isForeign20F) {
      lines.push(`### ${r.ticker} (${r.role})`);
      lines.push(`- **20-F 예외 응답 확인**: \`${r.apiError}\` (판정: **${r.overall}**)`);
      lines.push('');
      continue;
    }
    if (r.isNon10KFund) {
      lines.push(`### ${r.ticker} (${r.role})`);
      lines.push(`- **펀드(N-CSR) 예외 응답 확인**: \`${r.apiError}\` (판정: **${r.overall}**)`);
      lines.push('');
      continue;
    }
    if (r.isPre10KIpo) {
      lines.push(`### ${r.ticker} (${r.role})`);
      lines.push(`- **신규 상장(첫 10-K 미공시) 예외 응답 확인**: \`${r.apiError}\` (판정: **${r.overall}**)`);
      lines.push('');
      continue;
    }
    if (r.apiError) {
      lines.push(`### ${r.ticker} (${r.role})`);
      lines.push(`- **API 오류**: \`${r.apiError}\``);
      lines.push('');
      continue;
    }

    const comps = r.checks.SEC_GROUND_TRUTH?.comparisons || [];
    lines.push(
      `### ${r.ticker} (${r.role}) — 대조 연도: \`${(r.checks.SEC_GROUND_TRUTH?.sampledYears || []).join(', ')}\` (${r.checks.SEC_GROUND_TRUTH?.status})`
    );
    lines.push('');
    lines.push('| 항목 | 연도 | 엑셀 계정명 | 엑셀 수치($M) | SEC US-GAAP 태그 | SEC 공식($M) | 오차($M) | 판정 |');
    lines.push('|---|---|---|---|---|---:|---:|---|');
    for (const c of comps) {
      lines.push(
        `| ${c.metric} | ${c.year} | ${c.excelLabel} | ${formatNum(c.excelVal)} | \`${
          c.secTag
        }\` | ${formatNum(c.secValM)} | ${c.diff === null ? '-' : `${c.diff >= 0 ? '+' : ''}${c.diff}`} | ${
          c.status
        } |`
      );
    }
    lines.push('');
  }

  const totalFailures = results.reduce((acc, r) => acc + r.allFailures.length, 0);
  lines.push(`## 4. 전체 FAIL 상세 내역 (총 ${totalFailures}건)`);
  lines.push('');
  if (totalFailures === 0) {
    lines.push('FAIL 건수가 없습니다.');
  } else {
    for (const r of results) {
      if (r.allFailures.length === 0) continue;
      lines.push(`### ${r.ticker} (${r.allFailures.length}건)`);
      for (const f of r.allFailures) {
        lines.push(
          `- **[${f.checkName}]** ${f.sheet} | ${f.year} | \`${f.account}\` → 실제: ${f.actual} / 기대: ${f.expected}`
        );
      }
      lines.push('');
    }
  }

  lines.push('## 5. 동일 시트 내 중복 한글 라벨 자동 탐지 요약');
  lines.push('');
  for (const r of results) {
    const dups = r.translations?.duplicatesBySheet || [];
    if (dups.length === 0) continue;
    lines.push(`### ${r.ticker} (${dups.length}건)`);
    for (const d of dups) {
      const tagSummary = d.occurrences.map((o) => `\`${o.xbrlTag}\`("${o.engLabel}")`).join(', ');
      lines.push(`- **${d.sheetName}** \`${d.label}\` (${d.count}회): ${tagSummary}`);
    }
    lines.push('');
  }

  fs.writeFileSync(outPath, lines.join('\n'), 'utf8');
}

async function main() {
  const loadedEnvSource = loadDotEnvIfPresent();
  const opts = parseCliArgs(process.argv.slice(2));

  const envApiKey = (process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || '').trim();
  const hasRealGeminiKey = Boolean(envApiKey);
  const apiKey = envApiKey || 'STRESS_NO_GEMINI_KEY_FALLBACK';
  const userEmail = (process.env.SEC_USER_EMAIL || 'sec_user@gmail.com').trim();
  const geminiMode = hasRealGeminiKey
    ? `API Key 활성 (${loadedEnvSource ? `${loadedEnvSource} 로드됨` : '환경변수'})`
    : 'API Key 미설정 (fallback 모드)';

  let targetTickers = opts.tickers.length > 0 ? opts.tickers : DEFAULT_STRESS_TICKERS;
  if (opts.randomCount) {
    const sampled = await sampleRandomTickers(opts.randomCount, userEmail);
    targetTickers = sampled.map((s) => s.ticker);
    console.log(
      `무작위 선정 (${targetTickers.length}개): ${sampled.map((s) => `${s.ticker}(${s.title})`).join(', ')}`
    );
    if (opts.dryRun) {
      console.log('--dry-run 모드 종료.');
      process.exitCode = 0;
      return;
    }
  } else if (opts.dryRun) {
    console.log(`대상 종목 (${targetTickers.length}개): ${targetTickers.join(', ')}`);
    process.exitCode = 0;
    return;
  }

  console.log('====================================================================================================');
  console.log(' SEC 10-K 자동 스트레스 테스트 + 원본 교차대조 + 번역 품질 덤프 (stress-test.mjs)');
  console.log('====================================================================================================');
  console.log(` 점검 대상 (${targetTickers.length}개): ${targetTickers.join(', ')}`);
  console.log(` 호출 간격: ${opts.delayMs}ms | Gemini 모드: ${geminiMode}`);
  console.log(` 출력 파일: ${opts.reportPath} , ${opts.translationsPath}`);
  console.log('----------------------------------------------------------------------------------------------------');

  const results = [];
  for (let i = 0; i < targetTickers.length; i++) {
    const ticker = targetTickers[i];
    if (i > 0) await delay(opts.delayMs);
    process.stdout.write(`[${i + 1}/${targetTickers.length}] ${ticker} 스트레스 감사 실행 중... `);
    try {
      const res = await analyzeAndStressAuditTicker(ticker, apiKey, userEmail);
      results.push(res);
      if (res.isForeign20F || res.isNon10KFund) {
        console.log(`완료 (${res.elapsedSec}s, HTTP ${res.httpStatus}) => ${res.overall} [${res.apiError}]`);
      } else if (res.apiError) {
        console.log(`오류 (${res.elapsedSec}s) -> ${res.apiError}`);
      } else {
        console.log(
          `완료 (${res.elapsedSec}s, 연도: ${res.yearRange}, 행수: ${res.rowCounts}, 태그매칭: ${res.translations.matchedTagCount}/${res.translations.totalRows}) => ${res.overall}`
        );
      }
    } catch (err) {
      const errMsg = err?.message || String(err);
      console.log(`예외 발생 -> ${errMsg}`);
      results.push({
        ticker,
        role: TICKER_ROLE_LABELS[ticker] || '예외',
        elapsedSec: '0.0',
        yearRange: '-',
        rowCounts: '-',
        isForeign20F: false,
        httpStatus: 0,
        apiError: errMsg,
        checks: {
          BS_IDENTITY: { status: 'FAIL', failures: [] },
          SUBTOTALS: { status: 'FAIL', failures: [], breakdown: {} },
          SCALE_JUMP: { status: 'FAIL', failures: [] },
          OUTFLOW_SIGN: { status: 'FAIL', failures: [] },
          UNTRANSLATED: { status: 'FAIL', failures: [] },
          BRK_STRUCTURE: { status: 'N/A', failures: [] },
          IS_INTEGRITY: { status: 'FAIL', failures: [] },
          SEC_GROUND_TRUTH: { status: 'FAIL', failures: [], comparisons: [], sampledYears: [] },
          STRUCTURAL_ANOMALY: { status: 'FAIL', failures: [] },
        },
        translations: { entries: [], duplicatesBySheet: [], matchedTagCount: 0, totalRows: 0 },
        allFailures: [makeFailure(ticker, 'CRASH', '실행 예외', '-', 'ALL', '(예외)', errMsg, '정상 완료')],
        overall: 'FAIL',
      });
    }
  }

  writeTranslationsDumpFile(opts.translationsPath, results);
  writeMarkdownAuditReport(opts.reportPath, results, {
    delayMs: opts.delayMs,
    geminiMode,
  });

  console.log('\n====================================================================================================');
  console.log(' [요약표] 기업별 · 검증 1~8 감사 결과');
  console.log('====================================================================================================');

  const cols = [
    { header: '기업', width: 7 },
    { header: '연도범위', width: 10 },
    { header: '1.BS', width: 8 },
    { header: '2.소계', width: 8 },
    { header: '3.급변', width: 8 },
    { header: '4.유출', width: 8 },
    { header: '5.미번역', width: 9 },
    { header: '7.IS정합', width: 9 },
    { header: '8.SEC대조', width: 10 },
    { header: '구조/중복', width: 10 },
    { header: '최종', width: 11 },
  ];
  const headerLine = cols.map((c) => padEndDisplay(c.header, c.width)).join(' | ');
  console.log(headerLine);
  console.log('-'.repeat(displayWidth(headerLine)));
  for (const r of results) {
    console.log(
      [
        padEndDisplay(r.ticker, cols[0].width),
        padEndDisplay(r.yearRange, cols[1].width),
        padEndDisplay(formatCheckCell(r.checks.BS_IDENTITY), cols[2].width),
        padEndDisplay(formatCheckCell(r.checks.SUBTOTALS), cols[3].width),
        padEndDisplay(formatCheckCell(r.checks.SCALE_JUMP), cols[4].width),
        padEndDisplay(formatCheckCell(r.checks.OUTFLOW_SIGN), cols[5].width),
        padEndDisplay(formatCheckCell(r.checks.UNTRANSLATED), cols[6].width),
        padEndDisplay(formatCheckCell(r.checks.IS_INTEGRITY), cols[7].width),
        padEndDisplay(formatCheckCell(r.checks.SEC_GROUND_TRUTH), cols[8].width),
        padEndDisplay(formatCheckCell(r.checks.STRUCTURAL_ANOMALY), cols[9].width),
        padEndDisplay(r.overall, cols[10].width),
      ].join(' | ')
    );
  }

  const totalFailures = results.reduce((acc, r) => acc + r.allFailures.length, 0);
  const passCount = results.filter(
    (r) =>
      r.overall === 'PASS' ||
      r.overall === '20-F(PASS)' ||
      r.overall === 'FUND(PASS)' ||
      r.overall === 'PRE-10K(PASS)'
  ).length;
  console.log('\n----------------------------------------------------------------------------------------------------');
  console.log(` 최종 집계: ${passCount}/${results.length} PASS(20-F/펀드/신규IPO 정상예외 포함) | 총 FAIL: ${totalFailures}건`);
  console.log(` 저장 완료: ${opts.reportPath} , ${opts.translationsPath}`);
  console.log('====================================================================================================');

  process.exitCode = totalFailures > 0 ? 1 : 0;
}

main().catch((err) => {
  console.error('stress-test.mjs 치명적 오류:', err);
  process.exitCode = 1;
});

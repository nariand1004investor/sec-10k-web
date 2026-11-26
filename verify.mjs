// verify.mjs
// SEC 10-K 재무제표 자동 회계 감사 스크립트 (6개 대표 기업 기준선 및 회귀 검증용)
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

const origFetchForVerify = globalThis.fetch.bind(globalThis);
let cachedCompanyTickersResponse = null;
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
    const res = await origFetchForVerify(input, init);
    if (res.ok) {
      const buf = await res.arrayBuffer();
      const hdrs = {};
      res.headers.forEach((v, k) => {
        if (!['content-encoding', 'content-length', 'transfer-encoding'].includes(k.toLowerCase())) {
          hdrs[k] = v;
        }
      });
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
  return origFetchForVerify(input, init);
};

const { POST } = await import('./app/api/analyze/route.js');

const DEFAULT_TICKERS = ['AAPL', 'SFM', 'CME', 'MSA', 'RLI', 'BRK-A'];
const SHEET_NAMES = ['손익계산서', '재무상태표', '현금흐름표'];
const BS_IDENTITY_TOLERANCE = 3; // 단위: 백만 달러 ($M), 반올림 오차 허용 범위 ±3M
const SCALE_JUMP_THRESHOLD = 1000; // 직전 연도 대비 1,000배 이상 급변 탐지 기준
const REQUEST_DELAY_MS = 5500; // SEC EDGAR rate-limit 보호용 요청 간 대기 시간

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
  // 한글 번역 문장 내 포함된 고유 법인명·거래소명·티커 약어
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

  // AMGN처럼 원본 BS에 [[부채총계]] 행이 생략된 기업은 [유동부채 합계] + 비유동부채 세부항목으로 가상 부채합계 산출
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

  // --- BS 유동자산 / 유동부채 ---
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
      // 보험사·복합금융사 등 미구분 재무상태표(Unclassified Balance Sheet)
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
        /^\[?비유동자산\s*합계\]?$|^total non-?current assets$/i,
        /기타비유동자산/
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

  // --- CF 영업/투자/재무 현금흐름 ---
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
          /^(?:연결)?당기순이익$|^계속영업순이익$|^총?순이익$|^net\s+(?:income|earnings)/i.test(
            r.label
          )
      );
      const isExcludedOpRow = (lbl) =>
        /^\[.*\]$/.test(lbl) ||
        /^(?:\[)?조정항목\s*합계(?:\])?$|^total\s+adjustments(?:\s+to\s+reconcile.*)?$/i.test(lbl);
      let opCompRows;
      if (netIncIdx !== -1) {
        let effectiveOpStart = netIncIdx;
        const contOpsIdx = cfRows.findIndex(
          (r, idx) =>
            idx > netIncIdx &&
            idx <= Math.min(opIdx - 1, netIncIdx + 5) &&
            /^계속영업(?:순|당기)?손?이익$/i.test(r.label)
        );
        if (contOpsIdx !== -1) {
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
          if (verifiedWaterfall) {
            effectiveOpStart = contOpsIdx;
          }
        }
        const rawOpRows = cfRows.slice(effectiveOpStart, opIdx).filter((r) => !isExcludedOpRow(r.label));
        const skippedInterm = new Set();
        for (let i = 1; i < rawOpRows.length; i++) {
          const cand = rawOpRows[i];
          const mightBeIntermSub =
            cand.isBold ||
            /운전자본.*(?:변동전|순증감|합계)|before\s+changes\s+in\s+(?:operating|working)/i.test(
              cand.label
            );
          if (!mightBeIntermSub) continue;
          for (let start = 0; start < i; start++) {
            const slice = rawOpRows
              .slice(start, i)
              .filter((_, idxRel) => !skippedInterm.has(start + idxRel));
            if (slice.length < 2) continue;
            const tol = Math.max(3, Math.ceil((slice.length + 1) * 0.5));
            let checked = 0;
            let matched = 0;
            for (let yi = 0; yi < cfYears.length; yi++) {
              const cVal = cand.values[yi];
              if (cVal === null) continue;
              checked++;
              const sSum = slice.reduce((a, r) => a + (r.values[yi] ?? 0), 0);
              if (Math.abs(sSum - cVal) <= tol) matched++;
            }
            if (checked > 0 && matched === checked) {
              skippedInterm.add(i);
              break;
            }
          }
        }
        opCompRows = rawOpRows.filter((_, i) => !skippedInterm.has(i));
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
      const invCompRows = cfRows
        .slice(invStart, invIdx)
        .filter((r) => !/^\[.*\]$/.test(r.label));
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
      const finCompRows = cfRows
        .slice(finStart, finIdx)
        .filter((r) => !/^\[.*\]$/.test(r.label));
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

    let scanRows = rows;
    let invIdx = -1;
    let invStart = 0;
    if (sheetName === '현금흐름표') {
      const findCfIdx = (exact, re) => {
        const e = rows.findIndex((r) => r.label === exact);
        return e !== -1 ? e : rows.findIndex((r) => re.test(r.label));
      };
      const opIdx = findCfIdx('[영업활동현금흐름]', /영업활동\s*현금흐름|net cash.*operating activities/i);
      invIdx = findCfIdx('[투자활동현금흐름]', /투자활동\s*현금흐름|net cash.*investing activities/i);
      const finIdx = findCfIdx('[재무활동현금흐름]', /재무활동\s*현금흐름|net cash.*financing activities/i);
      const sorted = [opIdx, invIdx, finIdx].filter((x) => x !== -1).sort((a, b) => a - b);
      if (invIdx !== -1) {
        for (const s of sorted) {
          if (s < invIdx) invStart = s + 1;
        }
      }
      const lastAct = Math.max(opIdx, invIdx, finIdx);
      if (lastAct !== -1) scanRows = rows.slice(0, lastAct + 1);
    }

    for (const row of scanRows) {
      const isMaEventRow =
        sheetName === '현금흐름표' &&
        /기업인수로\s*인한\s*현금유출|사업매각으로\s*인한\s*현금유입|business\s+acquisitions?|acquisitions?,?\s*net\s+of\s+cash/i.test(
          row.label
        );
      for (let i = 1; i < years.length; i++) {
        const prev = row.values[i - 1];
        const curr = row.values[i];
        if (prev === null || curr === null || prev === 0 || curr === 0) continue;

        const maxAbs = Math.max(Math.abs(prev), Math.abs(curr));
        const minAbs = Math.min(Math.abs(prev), Math.abs(curr));
        const ratio = maxAbs / minAbs;

        if (ratio >= SCALE_JUMP_THRESHOLD && (minAbs >= 5 || maxAbs >= 100000)) {
          if (isMaEventRow && invIdx !== -1) {
            const invCompRows = rows.slice(invStart, invIdx).filter((r) => !/^\[.*\]$/.test(r.label));
            const invOkForYear = (yi) => {
              const subV = rows[invIdx].values[yi];
              const comps = invCompRows.map((r) => r.values[yi]).filter((v) => v !== null && v !== undefined);
              if (subV === null || comps.length === 0) return false;
              const sum = comps.reduce((a, b) => a + b, 0);
              const tol = Math.max(3, Math.ceil((comps.length + 1) * 0.5));
              return Math.abs(sum - subV) <= tol;
            };
            if (invOkForYear(i - 1) && invOkForYear(i)) continue;
          }
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

// [검증 4] 현금유출 부호 오류 탐지: CAPEX, 배당금 지급, 자사주 매입 등 명백한 현금유출이 양수(+)인 경우
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
    /순증감|증감|유입|유출입|순발행|순상환|순차입|순조달|순액|발행\(순\)|발행\s*대금|발행\s*금액|투자금\s*상환|(?:우선주|지분)\s*투자.*?상환|(?:발행|차입)\s*(?:및|·|\()\s*상환|처분|매각및|사업매각|만기도래|원금회수|만기회수|조기상환|만기상환|회수|조정|수익|상각|손익|손실|이익|미지급|^\[.*\]$|proceeds|net\s+change|increase|decrease|issuance|amortization|extinguishment\s+of\s+debt|sale\s+of|maturities|collections|accrued|payable|beginning|end\s+of/i;
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
            if (Math.abs(invSum - invSubVal) <= invTol && Math.abs(invSum - 2 * val - invSubVal) > invTol) {
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

// [검증 5] 미번역 영문 잔존 검사: A열(계정명) 내 영문 알파벳 및 XBRL 태그명 검사
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

// [검증 6] BRK-A 전용 다차원 복합 지주사 구조 & 2025년 앵커링·시계열 분기 행 품질 검사
function auditBrkAStructure(ticker, sheets) {
  if (ticker !== 'BRK-A') {
    return { status: 'N/A', failures: [] };
  }

  const failures = [];
  const pushFail = (sheet, year, account, actual, expected) => {
    failures.push(makeFailure(ticker, 'BRK_STRUCTURE', '6. BRK-A 구조', sheet, year, account, actual, expected));
  };

  const is = sheets['손익계산서'];
  const bs = sheets['재무상태표'];
  const cf = sheets['현금흐름표'];

  // --- IS 구조 검사 ---
  if (!is || is.rows.length === 0) {
    pushFail('손익계산서', 'ALL', '(시트 누락)', '손익계산서 없음', 'BRK-A 손익계산서 존재');
  } else {
    const { years, rows } = is;
    const revTotalRows = rows.filter((r) => r.label === '매출액');
    if (revTotalRows.length !== 1) {
      pushFail('손익계산서', 'ALL', '매출액', `연결 매출액 행 ${revTotalRows.length}개 존재`, '단일 연결 매출액 행 1개만 존재');
    } else {
      const revRow = revTotalRows[0];
      years.forEach((yr, yi) => {
        if (revRow.values[yi] === null) {
          pushFail('손익계산서', yr, '매출액', 'null', '2016~2025 전 연도 연결 매출액 존재');
        }
      });
    }

    const dupLeasing = rows.filter((r) => /^리스수익$/i.test(r.label));
    if (dupLeasing.length > 0) {
      pushFail('손익계산서', 'ALL', '리스수익', `비부문 중복 행 ${dupLeasing.length}개 잔존`, '부문([보험및기타] 리스수익) 외 중복 제거');
    }
    const legacyFinIS = rows.filter((r) => /^\[금융·금융상품/.test(r.label));
    if (legacyFinIS.length > 0) {
      pushFail('손익계산서', 'ALL', '[금융·금융상품] 부문', `소급 재작성 완료 연도에 구 부문 ${legacyFinIS.length}행 잔존`, '2018 10-K 재작성 기준 2개 부문([보험및기타], [철도·유틸리티·에너지])만 유지');
    }

    const totalRevIdx = rows.findIndex((r) => r.label === '매출액');
    const totalCostIdx = rows.findIndex((r) => r.label === '총원가및영업비용');
    if (totalRevIdx === -1 || totalCostIdx === -1) {
      pushFail('손익계산서', 'ALL', '매출액 / 총원가및영업비용', `revIdx=${totalRevIdx}, costIdx=${totalCostIdx}`, '연결 매출액 및 총원가및영업비용 행 존재');
    } else {
      const insRevSubIdx = rows.findIndex((r) => r.label === '[보험및기타 매출액 합계]');
      const rrRevSubIdx = rows.findIndex((r) => r.label === '[철도·유틸리티·에너지 매출액 합계]');
      const insCostSubIdx = rows.findIndex((r) => r.label === '[보험및기타 원가및영업비용 합계]');
      const rrCostSubIdx = rows.findIndex((r) => r.label === '[철도·유틸리티·에너지 원가및영업비용 합계]');

      if (insRevSubIdx === -1 || insRevSubIdx >= totalRevIdx || rrRevSubIdx === -1 || rrRevSubIdx >= totalRevIdx) {
        pushFail('손익계산서', 'ALL', '부문별 매출액 합계 위치', `보험=${insRevSubIdx}, 철도=${rrRevSubIdx}, 연결매출=${totalRevIdx}`, '연결 매출액 위쪽에 양대 부문 매출 정위치');
      }
      if (
        insCostSubIdx === -1 ||
        insCostSubIdx <= totalRevIdx ||
        insCostSubIdx >= totalCostIdx ||
        rrCostSubIdx === -1 ||
        rrCostSubIdx <= totalRevIdx ||
        rrCostSubIdx >= totalCostIdx
      ) {
        pushFail('손익계산서', 'ALL', '부문별 원가및영업비용 합계 위치', `보험=${insCostSubIdx}, 철도=${rrCostSubIdx}, 연결비용=${totalCostIdx}`, '연결 매출액 아래와 총원가및영업비용 위쪽 사이에 정위치');
      }

      if (insRevSubIdx !== -1 && rrRevSubIdx !== -1) {
        years.forEach((yr, yi) => {
          const r1 = rows[insRevSubIdx].values[yi];
          const r2 = rows[rrRevSubIdx].values[yi];
          const rt = rows[totalRevIdx].values[yi];
          if (rt !== null) {
            if (r1 === null || r2 === null || Math.abs((r1 + r2) - rt) > 3) {
              pushFail('손익계산서', yr, '부문 매출 합계 vs 연결 매출액', `${r1} + ${r2} = ${(r1 ?? 0) + (r2 ?? 0)}`, `매출액 = ${rt} (오차 <= 3)`);
            }
          }
        });
      }
      if (insCostSubIdx !== -1 && rrCostSubIdx !== -1) {
        years.forEach((yr, yi) => {
          const c1 = rows[insCostSubIdx].values[yi];
          const c2 = rows[rrCostSubIdx].values[yi];
          const ct = rows[totalCostIdx].values[yi];
          if (ct !== null) {
            if (c1 === null || c2 === null || Math.abs((c1 + c2) - ct) > 3) {
              pushFail('손익계산서', yr, '부문 비용 합계 vs 총원가및영업비용', `${c1} + ${c2} = ${(c1 ?? 0) + (c2 ?? 0)}`, `총원가및영업비용 = ${ct} (오차 <= 3)`);
            }
          }
        });
      }
    }

    // Class A / Class B EPS & Shares 구분 검사
    const classAEps = rows.find((r) => /Class A.*주당순이익/i.test(r.label));
    const classBEps = rows.find((r) => /Class B.*주당순이익/i.test(r.label));
    const classAShares = rows.find((r) => /Class A.*발행주식수/i.test(r.label));
    const classBShares = rows.find((r) => /Class B.*발행주식수/i.test(r.label));
    if (!classAEps || !classBEps || !classAShares || !classBShares) {
      pushFail(
        '손익계산서',
        'ALL',
        'Class A / Class B 구분',
        `A_EPS=${Boolean(classAEps)}, B_EPS=${Boolean(classBEps)}, A_SH=${Boolean(classAShares)}, B_SH=${Boolean(classBShares)}`,
        'Class A / Class B 주당순이익 및 발행주식수 4행 모두 구분 존재'
      );
    } else {
      const y2025Idx = years.indexOf('2025');
      if (y2025Idx !== -1 && !(classAEps.values[y2025Idx] > classBEps.values[y2025Idx])) {
        pushFail('손익계산서', '2025', 'Class A vs Class B EPS', `A=${classAEps.values[y2025Idx]}, B=${classBEps.values[y2025Idx]}`, 'Class A EPS > Class B EPS');
      }
      years.forEach((yr, yi) => {
        for (const cr of [classAEps, classBEps, classAShares, classBShares]) {
          if (cr.values[yi] === null) {
            pushFail('손익계산서', yr, cr.label, 'null', '2016~2025 전 연도 수치 존재');
          }
        }
      });
    }

    // 시계열 분기 행 분리 (파생상품 계약 손익 & 투자손익)
    const derivRow = rows.find((r) => r.label === '파생상품 계약 손익');
    const invRow = rows.find((r) => r.label === '투자손익');
    if (!derivRow || !invRow) {
      pushFail('손익계산서', 'ALL', '시계열 분기 행(투자손익 / 파생상품 계약 손익)', `inv=${Boolean(invRow)}, deriv=${Boolean(derivRow)}`, '2025 앵커 투자손익 및 과거 분기 파생상품 계약 손익 분리');
    } else {
      for (const y of ['2016', '2017', '2018', '2019']) {
        const yi = years.indexOf(y);
        if (yi !== -1 && derivRow.values[yi] === null) {
          pushFail('손익계산서', y, '파생상품 계약 손익', 'null', '과거 분리 공시 연도(2016~2019) 수치 존재');
        }
      }
      for (const y of ['2023', '2024', '2025']) {
        const yi = years.indexOf(y);
        if (yi !== -1 && derivRow.values[yi] !== null) {
          pushFail('손익계산서', y, '파생상품 계약 손익', String(derivRow.values[yi]), '2023~2025년 null (시계열 분기 행 분리)');
        }
        if (yi !== -1 && invRow.values[yi] === null) {
          pushFail('손익계산서', y, '투자손익', 'null', '최신(2023~2025) 앵커 수치 존재');
        }
      }
    }
  }

  // --- BS 구조 검사 ---
  if (!bs || bs.rows.length === 0) {
    pushFail('재무상태표', 'ALL', '(시트 누락)', '재무상태표 없음', 'BRK-A 재무상태표 존재');
  } else {
    const { years, rows } = bs;
    const lastRow = rows[rows.length - 1];
    if (lastRow?.label !== '[[부채와자본총계]]') {
      pushFail('재무상태표', 'ALL', lastRow?.label || '(없음)', `마지막 행="${lastRow?.label}"`, '마지막 행="[[부채와자본총계]]"');
    }

    const reqRows = [
      rows.find((r) => r.label === '[보험및기타] 현금및현금성자산'),
      rows.find((r) => r.label === '[철도·유틸리티·에너지] 현금및현금성자산'),
      rows.find((r) => r.label === '[보험및기타] 미국 국채 단기투자'),
    ];
    const reqNames = [
      '[보험및기타] 현금및현금성자산',
      '[철도·유틸리티·에너지] 현금및현금성자산',
      '[보험및기타] 미국 국채 단기투자',
    ];
    reqRows.forEach((row, idx) => {
      if (!row) {
        pushFail('재무상태표', '2018~2025', reqNames[idx], '행 누락', '2018~2025 부문별 현금 및 미국 국채 단기투자 복원');
      } else {
        for (let y = 2018; y <= 2025; y++) {
          const yi = years.indexOf(String(y));
          if (yi !== -1 && !(row.values[yi] > 0)) {
            pushFail('재무상태표', String(y), reqNames[idx], String(row.values[yi]), '> 0 (누락 없이 복원)');
          }
        }
      }
    });

    const segAssetTotals = rows.filter((r) => /^\[.*자산 합계\]$/.test(r.label));
    const totalAssetsRow = rows.find((r) => r.label === '[[자산총계]]');
    if (segAssetTotals.length < 2 || !totalAssetsRow) {
      pushFail('재무상태표', 'ALL', '부문별 자산 합계', `부문자산합계=${segAssetTotals.length}개`, '부문별 자산 합계 및 [[자산총계]] 존재');
    } else {
      years.forEach((yr, yi) => {
        const ta = totalAssetsRow.values[yi];
        if (ta === null) return;
        const sumSeg = segAssetTotals.reduce((acc, r) => acc + (r.values[yi] ?? 0), 0);
        if (Math.abs(sumSeg - ta) > 3) {
          pushFail('재무상태표', yr, '부문별 자산 합계 vs [[자산총계]]', String(sumSeg), `${ta} (오차 <= 3)`);
        }
      });
    }
  }

  // --- CF 구조 검사 ---
  if (!cf || cf.rows.length === 0) {
    pushFail('현금흐름표', 'ALL', '(시트 누락)', '현금흐름표 없음', 'BRK-A 현금흐름표 존재');
  } else {
    const { years, rows } = cf;
    const finIdx = rows.findIndex((r) => r.label === '[재무활동현금흐름]');
    const reqCfLabels = [
      '[보험및기타] 장기차입금의 차입',
      '[보험및기타] 장기차입금의 상환',
      '[철도·유틸리티·에너지] 장기차입금의 차입',
      '[철도·유틸리티·에너지] 장기차입금의 상환',
    ];
    for (const lbl of reqCfLabels) {
      const idx = rows.findIndex((r) => r.label === lbl);
      if (idx === -1 || finIdx === -1 || idx >= finIdx) {
        pushFail('현금흐름표', 'ALL', lbl, `행위치=${idx}, [재무활동현금흐름]=${finIdx}`, '[재무활동현금흐름] 위쪽에 정위치');
      }
    }

    const dupNciRows = rows.filter((r) =>
      /비지배지분\s*취득.*기타|acquisitions?\s+of\s+noncontrolling\s+interests?\s+and\s+other/i.test(
        r.label
      )
    );
    const y2016Idx = years.indexOf('2016');
    for (const dr of dupNciRows) {
      if (y2016Idx !== -1 && dr.values[y2016Idx] !== null && dr.values[y2016Idx] > 0) {
        pushFail('현금흐름표', '2016', dr.label, `+${dr.values[y2016Idx]}`, '중복 및 양수 오류 제거');
      }
    }
  }

  return {
    status: failures.length === 0 ? 'PASS' : 'FAIL',
    failures,
  };
}

// [검증 7] 손익계산서 정합성·부호·연간매출 급변 감사 (Tax Benefit 정상 환급 허용)
function auditIncomeStatementIntegrity(ticker, sheets) {
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
    failures.push(makeFailure(ticker, 'IS_INTEGRITY', '7. IS 정합성', '손익계산서', year, account, actual, expected));
  };

  const inferSignVectors = (componentRows, baseRow, targetRow, startIdx, signFn) => {
    // 2025 앵커 방향(최신 최대 3개년)에서 산식이 성립하는 부호 벡터 추정
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
    for (let m = 0; m < (1 << k); m++) {
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

  // --- 7A. 매출/매출총이익/영업이익 & 은행 손익계산서 ---
  const revIdx = findMain(['매출액', '영업수익']);
  const cogsIdx = findMain(['매출원가']);
  const gpIdx = findMain(['매출총이익']);
  const opExpIdx = findMain(['영업비용합계', '총비용합계', '총원가및영업비용', '영업비용']);
  const opIdx = findMain(['영업이익']);

  // 은행업 IS 지표
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
    // 은행 IS: 이자수익 ± 이자비용 = 순이자손익
    const intFit = inferSignVectors([mainRows[intExpIdx]], mainRows[intIncIdx], mainRows[netIntIdx], 0, (yi, v) =>
      mainRows[intIncIdx].values[yi] + v[0] * (mainRows[intExpIdx].values[yi] ?? 0)
    );
    if (intFit) {
      years.forEach((yr, yi) => {
        const a = mainRows[intIncIdx].values[yi];
        const b = mainRows[intExpIdx].values[yi];
        const t = mainRows[netIntIdx].values[yi];
        if (a !== null && b !== null && t !== null) {
          const calc = a + intFit.vec[0] * b;
          if (Math.abs(calc - t) > intFit.tol) {
            pushFail(yr, '이자수익→순이자손익', `계산=${formatNum(calc)}`, `순이자손익=${formatNum(t)} (차이: ${formatNum(calc - t)})`);
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
          pushFail(yr, '순이자손익+비이자수익=순영업수익', `합계=${formatNum(a + b)}`, `순영업수익=${formatNum(t)}`);
        }
      });
    }
    if (nonIntExpIdx !== -1 && preTaxIdx !== -1 && netRevIdx < nonIntExpIdx && nonIntExpIdx < preTaxIdx) {
      const comps = provIdx !== -1 && provIdx > netRevIdx && provIdx < nonIntExpIdx ? [mainRows[provIdx], mainRows[nonIntExpIdx]] : [mainRows[nonIntExpIdx]];
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
            pushFail(yr, '순영업수익→법인세차감전순이익', `계산=${formatNum(calc)}`, `세전이익=${formatNum(t)} (차이: ${formatNum(calc - t)})`);
          }
        });
      }
    }
  } else {
    // 일반 기업 IS: 매출 → 매출총이익 → 영업이익
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
              pushFail(yr, '매출액-매출원가=매출총이익', `계산=${formatNum(calc)}`, `매출총이익=${formatNum(g)} (차이: ${formatNum(calc - g)})`);
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
          expCompRows.reduce((acc, cr, ci) => acc + v[ci] * (cr.values[yi] ?? 0), mainRows[baseForOpIdx].values[yi])
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

  // --- 7B. 영업이익 → 법인세비용차감전순이익 (이자비용·영업외손익 부호 반전 탐지) ---
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
      const anchorCompIndices =
        nonOpTotIdx === -1
          ? comps
              .map((cr, ci) => (cr.values.slice(recentStartYi).some((v) => v !== null) ? ci : -1))
              .filter((ci) => ci !== -1)
          : [];
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

  // --- 7C. 법인세비용차감전순이익 ± 법인세비용 = (연결/계속영업)당기순이익 ---
  // 안전장치 2: 법인세 환급(음수/양수 전환) 자체는 오탐하지 않고 산식 불일치 시에만 FAIL
  const taxIdx = preTaxIdx !== -1 ? findMain(['법인세비용'], null, preTaxIdx + 1) : -1;
  if (preTaxIdx !== -1 && taxIdx !== -1) {
    let netCandidates = mainRows.filter(
      (r, i) =>
        i > taxIdx &&
        !isSegRow(r) &&
        (r.label === '계속영업순이익' || r.label === '연결당기순이익' || r.label === '당기순이익')
    );
    if (!netCandidates.some((r) => r.label === '계속영업순이익')) {
      const hasDiscRowInIs = rows.some((r) => /중단영업/.test(r.label));
      const contAfterMeta = rows.find((r) => !isSegRow(r) && r.label === '계속영업순이익');
      if (hasDiscRowInIs && contAfterMeta) {
        let mOk = 0;
        let pOk = 0;
        for (let yi = years.length - 1; yi >= Math.max(0, years.length - 4); yi--) {
          const pt = mainRows[preTaxIdx].values[yi];
          const tx = mainRows[taxIdx].values[yi];
          const cn = contAfterMeta.values[yi];
          if (pt === null || tx === null || cn === null) continue;
          if (Math.abs(pt - tx - cn) <= 3) mOk++;
          if (Math.abs(pt + tx - cn) <= 3) pOk++;
        }
        if (mOk > 0 || pOk > 0) {
          netCandidates = [contAfterMeta, ...netCandidates];
        }
      }
    }
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

  // --- 7D. 연간 매출액 비정상 급락/급변 (≥ 3배, 분기 컬럼 혼입·세분화 덮어쓰기 탐지) ---
  const primaryRevRowIdx = revIdx !== -1 ? revIdx : netRevIdx;
  if (primaryRevRowIdx !== -1) {
    const rRow = mainRows[primaryRevRowIdx];
    for (let i = 1; i < years.length; i++) {
      const prev = rRow.values[i - 1];
      const curr = rRow.values[i];
      if (prev === null || curr === null || prev <= 0 || curr <= 0) continue;
      const ratio = Math.max(prev, curr) / Math.min(prev, curr);
      if (ratio >= 3.0) {
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

// [신규 감사 A·B] 동일 시트 다개년 동일 수치 중복 행 탐지 & BS 주식수(sharesItemType) 혼입 탐지
function auditStructuralAnomalies(ticker, sheets) {
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
      let checkedYears = 0;
      let okYears = 0;
      let recentFail = false;
      const recentStartYi = Math.max(0, yearsLength - 3);
      for (let yi = 0; yi < yearsLength; yi++) {
        const subVal = rows[hi].values[yi];
        const compVal = rows[lo].values[yi];
        if (subVal === null || compVal === null) continue;
        checkedYears++;
        let fullSum = 0;
        for (const c of compRows) fullSum += c.row.values[yi] ?? 0;
        if (Math.abs(fullSum - subVal) <= tol) {
          okYears++;
        } else if (yi >= recentStartYi) {
          recentFail = true;
        }
      }
      if (checkedYears > 0 && !recentFail && okYears / checkedYears >= 0.8) return true;
    }
    return false;
  };

  // Rule B: BS 내 주식수(sharesItemType) 라벨 혼입 탐지
  const bs = sheets['재무상태표'];
  if (bs) {
    for (const row of bs.rows) {
      if (
        /주식수|주수|\bauthorized\s+shares\b|\bissued\s+shares\b|\boutstanding\s+shares\b|\btreasury.*shares\b/i.test(
          row.label
        )
      ) {
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
        if (sheetName === '현금흐름표' && cfNetChangeIdx !== -1 && (i > cfNetChangeIdx || j > cfNetChangeIdx)) {
          if (i > cfNetChangeIdx && j > cfNetChangeIdx) continue;
          const normA = ra.label.replace(/\s+/g, '');
          const normB = rb.label.replace(/\s+/g, '');
          if (normA !== normB) continue;
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
          }
        }
        const hasDecimalVal = Array.from(distinctVals).some((v) => !Number.isInteger(v));
        if (
          matchedCount >= 5 &&
          distinctVals.size >= 2 &&
          bothNonNullCount > 0 &&
          matchedCount / bothNonNullCount >= 0.85 &&
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

async function analyzeAndAuditTicker(ticker, apiKey, userEmail) {
  const startedAt = Date.now();
  const req = new Request('http://localhost/api/analyze', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ ticker, apiKey, userEmail }),
  });

  const res = await POST(req);
  const elapsedSec = ((Date.now() - startedAt) / 1000).toFixed(1);

  if (!res.ok) {
    let errMsg = `HTTP ${res.status}`;
    try {
      const text = await res.text();
      try {
        const parsed = JSON.parse(text);
        errMsg = `${errMsg}: ${parsed.error || text}`;
      } catch {
        errMsg = `${errMsg}: ${text}`;
      }
    } catch {
      // 응답 본문 읽기 실패 시 상태코드만 유지
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
      elapsedSec,
      yearRange: '-',
      rowCounts: '-',
      apiError: errMsg,
      checks: {
        BS_IDENTITY: { status: 'FAIL', failures: [apiFail] },
        SUBTOTALS: {
          status: 'FAIL',
          failures: [],
          breakdown: {
            유동자산: 'FAIL',
            유동부채: 'FAIL',
            영업현금흐름: 'FAIL',
            투자현금흐름: 'FAIL',
            재무현금흐름: 'FAIL',
          },
        },
        SCALE_JUMP: { status: 'FAIL', failures: [] },
        OUTFLOW_SIGN: { status: 'FAIL', failures: [] },
        UNTRANSLATED: { status: 'FAIL', failures: [] },
        BRK_STRUCTURE: ticker === 'BRK-A' ? { status: 'FAIL', failures: [] } : { status: 'N/A', failures: [] },
        IS_INTEGRITY: { status: 'FAIL', failures: [] },
        STRUCTURAL_ANOMALY: { status: 'FAIL', failures: [] },
      },
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
  const c7 = auditIncomeStatementIntegrity(ticker, sheets);
  const c8s = auditStructuralAnomalies(ticker, sheets);

  const allFailures = [
    ...c1.failures,
    ...c2.failures,
    ...c3.failures,
    ...c4.failures,
    ...c5.failures,
    ...c6.failures,
    ...c7.failures,
    ...c8s.failures,
  ];

  const overall =
    c1.status === 'PASS' &&
    c2.status === 'PASS' &&
    c3.status === 'PASS' &&
    c4.status === 'PASS' &&
    c5.status === 'PASS' &&
    c6.status !== 'FAIL' &&
    c7.status === 'PASS' &&
    c8s.status === 'PASS'
      ? 'PASS'
      : 'FAIL';

  return {
    ticker,
    elapsedSec,
    yearRange,
    rowCounts,
    apiError: null,
    checks: {
      BS_IDENTITY: c1,
      SUBTOTALS: c2,
      SCALE_JUMP: c3,
      OUTFLOW_SIGN: c4,
      UNTRANSLATED: c5,
      BRK_STRUCTURE: c6,
      IS_INTEGRITY: c7,
      STRUCTURAL_ANOMALY: c8s,
    },
    allFailures,
    overall,
  };
}

function formatCheckCell(checkRes) {
  if (!checkRes || checkRes.status === 'N/A') return '-';
  if (checkRes.status === 'PASS') return 'PASS';
  return `FAIL(${checkRes.failures.length})`;
}

async function main() {
  const loadedEnvSource = loadDotEnvIfPresent();

  const cliTickers = process.argv
    .slice(2)
    .map((t) => t.trim().toUpperCase())
    .filter(Boolean);
  const targetTickers = cliTickers.length > 0 ? cliTickers : DEFAULT_TICKERS;

  const envApiKey = (process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || '').trim();
  const hasRealGeminiKey = Boolean(envApiKey);
  const apiKey = envApiKey || 'VERIFY_NO_GEMINI_KEY_FALLBACK';
  const userEmail = (process.env.SEC_USER_EMAIL || 'sec_user@gmail.com').trim();

  console.log('====================================================================================================');
  console.log(' SEC 10-K 자동 회계 감사 스크립트 (verify.mjs)');
  console.log('====================================================================================================');
  console.log(` 점검 대상 (${targetTickers.length}개): ${targetTickers.join(', ')}`);
  console.log(
    ` Gemini 번역 모드: ${
      hasRealGeminiKey
        ? `API Key 활성 (${loadedEnvSource ? `${loadedEnvSource} 로드됨` : '환경변수'}) - 실시간 AI 번역 포함`
        : 'API Key 미설정 (사전 매핑 + fallbackKoreanLabelBuilder 기준선 모드)'
    }`
  );
  console.log(
    ` 허용 오차 기준  : [1] 재무상태표 항등식 ±${BS_IDENTITY_TOLERANCE}M | [2] 주요 소계 max(3, ceil((항목수+1)*0.5))M | [3] 급변 ≥${SCALE_JUMP_THRESHOLD}배`
  );
  console.log('----------------------------------------------------------------------------------------------------');

  const results = [];
  for (let i = 0; i < targetTickers.length; i++) {
    const ticker = targetTickers[i];
    if (i > 0) await delay(REQUEST_DELAY_MS);
    process.stdout.write(`[${i + 1}/${targetTickers.length}] ${ticker} 분석 및 감사 실행 중... `);
    try {
      const res = await analyzeAndAuditTicker(ticker, apiKey, userEmail);
      results.push(res);
      if (res.apiError) {
        console.log(`오류 (${res.elapsedSec}s) -> ${res.apiError}`);
      } else {
        console.log(
          `완료 (${res.elapsedSec}s, 연도: ${res.yearRange}, 행수 IS/BS/CF: ${res.rowCounts}) => ${res.overall}`
        );
      }
    } catch (err) {
      const errMsg = err?.message || String(err);
      console.log(`예외 발생 -> ${errMsg}`);
      const crashFail = makeFailure(
        ticker,
        'CRASH',
        '실행 예외',
        '-',
        'ALL',
        '(실행 중 예외)',
        errMsg,
        '정상 완료'
      );
      results.push({
        ticker,
        elapsedSec: '0.0',
        yearRange: '-',
        rowCounts: '-',
        apiError: errMsg,
        checks: {
          BS_IDENTITY: { status: 'FAIL', failures: [crashFail] },
          SUBTOTALS: {
            status: 'FAIL',
            failures: [],
            breakdown: {
              유동자산: 'FAIL',
              유동부채: 'FAIL',
              영업현금흐름: 'FAIL',
              투자현금흐름: 'FAIL',
              재무현금흐름: 'FAIL',
            },
          },
          SCALE_JUMP: { status: 'FAIL', failures: [] },
          OUTFLOW_SIGN: { status: 'FAIL', failures: [] },
          UNTRANSLATED: { status: 'FAIL', failures: [] },
          BRK_STRUCTURE: ticker === 'BRK-A' ? { status: 'FAIL', failures: [] } : { status: 'N/A', failures: [] },
          IS_INTEGRITY: { status: 'FAIL', failures: [] },
          STRUCTURAL_ANOMALY: { status: 'FAIL', failures: [] },
        },
        allFailures: [crashFail],
        overall: 'FAIL',
      });
    }
  }

  console.log('\n====================================================================================================');
  console.log(' [요약표] 기업별 · 검증 항목별 감사 결과');
  console.log('====================================================================================================');

  const cols = [
    { header: '기업', width: 8 },
    { header: '연도범위', width: 11 },
    { header: '1.BS항등식', width: 12 },
    { header: '2.주요소계', width: 12 },
    { header: '3.단위급변', width: 12 },
    { header: '4.유출부호', width: 12 },
    { header: '5.미번역영문', width: 14 },
    { header: '6.BRK구조', width: 12 },
    { header: '7.IS정합성', width: 12 },
    { header: '구조/중복', width: 11 },
    { header: '종합', width: 8 },
  ];

  const headerLine = cols.map((c) => padEndDisplay(c.header, c.width)).join(' | ');
  console.log(headerLine);
  console.log('-'.repeat(displayWidth(headerLine)));

  for (const r of results) {
    const line = [
      padEndDisplay(r.ticker, cols[0].width),
      padEndDisplay(r.yearRange, cols[1].width),
      padEndDisplay(formatCheckCell(r.checks.BS_IDENTITY), cols[2].width),
      padEndDisplay(formatCheckCell(r.checks.SUBTOTALS), cols[3].width),
      padEndDisplay(formatCheckCell(r.checks.SCALE_JUMP), cols[4].width),
      padEndDisplay(formatCheckCell(r.checks.OUTFLOW_SIGN), cols[5].width),
      padEndDisplay(formatCheckCell(r.checks.UNTRANSLATED), cols[6].width),
      padEndDisplay(formatCheckCell(r.checks.BRK_STRUCTURE), cols[7].width),
      padEndDisplay(formatCheckCell(r.checks.IS_INTEGRITY), cols[8].width),
      padEndDisplay(formatCheckCell(r.checks.STRUCTURAL_ANOMALY), cols[9].width),
      padEndDisplay(r.overall, cols[10].width),
    ].join(' | ');
    console.log(line);
  }

  console.log('\n----------------------------------------------------------------------------------------------------');
  console.log(' [세부표] 검증 2 (주요 소계 합산) 항목별 상태');
  console.log('----------------------------------------------------------------------------------------------------');
  const subCols = [
    { header: '기업', width: 8 },
    { header: '유동자산', width: 14 },
    { header: '비유동자산', width: 14 },
    { header: '유동부채', width: 14 },
    { header: '비유동부채', width: 14 },
    { header: '영업CF', width: 10 },
    { header: '투자CF', width: 10 },
    { header: '재무CF', width: 10 },
  ];
  const subHeader = subCols.map((c) => padEndDisplay(c.header, c.width)).join(' | ');
  console.log(subHeader);
  console.log('-'.repeat(displayWidth(subHeader)));
  for (const r of results) {
    const b = r.checks.SUBTOTALS.breakdown || {};
    console.log(
      [
        padEndDisplay(r.ticker, subCols[0].width),
        padEndDisplay(b['유동자산'] || '-', subCols[1].width),
        padEndDisplay(b['비유동자산'] || '-', subCols[2].width),
        padEndDisplay(b['유동부채'] || '-', subCols[3].width),
        padEndDisplay(b['비유동부채'] || '-', subCols[4].width),
        padEndDisplay(b['영업현금흐름'] || '-', subCols[5].width),
        padEndDisplay(b['투자현금흐름'] || '-', subCols[6].width),
        padEndDisplay(b['재무현금흐름'] || '-', subCols[7].width),
      ].join(' | ')
    );
  }

  const totalFailures = results.reduce((acc, r) => acc + r.allFailures.length, 0);
  const failedTickers = results.filter((r) => r.overall === 'FAIL');

  if (totalFailures > 0) {
    console.log('\n====================================================================================================');
    console.log(` [FAIL 상세 리포트] 총 ${totalFailures}건 (실패 기업: ${failedTickers.length}/${results.length})`);
    console.log('====================================================================================================');

    for (const r of results) {
      if (r.allFailures.length === 0) continue;
      console.log(`\n▶ ${r.ticker} (총 ${r.allFailures.length}건 FAIL)`);
      const grouped = new Map();
      for (const f of r.allFailures) {
        if (!grouped.has(f.checkName)) grouped.set(f.checkName, []);
        grouped.get(f.checkName).push(f);
      }

      for (const [checkName, items] of grouped.entries()) {
        console.log(`  [${checkName}] (${items.length}건)`);
        for (const f of items) {
          console.log(
            `    - 기업: ${f.ticker} | 시트: ${f.sheet} | 연도: ${f.year} | 계정명: "${f.account}"`
          );
          console.log(`      실제값: ${f.actual}`);
          console.log(`      기대값: ${f.expected}`);
        }
      }
    }
  } else {
    console.log('\n모든 검증 항목(5대 회계 감사 + BRK-A 구조 검증)을 100% 통과했습니다!');
  }

  const passCount = results.filter((r) => r.overall === 'PASS').length;
  console.log('\n====================================================================================================');
  console.log(
    ` 최종 집계: ${padStartDisplay(passCount, 1)}/${results.length}개 기업 PASS | 총 FAIL 건수: ${totalFailures}건`
  );
  console.log('====================================================================================================');

  process.exitCode = failedTickers.length > 0 || totalFailures > 0 ? 1 : 0;
}

main().catch((err) => {
  console.error('verify.mjs 치명적 오류:', err);
  process.exitCode = 1;
});

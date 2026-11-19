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

const { POST } = await import('./app/api/analyze/route.js');

const DEFAULT_TICKERS = ['AAPL', 'SFM', 'CME', 'MSA', 'RLI', 'BRK-A'];
const SHEET_NAMES = ['손익계산서', '재무상태표', '현금흐름표'];
const BS_IDENTITY_TOLERANCE = 3; // 단위: 백만 달러 ($M), 반올림 오차 허용 범위 ±3M
const SCALE_JUMP_THRESHOLD = 1000; // 직전 연도 대비 1,000배 이상 급변 탐지 기준
const REQUEST_DELAY_MS = 400; // SEC EDGAR rate-limit 보호용 요청 간 대기 시간

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
  const liabRow = findRow('[[부채총계]]', /^\[*부채총계\]*$|^총부채$/i, /자본/);
  const equityRow = findRow('[[자본총계]]', /^\[*자본총계\]*$|^총자본$/i, /부채/, true);
  const liabAndEqRow = findRow('[[부채와자본총계]]', /부채(?:와|및)자본총계/i);
  const mezzanineRows = rows.filter((r) =>
    /상환가능비지배지분|상환가능지분|메자닌|redeemable\s+noncontrolling|temporary\s+equity/i.test(
      r.label
    )
  );

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

// [검증 2] 주요 소계 합산 일치: 유동자산, 유동부채, 영업CF, 투자CF, 재무CF
function auditMajorSubtotals(ticker, sheets) {
  const failures = [];
  const breakdown = {
    유동자산: 'FAIL',
    유동부채: 'FAIL',
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
      breakdown['유동부채'] = 'N/A(미구분BS)';
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
        const caCompRows = bsRows
          .slice(0, caIdx)
          .filter((r) => !/^\[.*합계\]$/.test(r.label));
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
      const netIncIdx = cfRows.findIndex(
        (r, idx) =>
          idx < opIdx &&
          /^(?:연결)?당기순이익$|^계속영업순이익$|^총?순이익$|^net\s+(?:income|earnings)/i.test(
            r.label
          )
      );
      let opCompRows;
      if (netIncIdx !== -1) {
        opCompRows = cfRows
          .slice(netIncIdx, opIdx)
          .filter((r) => !/^\[.*\]$/.test(r.label));
      } else {
        opCompRows = cfRows
          .slice(0, opIdx)
          .filter(
            (r) =>
              !/^\[.*\]$/.test(r.label) &&
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
      const invStart = opIdx !== -1 && opIdx < invIdx ? opIdx + 1 : 0;
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
      let finStart = 0;
      if (invIdx !== -1 && invIdx < finIdx) {
        finStart = invIdx + 1;
      } else if (opIdx !== -1 && opIdx < finIdx) {
        finStart = opIdx + 1;
      }
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

    for (const row of rows) {
      for (let i = 1; i < years.length; i++) {
        const prev = row.values[i - 1];
        const curr = row.values[i];
        if (prev === null || curr === null || prev === 0 || curr === 0) continue;

        const maxAbs = Math.max(Math.abs(prev), Math.abs(curr));
        const minAbs = Math.min(Math.abs(prev), Math.abs(curr));
        const ratio = maxAbs / minAbs;

        if (ratio >= SCALE_JUMP_THRESHOLD) {
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
    /순증감|증감|유입|처분|매각및|사업매각|만기도래|원금회수|만기회수|조기상환|만기상환|회수|조정|수익|미지급|^\[|proceeds|net\s+change|increase|decrease|issuance|sale\s+of|maturities|collections|accrued|payable|beginning|end\s+of/i;
  const outflowRegex =
    /(?:CAPEX|자본적\s*지출|유형자산의?\s*(?:취득|매입)|생산설비(?:자산)?의?\s*취득|무형자산의?\s*취득|배당금의?\s*지급|현금배당(?:금)?\s*지급|자사주\s*매입|자기주식의?\s*(?:취득|매입)|(?:채무증권|지분증권|유가증권|단기금융상품|투자자산|대출채권|리스채권|비지배지분).*?(?:취득|매입)|기업인수로\s*인한\s*현금유출|관련\s*현금유출|(?:차입금|사채|부채|리스부채|우선주).*?상환|상환에\s*따른\s*현금유출|원천징수세액\s*납부|이연금융원가의?\s*지급|^(?:purchases?\s+of|additions?\s+to|capital\s+expenditures?|payments?\s+(?:to\s+acquire|for\s+repurchase|of\s+dividends|for\s+dividends|on\s+long-?term\s+debt|on\s+.*lease)|repayments?\s+of|repurchases?\s+of|common\s+stock\s+repurchased|treasury\s+stock\s+purchases?|acquisitions?,?\s*net\s+of\s+cash|business\s+acquisitions?|cash\s+dividends\s+paid|dividends\s+paid|redemption\s+of|retirement\s+of)|\bcapital\s+expenditures?\b|\bdividends?\s+paid\b|\brepurchases?\s+of\s+(?:common|treasury|capital)\s+stock\b)/i;

  const { years, rows } = cf;
  const finIdx = rows.findIndex(
    (r) =>
      r.label === '[재무활동현금흐름]' ||
      /재무활동\s*현금흐름|net cash.*financing activities/i.test(r.label)
  );
  const activityRows = finIdx !== -1 ? rows.slice(0, finIdx) : rows;
  for (const row of activityRows) {
    if (/^\[.*\]$/.test(row.label)) continue;
    if (exclusionRegex.test(row.label)) continue;
    if (!outflowRegex.test(row.label)) continue;

    for (let yi = 0; yi < years.length; yi++) {
      const val = row.values[yi];
      if (val !== null && val > 0) {
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

  const allFailures = [
    ...c1.failures,
    ...c2.failures,
    ...c3.failures,
    ...c4.failures,
    ...c5.failures,
  ];

  const overall =
    c1.status === 'PASS' &&
    c2.status === 'PASS' &&
    c3.status === 'PASS' &&
    c4.status === 'PASS' &&
    c5.status === 'PASS'
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
    },
    allFailures,
    overall,
  };
}

function formatCheckCell(checkRes) {
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
      padEndDisplay(r.overall, cols[7].width),
    ].join(' | ');
    console.log(line);
  }

  console.log('\n----------------------------------------------------------------------------------------------------');
  console.log(' [세부표] 검증 2 (주요 소계 합산) 항목별 상태');
  console.log('----------------------------------------------------------------------------------------------------');
  const subCols = [
    { header: '기업', width: 8 },
    { header: '유동자산', width: 14 },
    { header: '유동부채', width: 14 },
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
        padEndDisplay(b['유동부채'] || '-', subCols[2].width),
        padEndDisplay(b['영업현금흐름'] || '-', subCols[3].width),
        padEndDisplay(b['투자현금흐름'] || '-', subCols[4].width),
        padEndDisplay(b['재무현금흐름'] || '-', subCols[5].width),
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
    console.log('\n모든 검증 항목(5대 회계 감사)을 100% 통과했습니다!');
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

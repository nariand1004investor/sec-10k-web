// app/api/tickers/route.js
import { NextResponse } from 'next/server';

let cachedTickers = null;

// 복사 시 URL 깨짐 방지용 헬퍼
function makeUrl(hostParts, pathPart) {
  return ['https:', '', hostParts.join('.'), pathPart].join('/');
}

export async function GET() {
  try {
    if (cachedTickers) {
      return NextResponse.json(cachedTickers);
    }

    const tickerUrl = makeUrl(['www', 'sec', 'gov'], 'files/company_tickers.json');
    const response = await fetch(tickerUrl, {
      headers: {
        'User-Agent': 'SecExcelWebApp sec_user@gmail.com',
        'Accept-Encoding': 'gzip, deflate',
        Accept: 'application/json',
      },
    });

    if (!response.ok) {
      throw new Error(`SEC 티커 목록 응답 오류: ${response.status}`);
    }

    const data = await response.json();
    cachedTickers = Object.values(data).map((item) => ({
      ticker: item.ticker,
      title: item.title,
      cik: String(item.cik_str).padStart(10, '0'),
    }));

    return NextResponse.json(cachedTickers);
  } catch (error) {
    console.error('티커 목록 로드 실패:', error.message);
    return NextResponse.json([], { status: 200 });
  }
}
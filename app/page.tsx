"use client";

import React, { useState, useEffect, useRef } from 'react';

// 복사 시 URL 마크다운 변환 방지용 헬퍼
function makeExternalUrl(hostParts: string[], pathPart: string) {
  return ['https:', '', hostParts.join('.'), pathPart].join('/');
}

export default function FinancialAnalyzerApp() {
  const [tickerInput, setTickerInput] = useState('');
  const [selectedCompany, setSelectedCompany] = useState<any>(null);
  const [allTickers, setAllTickers] = useState<any[]>([]);
  const [suggestions, setSuggestions] = useState<any[]>([]);
  const [showDropdown, setShowDropdown] = useState(false);
  const [highlightIndex, setHighlightIndex] = useState(-1);

  const [status, setStatus] = useState<'idle' | 'processing' | 'success'>('idle');
  const [progressText, setProgressText] = useState('');
  const [progressPercent, setProgressPercent] = useState(0);
  const [savedFileName, setSavedFileName] = useState('');

  const [isSettingsOpen, setIsSettingsOpen] = useState(false);
  const [apiKey, setApiKey] = useState('');
  const [userEmail, setUserEmail] = useState('');

  const [alertModal, setAlertModal] = useState({
    isOpen: false,
    title: '',
    message: '',
  });

  const dropdownRef = useRef<HTMLDivElement>(null);
  const abortControllerRef = useRef<AbortController | null>(null);

  // 1. 초기 로드: 저장된 API 키 불러오기 및 웹 서버에서 SEC 티커 목록 로드
  useEffect(() => {
    const savedKey = localStorage.getItem('gemini_api_key') || '';
    const savedEmail = localStorage.getItem('sec_user_email') || '';
    setApiKey(savedKey);
    setUserEmail(savedEmail);

    fetch('/api/tickers')
      .then((res) => res.json())
      .then((list) => {
        if (Array.isArray(list)) setAllTickers(list);
      })
      .catch(() => {});
  }, []);

  // 외부 클릭 시 자동완성 드롭다운 닫기
  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (dropdownRef.current && !dropdownRef.current.contains(e.target as Node)) {
        setShowDropdown(false);
      }
    };
    document.addEventListener('mousedown', handleClickOutside);
    return () => document.removeEventListener('mousedown', handleClickOutside);
  }, []);

  const handleInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const value = e.target.value;
    setTickerInput(value);
    setSelectedCompany(null);

    if (!value.trim()) {
      setSuggestions([]);
      setShowDropdown(false);
      return;
    }

    const query = value.trim().toUpperCase();
    const filtered = allTickers
      .filter(
        (item) =>
          item.ticker.toUpperCase().startsWith(query) ||
          item.title.toUpperCase().includes(query)
      )
      .slice(0, 8);

    setSuggestions(filtered);
    setShowDropdown(filtered.length > 0);
    setHighlightIndex(-1);
  };

  const handleSelectSuggestion = (item: any) => {
    setTickerInput(item.ticker);
    setSelectedCompany(item);
    setShowDropdown(false);
  };

  const handleKeyDown = (e: React.KeyboardEvent<HTMLInputElement>) => {
    if (!showDropdown || suggestions.length === 0) return;

    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setHighlightIndex((prev) => (prev < suggestions.length - 1 ? prev + 1 : 0));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setHighlightIndex((prev) => (prev > 0 ? prev - 1 : suggestions.length - 1));
    } else if (e.key === 'Enter' && highlightIndex >= 0) {
      e.preventDefault();
      handleSelectSuggestion(suggestions[highlightIndex]);
    } else if (e.key === 'Escape') {
      setShowDropdown(false);
    }
  };

  const handleSaveSettings = (e: React.FormEvent) => {
    e.preventDefault();
    localStorage.setItem('gemini_api_key', apiKey.trim());
    localStorage.setItem('sec_user_email', userEmail.trim());
    setIsSettingsOpen(false);
  };

  // 분석 및 엑셀 다운로드 실행 (웹 브라우저 방식)
  const handleStartAnalysis = async (e: React.FormEvent) => {
    e.preventDefault();
    const targetTicker = tickerInput.trim().toUpperCase();

    if (!targetTicker) {
      setAlertModal({
        isOpen: true,
        title: '티커 입력 필요',
        message: '분석할 미국 상장기업의 티커(예: AAPL, SFM, MSFT)를 입력해 주세요.',
      });
      return;
    }

    if (!apiKey.trim()) {
      setAlertModal({
        isOpen: true,
        title: 'Gemini API 키 설정 필요',
        message: '우측 상단의 [환경설정] 버튼을 눌러 본인의 Gemini API 키를 먼저 등록해 주세요.',
      });
      return;
    }

    setShowDropdown(false);
    setStatus('processing');
    setProgressText('SEC EDGAR에서 10년 치 10-K 재무제표 수집 중 ⏳');
    setProgressPercent(25);

    const controller = new AbortController();
    abortControllerRef.current = controller;

    // 진행률 게이지 애니메이션 타이머
    const timer1 = setTimeout(() => {
      setProgressText('Gemini AI가 영문 계정명을 한글 회계 계정명으로 번역 중 ⚙️');
      setProgressPercent(65);
    }, 2200);

    const timer2 = setTimeout(() => {
      setProgressText('손익계산서·재무상태표·현금흐름표 엑셀 파일 생성 중 📊');
      setProgressPercent(88);
    }, 5000);

    try {
      const response = await fetch('/api/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          ticker: targetTicker,
          apiKey: apiKey.trim(),
          userEmail: userEmail.trim() || 'sec_user@gmail.com',
        }),
        signal: controller.signal,
      });

      clearTimeout(timer1);
      clearTimeout(timer2);

      if (!response.ok) {
        const errData = await response.json().catch(() => ({}));
        throw new Error(errData.error || '분석 중 서버 오류가 발생했습니다.');
      }

      // 엑셀 파일(Blob) 브라우저 즉시 다운로드 처리
      const blob = await response.blob();
      const encodedName = response.headers.get('X-File-Name');
      const fileName = encodedName
        ? decodeURIComponent(encodedName)
        : `${targetTicker}_10-K_재무제표.xlsx`;

      const downloadUrl = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = downloadUrl;
      link.download = fileName;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(downloadUrl);

      setProgressPercent(100);
      setSavedFileName(fileName);
      setStatus('success');
    } catch (err: any) {
      clearTimeout(timer1);
      clearTimeout(timer2);

      if (err.name === 'AbortError') {
        setStatus('idle');
        setProgressPercent(0);
        return;
      }

      setStatus('idle');
      setProgressPercent(0);
      setAlertModal({
        isOpen: true,
        title: '분석 안내',
        message: err.message || '통신 중 오류가 발생했습니다.',
      });
    } finally {
      abortControllerRef.current = null;
    }
  };

  const handleCancel = () => {
    if (abortControllerRef.current) {
      abortControllerRef.current.abort();
    }
    setStatus('idle');
    setProgressPercent(0);
  };

  const handleReset = () => {
    setStatus('idle');
    setTickerInput('');
    setSelectedCompany(null);
    setProgressPercent(0);
    setSavedFileName('');
  };

  const openLink = (url: string) => {
    window.open(url, '_blank', 'noopener,noreferrer');
  };

  const blogUrl = makeExternalUrl(['blog', 'naver', 'com'], 'jiniand1004');
  const apiKeyUrl = makeExternalUrl(['aistudio', 'google', 'com'], 'app/apikey');
  const bookUrl = makeExternalUrl(['www', 'yes24', 'com'], 'product/goods/124838987');

  return (
    <div className="min-h-screen bg-white text-gray-800 flex flex-col justify-between p-8 select-none">
      {/* 상단 네비게이션 */}
      <header className="max-w-4xl w-full mx-auto flex justify-between items-center border-b border-gray-100 pb-4">
        <div className="flex items-center space-x-2">
          <span className="w-3 h-3 rounded-full bg-blue-600"></span>
          <span className="font-semibold text-gray-700 text-sm">SEC EDGAR 10-K Analyzer</span>
        </div>

        <div className="flex items-center space-x-3">
          <button
            type="button"
            onClick={() => openLink(blogUrl)}
            className="text-xs font-semibold text-blue-600 hover:underline bg-blue-50 px-3 py-2 rounded-lg border border-blue-100 transition"
          >
            Made by 현명한직장인 ↗
          </button>

          <button
            type="button"
            onClick={() => setIsSettingsOpen(true)}
            className="flex items-center space-x-1.5 text-sm font-medium text-gray-600 hover:text-blue-600 bg-gray-50 hover:bg-blue-50 px-3.5 py-2 rounded-lg border border-gray-200 transition"
          >
            <span>⚙️ 환경설정 (API Key)</span>
            {!apiKey && <span className="w-2 h-2 rounded-full bg-red-500"></span>}
          </button>
        </div>
      </header>

      {/* 메인 컨텐츠 영역 */}
      <main className="max-w-2xl w-full mx-auto my-auto py-12 text-center space-y-8">
        <div className="space-y-3">
          <h1 className="text-3xl font-extrabold text-gray-900 tracking-tight">
            미국 상장기업 <span className="text-blue-600">10-K 재무제표</span> 한글 엑셀 변환
          </h1>
          <p className="text-gray-500 text-base">
            최근 결산월 기준 과거 10개년도 손익계산서·재무상태표·현금흐름표의 영문 계정명을 AI로 한글화하여 엑셀로 내려받습니다.
          </p>
        </div>

        {status === 'idle' && (
          <form onSubmit={handleStartAnalysis} className="space-y-4">
            <div className="relative" ref={dropdownRef}>
              <div className="flex shadow-sm rounded-xl overflow-hidden border-2 border-gray-200 focus-within:border-blue-600 transition">
                <input
                  type="text"
                  value={tickerInput}
                  onChange={handleInputChange}
                  onKeyDown={handleKeyDown}
                  onFocus={() => suggestions.length > 0 && setShowDropdown(true)}
                  placeholder="티커 또는 기업명 입력 (예: AAPL, SFM, MSFT)"
                  className="w-full px-5 py-4 text-lg text-gray-900 placeholder-gray-400 focus:outline-none"
                />
                <button
                  type="submit"
                  className="bg-blue-600 hover:bg-blue-700 text-white font-semibold px-8 py-4 text-base transition shrink-0"
                >
                  분석 및 엑셀 다운로드
                </button>
              </div>

              {showDropdown && (
                <ul className="absolute left-0 right-0 mt-1 bg-white border border-gray-200 rounded-xl shadow-lg z-40 max-h-64 overflow-y-auto text-left divide-y divide-gray-100">
                  {suggestions.map((item, idx) => (
                    <li
                      key={`${item.cik}-${item.ticker}`}
                      onClick={() => handleSelectSuggestion(item)}
                      className={`px-5 py-3 cursor-pointer flex justify-between items-center transition ${
                        idx === highlightIndex ? 'bg-blue-50 text-blue-700' : 'hover:bg-gray-50'
                      }`}
                    >
                      <div>
                        <span className="font-bold text-gray-900 mr-2">{item.ticker}</span>
                        <span className="text-sm text-gray-600">{item.title}</span>
                      </div>
                      <span className="text-xs text-gray-400">CIK: {item.cik}</span>
                    </li>
                  ))}
                </ul>
              )}
            </div>

            {selectedCompany && (
              <p className="text-sm text-blue-600 font-medium">
                선택된 기업: {selectedCompany.title} ({selectedCompany.ticker})
              </p>
            )}
          </form>
        )}

        {status === 'processing' && (
          <div className="bg-gray-50 border border-gray-200 rounded-2xl p-8 space-y-6 shadow-sm">
            <div className="space-y-2">
              <p className="text-lg font-bold text-gray-800">{progressText}</p>
              <p className="text-sm text-gray-500">
                10년 치 재무제표 수집 및 한글 계정명 변환을 진행하고 있습니다...
              </p>
            </div>

            <div className="w-full bg-gray-200 h-4 rounded-full overflow-hidden">
              <div
                className="bg-blue-600 h-full transition-all duration-500 ease-out"
                style={{ width: `${progressPercent}%` }}
              ></div>
            </div>
            <p className="text-sm font-semibold text-blue-600">{progressPercent}% 완료</p>

            <button
              type="button"
              onClick={handleCancel}
              className="px-6 py-2.5 rounded-lg border border-gray-300 bg-white hover:bg-gray-100 text-gray-700 text-sm font-medium transition"
            >
              작업 취소 (Cancel)
            </button>
          </div>
        )}

        {status === 'success' && (
          <div className="bg-blue-50/60 border border-blue-200 rounded-2xl p-8 space-y-5 shadow-sm">
            <div className="text-4xl">✅</div>
            <div className="space-y-1">
              <h2 className="text-xl font-bold text-gray-900">
                한글 재무제표 엑셀 다운로드가 완료되었습니다!
              </h2>
              <p className="text-sm text-blue-700 font-semibold">{savedFileName}</p>
              <p className="text-xs text-gray-500">
                브라우저의 [다운로드] 폴더에서 생성된 엑셀 파일을 확인하세요.
              </p>
            </div>
            <div className="pt-2">
              <button
                type="button"
                onClick={handleReset}
                className="bg-blue-600 hover:bg-blue-700 text-white font-semibold px-6 py-3 rounded-lg text-sm transition"
              >
                다른 기업 검색하기
              </button>
            </div>
          </div>
        )}
      </main>

      {/* 하단 푸터 영역 (도서 구매 버튼 포함) */}
      <footer className="max-w-4xl w-full mx-auto text-center space-y-3 border-t border-gray-100 pt-4">
        <div>
          <button
            type="button"
            onClick={() => openLink(bookUrl)}
            className="inline-flex items-center space-x-2 text-xs font-bold text-emerald-700 bg-emerald-50 hover:bg-emerald-100 px-4 py-2 rounded-lg border border-emerald-200 transition shadow-sm"
          >
            <span>📚 실전 재무제표 분석방법 도서 구매 ↗</span>
          </button>
        </div>
        <p className="text-xs text-gray-400">
          단위: 백만 달러(In Millions, 정수 반올림) · 데이터 출처: SEC EDGAR API
        </p>
      </footer>

      {/* 환경설정 모달 */}
      {isSettingsOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-xl space-y-5">
            <div className="flex justify-between items-center border-b pb-3">
              <h3 className="text-lg font-bold text-gray-900">환경설정 (API Key 관리)</h3>
              <button
                type="button"
                onClick={() => setIsSettingsOpen(false)}
                className="text-gray-400 hover:text-gray-600 text-xl leading-none"
              >
                ×
              </button>
            </div>

            <form onSubmit={handleSaveSettings} className="space-y-4">
              <div className="space-y-1.5 text-left">
                <div className="flex justify-between items-center">
                  <label className="block text-sm font-semibold text-gray-700">
                    Google Gemini API Key <span className="text-red-500">*</span>
                  </label>
                  <button
                    type="button"
                    onClick={() => openLink(apiKeyUrl)}
                    className="text-xs font-semibold text-blue-600 hover:underline bg-blue-50 px-2.5 py-1 rounded border border-blue-200 transition"
                  >
                    🔑 API키 받기 ↗
                  </button>
                </div>
                <input
                  type="password"
                  value={apiKey}
                  onChange={(e) => setApiKey(e.target.value)}
                  placeholder="AIzaSy..."
                  className="w-full px-3.5 py-2.5 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-blue-600"
                  required
                />
                <p className="text-xs text-gray-500">
                  입력한 키는 각 사용자의 브라우저(localStorage)에만 안전하게 저장됩니다.
                </p>
              </div>

              <div className="space-y-1.5 text-left">
                <label className="block text-sm font-semibold text-gray-700">
                  사용자 이메일 (SEC API 식별용)
                </label>
                <input
                  type="email"
                  value={userEmail}
                  onChange={(e) => setUserEmail(e.target.value)}
                  placeholder="example@email.com"
                  className="w-full px-3.5 py-2.5 border border-gray-300 rounded-lg text-sm focus:outline-none focus:border-blue-600"
                />
              </div>

              <div className="flex justify-end space-x-2 pt-2">
                <button
                  type="button"
                  onClick={() => setIsSettingsOpen(false)}
                  className="px-4 py-2 rounded-lg border border-gray-300 text-sm text-gray-700 hover:bg-gray-50"
                >
                  취소
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-lg bg-blue-600 hover:bg-blue-700 text-sm font-semibold text-white"
                >
                  저장하기
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* 알림 모달 */}
      {alertModal.isOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl max-w-sm w-full p-6 shadow-xl text-center space-y-4">
            <h3 className="text-lg font-bold text-gray-900">{alertModal.title}</h3>
            <p className="text-sm text-gray-600 whitespace-pre-line">{alertModal.message}</p>
            <button
              type="button"
              onClick={() => setAlertModal({ isOpen: false, title: '', message: '' })}
              className="w-full py-2.5 rounded-lg bg-blue-600 hover:bg-blue-700 text-white font-semibold text-sm transition"
            >
              확인
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
// app/api/analyze/route.js
import { NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export const maxDuration = 60;

// 복사 시 마크다운 URL 변환 깨짐 방지 헬퍼
function makeUrl(hostParts, pathPart) {
  return ['https:', '', hostParts.join('.'), pathPart].join('/');
}

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// 1. 재무상태표(BS)에서 제외할 순수 취득원가(Gross) 및 단독 감가상각누계액 태그 (수치 로직 100% 동일)
const BS_GROSS_AND_ACCUM_EXCLUDE = new Set([
  'PropertyPlantAndEquipmentGross',
  'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment',
  'FiniteLivedIntangibleAssetsGross',
  'FiniteLivedIntangibleAssetsAccumulatedAmortization',
  'IntangibleAssetsGrossExcludingGoodwill',
  'AllowanceForDoubtfulAccountsReceivableCurrent',
]);

// 2. XBRL 1:1 동의어 태그 그룹 (수치 로직 100% 동일)
const HISTORICAL_SYNONYM_GROUPS = [
  [
    'RevenueFromContractWithCustomerExcludingAssessedTax',
    'RevenueFromContractWithCustomerIncludingAssessedTax',
    'Revenues',
    'SalesRevenueNet',
    'SalesRevenueGoodsNet',
  ],
  ['CostOfGoodsAndServicesSold', 'CostOfRevenue', 'CostOfGoodsSold'],
  ['InterestAndDividendIncomeOperating', 'InterestIncomeOperating', 'InvestmentIncomeInterest'],
  [
    'InterestExpense',
    'InterestExpenseOperating',
    'InterestExpenseNonoperating',
    'InterestIncomeExpenseNonoperatingNet',
    'InterestIncomeExpenseNet',
  ],
  ['RestructuringCharges', 'RestructuringSettlementAndImpairmentProvisions'],
  [
    'ProvisionForCreditLosses',
    'ProvisionForLoanLeaseAndOtherLosses',
    'ProvisionForLoanAndLeaseLosses',
    'ProvisionForLoanLossesExpensed',
  ],
  [
    'AvailableForSaleSecuritiesDebtSecurities',
    'AvailableForSaleSecurities',
    'AvailableForSaleDebtSecuritiesAmortizedCostBasis',
  ],
  [
    'DebtSecuritiesAvailableForSaleRealizedGainLoss',
    'DebtSecuritiesRealizedGainLoss',
    'AvailableForSaleSecuritiesGrossRealizedGainLossNet',
    'GainLossOnSalesOfMortgageBackedSecuritiesMBS',
  ],
  ['HeldToMaturitySecurities', 'HeldToMaturitySecuritiesDebtMaturitiesNextRollingTwelveMonthsFairValue'],
  ['MarketableSecurities', 'DebtSecurities'],
  [
    'LoansAndLeasesReceivableNetOfDeferredIncome',
    'LoansAndLeasesReceivableGrossCarryingAmount',
    'FinancingReceivableRecordedInvestmentNotPastDue',
  ],
  [
    'FinancingReceivableAllowanceForCreditLosses',
    'LoansAndLeasesReceivableAllowance',
    'AllowanceForLoanAndLeaseLossesRealEstate',
  ],
  ['LoansAndLeasesReceivableNetReportedAmount', 'NotesReceivableNet'],
  ['TimeDepositsPlacedAndOtherShortTermInvestments', 'OtherShortTermInvestments'],
  [
    'PropertyPlantAndEquipmentNet',
    'PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization',
  ],
  [
    'IntangibleAssetsNetExcludingGoodwill',
    'FiniteLivedIntangibleAssetsNet',
  ],
  ['AccountsPayableCurrent', 'AccountsPayableTradeCurrent'],
  ['AccountsPayableAndAccruedLiabilitiesCurrent', 'AccountsPayableAndOtherAccruedLiabilitiesCurrent'],
  ['AccruedLiabilitiesCurrent', 'OtherAccruedLiabilitiesCurrent'],
  ['PaymentsToAcquireMarketableSecurities', 'PaymentsToAcquireAvailableForSaleSecuritiesEquity'],
  ['ProceedsFromSaleAndMaturityOfMarketableSecurities', 'ProceedsFromSaleOfAvailableForSaleSecuritiesEquity'],
  ['LongTermDebtCurrent', 'LongTermDebtAndCapitalLeaseObligationsCurrent'],
  ['LongTermDebtNoncurrent', 'LongTermDebtAndCapitalLeaseObligations', 'LongTermDebt'],
  ['Liabilities', 'LiabilitiesIncludingDeferredIncomeTaxes'],
  [
    'ProceedsFromPaymentsForOtherFinancingActivities',
    'PaymentsToMinorityShareholdersAndProceedsFromPaymentsForOtherFinancingActivities',
  ],
  ['OtherNoncashIncomeExpense', 'OtherNonCashIncomeExpenseIncludingAssetImpairmentCharges', 'OtherNoncashIncome'],
  [
    'NetCashProvidedByUsedInOperatingActivities',
    'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
  ],
  [
    'NetCashProvidedByUsedInInvestingActivities',
    'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
  ],
  [
    'NetCashProvidedByUsedInFinancingActivities',
    'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
  ],
  ['FinanceLeaseLiabilityCurrent', 'CapitalLeaseObligationsCurrent'],
  ['FinanceLeaseLiabilityNoncurrent', 'CapitalLeaseObligationsNoncurrent'],
  ['IncreaseDecreaseInTradingAssets', 'IncreaseDecreaseInTradingAccountAssets'],
  [
    'PaymentsToAcquireBusinessesNetOfCashAcquired',
    'PaymentsToAcquireBusinessesAndInterestInAffiliates',
    'PaymentsToAcquireBusinessesGross',
  ],
  [
    'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect',
    'CashAndCashEquivalentsPeriodIncreaseDecrease',
  ],
  ['ContractWithCustomerLiabilityCurrent', 'DeferredRevenueCurrent'],
  ['ContractWithCustomerLiabilityNoncurrent', 'DeferredRevenueNoncurrent'],
  ['OperatingExpenses', 'CostsAndExpenses'],
  ['SellingGeneralAndAdministrativeExpense', 'GeneralAndAdministrativeExpense'],
  ['LaborAndRelatedExpense', 'SalariesAndWages'],
  ['CommunicationsAndInformationTechnology', 'InformationTechnologyAndDataProcessing'],
  ['OtherAssets', 'OtherAssetsNoncurrent'],
  ['OtherLiabilities', 'OtherLiabilitiesNoncurrent'],
  [
    'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
    'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments',
  ],
  ['EarningsPerShareBasic', 'IncomeLossFromContinuingOperationsPerBasicShare'],
  ['EarningsPerShareDiluted', 'IncomeLossFromContinuingOperationsPerDilutedShare'],
];

// 다차원 부문(동일 dimMember) 전용 동의어 그룹 — 단일 부문 기업(5대 기준 종목)과 완전히 격리됨
const DIM_SYNONYM_GROUPS = [
  ['SalesAndServiceRevenue', 'RevenueFromContractWithCustomerExcludingAssessedTax', 'SalesRevenueNet'],
  ['OperatingLeaseLeaseIncome', 'OperatingLeasesIncomeStatementLeaseRevenue'],
  [
    'LiabilityForUnpaidClaimsAndClaimsAdjustmentExpenseIncurredClaims1',
    'IncurredClaimsPropertyCasualtyAndLiability',
  ],
  [
    'PolicyholderBenefitsAndClaimsIncurredNet',
    'PolicyholderBenefitsAndClaimsIncurredLifeAnnuityAndHealth',
    'PolicyholderBenefitsAndClaimsIncurredLifeAndAnnuity',
  ],
  ['RailroadTransportationRevenues', 'FreightRailTransportationRevenues'],
  ['UtilityAndEnergyOperatingRevenues', 'EnergyOperatingRevenues'],
  ['LiabilitiesExcludingIncomeTaxesPrincipallyDeferred', 'Liabilities'],
  ['OtherPayablesToBrokerDealersAndClearingOrganizations', 'PayableForPurchaseOfUSTreasuryBills'],
  ['EquitySecuritiesFvNi', 'AvailableForSaleSecuritiesEquitySecurities'],
  [
    'LiabilityForClaimsAndClaimsAdjustmentExpense',
    'LiabilityForClaimsAndClaimsAdjustmentExpensePropertyCasualtyLiability',
  ],
  [
    'LiabilityForUnpaidClaimsAndClaimAdjustmentExpenseRetroactiveReinsurance',
    'LiabilityForUnpaidClaimsAndClaimAdjustmentExpenseRetroactiveReinsuranceContracts',
  ],
  ['NotesReceivableNet', 'LoansAndLeasesReceivableNetReportedAmount'],
];

function canonicalizeDimensionMember(axis, rawLabel) {
  const cleaned = String(rawLabel || '')
    .replace(/\[member\]/gi, '')
    .replace(/\s+/g, ' ')
    .trim();
  if (!cleaned) return '';

  if (/StatementClassOfStockAxis/i.test(axis)) {
    if (/\bclass\s*a\b/i.test(cleaned)) return 'CLASS_A';
    if (/\bclass\s*b\b/i.test(cleaned)) return 'CLASS_B';
    return 'CLASS_' + cleaned.toUpperCase().replace(/[^A-Z0-9]+/g, '_');
  }

  if (/ProductOrServiceAxis/i.test(axis)) {
    if (/insurance\s+and\s+other/i.test(cleaned)) return 'SEG_INSURANCE_OTHER';
    if (/railroad.*utilities.*energy/i.test(cleaned)) return 'SEG_RAILROAD_UTILITIES_ENERGY';
    if (/^(?:freight\s+rail\s+transportation|railroad)$/i.test(cleaned)) return 'SEG_SUB_RAILROAD';
    if (/^utilities\s+and\s+energy$/i.test(cleaned)) return 'SEG_SUB_UTILITIES_ENERGY';
    if (/finance\s+and\s+financial\s+products/i.test(cleaned)) return 'SEG_FINANCE_PRODUCTS';
    return 'SEG_' + cleaned.toUpperCase().replace(/[^A-Z0-9]+/g, '_');
  }

  return 'DIM_' + cleaned.toUpperCase().replace(/[^A-Z0-9]+/g, '_');
}

// 3. 표준 US-GAAP 태그 사전 (수치 및 행 매칭 로직 100% 동일 유지)
const STANDARD_KOR_DICT = {
  // === 손익계산서 (IS) ===
  RevenueFromContractWithCustomerExcludingAssessedTax: { kor: '매출액', sign: 'raw', bold: true },
  RevenueFromContractWithCustomerIncludingAssessedTax: { kor: '매출액', sign: 'raw', bold: true },
  Revenues: { kor: '매출액', sign: 'raw', bold: true },
  SalesRevenueNet: { kor: '매출액', sign: 'raw', bold: true },
  RegulatedAndUnregulatedOperatingRevenue: { kor: '영업수익', sign: 'raw', bold: true },
  RegulatedOperatingRevenue: { kor: '영업수익', sign: 'raw', bold: true },
  SalesRevenueGoodsNet: { kor: '상품매출액', sign: 'raw' },
  SalesRevenueServicesNet: { kor: '용역매출액', sign: 'raw' },
  CostOfGoodsAndServicesSold: { kor: '매출원가', sign: 'pos' },
  CostOfRevenue: { kor: '매출원가', sign: 'pos' },
  CostOfGoodsSold: { kor: '매출원가', sign: 'pos' },
  GrossProfit: { kor: '매출총이익', sign: 'raw', bold: true },
  ResearchAndDevelopmentExpense: { kor: '연구개발비', sign: 'pos' },
  SellingGeneralAndAdministrativeExpense: { kor: '판매비와관리비', sign: 'pos' },
  SellingAndMarketingExpense: { kor: '판매및마케팅비', sign: 'pos' },
  SellingExpense: { kor: '판매비', sign: 'pos' },
  GeneralAndAdministrativeExpense: { kor: '일반관리비', sign: 'pos' },
  DepreciationAndAmortization: { kor: '감가상각비및무형자산상각비', sign: 'pos' },
  Depreciation: { kor: '감가상각비', sign: 'pos' },
  AmortizationOfIntangibleAssets: { kor: '무형자산상각비', sign: 'pos' },
  RestructuringCharges: { kor: '구조조정및기타비용(순액)', sign: 'raw' },
  RestructuringSettlementAndImpairmentProvisions: { kor: '구조조정및기타비용(순액)', sign: 'raw' },
  AssetImpairmentCharges: { kor: '자산손상차손', sign: 'raw' },
  OperatingExpenses: { kor: '영업비용합계', sign: 'pos', bold: true },
  CostsAndExpenses: { kor: '총원가및영업비용', sign: 'pos', bold: true },
  OtherOperatingIncomeExpenseNet: { kor: '기타영업손익(순)', sign: 'raw' },
  OperatingIncomeLoss: { kor: '영업이익', sign: 'raw', bold: true },

  // 금융업 수익·비용
  InterestAndFeeIncomeLoansAndLeases: { kor: '대출채권및리스이자수익', sign: 'raw' },
  InterestAndDividendIncomeSecurities: { kor: '유가증권이자및배당수익', sign: 'raw' },
  InterestIncomeOperating: { kor: '이자수익', sign: 'raw' },
  InterestAndDividendIncomeOperating: { kor: '이자수익', sign: 'raw' },
  InterestExpenseDeposits: { kor: '예수금이자비용', sign: 'raw' },
  InterestExpenseShortTermBorrowings: { kor: '단기차입금이자비용', sign: 'raw' },
  InterestExpenseLongTermDebt: { kor: '장기차입금이자비용', sign: 'raw' },
  InterestExpenseOperating: { kor: '이자비용', sign: 'raw' },
  InterestIncomeExpenseNet: { kor: '순이자손익', sign: 'raw', bold: true },
  NoninterestIncome: { kor: '비이자수익합계', sign: 'raw', bold: true },
  RevenuesNetOfInterestExpense: { kor: '순영업수익(이자비용차감후총수익)', sign: 'raw', bold: true },
  ProvisionForLoanLeaseAndOtherLosses: { kor: '신용손실충당금전입액(대손상각비)', sign: 'raw' },
  ProvisionForCreditLosses: { kor: '신용손실충당금전입액(대손상각비)', sign: 'raw' },
  ProvisionForLoanAndLeaseLosses: { kor: '신용손실충당금전입액(대손상각비)', sign: 'raw' },
  LaborAndRelatedExpense: { kor: '급여및복리후생비', sign: 'raw' },
  OccupancyNet: { kor: '임차료및설비비', sign: 'raw' },
  CommunicationsAndInformationTechnology: { kor: '정보처리및통신비', sign: 'raw' },
  InformationTechnologyAndDataProcessing: { kor: '정보처리및통신비', sign: 'raw' },
  ProfessionalFees: { kor: '지급수수료(전문가보수)', sign: 'raw' },
  MarketingAndAdvertisingExpense: { kor: '마케팅및광고선전비', sign: 'raw' },
  OtherGeneralExpense: { kor: '기타일반영업비용', sign: 'raw' },
  NoninterestExpense: { kor: '비이자비용합계', sign: 'raw', bold: true },

  InvestmentIncomeInterest: { kor: '이자수익', sign: 'pos' },
  InterestIncomeNonoperating: { kor: '이자수익', sign: 'pos' },
  InterestAndOtherIncome: { kor: '이자및기타수익', sign: 'pos' },
  InterestExpense: { kor: '이자비용', sign: 'raw' },
  InterestExpenseNonoperating: { kor: '이자비용', sign: 'raw' },
  InterestIncomeExpenseNonoperatingNet: { kor: '순이자비용(수익)', sign: 'raw' },
  GainsLossesOnExtinguishmentOfDebt: { kor: '채무상환손익(조기상환손실)', sign: 'raw' },
  OtherNonoperatingIncomeExpense: { kor: '기타영업외손익(순)', sign: 'raw' },
  NonoperatingIncomeExpense: { kor: '영업외손익합계', sign: 'raw' },
  IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest: {
    kor: '법인세비용차감전순이익',
    sign: 'raw',
    bold: true,
  },
  IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments: {
    kor: '법인세비용차감전순이익',
    sign: 'raw',
    bold: true,
  },
  IncomeTaxExpenseBenefit: { kor: '법인세비용', sign: 'pos' },
  IncomeLossFromEquityMethodInvestments: { kor: '지분법손익', sign: 'raw' },
  IncomeLossFromContinuingOperations: { kor: '계속영업순이익', sign: 'raw' },
  ProfitLoss: { kor: '연결당기순이익', sign: 'raw', bold: true },
  NetIncomeLossAttributableToNoncontrollingInterest: { kor: '비지배지분귀속순이익', sign: 'raw' },
  NetIncomeLoss: { kor: '당기순이익', sign: 'raw', bold: true },
  PreferredStockDividendsAndOtherAdjustments: { kor: '우선주배당금', sign: 'raw' },
  PreferredStockDividendsIncomeStatementImpact: { kor: '우선주배당금', sign: 'raw' },
  NetIncomeLossAvailableToCommonStockholdersBasic: { kor: '보통주귀속당기순이익', sign: 'raw', bold: true },
  EarningsPerShareBasic: { kor: '기본 주당순이익(달러)', sign: 'eps' },
  EarningsPerShareDiluted: { kor: '희석 주당순이익(달러)', sign: 'eps' },
  IncomeLossFromContinuingOperationsPerBasicShare: { kor: '기본 주당순이익(달러)', sign: 'eps' },
  IncomeLossFromContinuingOperationsPerDilutedShare: { kor: '희석 주당순이익(달러)', sign: 'eps' },
  WeightedAverageNumberOfSharesOutstandingBasic: { kor: '기본 가중평균발행주식수(백만주)', sign: 'shares' },
  WeightedAverageNumberOfDilutedSharesOutstanding: { kor: '희석 가중평균발행주식수(백만주)', sign: 'shares' },

  // === 재무상태표 (BS) ===
  CashAndDueFromBanks: { kor: '현금및타행예치금', sign: 'raw' },
  InterestBearingDepositsInBanks: { kor: '연방준비은행및타행이자부예치금', sign: 'raw' },
  CashAndCashEquivalentsAtCarryingValue: { kor: '현금및현금성자산', sign: 'raw' },
  RestrictedCashCurrent: { kor: '사용제한현금(유동)', sign: 'raw' },
  TimeDepositsPlacedAndOtherShortTermInvestments: { kor: '정기예치금및기타단기금융상품', sign: 'raw' },
  FederalFundsSoldAndSecuritiesPurchasedUnderAgreementsToResell: {
    kor: '연방기금매도및환매조건부매수채권',
    sign: 'raw',
  },
  TradingAssets: { kor: '단기매매금융자산', sign: 'raw' },
  DerivativeAssets: { kor: '파생상품자산', sign: 'raw' },
  ShortTermInvestments: { kor: '단기금융상품', sign: 'raw' },
  MarketableSecuritiesCurrent: { kor: '단기유가증권', sign: 'raw' },
  AvailableForSaleSecuritiesDebtSecurities: { kor: '공정가치측정채무증권(매도가능)', sign: 'raw' },
  AvailableForSaleSecurities: { kor: '공정가치측정채무증권(매도가능)', sign: 'raw' },
  AvailableForSaleSecuritiesDebtSecuritiesCurrent: { kor: '단기매도가능채무증권', sign: 'raw' },
  HeldToMaturitySecurities: { kor: '만기보유채무증권(상각후원가)', sign: 'raw' },
  MarketableSecurities: { kor: '채무증권(유가증권)합계', sign: 'raw', bold: true },
  LoansAndLeasesReceivableNetOfDeferredIncome: { kor: '대출채권및리스채권', sign: 'raw' },
  FinancingReceivableAllowanceForCreditLosses: { kor: '대출채권및리스대손충당금', sign: 'neg' },
  LoansAndLeasesReceivableAllowance: { kor: '대출채권및리스대손충당금', sign: 'neg' },
  LoansAndLeasesReceivableNetReportedAmount: { kor: '대출채권및리스채권(순액)', sign: 'raw', bold: true },
  LoansReceivableHeldForSaleAmount: { kor: '매각예정대출채권', sign: 'raw' },
  CashCashEquivalentsAndShortTermInvestments: { kor: '현금및단기금융자산', sign: 'raw' },
  AccountsReceivableNetCurrent: { kor: '매출채권(순액)', sign: 'raw' },
  ReceivablesNetCurrent: { kor: '매출채권(순액)', sign: 'raw' },
  AccruedInvestmentIncomeReceivable: { kor: '고객및기타수취채권', sign: 'raw' },
  NontradeReceivablesCurrent: { kor: '미수금(비매출채권)', sign: 'raw' },
  InventoryNet: { kor: '재고자산', sign: 'raw' },
  PrepaidExpenseCurrent: { kor: '선급비용', sign: 'raw' },
  PrepaidExpenseAndOtherAssetsCurrent: { kor: '선급비용및기타유동자산', sign: 'raw' },
  OtherAssetsCurrent: { kor: '기타유동자산', sign: 'raw' },
  AssetsCurrent: { kor: '[유동자산 합계]', sign: 'raw', bold: true },
  PropertyPlantAndEquipmentNet: { kor: '유형자산(순액)', sign: 'raw' },
  PropertyPlantAndEquipmentAndFinanceLeaseRightOfUseAssetAfterAccumulatedDepreciationAndAmortization: {
    kor: '유형자산(순액)',
    sign: 'raw',
  },
  OperatingLeaseRightOfUseAsset: { kor: '영업리스사용권자산', sign: 'raw' },
  FinanceLeaseRightOfUseAsset: { kor: '금융리스사용권자산', sign: 'raw' },
  Goodwill: { kor: '영업권', sign: 'raw' },
  IntangibleAssetsNetExcludingGoodwill: { kor: '무형자산(순액)', sign: 'raw' },
  FiniteLivedIntangibleAssetsNet: { kor: '무형자산(순액)', sign: 'raw' },
  IndefiniteLivedIntangibleAssetsExcludingGoodwill: { kor: '비한정내용연수무형자산', sign: 'raw' },
  MarketableSecuritiesNoncurrent: { kor: '장기유가증권', sign: 'raw' },
  LongTermInvestments: { kor: '장기투자자산', sign: 'raw' },
  EquityMethodInvestments: { kor: '지분법적용투자주식', sign: 'raw' },
  DeferredIncomeTaxAssetsNet: { kor: '이연법인세자산', sign: 'raw' },
  OtherAssetsNoncurrent: { kor: '기타비유동자산', sign: 'raw' },
  OtherAssets: { kor: '기타자산', sign: 'raw' },
  AssetsNoncurrent: { kor: '[비유동자산 합계]', sign: 'raw', bold: true },
  Assets: { kor: '[[자산총계]]', sign: 'raw', bold: true },

  // 부채 및 자본
  NoninterestBearingDepositLiabilities: { kor: '무이자부예수금', sign: 'raw' },
  InterestBearingDepositLiabilities: { kor: '이자부예수금', sign: 'raw' },
  NoninterestBearingDepositLiabilitiesDomestic: { kor: '미국내 무이자부예수금', sign: 'raw' },
  InterestBearingDepositLiabilitiesDomestic: { kor: '미국내 이자부예수금', sign: 'raw' },
  NoninterestBearingDepositLiabilitiesForeign: { kor: '해외 무이자부예수금', sign: 'raw' },
  InterestBearingDepositLiabilitiesForeign: { kor: '해외 이자부예수금', sign: 'raw' },
  DepositsDomestic: { kor: '미국내예수금', sign: 'raw' },
  DepositsForeign: { kor: '해외예수금', sign: 'raw' },
  Deposits: { kor: '[예수금 합계]', sign: 'raw', bold: true },
  FederalFundsPurchasedAndSecuritiesSoldUnderAgreementsToRepurchase: {
    kor: '연방기금매입및환매조건부매도채무',
    sign: 'raw',
  },
  TradingLiabilities: { kor: '단기매매금융부채', sign: 'raw' },
  DerivativeLiabilities: { kor: '파생상품부채', sign: 'raw' },
  AccountsPayableCurrent: { kor: '매입채무', sign: 'raw' },
  AccountsPayableTradeCurrent: { kor: '매입채무', sign: 'raw' },
  AccountsPayableAndAccruedLiabilitiesCurrent: { kor: '매입채무및미지급비용', sign: 'raw' },
  AccruedLiabilitiesCurrent: { kor: '미지급비용', sign: 'raw' },
  OtherAccruedLiabilitiesCurrent: { kor: '미지급비용', sign: 'raw' },
  AccruedLiabilitiesAndOtherLiabilities: { kor: '미지급비용및기타부채', sign: 'raw' },
  EmployeeRelatedLiabilitiesCurrent: { kor: '미지급급여및복리후생비', sign: 'raw' },
  AccruedIncomeTaxesCurrent: { kor: '미지급법인세(유동)', sign: 'raw' },
  ContractWithCustomerLiabilityCurrent: { kor: '계약부채(유동)', sign: 'raw' },
  DeferredRevenueCurrent: { kor: '이연수익(유동)', sign: 'raw' },
  CommercialPaper: { kor: '기업어음(CP)', sign: 'raw' },
  ShortTermBorrowings: { kor: '단기차입금', sign: 'raw' },
  LongTermDebtCurrent: { kor: '유동성장기부채및금융리스부채', sign: 'raw' },
  LongTermDebtAndCapitalLeaseObligationsCurrent: { kor: '유동성장기부채및금융리스부채', sign: 'raw' },
  OperatingLeaseLiabilityCurrent: { kor: '유동성영업리스부채', sign: 'raw' },
  FinanceLeaseLiabilityCurrent: { kor: '유동성금융리스부채', sign: 'raw' },
  CapitalLeaseObligationsCurrent: { kor: '유동성금융리스부채', sign: 'raw' },
  OtherLiabilitiesCurrent: { kor: '기타유동부채', sign: 'raw' },
  LiabilitiesCurrent: { kor: '[유동부채 합계]', sign: 'raw', bold: true },
  LongTermDebtNoncurrent: { kor: '장기차입금및금융리스부채', sign: 'raw' },
  LongTermDebt: { kor: '장기차입금및금융리스부채', sign: 'raw' },
  LongTermDebtAndCapitalLeaseObligations: { kor: '장기차입금및금융리스부채', sign: 'raw' },
  CapitalizationLongtermDebtAndEquity: { kor: '총자본화금액', sign: 'raw' },
  AssetRetirementObligationsNoncurrent: { kor: '자산복구충당부채', sign: 'raw' },
  AssetRetirementObligation: { kor: '자산복구충당부채', sign: 'raw' },
  OperatingLeaseLiabilityNoncurrent: { kor: '장기영업리스부채', sign: 'raw' },
  FinanceLeaseLiabilityNoncurrent: { kor: '비유동금융리스부채', sign: 'raw' },
  CapitalLeaseObligationsNoncurrent: { kor: '비유동금융리스부채', sign: 'raw' },
  ContractWithCustomerLiabilityNoncurrent: { kor: '계약부채(비유동)', sign: 'raw' },
  DeferredRevenueNoncurrent: { kor: '이연수익(비유동)', sign: 'raw' },
  DeferredIncomeTaxLiabilitiesNet: { kor: '이연법인세부채', sign: 'raw' },
  DeferredIncomeTaxesAndOtherLiabilitiesNoncurrent: { kor: '이연법인세및기타비유동부채', sign: 'raw' },
  AccruedIncomeTaxesNoncurrent: { kor: '장기미지급법인세', sign: 'raw' },
  OtherLiabilitiesNoncurrent: { kor: '기타비유동부채', sign: 'raw' },
  OtherLiabilities: { kor: '기타부채', sign: 'raw' },
  LiabilitiesNoncurrent: { kor: '[비유동부채 합계]', sign: 'raw', bold: true },
  Liabilities: { kor: '[[부채총계]]', sign: 'raw', bold: true },
  LiabilitiesIncludingDeferredIncomeTaxes: { kor: '[[부채총계]]', sign: 'raw', bold: true },
  LiabilitiesSubjectToCompromise: { kor: '화해대상부채(파산보호)', sign: 'raw' },
  CommitmentsAndContingencies: { kor: '우발부채및약정사항', sign: 'raw' },
  PreferredStockValue: { kor: '우선주자본금', sign: 'raw' },
  CommonStockValue: { kor: '보통주자본금', sign: 'raw' },
  CommonStocksIncludingAdditionalPaidInCapital: { kor: '보통주자본금및자본잉여금', sign: 'raw' },
  AdditionalPaidInCapital: { kor: '자본잉여금', sign: 'raw' },
  AdditionalPaidInCapitalCommonStock: { kor: '자본잉여금(보통주)', sign: 'raw' },
  RetainedEarningsAccumulatedDeficit: { kor: '이익잉여금(결손금)', sign: 'raw' },
  TreasuryStockValue: { kor: '자기주식', sign: 'neg' },
  TreasuryStockCommonValue: { kor: '자기주식(보통주)', sign: 'neg' },
  AccumulatedOtherComprehensiveIncomeLossNetOfTax: { kor: '기타포괄손익누계액', sign: 'raw' },
  StockholdersEquity: { kor: '[[자본총계]]', sign: 'raw', bold: true },
  MinorityInterest: { kor: '비지배지분', sign: 'raw' },
  RedeemableNoncontrollingInterestEquityCarryingAmount: { kor: '상환가능비지배지분', sign: 'raw' },
  RedeemableNoncontrollingInterestEquityCommonCarryingAmount: { kor: '상환가능비지배지분', sign: 'raw' },
  RedeemableNoncontrollingInterestEquityPreferredCarryingAmount: { kor: '상환가능비지배지분', sign: 'raw' },
  TemporaryEquityCarryingAmountAttributableToParent: { kor: '상환가능지분(메자닌자본)', sign: 'raw' },
  TemporaryEquityCarryingAmountIncludingPortionAttributableToNoncontrollingInterests: {
    kor: '상환가능비지배지분',
    sign: 'raw',
  },
  StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest: {
    kor: '[[자본총계]]',
    sign: 'raw',
    bold: true,
  },
  PartnersCapital: { kor: '[[자본총계]]', sign: 'raw', bold: true },
  PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest: {
    kor: '[[자본총계]]',
    sign: 'raw',
    bold: true,
  },
  LiabilitiesAndStockholdersEquity: { kor: '[[부채와자본총계]]', sign: 'raw', bold: true },

  // === 현금흐름표 (CF) ===
  DepreciationDepletionAndAmortization: { kor: '감가상각비및무형자산상각비', sign: 'raw' },
  OperatingLeaseRightOfUseAssetAmortizationExpense: { kor: '운용리스사용권자산 상각비', sign: 'raw' },
  AmortizationOfFinancingCostsAndDiscounts: { kor: '금융비용및사채발행비 상각액', sign: 'raw' },
  AccretionAmortizationOfDiscountsAndPremiumsInvestments: {
    kor: '채무증권할증·할인차금상각(순)',
    sign: 'invert',
  },
  ShareBasedCompensation: { kor: '주식기반보상비용', sign: 'raw' },
  DeferredIncomeTaxExpenseBenefit: { kor: '이연법인세비용(혜택)', sign: 'raw' },
  DebtSecuritiesAvailableForSaleRealizedGainLoss: { kor: '채무증권처분손익(순)', sign: 'invert' },
  DebtSecuritiesRealizedGainLoss: { kor: '채무증권처분손익(순)', sign: 'invert' },
  AvailableForSaleSecuritiesGrossRealizedGainLossNet: { kor: '채무증권처분손익(순)', sign: 'invert' },
  GainLossOnDispositionOfAssets: { kor: '자산처분손익', sign: 'invert' },
  GainLossOnSaleOfPropertyPlantEquipment: { kor: '유형자산처분손익', sign: 'invert' },
  OtherNoncashIncomeExpense: { kor: '기타비현금성손익', sign: 'invert' },
  OtherOperatingActivitiesCashFlowStatement: { kor: '기타영업활동현금흐름(순)', sign: 'raw' },
  PaymentsForOriginationAndPurchasesOfLoansHeldForSale: {
    kor: '매각예정대출채권의 실행및매입',
    sign: 'neg',
  },
  ProceedsFromSaleAndCollectionOfLoansHeldForSale: {
    kor: '매각예정대출채권의 매각및회수',
    sign: 'pos',
  },
  IncreaseDecreaseInLoansHeldForSale: { kor: '매각예정대출채권의 순증감', sign: 'invert' },
  IncreaseDecreaseInTradingAssets: { kor: '단기매매및파생상품자산·부채의 순증감', sign: 'invert' },
  IncreaseDecreaseInTradingAccountAssets: { kor: '단기매매및파생상품자산·부채의 순증감', sign: 'invert' },
  IncreaseDecreaseInAccountsReceivable: { kor: '매출채권의 증감', sign: 'invert' },
  IncreaseDecreaseInOtherReceivables: { kor: '기타수취채권의 증감', sign: 'invert' },
  IncreaseDecreaseInInventories: { kor: '재고자산의 증감', sign: 'invert' },
  IncreaseDecreaseInPrepaidDeferredExpenseAndOtherAssets: { kor: '선급비용및기타유동자산의 증감', sign: 'invert' },
  IncreaseDecreaseInOtherOperatingAssets: { kor: '기타자산의 증감', sign: 'invert' },
  IncreaseDecreaseInAccountsPayable: { kor: '매입채무의 증감', sign: 'raw' },
  IncreaseDecreaseInAccountsPayableAndAccruedLiabilities: { kor: '매입채무및미지급비용의 증감', sign: 'raw' },
  IncreaseDecreaseInAccruedLiabilities: { kor: '미지급비용의 증감', sign: 'raw' },
  IncreaseDecreaseInEmployeeRelatedLiabilities: { kor: '미지급급여및복리후생비의 증감', sign: 'raw' },
  IncreaseDecreaseInAccruedIncomeTaxesPayable: { kor: '미지급법인세의 증감', sign: 'raw' },
  IncreaseDecreaseInAccruedLiabilitiesAndOtherOperatingLiabilities: {
    kor: '미지급비용및기타부채의 증감',
    sign: 'raw',
  },
  IncreaseDecreaseInContractWithCustomerLiability: { kor: '계약부채의 증감', sign: 'raw' },
  IncreaseDecreaseInDeferredRevenue: { kor: '이연수익의 증감', sign: 'raw' },
  IncreaseDecreaseInOperatingLeaseLiability: { kor: '영업리스부채의 증감', sign: 'raw' },
  IncreaseDecreaseInOtherOperatingLiabilities: { kor: '기타비유동부채의 증감', sign: 'raw' },
  IncreaseDecreaseInOtherOperatingCapitalNet: { kor: '기타영업자산·부채의 순증감', sign: 'invert' },
  IncreaseDecreaseInOperatingCapital: { kor: '운전자본의 순증감', sign: 'invert' },
  NetCashProvidedByUsedInOperatingActivities: { kor: '[영업활동현금흐름]', sign: 'raw', bold: true },
  NetCashProvidedByUsedInOperatingActivitiesContinuingOperations: {
    kor: '[영업활동현금흐름]',
    sign: 'raw',
    bold: true,
  },
  PaymentsForProceedsFromShortTermInvestments: {
    kor: '정기예치금및기타단기금융상품의 순증감',
    sign: 'invert',
  },
  PaymentsForProceedsFromFederalFundsSoldAndSecuritiesPurchasedUnderAgreementsToResell: {
    kor: '연방기금매도및환매조건부매수채권의 순증감',
    sign: 'invert',
  },
  PaymentsToAcquirePropertyPlantAndEquipment: { kor: '유형자산의 취득(CAPEX)', sign: 'neg' },
  PaymentsToAcquireProductiveAssets: { kor: '생산설비자산의 취득', sign: 'neg' },
  PaymentsToAcquireIntangibleAssets: { kor: '무형자산의 취득', sign: 'neg' },
  ProceedsFromSaleOfPropertyPlantAndEquipment: { kor: '유형자산의 처분', sign: 'pos' },
  PaymentsToAcquireBusinessesNetOfCashAcquired: { kor: '기업인수로 인한 현금유출', sign: 'neg' },
  PaymentsToAcquireBusinessesAndInterestInAffiliates: { kor: '기업인수로 인한 현금유출', sign: 'neg' },
  PaymentsToAcquireBusinessesGross: { kor: '기업인수로 인한 현금유출', sign: 'neg' },
  ProceedsFromDivestitureOfBusinesses: { kor: '사업매각으로 인한 현금유입', sign: 'pos' },
  PaymentsToAcquireAvailableForSaleSecuritiesDebt: { kor: '공정가치측정 채무증권의 취득', sign: 'neg' },
  ProceedsFromMaturitiesPrepaymentsAndCallsOfAvailableForSaleSecurities: {
    kor: '공정가치측정 채무증권의 만기도래 및 조기상환',
    sign: 'pos',
  },
  ProceedsFromSaleOfAvailableForSaleSecuritiesDebt: { kor: '공정가치측정 채무증권의 처분', sign: 'pos' },
  PaymentsToAcquireHeldToMaturitySecurities: { kor: '만기보유채무증권의 취득', sign: 'neg' },
  ProceedsFromMaturitiesPrepaymentsAndCallsOfHeldToMaturitySecurities: {
    kor: '만기보유채무증권의 원금회수및만기상환',
    sign: 'pos',
  },
  ProceedsFromSaleAndCollectionOfLoansReceivable: {
    kor: '투자목적분류대출채권의 매각및원금회수',
    sign: 'pos',
  },
  PaymentsToAcquireLoansReceivable: { kor: '대출채권및리스채권의 매입(취득)', sign: 'neg' },
  PaymentsForProceedsFromLoansAndLeases: { kor: '대출채권및리스채권의 기타순증감', sign: 'invert' },
  PaymentsToAcquireMarketableSecurities: { kor: '유가증권의 취득', sign: 'neg' },
  ProceedsFromSaleAndMaturityOfMarketableSecurities: { kor: '유가증권의 처분및만기회수', sign: 'pos' },
  PaymentsToAcquireShortTermInvestments: { kor: '단기금융상품의 취득', sign: 'neg' },
  ProceedsFromSaleAndMaturityOfShortTermInvestments: { kor: '단기금융상품의 처분및만기회수', sign: 'pos' },
  PaymentsToAcquireInvestments: { kor: '투자자산의 취득', sign: 'neg' },
  ProceedsFromSaleMaturityAndCollectionsOfInvestments: { kor: '투자자산의 처분및회수', sign: 'pos' },
  PaymentsForProceedsFromOtherInvestingActivities: { kor: '기타투자활동현금흐름(순)', sign: 'invert' },
  NetCashProvidedByUsedInInvestingActivities: { kor: '[투자활동현금흐름]', sign: 'raw', bold: true },
  NetCashProvidedByUsedInInvestingActivitiesContinuingOperations: {
    kor: '[투자활동현금흐름]',
    sign: 'raw',
    bold: true,
  },
  IncreaseDecreaseInDeposits: { kor: '예수금의 순증감', sign: 'raw' },
  IncreaseDecreaseInFederalFundsPurchasedAndSecuritiesSoldUnderAgreementsToRepurchase: {
    kor: '연방기금매입및환매조건부매도채무의 순증감',
    sign: 'raw',
  },
  ProceedsFromLinesOfCredit: { kor: '회전신용한도 차입금의 유입', sign: 'pos' },
  RepaymentsOfLinesOfCredit: { kor: '회전신용한도 차입금의 상환', sign: 'neg' },
  RepaymentsOfLongTermCapitalLeaseObligations: { kor: '자본리스및금융리스부채 상환액', sign: 'neg' },
  FinanceLeasePrincipalPayments: { kor: '자본리스및금융리스부채 상환액', sign: 'neg' },
  PaymentsOfFinancingCosts: { kor: '이연금융원가의 지급', sign: 'neg' },
  ProceedsFromIssuanceOfLongTermDebt: { kor: '장기차입금의 차입', sign: 'pos' },
  RepaymentsOfLongTermDebt: { kor: '장기차입금의 상환', sign: 'neg' },
  ProceedsFromRepaymentsOfShortTermDebt: { kor: '단기차입금의 순증감', sign: 'raw' },
  ProceedsFromRepaymentsOfCommercialPaper: { kor: '기업어음(CP)의 순증감', sign: 'raw' },
  ProceedsFromShortTermDebtMaturingInMoreThanThreeMonths: { kor: '단기차입금(3개월초과)의 차입', sign: 'pos' },
  RepaymentsOfShortTermDebtMaturingInMoreThanThreeMonths: { kor: '단기차입금(3개월초과)의 상환', sign: 'neg' },
  ProceedsFromIssuanceOfPreferredStockAndPreferenceStock: { kor: '우선주 발행에 따른 현금유입', sign: 'pos' },
  PaymentsForRepurchaseOfPreferredStockAndPreferenceStock: { kor: '우선주 상환에 따른 현금유출', sign: 'neg' },
  ProceedsFromIssuanceOfCommonStock: { kor: '보통주 발행에 따른 현금유입', sign: 'pos' },
  ProceedsFromStockOptionsExercised: { kor: '주식매수선택권 행사에 따른 유입', sign: 'pos' },
  PaymentsForRepurchaseOfCommonStock: { kor: '자기주식의 취득(자사주매입)', sign: 'neg' },
  PaymentsRelatedToTaxWithholdingForShareBasedCompensation: {
    kor: '주식보상 관련 원천징수세액 납부',
    sign: 'neg',
  },
  PaymentsOfDividends: { kor: '현금배당금의 지급', sign: 'neg' },
  PaymentsOfDividendsCommonStock: { kor: '보통주 배당금의 지급', sign: 'neg' },
  PaymentsOfDividendsPreferredStockAndPreferenceStock: { kor: '우선주 배당금의 지급', sign: 'neg' },
  PaymentsOfDividendsMinorityInterest: { kor: '비지배지분 배당금의 지급', sign: 'neg' },
  PaymentsToMinorityShareholders: { kor: '연결종속기업 비지배지분 취득', sign: 'neg' },
  ProceedsFromPaymentsForOtherFinancingActivities: { kor: '기타재무활동현금흐름(순)', sign: 'raw' },
  NetCashProvidedByUsedInFinancingActivities: { kor: '[재무활동현금흐름]', sign: 'raw', bold: true },
  NetCashProvidedByUsedInFinancingActivitiesContinuingOperations: {
    kor: '[재무활동현금흐름]',
    sign: 'raw',
    bold: true,
  },
  EffectOfExchangeRateOnCashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents: {
    kor: '현금에 대한 환율변동효과',
    sign: 'raw',
  },
  CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect: {
    kor: '현금및현금성자산의 순증감',
    sign: 'raw',
    bold: true,
  },
  CashAndCashEquivalentsPeriodIncreaseDecrease: {
    kor: '현금및현금성자산의 순증감',
    sign: 'raw',
    bold: true,
  },
};

// 4. R*.htm 하단 내장 XBRL 정의 테이블(authRefData) 파싱 (lookahead 분할로 id 속성 보존 및 장문 Reference 뒤 Details까지 전수 탐색)
function parseEmbeddedXbrlAuthRefs(html) {
  const authMap = {};
  const blocks = html.split(/(?=<table[^>]*class=["']authRefData["'][^>]*>)/i);
  for (let i = 1; i < blocks.length; i++) {
    const chunk = blocks[i];
    const openTagEnd = chunk.indexOf('>');
    const openTag = openTagEnd !== -1 ? chunk.slice(0, openTagEnd + 1) : chunk.slice(0, 300);
    const idMatch = openTag.match(/id=["'](defref_[^"']+)["']/i);
    const nameMatch = chunk.match(
      /<tr>\s*<td[^>]*>\s*(?:<[^>]+>\s*)*Name:\s*(?:<\/[^>]+>\s*)*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i
    );
    const typeMatch = chunk.match(
      /<tr>\s*<td[^>]*>\s*(?:<[^>]+>\s*)*(?:Data\s+)?Type:\s*(?:<\/[^>]+>\s*)*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i
    );
    const balMatch = chunk.match(
      /<tr>\s*<td[^>]*>\s*(?:<[^>]+>\s*)*Balance(?:\s+Type)?:\s*(?:<\/[^>]+>\s*)*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i
    );
    const perMatch = chunk.match(
      /<tr>\s*<td[^>]*>\s*(?:<[^>]+>\s*)*Period(?:\s+Type)?:\s*(?:<\/[^>]+>\s*)*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i
    );

    const key = idMatch ? idMatch[1].trim() : nameMatch ? `defref_${nameMatch[1].trim().replace(':', '_')}` : null;
    if (key) {
      authMap[key] = {
        dataType: typeMatch ? typeMatch[1].trim() : '',
        balance: balMatch ? balMatch[1].trim().toLowerCase() : '',
        periodType: perMatch ? perMatch[1].trim().toLowerCase() : '',
      };
    }
  }
  return authMap;
}

// 5. 라벨 내 가변 주석 문구 정제 (수치 로직 100% 동일)
function stripDynamicParentheticals(label) {
  return String(label || '')
    .replace(/\[.*?\]/g, '')
    .replace(
      /,?\s*net of (?:accumulated depreciation|accumulated amortization|allowance|unearned|deferred|current portion)[\s\S]*$/i,
      ', net'
    )
    .replace(
      /,?\s*less (?:accumulated depreciation|accumulated amortization|allowance|current portion)[\s\S]*$/i,
      ', net'
    )
    .replace(/\((?:exclusive of|excluding)[^()]*?\)/gi, '')
    .replace(
      /\((?:[^()]*?(?:\$|shares|par value|fair value|amortized cost|includes|note \d|vie|accumulated|allowance|\d{3,})[^()]*?)\)/gi,
      ''
    )
    .replace(/[,;]\s*(?:\$?\d+(?:\.\d+)?\s+)?par value[\s\S]*$/i, '')
    .replace(/[,;]\s*authorized[\s\S]*$/i, '')
    .replace(/\s+/g, ' ')
    .trim();
}

// 6. 행 속성('EPS' | 'SHARES' | 'USD') 판별 (수치 로직 100% 동일)
function detectRowType(tag, cleanLabel, currentSection, stmtType, xbrlDataType = '') {
  if (/perShareItemType/i.test(xbrlDataType)) return 'EPS';
  if (/sharesItemType/i.test(xbrlDataType)) return 'SHARES';

  const t = String(tag || '');
  const lbl = String(cleanLabel || '').trim();
  const sec = String(currentSection || '').trim();

  if (
    stmtType === 'IS' &&
    !/\bnumerator\b/i.test(lbl) &&
    /\bper\s+(?:common\s+|diluted\s+|basic\s+)?share\b/i.test(
      lbl.replace(/\([^()]*numerator[^()]*\)/gi, '').trim()
    )
  ) {
    return 'EPS';
  }

  if (/monetaryItemType/i.test(xbrlDataType)) return 'USD';

  if (
    /^NetIncomeLossAvailableToCommonStockholders(?:Basic|Diluted)?$/i.test(t) ||
    /^IncomeLossFromContinuingOperationsAttributableTo/i.test(t)
  ) {
    return 'USD';
  }

  if (
    /WeightedAverageNumberOf|SharesOutstanding|SharesIssued|SharesAuthorized|TreasuryStock(?:Common|Preferred)?Shares|AverageCommonShares|CommonStockShares/i.test(
      t
    ) &&
    !/PerShare|PerBasic|PerDiluted/i.test(t)
  ) {
    return 'SHARES';
  }
  if (
    /EarningsPerShare|PerShare|PerBasic|PerDiluted|IncomeLossFromContinuingOperationsPer/i.test(t)
  ) {
    return 'EPS';
  }

  if (stmtType === 'IS') {
    if (
      !/\bnumerator\b/i.test(lbl) &&
      /weighted[- ]average.*shares|shares\s+used\s+in|average\s+common\s+shares|number\s+of\s+shares/i.test(
        lbl
      )
    ) {
      return 'SHARES';
    }
    const lblNoNumerator = lbl.replace(/\([^()]*numerator[^()]*\)/gi, '').trim();
    if (
      !/\bnumerator\b/i.test(lbl) &&
      /\bper\s+(?:common\s+|diluted\s+|basic\s+)?share\b/i.test(lblNoNumerator)
    ) {
      return 'EPS';
    }
    if (/^(?:basic|diluted|basic and diluted):?$/i.test(lbl)) {
      if (
        /weighted|shares\s+used|number\s+of\s+shares|average.*shares/i.test(sec) &&
        !/per\s+share|earnings\s+per|income\s+per|loss\s+per/i.test(sec)
      ) {
        return 'SHARES';
      }
      if (/per\s+(?:common\s+)?share|earnings\s+per|income\s+per|loss\s+per/i.test(sec)) {
        return 'EPS';
      }
    }
  }

  return 'USD';
}

// 7. 소계/총계 행 판별 함수 (수치 검산 경계 보호를 위해 MSA Safety 버전과 100% 동일하게 고정!)
function isTrueSubtotalRow(cleanLabel, tag, stmtType, rowType) {
  if (rowType !== 'USD') return false;
  const lbl = String(cleanLabel || '').trim();
  const t = String(tag || '');

  if (stmtType === 'BS') {
    if (/^total\s+|,\s*total$/i.test(lbl)) return true;
    return [
      'AssetsCurrent',
      'AssetsNoncurrent',
      'Assets',
      'MarketableSecurities',
      'LoansAndLeasesReceivableNetReportedAmount',
      'LiabilitiesCurrent',
      'LiabilitiesNoncurrent',
      'Deposits',
      'Liabilities',
      'LiabilitiesIncludingDeferredIncomeTaxes',
      'LiabilitiesExcludingIncomeTaxesPrincipallyDeferred',
      'StockholdersEquity',
      'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
      'PartnersCapital',
      'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest',
      'LiabilitiesAndStockholdersEquity',
    ].includes(t);
  }

  if (stmtType === 'CF') {
    const CF_TOTAL_TAGS = [
      'NetCashProvidedByUsedInOperatingActivities',
      'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
      'NetCashProvidedByUsedInInvestingActivities',
      'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
      'NetCashProvidedByUsedInFinancingActivities',
      'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
      'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect',
      'CashAndCashEquivalentsPeriodIncreaseDecrease',
    ];
    if (CF_TOTAL_TAGS.includes(t)) return true;
    if (
      /DiscontinuedOperations/i.test(t) ||
      (/\bdiscontinued\b/i.test(lbl) && !/\bincluding\s+discontinued\b/i.test(lbl)) ||
      /^(?:PaymentsToAcquire|PaymentsFor|ProceedsFrom|RepaymentsOf)/i.test(t)
    ) {
      return false;
    }
    if (
      /^IncreaseDecrease.*(?:Collateral|Escrow|Margin|Deposit|Payable|Receivable)/i.test(t)
    ) {
      return false;
    }
    if (
      /^net cash (?:provided by|used in)|^cash (?:flows? from|generated by|used in)|^(?:net )?increase \(decrease\) in cash(?![\s-]*(?:collateral|held\s+as|deposited|escrow|margin|advance|balance\s+with))/i.test(
        lbl
      )
    ) {
      return true;
    }
    return false;
  }

  if (stmtType === 'IS') {
    if (
      /^total\s+|^gross profit$|^gross margin$|^operating income|^operating profit|^net interest income$|^income (?:\(loss\) )?before|^earnings before|^net income|^net earnings/i.test(
        lbl
      )
    ) {
      return true;
    }
    return [
      'GrossProfit',
      'OperatingExpenses',
      'CostsAndExpenses',
      'OperatingIncomeLoss',
      'InterestAndDividendIncomeOperating',
      'InterestExpenseOperating',
      'InterestIncomeExpenseNet',
      'NoninterestIncome',
      'RevenuesNetOfInterestExpense',
      'NoninterestExpense',
      'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
      'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments',
      'IncomeLossFromContinuingOperations',
      'ProfitLoss',
      'NetIncomeLoss',
      'NetIncomeLossAvailableToCommonStockholdersBasic',
    ].includes(t);
  }

  return false;
}

// 7b. 현금흐름표 영업활동 중간 소계('조정항목 합계' 등) 판별 (총계 격상 없이 검산 이중합산 방지)
function isIntermediateCfAdjustmentSubtotal(tag = '', label = '') {
  const t = String(tag || '');
  const lbl = String(label || '').trim();
  return (
    t === 'AdjustmentsToReconcileNetIncomeLossToCashProvidedByUsedInOperatingActivities' ||
    /^total\s+adjustments(?:\s+to\s+reconcile.*)?$/i.test(lbl)
  );
}
function isExcludedGrossOrAccumRow(stmtType, tag, rawLabel) {
  if (stmtType !== 'BS') return false;
  if (/Net|RightOfUse|Goodwill/i.test(tag || '')) return false;
  if (/\bnet\b/i.test(rawLabel || '')) return false;

  if (tag && BS_GROSS_AND_ACCUM_EXCLUDE.has(tag)) return true;
  if (/^(?:less:?\s*)?accumulated (?:depreciation|amortization)/i.test(rawLabel || '')) return true;
  if (/property,?\s*plant and equipment,?\s*gross|property and equipment,?\s*gross/i.test(rawLabel || '')) {
    return true;
  }
  return false;
}

// 9. 문맥 라벨 생성 함수 (수치 및 행 매칭 로직 보호를 위해 MSA Safety 버전과 100% 동일하게 고정!)
function buildContextualLabel(
  cleanLabel,
  sectionHeader,
  rowType,
  stmtType = '',
  cashFlowCashIndex = 0,
  isMultiSegment = false
) {
  const lbl = stripDynamicParentheticals(cleanLabel).replace(/:$/, '').trim();
  const sec = stripDynamicParentheticals(sectionHeader || '').replace(/:$/, '').trim();

  if (rowType === 'EPS' && /^basic$/i.test(lbl)) return 'Earnings per share - Basic';
  if (rowType === 'EPS' && /^diluted$/i.test(lbl)) return 'Earnings per share - Diluted';
  if (rowType === 'SHARES' && /^basic$/i.test(lbl)) return 'Weighted average shares - Basic';
  if (rowType === 'SHARES' && /^diluted$/i.test(lbl)) return 'Weighted average shares - Diluted';

  if (stmtType === 'CF') {
    const isRestrictedFootnoteRow =
      isMultiSegment && /restricted\s+cash,\s*included/i.test(lbl);
    if (
      !isRestrictedFootnoteRow &&
      (/beginning of (?:period|year)|at beginning of/i.test(lbl) || cashFlowCashIndex === 1)
    ) {
      return 'Cash and cash equivalents, beginning of period';
    }
    if (
      !isRestrictedFootnoteRow &&
      (/end of (?:period|year)|at end of/i.test(lbl) || cashFlowCashIndex === 2)
    ) {
      return 'Cash and cash equivalents, end of period';
    }
  }

  if (!sec) return lbl;

  const lowerLbl = lbl.toLowerCase();
  const lowerSec = sec.toLowerCase();

  if (/^carried at fair value$/i.test(lowerLbl) && /debt securities/i.test(lowerSec)) {
    return 'Available-for-sale debt securities carried at fair value';
  }
  if (/^held-?to-?maturity/i.test(lowerLbl) && /debt securities/i.test(lowerSec)) {
    return 'Held-to-maturity debt securities, at cost';
  }
  if (/^(non-?interest-?bearing|interest-?bearing)/i.test(lowerLbl)) {
    if (/non-?u\.?s\.?/i.test(lowerSec)) return `Non-U.S. ${lbl} deposits`;
    if (/u\.?s\.?/i.test(lowerSec)) return `U.S. ${lbl} deposits`;
  }
  if (
    /^(proceeds from sales|proceeds from paydowns and maturities|proceeds from maturities|purchases|other changes)/i.test(
      lowerLbl
    )
  ) {
    if (/available-?for-?sale|carried at fair value/i.test(lowerSec)) {
      return `${lbl} of available-for-sale debt securities`;
    }
    if (/held-?to-?maturity/i.test(lowerSec)) {
      return `${lbl} of held-to-maturity debt securities`;
    }
    if (/loans and leases/i.test(lowerSec) && !/loans/i.test(lowerLbl)) {
      return `${lbl} of loans and leases`;
    }
  }
  return lbl;
}

// 10. 1:1 한글 계정명 매핑 함수 (MSA Safety 버전과 100% 동일하게 고정!)
function matchExactKoreanAccountName(contextualLabel, tag = '', stmtType = '', rowType = 'USD') {
  if (stmtType === 'BS' && tag === 'LiabilitiesAndStockholdersEquity') {
    return '[[부채와자본총계]]';
  }
  if (tag === 'ProfitLoss') {
    return '연결당기순이익';
  }

  const raw = stripDynamicParentheticals(contextualLabel);
  if (!raw) return null;

  if (rowType === 'EPS') {
    if (/basic/i.test(raw) && !/diluted/i.test(raw)) return '기본 주당순이익(달러)';
    if (/diluted/i.test(raw)) return '희석 주당순이익(달러)';
  }
  if (rowType === 'SHARES') {
    if (/basic/i.test(raw) && !/diluted/i.test(raw)) return '기본 가중평균발행주식수(백만주)';
    if (/diluted/i.test(raw)) return '희석 가중평균발행주식수(백만주)';
  }

  if (!/[a-zA-Z]/.test(raw)) return raw;

  if (
    stmtType === 'CF' &&
    /IncreaseDecreaseIn(?:Accounts)?Receivables?/i.test(tag) &&
    /^(?:trade\s+)?(?:accounts\s+)?receivables?$/i.test(raw)
  ) {
    return '매출채권의 증감';
  }
  if (
    stmtType === 'CF' &&
    tag === 'PaymentsToMinorityShareholders' &&
    /\b(?:distributions?|dividends?)\b/i.test(raw)
  ) {
    return /\bdividends?\b/i.test(raw) ? '비지배지분 배당금의 지급' : '비지배지분 분배금 지급';
  }

  const exactRules = [
    [/^interest and fees on loans and leases$/i, '대출채권및리스이자수익'],
    [/^interest on debt securities$|^interest and dividends on securities$/i, '채무증권이자수익'],
    [/^federal funds sold and securities borrowed or purchased under agreements to resell$/i,
      stmtType === 'IS'
        ? '연방기금매도및환매조건부매수이자'
        : stmtType === 'CF'
        ? '연방기금매도및환매조건부매수채권의 순증감'
        : '연방기금매도및환매조건부매수채권'],
    [/^trading account assets$/i, stmtType === 'IS' ? '단기매매금융자산이자수익' : '단기매매금융자산'],
    [/^other interest income$/i, '기타이자수익'],
    [/^total interest income$|^interest income$/i, '이자수익'],
    [/^deposits$/i, stmtType === 'IS' ? '예수금이자비용' : stmtType === 'CF' ? '예수금의 순증감' : '예수금'],
    [/^short-?term borrowings$/i,
      stmtType === 'IS' ? '단기차입금이자비용' : stmtType === 'CF' ? '단기차입금의 순증감' : '단기차입금'],
    [/^trading account liabilities$/i, stmtType === 'IS' ? '단기매매금융부채이자비용' : '단기매매금융부채'],
    [/^long-?term debt(?:,?\s*net)?$/i, stmtType === 'IS' ? '장기차입금이자비용' : '장기차입금및금융리스부채'],
    [/^total interest expense$/i, '이자비용'],
    [/^net interest income$/i, '순이자손익'],
    [/^interest\s*(?:\(income\)\s*expense|expense\s*\(income\)|expense|income),?\s*net$|^interest,?\s*net$/i,
      '순이자비용(수익)'],
    [/^interest expense$/i, '이자비용'],
    [/^(?:loss|gain \(loss\)|gains \(losses\)) on extinguishment of debt.*$/i,
      '채무상환손익(조기상환손실)'],
    [/^store closure and other costs.*$|^store closure and exit costs.*$|^store closure costs.*$/i,
      '매장폐쇄및기타비용(순액)'],
    [/^direct store expenses?$/i, '점포직접비'],
    [/^store pre-?opening costs?$/i, '점포개점전비용'],
    [/^clearing and transaction fees?$/i, '청산및거래수수료수익'],
    [/^market data and information services$/i, '시장데이터및정보서비스수익'],
    [/^card income$/i, '카드수수료수익'],
    [/^service charges$/i, '서비스수수료수익'],
    [/^investment and brokerage services$/i, '투자및중개수수료수익'],
    [/^investment banking fees$|^investment banking income$/i, '투자은행(IB)수수료수익'],
    [/^total fees and commissions$|^fees and commissions$/i, '수수료및커미션수익'],
    [/^market making and similar activities$|^trading account profits$|^trading account income$/i,
      '시장조성및유사활동수익'],
    [/^total non-?interest income$|^non-?interest income$/i, '비이자수익합계'],
    [/^total revenue,?\s*net of interest expense$/i, '순영업수익(이자비용차감후총수익)'],
    [/^provision for credit losses$|^provision \(benefit\) for credit losses$|^provision for loan and lease losses$/i,
      '신용손실충당금전입액(대손상각비)'],
    [/^compensation and benefits$|^salaries and employee benefits$|^personnel$/i, '급여및복리후생비'],
    [/^occupancy and equipment$|^net occupancy$|^occupancy$/i, '임차료및설비비'],
    [/^information processing and communications$|^technology and communications$|^data processing$|^technology$/i,
      '정보처리및통신비'],
    [/^licensing and other fee agreements$/i, '라이선스및기타수수료약정비용'],
    [/^product delivery and transaction related$/i, '상품제공및거래관련비용'],
    [/^professional fees and outside services$|^professional fees$|^professional services$/i,
      '지급수수료(전문가보수)'],
    [/^marketing$|^marketing and advertising$/i, '마케팅및광고선전비'],
    [/^other general operating$|^other operating expense$/i, '기타일반영업비용'],
    [/^total non-?interest expense$|^non-?interest expense$/i, '비이자비용합계'],
    [/^preferred stock dividends$/i, '우선주배당금'],
    [/^net income applicable to common shareholders$|^net income available to common stockholders$|^net income attributable to common.*$/i,
      '보통주귀속당기순이익'],

    [/^cash and due from banks$/i, '현금및타행예치금'],
    [/^interest-?bearing deposits with the federal reserve.*|^interest-?bearing deposits in banks$/i,
      '연방준비은행및타행이자부예치금'],
    [/^cash and cash equivalents,?\s*beginning of (?:period|year)$/i, '기초 현금및현금성자산'],
    [/^cash and cash equivalents,?\s*end of (?:period|year)$/i, '기말 현금및현금성자산'],
    [/^cash and cash equivalents$/i, '현금및현금성자산'],
    [/^restricted cash$/i, '사용제한현금'],
    [/^marketable securities$/i, /Noncurrent/i.test(tag) ? '장기유가증권' : '단기유가증권'],
    [/^(?:cash )?performance bonds and (?:guaranty|guarantee) fund contributions$/i,
      '계약이행보증금및보증기금'],
    [/^time deposits placed and other short-?term investments$/i,
      stmtType === 'CF' ? '정기예치금및기타단기금융상품의 순증감' : '정기예치금및기타단기금융상품'],
    [/^derivative assets$/i, '파생상품자산'],
    [/^derivative liabilities$/i, '파생상품부채'],
    [/^available-?for-?sale debt securities.*|^debt securities carried at fair value$|^carried at fair value$/i,
      '공정가치측정채무증권(매도가능)'],
    [/^held-?to-?maturity debt securities.*|^held-?to-?maturity,?\s*at cost$/i,
      '만기보유채무증권(상각후원가)'],
    [/^total debt securities$|^total marketable securities$/i, '채무증권(유가증권)합계'],
    [/^loans and leases$/i, stmtType === 'CF' ? '대출채권및리스채권의 순증감' : '대출채권및리스채권'],
    [/^allowance for loan and lease losses$|^allowance for credit losses$/i, '대출채권및리스대손충당금'],
    [/^loans and leases,?\s*net of allowance$|^loans and leases,?\s*net$/i, '대출채권및리스채권(순액)'],
    [/^premises and equipment,?\s*net$|^property,?\s*plant and equipment,?\s*net$|^property and equipment,?\s*net$/i,
      '유형자산(순액)'],
    [/^operating lease right-?of-?use assets?,?\s*net$|^operating lease right-?of-?use assets?$/i,
      '영업리스사용권자산'],
    [/^finance lease right-?of-?use assets?,?\s*net$|^finance lease right-?of-?use assets?$/i,
      '금융리스사용권자산'],
    [/^goodwill$/i, '영업권'],
    [/^intangible assets\s*[—\-–]\s*trading products$/i, '무형자산(거래상품)'],
    [/^intangible assets\s*[—\-–]\s*other,?\s*net$|^intangible assets,?\s*net$|^other intangible assets,?\s*net$|^intangible assets$/i,
      '무형자산(순액)'],
    [/^loans held-?for-?sale$/i, stmtType === 'CF' ? '매각예정대출채권의 순증감' : '매각예정대출채권'],
    [/^accounts receivable,?\s*net$|^trade accounts receivable,?\s*net$/i,
      stmtType === 'CF' ? '매출채권의 증감' : '매출채권(순액)'],
    [/^inventories$|^inventory$|^merchandise inventories$/i,
      stmtType === 'CF' ? '재고자산의 증감' : '재고자산'],
    [/^prepaid expenses and other current assets$|^prepaid expenses$/i,
      stmtType === 'CF' ? '선급비용및기타유동자산의 증감' : '선급비용및기타유동자산'],
    [/^other current assets$/i, stmtType === 'CF' ? '기타유동자산의 증감' : '기타유동자산'],
    [/^other non-?current assets$|^other assets$/i, stmtType === 'CF' ? '기타자산의 증감' : '기타비유동자산'],
    [/^customer and other receivables$|^accrued interest and other receivables$/i, '고객및기타수취채권'],
    [/^u\.?s\.?\s*non-?interest-?bearing deposits$|^non-?interest-?bearing deposits$/i, '미국내 무이자부예수금'],
    [/^u\.?s\.?\s*interest-?bearing deposits$|^interest-?bearing deposits$/i, '미국내 이자부예수금'],
    [/^non-?u\.?s\.?\s*non-?interest-?bearing deposits$/i, '해외 무이자부예수금'],
    [/^non-?u\.?s\.?\s*interest-?bearing deposits$/i, '해외 이자부예수금'],
    [/^total deposits$/i, '[예수금 합계]'],
    [/^federal funds purchased and securities loaned or sold under agreements to repurchase$/i,
      stmtType === 'CF'
        ? '연방기금매입및환매조건부매도채무의 순증감'
        : '연방기금매입및환매조건부매도채무'],
    [/^trade accounts payable and other accrued liabilities$|^accounts payable and accrued liabilities$/i,
      stmtType === 'CF' ? '매입채무및미지급비용의 증감' : '매입채무및미지급비용'],
    [/^accounts payable$|^trade accounts payable$/i,
      stmtType === 'CF' ? '매입채무의 증감' : '매입채무'],
    [/^accrued liabilities$|^other accrued liabilities$|^accrued expenses$/i,
      stmtType === 'CF' ? '미지급비용의 증감' : '미지급비용'],
    [/^accrued salaries and benefits$|^accrued compensation and benefits$/i,
      stmtType === 'CF' ? '미지급급여및복리후생비의 증감' : '미지급급여및복리후생비'],
    [/^accrued expenses and other liabilities$/i,
      stmtType === 'CF' ? '미지급비용및기타부채의 증감' : '미지급비용및기타부채'],
    [/^other current liabilities$/i, stmtType === 'CF' ? '기타유동부채의 증감' : '기타유동부채'],
    [/^current portion of operating lease liabilities$|^operating lease liabilities,?\s*current$/i,
      '유동성영업리스부채'],
    [/^current portion of (?:capital and financing|capital|finance) lease.*$/i,
      '유동성금융리스부채'],
    [/^current portion of long-?term debt.*|^short-?term debt and current portion of long-?term debt$/i,
      '유동성장기부채및금융리스부채'],
    [/^long-?term operating lease liabilities.*|^operating lease liabilities,?\s*non-?current$|^operating lease liabilities,?\s*net$/i,
      '장기영업리스부채'],
    [/^(?:long-?term )?(?:capital and financing|capital|finance) lease (?:obligations|liabilities)(?:,?\s*net.*)?$/i,
      '비유동금융리스부채'],
    [/^long-?term debt and (?:other financing|capital lease|finance lease|capital and financing lease).*|^long-?term debt,?\s*net.*|^long-?term debt,?\s*non-?current$/i,
      '장기차입금및금융리스부채'],
    [/^other non-?current liabilities$|^other long-?term liabilities$|^other liabilities$/i,
      stmtType === 'CF' ? '기타부채의 증감' : '기타비유동부채'],
    [/^deferred income tax assets?,?\s*net$|^deferred income tax assets?$|^deferred tax assets?,?\s*net$|^deferred tax assets?$/i,
      stmtType === 'BS' ? '이연법인세자산' : '이연법인세비용(혜택)'],
    [/^deferred income tax liabilities,?\s*net$|^deferred income tax liabilities$|^deferred income taxes$/i,
      stmtType === 'BS'
        ? /Deferred(?:Income)?TaxAssets/i.test(tag)
          ? '이연법인세자산'
          : '이연법인세부채'
        : '이연법인세비용(혜택)'],
    [/^common stock$/i, '보통주자본금'],
    [/^preferred stock$/i, '우선주자본금'],
    [/^additional paid-?in capital$/i, '자본잉여금(보통주)'],
    [/^common stock and additional paid-?in capital$/i, '보통주자본금및자본잉여금'],
    [/^retained earnings$|^retained earnings \(accumulated deficit\)$|^accumulated deficit$/i,
      '이익잉여금(결손금)'],
    [/^treasury stock$|^treasury stock,?\s*at cost$/i, '자기주식'],
    [/^accumulated other comprehensive income \(loss\)$|^accumulated other comprehensive loss$/i,
      '기타포괄손익누계액'],

    [/^net \(gains\) losses on sales of (?:available-?for-?sale )?debt securities$|^gains \(losses\) on sales of debt securities$/i,
      '채무증권처분손익(순)'],
    [/^depreciation and premises improvements amortization$|^depreciation and amortization.*$/i,
      '감가상각비및무형자산상각비'],
    [/^operating lease asset amortization$|^amortization of operating lease right-?of-?use assets$/i,
      '운용리스사용권자산 상각비'],
    [/^amortization of financing fees and debt issuance costs$|^amortization of debt issuance costs$/i,
      '금융비용및사채발행비 상각액'],
    [/^amortization of intangibles$|^amortization of intangible assets$/i, '무형자산상각비'],
    [/^net amortization of premium\/discount on debt securities$/i, '채무증권할증·할인차금상각(순)'],
    [/^stock-?based compensation$|^share-?based compensation$/i, '주식기반보상비용'],
    [/^originations and purchases$|^originations and purchases of loans held-?for-?sale$/i,
      '매각예정대출채권의 실행및매입'],
    [/^proceeds from sales and paydowns of loans originally classified as held for sale.*/i,
      '매각예정대출채권의 매각및회수'],
    [/^net change in loans held-?for-?sale$/i, '매각예정대출채권의 순증감'],
    [/^trading and derivative assets\/liabilities.*|^net change in trading and derivative.*/i,
      '단기매매및파생상품자산·부채의 순증감'],
    [/^other operating activities,?\s*net$/i, '기타영업활동현금흐름(순)'],
    [/^net change in time deposits placed and other short-?term investments$/i,
      '정기예치금및기타단기금융상품의 순증감'],
    [/^net change in federal funds sold and securities borrowed or purchased under agreements to resell$/i,
      '연방기금매도및환매조건부매수채권의 순증감'],
    [/^proceeds from sales of available-?for-?sale debt securities$/i, '공정가치측정 채무증권의 처분'],
    [/^proceeds from paydowns and maturities of available-?for-?sale debt securities$|^proceeds from maturities,?\s*calls and mandatory redemptions of available-?for-?sale debt securities$/i,
      '공정가치측정 채무증권의 만기도래 및 조기상환'],
    [/^purchases of available-?for-?sale debt securities$/i, '공정가치측정 채무증권의 취득'],
    [/^proceeds from paydowns and maturities of held-?to-?maturity debt securities$|^proceeds from maturities of held-?to-?maturity debt securities$/i,
      '만기보유채무증권의 원금회수및만기상환'],
    [/^purchases of held-?to-?maturity debt securities$/i, '만기보유채무증권의 취득'],
    [/^proceeds from sales of loans originally classified as held for investment.*|^proceeds from sales and paydowns of loans originally classified as held for investment.*/i,
      '투자목적분류대출채권의 매각및원금회수'],
    [/^purchases of loans and leases$/i, '대출채권및리스채권의 매입(취득)'],
    [/^other changes in loans and leases,?\s*net$|^net change in loans and leases$/i,
      '대출채권및리스채권의 기타순증감'],
    [/^additions to property and equipment$|^capital expenditures$|^purchases of property,?\s*plant and equipment$/i,
      '유형자산의 취득(CAPEX)'],
    [/^acquisitions?,?\s*net of cash acquired$|^payments to acquire businesses.*$|^business acquisitions?,?\s*net.*$/i,
      '기업인수로 인한 현금유출'],
    [/^other investing activities,?\s*net$/i, '기타투자활동현금흐름(순)'],
    [/^net change in deposits$/i, '예수금의 순증감'],
    [/^net change in federal funds purchased and securities loaned or sold under agreements to repurchase$/i,
      '연방기금매입및환매조건부매도채무의 순증감'],
    [/^net change in short-?term borrowings$/i, '단기차입금의 순증감'],
    [/^proceeds from revolving credit facilities$|^borrowings under revolving credit facilities$/i,
      '회전신용한도 차입금의 유입'],
    [/^payments on revolving credit facilities$|^repayments of revolving credit facilities$/i,
      '회전신용한도 차입금의 상환'],
    [/^payments on (?:capital and financing|capital|finance) lease obligations$/i,
      '자본리스및금융리스부채 상환액'],
    [/^proceeds from issuance of long-?term debt$/i, '장기차입금의 차입'],
    [/^retirement of long-?term debt$|^repayments of long-?term debt$|^payments on long-?term debt$/i,
      '장기차입금의 상환'],
    [/^proceeds from issuance of preferred stock$/i, '우선주 발행에 따른 현금유입'],
    [/^redemption of preferred stock$/i, '우선주 상환에 따른 현금유출'],
    [/^common stock repurchased$|^repurchases of common stock$|^treasury stock purchases?$|^purchases? of treasury stock$/i,
      '자기주식의 취득(자사주매입)'],
    [/^cash dividends paid$|^dividends paid$|^dividends paid on common stock$/i, '현금배당금의 지급'],
    [/^dividends paid to noncontrolling interests?(?:\s+holders)?$/i, '비지배지분 배당금의 지급'],
    [/^distributions to noncontrolling interests?(?:\s+holders)?$/i, '비지배지분 분배금 지급'],
    [/^other financing activities,?\s*net$/i, '기타재무활동현금흐름(순)'],

    [/^revenues?$|^net sales$|^total net sales$|^total revenues?$/i, '매출액'],
    [/^products?$/i, '제품매출액'],
    [/^services?$/i, '서비스매출액'],
    [/^cost of sales,?\s*buying and occupancy$/i, '매출원가, 매입 및 점포임차비용'],
    [/^cost of sales$|^cost of goods sold$|^cost of revenues?$|^total cost of sales$|^cost of sales \(exclusive of.*$/i,
      '매출원가'],
    [/^gross profit$|^gross margin$/i, '매출총이익'],
    [/^research and development$/i, '연구개발비'],
    [/^selling,?\s*general and administrative.*$/i, '판매비와관리비'],
    [/^total operating expenses$|^operating expenses$/i, '영업비용합계'],
    [/^operating income$|^operating income \(loss\)$|^operating profit$/i, '영업이익'],
    [/^other income \(loss\)$|^other income,?\s*net$|^other income \(expense\),?\s*net$/i, '기타손익(순)'],
    [/^other,?\s*net$/i,
      stmtType === 'CF'
        ? /Investing/i.test(tag)
          ? '기타투자활동현금흐름(순)'
          : /Financing/i.test(tag)
          ? '기타재무활동현금흐름(순)'
          : /Noncash/i.test(tag)
          ? '기타비현금성손익'
          : /Operating/i.test(tag)
          ? '기타영업활동현금흐름(순)'
          : null
        : '기타손익(순)'],
    [/^income before income taxes$|^income \(loss\) before income taxes$|^income before provision for income taxes$|^earnings before income taxes$/i,
      '법인세비용차감전순이익'],
    [/^provision for income taxes$|^income tax provision$|^income tax expense$/i, '법인세비용'],
    [/^income taxes$/i,
      stmtType === 'BS'
        ? /Noncurrent/i.test(tag)
          ? '장기미지급법인세'
          : /Receivable|Asset/i.test(tag)
          ? '미수법인세'
          : '미지급법인세(유동)'
        : stmtType === 'CF'
        ? /IncomeTaxesPaid/i.test(tag)
          ? '법인세납부액'
          : /Deferred/i.test(tag)
          ? '이연법인세비용(혜택)'
          : '미지급법인세의 증감'
        : '법인세비용'],
    [/^net income$|^net earnings$|^net income \(loss\)$/i, '당기순이익'],
    [/^net increase \(decrease\) in cash and cash equivalents.*$/i, '현금및현금성자산의 순증감'],
    [/^effect of exchange rate changes on cash.*/i, '현금에 대한 환율변동효과'],
    [/^total current assets$/i, '[유동자산 합계]'],
    [/^total non-?current assets$|^total other assets$/i, '[비유동자산 합계]'],
    [/^total assets$/i, '[[자산총계]]'],
    [/^total current liabilities$/i, '[유동부채 합계]'],
    [/^total non-?current liabilities$|^total long-?term liabilities$/i, '[비유동부채 합계]'],
    [/^total liabilities$/i, '[[부채총계]]'],
    [/^redeemable noncontrolling interests?$/i, '상환가능비지배지분'],
    [/^total liabilities(?:,?\s*redeemable noncontrolling interests?)?(?:,?\s*and|\s+and)\s+(?:stockholders|shareholders)'? equity$|^total liabilities and equity$/i,
      '[[부채와자본총계]]'],
    [/^total stockholders'? equity$|^total shareholders'? equity$|^total equity$|^total (?:equity and )?partners'? capital$/i, '[[자본총계]]'],
    [/^total (?!liabilities\b).*?(?:stockholders|shareholders)'? equity$|^total parent partners'? capital$/i, '지배기업주주지분'],
    [/^net cash provided by \(used in\) operating activities$|^cash generated by operating activities$|^cash flows? from operating activities$/i,
      '[영업활동현금흐름]'],
    [/^net cash provided by \(used in\) investing activities$|^cash used in investing activities$|^cash flows? from investing activities$/i,
      '[투자활동현금흐름]'],
    [/^net cash provided by \(used in\) financing activities$|^cash used in financing activities$|^cash flows? from financing activities$/i,
      '[재무활동현금흐름]'],
  ];

  for (const [regex, kor] of exactRules) {
    if (regex.test(raw)) return kor;
  }

  return null;
}

function normalizeLabel(lbl) {
  return stripDynamicParentheticals(lbl)
    .toLowerCase()
    .replace(/^cash\s+(?=performance bonds)/g, '')
    .replace(/,?\s*net$/g, '')
    .replace(/\bexpenses?\b/g, '')
    .replace(/[^a-z0-9]/g, '');
}

// 11. 현금흐름표 유출/유입 판별 함수 (수치 로직 100% 동일)
function classifyCashFlowRowPolarity(tag = '', contextualLabel = '', korName = '') {
  const t = String(tag || '');
  const lbl = String(contextualLabel || '').toLowerCase();
  const kor = String(korName || '');

  if (
    /ProceedsFrom(?:Payments|Repayments|Repurchase)|PaymentsForProceeds|IncreaseDecrease|NetChange|^CashFlowFrom|DisbursementsAndDeposits|DepositsAndDisbursements|ReceiptsAndDisbursements|DepositsAndWithdrawals|WithdrawalsAndDeposits/i.test(
      t
    )
  ) {
    return 'NEUTRAL';
  }
  if (
    /\bnet\s+change\b|\bincrease\s*\(decrease\)|\bdisbursements?\s*\(deposits?\)|\bdeposits?\s*\(disbursements?\)|\bproceeds?\s*\((?:payments?|deposits?|outflows?)\)|\brefunds?,?\s*net\b|순증감|증감|기타.*순/i.test(
      `${lbl} ${kor}`
    )
  ) {
    return 'NEUTRAL';
  }
  // 영업·투자활동의 대출채권/대여금·투자자산·우선주 회수성 RepaymentsOf*는 차입금 상환이 아니므로 원표기 부호 유지
  if (
    /RepaymentsOf.*(?:Loans|Advances|NotesReceivable|PreferredEquity|Investments?)/i.test(t) &&
    !/\[재무활동\]/i.test(`${lbl} ${kor}`) &&
    !/(?:Debt|Borrowing|Credit|Bank|TermLoan|LineOfCredit|LongTerm|ShortTerm|Senior|Subordinated)/i.test(
      t
    )
  ) {
    return 'NEUTRAL';
  }

  if (/^ProceedsFrom/i.test(t)) {
    return 'INFLOW';
  }
  if (
    /^(?:PaymentsToAcquire|PaymentsForRepurchase|PaymentsOfDividends|PaymentsOf|PaymentsFor|PaymentsRelatedTo|PaymentsToMinority|RepaymentsOf|FinanceLeasePrincipalPayments)/i.test(
      t
    )
  ) {
    return 'OUTFLOW';
  }
  if (
    /^(?:purchases? of|additions? to|acquisitions? of|business acquisitions?|acquisitions?,?\s*net of cash|capital expenditures|repayments? of|payments? (?:on|of|for|to)|repurchases? of|common stock repurchased|redemption of|retirement of|cash dividends paid|dividends paid)/i.test(
      lbl
    )
  ) {
    return 'OUTFLOW';
  }
  if (
    /의\s*취득|인수로\s*인한\s*현금유출|상환에\s*따른\s*현금유출|의\s*상환|상환액|배당금의?\s*지급|원천징수세액\s*납부|출연금\s*지급|배상금\s*지급|원가의\s*지급|차입부채발행비용|비지배지분\s*취득|자기주식\s*취득/.test(
      kor
    )
  ) {
    return 'OUTFLOW';
  }

  if (/^(?:proceeds from|borrowings under|issuance of)/i.test(lbl)) {
    return 'INFLOW';
  }
  if (/현금유입|차입금의\s*(?:차입|유입)|행사에\s*따른\s*유입|매입제도로\s*인한\s*유입|만기회수|만기도래\s*및\s*조기상환/.test(kor)) {
    return 'INFLOW';
  }

  return 'NEUTRAL';
}

function isMezzanineOrTemporaryEquityRow(tag = '', label = '') {
  if (
    tag === 'LiabilitiesAndStockholdersEquity' ||
    /^total\s+liabilities\b/i.test(String(label || '').trim())
  ) {
    return false;
  }
  return (
    /RedeemableNoncontrollingInterest|TemporaryEquity/i.test(String(tag || '')) ||
    /redeemable\s+noncontrolling|temporary\s+equity|상환가능비지배지분|상환가능지분|메자닌/i.test(
      String(label || '')
    )
  );
}

function isSpecialStandaloneBsLiabilityRow(tag = '', label = '') {
  return (
    /^LiabilitiesSubjectToCompromise$/i.test(String(tag || '')) ||
    /\bliabilities\s+subject\s+to\s+compromise\b|화해대상부채/i.test(String(label || ''))
  );
}

function findOperatingCfStartIdx(rows, subIdx) {
  const netIncIdx = rows.findIndex(
    (r, idx) =>
      idx < subIdx &&
      (r.tag === 'NetIncomeLoss' ||
        r.tag === 'ProfitLoss' ||
        /^net income|^net earnings/i.test(r.companyLabel || ''))
  );
  const contOpsIdx = rows.findIndex(
    (r, idx) =>
      idx < subIdx &&
      (r.tag === 'IncomeLossFromContinuingOperations' ||
        /income\s*\(?loss\)?\s*from\s+continuing\s+operations/i.test(r.companyLabel || ''))
  );
  if (
    contOpsIdx !== -1 &&
    (netIncIdx === -1 || (contOpsIdx > netIncIdx && contOpsIdx - netIncIdx <= 3))
  ) {
    return contOpsIdx;
  }
  return netIncIdx !== -1 ? netIncIdx : 0;
}

// 12. 소계 항등식 역산을 통한 자동 부호 교정 엔진 (Phase A 안전장치 + IS 보수적 부호 안전망)
function autoFixRowSignsBySubtotalReconciliation(rows, targetYears, stmtType) {
  if (stmtType === 'IS') {
    if (rows.some((r) => r.dimAxis === 'ProductOrServiceAxis')) return;
    const metaIdx = rows.findIndex((r) => r.rowType === 'EPS' || r.rowType === 'SHARES');
    const limit = metaIdx !== -1 ? metaIdx : rows.length;
    const findMainIdx = (tags) =>
      rows.findIndex((r, idx) => idx < limit && r.isMainRow && r.rowType === 'USD' && tags.includes(r.tag));

    const opIdx = findMainIdx(['OperatingIncomeLoss']);
    const preTaxIdx = findMainIdx([
      'IncomeLossFromContinuingOperationsBeforeIncomeTaxesExtraordinaryItemsNoncontrollingInterest',
      'IncomeLossFromContinuingOperationsBeforeIncomeTaxesMinorityInterestAndIncomeLossFromEquityMethodInvestments',
    ]);
    const taxIdx = findMainIdx(['IncomeTaxExpenseBenefit']);
    const netIdx = rows.findIndex(
      (r, idx) =>
        idx > (taxIdx !== -1 ? taxIdx : preTaxIdx) &&
        idx < limit &&
        r.rowType === 'USD' &&
        (r.tag === 'IncomeLossFromContinuingOperations' ||
          r.tag === 'ProfitLoss' ||
          r.tag === 'NetIncomeLoss')
    );

    const recentYears = targetYears.slice(-3);

    // 1) 영업이익 -> 법인세비용차감전순이익 단일 비영업/이자비용 행 부호 안전망
    if (opIdx !== -1 && preTaxIdx !== -1 && opIdx < preTaxIdx) {
      const comps = [];
      for (let i = opIdx + 1; i < preTaxIdx; i++) {
        if (rows[i].rowType === 'USD' && !rows[i].isTotalRow) comps.push(i);
      }
      if (comps.length > 0 && comps.length <= 5) {
        let anchorOkCount = 0;
        for (const y of recentYears) {
          const op = rows[opIdx].rValuesByYear?.[y];
          const pt = rows[preTaxIdx].rValuesByYear?.[y];
          if (op === undefined || pt === undefined) continue;
          const sumComps = comps.reduce((acc, i) => acc + (rows[i].rValuesByYear?.[y] ?? 0), 0);
          if (Math.abs(op + sumComps - pt) <= 3) anchorOkCount++;
        }
        if (anchorOkCount >= 1) {
          for (const y of targetYears) {
            const op = rows[opIdx].rValuesByYear?.[y];
            const pt = rows[preTaxIdx].rValuesByYear?.[y];
            if (op === undefined || pt === undefined) continue;
            const sumComps = comps.reduce((acc, i) => acc + (rows[i].rValuesByYear?.[y] ?? 0), 0);
            const curErr = Math.abs(op + sumComps - pt);
            if (curErr <= 3) continue;
            for (const cIdx of comps) {
              const val = rows[cIdx].rValuesByYear?.[y];
              if (val === undefined || val === 0) continue;
              const anchorHasSameSignAsFlip = recentYears.some(
                (ry) =>
                  rows[cIdx].rValuesByYear?.[ry] !== undefined &&
                  rows[cIdx].rValuesByYear?.[ry] !== 0 &&
                  Math.sign(rows[cIdx].rValuesByYear[ry]) === Math.sign(-val)
              );
              if (!anchorHasSameSignAsFlip) continue;
              if (Math.abs(op + sumComps - 2 * val - pt) <= 3) {
                rows[cIdx].rValuesByYear[y] = -val;
                break;
              }
            }
          }
        }
      }
    }

    // 2) 세전이익 +- 법인세 = 순이익 안전망 (단, 앵커 산식과 어긋날 때만 교정; 세금~순이익 사이 지분법손익 등 중간행 존재 시 함께 반영)
    if (preTaxIdx !== -1 && taxIdx !== -1 && netIdx !== -1 && preTaxIdx < taxIdx && taxIdx < netIdx) {
      const betweenTaxNet = [];
      for (let i = taxIdx + 1; i < netIdx; i++) {
        if (rows[i].rowType === 'USD' && !rows[i].isTotalRow && !rows[i].dimMember) {
          betweenTaxNet.push(i);
        }
      }
      if (betweenTaxNet.length === 0) {
        let minusOk = 0;
        let plusOk = 0;
        for (const y of recentYears) {
          const pt = rows[preTaxIdx].rValuesByYear?.[y];
          const tx = rows[taxIdx].rValuesByYear?.[y];
          const nt = rows[netIdx].rValuesByYear?.[y];
          if (pt === undefined || tx === undefined || tx === 0 || nt === undefined) continue;
          if (Math.abs(pt - tx - nt) <= 3) minusOk++;
          if (Math.abs(pt + tx - nt) <= 3) plusOk++;
        }
        if ((minusOk > 0) !== (plusOk > 0)) {
          const taxOp = minusOk > plusOk ? -1 : 1;
          for (const y of targetYears) {
            const pt = rows[preTaxIdx].rValuesByYear?.[y];
            const tx = rows[taxIdx].rValuesByYear?.[y];
            const nt = rows[netIdx].rValuesByYear?.[y];
            if (pt === undefined || tx === undefined || tx === 0 || nt === undefined) continue;
            if (Math.abs(pt + taxOp * tx - nt) > 3 && Math.abs(pt - taxOp * tx - nt) <= 3) {
              rows[taxIdx].rValuesByYear[y] = -tx;
            }
          }
        }
      } else {
        const vecs = [
          [-1, 1],
          [1, 1],
          [-1, -1],
          [1, -1],
        ];
        const validVecs = vecs.filter(([tOp, bOp]) => {
          let ok = 0;
          for (const y of recentYears) {
            const pt = rows[preTaxIdx].rValuesByYear?.[y];
            const tx = rows[taxIdx].rValuesByYear?.[y];
            const nt = rows[netIdx].rValuesByYear?.[y];
            if (pt === undefined || tx === undefined || tx === 0 || nt === undefined) continue;
            const bSum = betweenTaxNet.reduce((acc, idx) => acc + (rows[idx].rValuesByYear?.[y] ?? 0), 0);
            if (Math.abs(pt + tOp * tx + bOp * bSum - nt) <= 3) ok++;
            else return false;
          }
          return ok > 0;
        });
        if (validVecs.length === 1) {
          const [taxOp, bOp] = validVecs[0];
          for (const y of targetYears) {
            const pt = rows[preTaxIdx].rValuesByYear?.[y];
            const tx = rows[taxIdx].rValuesByYear?.[y];
            const nt = rows[netIdx].rValuesByYear?.[y];
            if (pt === undefined || tx === undefined || tx === 0 || nt === undefined) continue;
            const bSum = betweenTaxNet.reduce((acc, idx) => acc + (rows[idx].rValuesByYear?.[y] ?? 0), 0);
            if (
              Math.abs(pt + taxOp * tx + bOp * bSum - nt) > 3 &&
              Math.abs(pt - taxOp * tx + bOp * bSum - nt) <= 3
            ) {
              rows[taxIdx].rValuesByYear[y] = -tx;
            }
          }
        }
      }
    }
    return;
  }

  if (stmtType !== 'CF' && stmtType !== 'BS') return;

  const assetsTotalIdx =
    stmtType === 'BS'
      ? rows.findIndex(
          (r) =>
            r.isTotalRow &&
            !r.dimMember &&
            (r.tag === 'Assets' || /^total assets$/i.test(r.companyLabel || ''))
        )
      : -1;

  const isFlipForbidden = (row, val, sIdx = -1) => {
    if (stmtType === 'CF') {
      if (['ProfitLoss', 'NetIncomeLoss', 'IncomeLossFromContinuingOperations'].includes(row.tag)) {
        return true;
      }
      const stdRule = STANDARD_KOR_DICT[row.tag] || null;
      const polarity = classifyCashFlowRowPolarity(
        row.tag,
        row.contextualLabel,
        row.exactKor || stdRule?.kor || ''
      );
      if ((polarity === 'OUTFLOW' || stdRule?.sign === 'neg') && -val > 0) return true;
      if ((polarity === 'INFLOW' || stdRule?.sign === 'pos') && -val < 0) return true;
      return false;
    }
    if (stmtType === 'BS' && val > 0) {
      const stdRule = STANDARD_KOR_DICT[row.tag] || null;
      const isContraOrFlexible =
        stdRule?.sign === 'neg' ||
        /Allowance|AccumulatedDepreciation|AccumulatedAmortization|TreasuryStock|RetainedEarnings|AccumulatedOtherComprehensive|AccumulatedDeficit|Valuation|Discount|Contra/i.test(
          row.tag || ''
        ) ||
        /\b(?:allowance|accumulated\s+(?:depreciation|amortization|deficit|other\s+comprehensive)|treasury\s+stock|retained\s+earnings|less:)\b/i.test(
          row.contextualLabel || ''
        );
      if (!isContraOrFlexible) {
        if (assetsTotalIdx !== -1 && sIdx !== -1 && sIdx <= assetsTotalIdx) return true;
        const allVals = Object.values(row.rValuesByYear || {}).filter(
          (v) => typeof v === 'number' && v !== 0
        );
        const posCount = allVals.filter((v) => v > 0).length;
        const negCount = allVals.filter((v) => v < 0).length;
        if (posCount >= 2 && negCount === 0) return true;
      }
    }
    return false;
  };

  let prevTotalIdx = -1;
  for (let sIdx = 0; sIdx < rows.length; sIdx++) {
    if (!rows[sIdx].isTotalRow || rows[sIdx].rowType !== 'USD') continue;
    if (stmtType === 'CF' && isIntermediateCfAdjustmentSubtotal(rows[sIdx].tag, rows[sIdx].companyLabel)) {
      continue;
    }

    const subTotalRow = rows[sIdx];
    let compStartIdx = prevTotalIdx + 1;

    if (stmtType === 'CF' && prevTotalIdx === -1) {
      compStartIdx = findOperatingCfStartIdx(rows, sIdx);
    }
    if (
      stmtType === 'BS' &&
      (subTotalRow.tag === 'AssetsNoncurrent' || subTotalRow.tag === 'LiabilitiesNoncurrent')
    ) {
      const targetAbs =
        subTotalRow.tag === 'AssetsNoncurrent' ? 'ASSETS_NONCURRENT' : 'LIABILITIES_NONCURRENT';
      const absStart = findTightenedBsSectionStart(rows, compStartIdx, sIdx, targetAbs);
      if (absStart > compStartIdx) compStartIdx = absStart;
    }

    const compIndices = [];
    for (let i = compStartIdx; i < sIdx; i++) {
      if (rows[i].rowType !== 'USD' || rows[i].isTotalRow) continue;
      if (stmtType === 'CF' && isIntermediateCfAdjustmentSubtotal(rows[i].tag, rows[i].companyLabel)) {
        continue;
      }
      if (
        stmtType === 'BS' &&
        (isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].companyLabel) ||
          isSpecialStandaloneBsLiabilityRow(rows[i].tag, rows[i].companyLabel))
      ) {
        continue;
      }
      compIndices.push(i);
    }

    if (compIndices.length > 0) {
      for (const y of targetYears) {
        const subVal = subTotalRow.rValuesByYear?.[y];
        if (subVal === undefined || subVal === null) continue;

        const activeIndices = compIndices.filter(
          (idx) =>
            rows[idx].rValuesByYear?.[y] !== undefined &&
            rows[idx].rValuesByYear?.[y] !== null &&
            rows[idx].rValuesByYear?.[y] !== 0
        );
        if (activeIndices.length === 0) continue;

        const currentSum = activeIndices.reduce(
          (acc, idx) => acc + rows[idx].rValuesByYear[y],
          0
        );
        const currentErr = Math.abs(currentSum - subVal);

        if (currentErr <= 3) continue;

        let bestFlipIdx = -1;
        let bestNewErr = currentErr;

        for (const idx of activeIndices) {
          const val = rows[idx].rValuesByYear[y];
          if (isFlipForbidden(rows[idx], val, sIdx)) continue;
          const flippedSum = currentSum - 2 * val;
          const flippedErr = Math.abs(flippedSum - subVal);

          if (flippedErr <= 3 && flippedErr < bestNewErr) {
            bestNewErr = flippedErr;
            bestFlipIdx = idx;
          }
        }

        if (bestFlipIdx !== -1) {
          rows[bestFlipIdx].rValuesByYear[y] = -rows[bestFlipIdx].rValuesByYear[y];
          continue;
        }

        // 영업활동현금흐름에서 처분이익 등 비현금 손익 2개 행이 동시에 반대 부호로 표기된 경우 보정
        if (stmtType === 'CF' && prevTotalIdx === -1) {
          const gainLossCandidates = activeIndices.filter((idx) => {
            if (idx === compStartIdx) return false;
            const r = rows[idx];
            if (isFlipForbidden(r, r.rValuesByYear[y], sIdx)) return false;
            return (
              /Gains?Loss(?:es)?|DeconsolidationGain|RealizedGain|UnrealizedGain/i.test(r.tag || '') ||
              /\bgains?\b|\bloss(?:es)?\b|처분손익|평가손익|투자손익/i.test(
                `${r.contextualLabel || ''} ${r.exactKor || ''}`
              )
            );
          });

          let bestPair = null;
          let bestPairErr = currentErr;
          for (let a = 0; a < gainLossCandidates.length; a++) {
            for (let b = a + 1; b < gainLossCandidates.length; b++) {
              const idxA = gainLossCandidates[a];
              const idxB = gainLossCandidates[b];
              const pairSum =
                currentSum - 2 * rows[idxA].rValuesByYear[y] - 2 * rows[idxB].rValuesByYear[y];
              const pairErr = Math.abs(pairSum - subVal);
              if (pairErr <= 3 && pairErr < bestPairErr) {
                bestPairErr = pairErr;
                bestPair = [idxA, idxB];
              }
            }
          }

          if (bestPair) {
            rows[bestPair[0]].rValuesByYear[y] = -rows[bestPair[0]].rValuesByYear[y];
            rows[bestPair[1]].rValuesByYear[y] = -rows[bestPair[1]].rValuesByYear[y];
          }
        }
      }
    }

    prevTotalIdx = sIdx;
  }
}

// 13. 동일 XBRL 태그·동일 연도 1:1 전수 대조 스케일 확정 (수치 로직 100% 동일)
function determineExactTableScaleFromFacts(rows, colYears, factsObj, accessionNumber, headerDefaultScale) {
  const usGaap = factsObj?.['us-gaap'] || {};
  const votes = { 1: 0, 0.001: 0, 0.000001: 0 };

  for (const r of rows) {
    if (!r.isMainRow || r.rowType !== 'USD' || !r.tag) continue;
    const concept = factsObj?.[r.prefix]?.[r.tag] || usGaap[r.tag];
    const usdRecords = concept?.units?.USD;
    if (!Array.isArray(usdRecords)) continue;

    for (const yr of colYears) {
      const rawCellVal = r.rawNumbersByYear?.[yr];
      if (!rawCellVal || Math.abs(rawCellVal) < 1) continue;

      let matchedRecord = usdRecords.find((rec) => {
        if (rec.accn !== accessionNumber) return false;
        let endYr = rec.end ? parseInt(rec.end.slice(0, 4), 10) : rec.fy;
        if (rec.end && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(rec.end)) endYr -= 1;
        return endYr === yr;
      });

      if (!matchedRecord) {
        matchedRecord = usdRecords.find((rec) => {
          if (rec.form !== '10-K' && rec.form !== '10-K/A') return false;
          const isEarlyFye = Boolean(
            rec.end && /^20\d{2}-(?:01-(?:1[6-9]|2\d|3[01])|0[2-5]-\d{2})$/.test(rec.end)
          );
          if (!isEarlyFye && rec.frame && new RegExp(`^CY${yr}(?:Q4)?I?$`).test(rec.frame)) return true;
          let endYr = rec.end ? parseInt(rec.end.slice(0, 4), 10) : rec.fy;
          if (rec.end && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(rec.end)) endYr -= 1;
          return endYr === yr;
        });
      }

      if (!matchedRecord || !matchedRecord.val) continue;

      const targetMillions = Math.abs(matchedRecord.val) / 1e6;
      if (targetMillions < 0.5) continue;

      const rawAbs = Math.abs(rawCellVal);
      const ratio = rawAbs / targetMillions;

      if (ratio > 0.4 && ratio < 2.5) {
        votes[1] += 1;
      } else if (ratio > 400 && ratio < 2500) {
        votes[0.001] += 1;
      } else if (ratio > 400000 && ratio < 2500000) {
        votes[0.000001] += 1;
      }
    }
  }

  const maxVote = Math.max(votes[1], votes[0.001], votes[0.000001]);
  if (maxVote === 0) return headerDefaultScale;
  if (votes[1] === maxVote) return 1;
  if (votes[0.001] === maxVote) return 0.001;
  return 0.000001;
}

function normalizeCorpCoreName(name = '') {
  return String(name || '')
    .toUpperCase()
    .replace(/[/.,'’"&-]+/g, ' ')
    .replace(
      /\b(?:INCORPORATED|INC|CORPORATION|CORP|HOLDINGS?|HOLDCO|COMPANY|CO|LIMITED|LTD|LLC|PLC|LP|NV|SA|THE|OF|DELAWARE|NEW)\b/g,
      ' '
    )
    .replace(/\s+/g, '')
    .trim();
}

async function extract10KsFromSubmissionsData(subData, headers, maxCount = 12) {
  const accMap = new Map();
  const collectFromObj = (obj) => {
    if (!obj || !Array.isArray(obj.form)) return;
    for (let i = 0; i < obj.form.length; i++) {
      if (obj.form[i] === '10-K') {
        const accn = obj.accessionNumber?.[i];
        const filed = obj.filingDate?.[i] || '';
        const repDate = obj.reportDate?.[i] || '';
        let repYear = repDate
          ? parseInt(repDate.slice(0, 4), 10)
          : parseInt(filed.slice(0, 4), 10) - 1;
        if (repDate && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(repDate)) repYear -= 1;
        if (accn && !accMap.has(accn)) {
          accMap.set(accn, { accn, filed, reportYear: repYear });
        }
      }
    }
  };

  collectFromObj(subData?.filings?.recent);

  const archiveFiles = subData?.filings?.files || [];
  for (let f = 0; f < archiveFiles.length && accMap.size < maxCount; f++) {
    const fileName = archiveFiles[f]?.name;
    if (!fileName) continue;
    try {
      const archUrl = makeUrl(['data', 'sec', 'gov'], `submissions/${fileName}`);
      const archRes = await fetch(archUrl, { headers });
      if (archRes.ok) {
        const archData = await archRes.json();
        collectFromObj(archData);
      }
    } catch (e) {}
  }

  return Array.from(accMap.values()).sort((a, b) => (b.filed || '').localeCompare(a.filed || ''));
}

async function findVerifiedPredecessor10KCik({
  targetCik,
  targetEntityName = '',
  targetTicker = '',
  targetSubData = null,
  headers,
  requiredMaxReportYear = null,
}) {
  try {
    const rawName = targetSubData?.name || targetEntityName || '';
    const targetCore = normalizeCorpCoreName(rawName);
    if (targetCore.length < 4) return null;

    const searchPhrase = String(rawName)
      .replace(/([a-z])([A-Z])/g, '$1 $2')
      .replace(/[/.,'’"&-]+/g, ' ')
      .replace(
        /\b(?:INCORPORATED|INC|CORPORATION|CORP|HOLDINGS?|HOLDCO|COMPANY|CO|LIMITED|LTD|LLC|PLC|THE)\b/gi,
        ' '
      )
      .replace(/\s+/g, ' ')
      .trim();
    if (!searchPhrase || searchPhrase.length < 3) return null;

    const endDt = requiredMaxReportYear
      ? `${requiredMaxReportYear + 1}-12-31`
      : new Date().toISOString().slice(0, 10);
    const startDt = requiredMaxReportYear
      ? `${requiredMaxReportYear - 3}-01-01`
      : `${new Date().getFullYear() - 5}-01-01`;

    const queryUrls = [
      makeUrl(
        ['efts', 'sec', 'gov'],
        `LATEST/search-index?entityName=${encodeURIComponent(searchPhrase)}&forms=10-K&dateRange=custom&startdt=${startDt}&enddt=${endDt}`
      ),
      makeUrl(
        ['efts', 'sec', 'gov'],
        `LATEST/search-index?q=${encodeURIComponent(`"${searchPhrase}"`)}&forms=10-K&dateRange=custom&startdt=${startDt}&enddt=${endDt}`
      ),
    ];

    const testedCiks = new Set([targetCik]);
    const targetSic = targetSubData?.sic ? String(targetSubData.sic).trim() : '';

    for (const qUrl of queryUrls) {
      let eftsData = null;
      try {
        const res = await fetch(qUrl, { headers });
        if (!res.ok) continue;
        eftsData = await res.json();
      } catch {
        continue;
      }

      const hits = eftsData?.hits?.hits || [];
      const cikCounts = new Map();
      for (const h of hits) {
        if (h?._source?.form && h._source.form !== '10-K') continue;
        const ciks = Array.isArray(h?._source?.ciks) ? h._source.ciks : [];
        for (const c of ciks) {
          const padded = String(c).padStart(10, '0');
          if (testedCiks.has(padded)) continue;
          cikCounts.set(padded, (cikCounts.get(padded) || 0) + 1);
        }
      }

      const rankedCiks = Array.from(cikCounts.entries())
        .sort((a, b) => b[1] - a[1])
        .slice(0, 8)
        .map(([c]) => c);

      for (const candCik of rankedCiks) {
        testedCiks.add(candCik);
        try {
          const subUrl = makeUrl(['data', 'sec', 'gov'], `submissions/CIK${candCik}.json`);
          const subRes = await fetch(subUrl, { headers });
          if (!subRes.ok) continue;
          const candSub = await subRes.json();

          // 1) SIC 코드 일치 확인 (양측 존재 시)
          const candSic = candSub?.sic ? String(candSub.sic).trim() : '';
          if (targetSic && candSic && targetSic !== candSic) continue;

          // 2) 타 상장사 활성 티커 충돌 방지
          const candTickers = Array.isArray(candSub?.tickers) ? candSub.tickers : [];
          if (
            candTickers.length > 0 &&
            targetTicker &&
            !candTickers.some((t) => String(t).toUpperCase() === String(targetTicker).toUpperCase())
          ) {
            continue;
          }

          // 3) 현재명 또는 과거 사명(formerNames)의 정규화 핵심 법인명 완전 일치
          const candNames = [
            candSub?.name,
            ...((candSub?.formerNames || []).map((f) => f?.name) || []),
          ].filter(Boolean);
          const nameMatched = candNames.some((n) => normalizeCorpCoreName(n) === targetCore);
          if (!nameMatched) continue;

          // 4) 실제 10-K 공시 보유 확인
          const filings10K = await extract10KsFromSubmissionsData(candSub, headers, 12);
          if (filings10K.length === 0) continue;
          if (
            requiredMaxReportYear !== null &&
            !filings10K.some((f) => f.reportYear <= requiredMaxReportYear)
          ) {
            continue;
          }

          return { cik: candCik, subData: candSub, filings10K };
        } catch {
          continue;
        }
      }
    }
  } catch (e) {}
  return null;
}

// 14. 최근 10개년 10-K 공시 목록 추출 (수치 로직 100% 동일)
async function discoverAll10KFilings(cik, factsObj, headers, preloadedSubData = null) {
  const accMap = new Map();

  const usGaap = factsObj?.['us-gaap'] || {};
  const dei = factsObj?.['dei'] || {};
  const scanTags = [
    'Assets',
    'Liabilities',
    'StockholdersEquity',
    'NetIncomeLoss',
    'Revenues',
    'RevenueFromContractWithCustomerExcludingAssessedTax',
    'SalesRevenueNet',
    'InterestIncomeExpenseNet',
    'CashAndCashEquivalentsAtCarryingValue',
    'EntityCommonStockSharesOutstanding',
  ];

  for (const tag of scanTags) {
    const concept = usGaap[tag] || dei[tag];
    if (!concept || !concept.units) continue;
    for (const records of Object.values(concept.units)) {
      if (!Array.isArray(records)) continue;
      for (const r of records) {
        if (r.form === '10-K' && r.accn && r.filed) {
          const filedYear = parseInt(r.filed.slice(0, 4), 10);
          let endYear = r.end ? parseInt(r.end.slice(0, 4), 10) : filedYear - 1;
          if (r.end && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(r.end)) endYear -= 1;
          const existing = accMap.get(r.accn);
          if (!existing || endYear > existing.reportYear) {
            accMap.set(r.accn, {
              accn: r.accn,
              filed: r.filed,
              reportYear: endYear,
            });
          }
        }
      }
    }
  }

  try {
    let subData = preloadedSubData;
    if (!subData) {
      const subUrl = makeUrl(['data', 'sec', 'gov'], `submissions/CIK${cik}.json`);
      const subRes = await fetch(subUrl, { headers });
      if (subRes.ok) subData = await subRes.json();
    }
    if (subData) {
      const sub10Ks = await extract10KsFromSubmissionsData(subData, headers, 11);
      for (const item of sub10Ks) {
        if (!accMap.has(item.accn)) accMap.set(item.accn, item);
      }
    }
  } catch (e) {}

  const allFilings = Array.from(accMap.values()).sort((a, b) =>
    (b.filed || '').localeCompare(a.filed || '')
  );

  const uniqueByYear = [];
  const seenFiledYear = new Set();
  for (const item of allFilings) {
    const fYr = item.filed ? item.filed.slice(0, 4) : String(item.reportYear);
    if (seenFiledYear.has(fYr)) continue;
    seenFiledYear.add(fYr);
    uniqueByYear.push(item);
    if (uniqueByYear.length >= 10) break;
  }

  return uniqueByYear;
}

// 15. 단일 10-K 공시의 R*.htm 추출 (MSA Safety 버전과 100% 동일하게 고정!)
async function fetchSingle10KRStatements(cikUnpadded, accessionNumber, factsObj, headers) {
  try {
    const accNoDashes = accessionNumber.replace(/-/g, '');
    const baseArchivePath = `Archives/edgar/data/${cikUnpadded}/${accNoDashes}`;

    const summaryUrl = makeUrl(['www', 'sec', 'gov'], `${baseArchivePath}/FilingSummary.xml`);
    const summaryRes = await fetch(summaryUrl, { headers });
    if (!summaryRes.ok) return null;

    const summaryXml = await summaryRes.text();
    const reportBlocks = summaryXml.match(/<Report[\s\S]*?<\/Report>/gi) || [];

    const statementFiles = { IS: null, BS: null, CF: null };
    let fallbackComprehensiveIS = null;

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
        if (
          !fallbackComprehensiveIS &&
          !/Tax\s*Effect|Reclassification/i.test(shortName) &&
          /Comprehensive\s+(?:Income|Loss|Earnings)/i.test(shortName)
        ) {
          fallbackComprehensiveIS = htmlFile;
        }
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
    if (!statementFiles.IS && fallbackComprehensiveIS) {
      statementFiles.IS = fallbackComprehensiveIS;
    }

    const parseRHtmlTable = async (htmlFile, stmtType) => {
      if (!htmlFile) return [];
      const rUrl = makeUrl(['www', 'sec', 'gov'], `${baseArchivePath}/${htmlFile}`);
      const rRes = await fetch(rUrl, { headers });
      if (!rRes.ok) return [];

      const html = await rRes.text();

      const xbrlAuthMap = parseEmbeddedXbrlAuthRefs(html);

      const headerHtmlChunk = html.split(/<td[^>]*class=["']nump?["']/i)[0] || html.slice(0, 4500);
      const headerMetaText = headerHtmlChunk
        .replace(/<[^>]+>/g, ' ')
        .replace(/&#160;|&nbsp;/gi, ' ')
        .replace(/\s+/g, ' ');

      const dollarMeta = headerMetaText.replace(/shares\s+in\s+(?:thousands|millions|billions)/gi, ' ');

      let headerDefaultScale = 1;
      if (/\bin\s+thousands\b/i.test(dollarMeta)) {
        headerDefaultScale = 0.001;
      }

      let shareScaleToMillions = headerDefaultScale;
      if (/shares\s+in\s+thousands/i.test(headerMetaText)) {
        shareScaleToMillions = 0.001;
      } else if (/shares\s+in\s+millions/i.test(headerMetaText)) {
        shareScaleToMillions = 1;
      } else if (/except\s+(?:share|number\s+of\s+shares)/i.test(headerMetaText)) {
        shareScaleToMillions = 0.000001;
      }

      const trBlocks = html.match(/<tr[\s\S]*?<\/tr>/gi) || [];

      const headerRows = [];
      let durationBySlot = [];
      for (const tr of trBlocks) {
        if (/class=["']nump?["']/i.test(tr)) break;
        if (!/<th[\s>]/i.test(tr)) continue;
        const thCells = tr.match(/<th[\s\S]*?<\/th>/gi) || [];
        const extractedCells = [];
        const slotDurations = [];
        let hasDurationCell = false;
        let slot = 0;

        for (const th of thCells) {
          if (/class=["']tl["']/i.test(th)) continue;
          const colspanMatch = th.match(/colspan=["']?(\d+)["']?/i);
          const colspan = colspanMatch ? Math.max(1, parseInt(colspanMatch[1], 10)) : 1;
          const text = th
            .replace(/<[^>]+>/g, ' ')
            .replace(/&#160;|&nbsp;/gi, ' ')
            .replace(/\s+/g, ' ')
            .trim();
          const rowStartSlot = slot;
          slot += colspan;
          if (/^\d+\s*(?:months?|weeks?)\s*ended$/i.test(text)) {
            hasDurationCell = true;
            for (let k = 0; k < colspan; k++) slotDurations[rowStartSlot + k] = text;
            continue;
          }
          const isFootnoteHeader = /^\[\d+\](?:\s*,\s*\[\d+\])*$/.test(text);
          extractedCells.push({ text, slot: rowStartSlot, colspan, isFootnoteHeader });
        }
        if (hasDurationCell && extractedCells.every((c) => !c.text || c.isFootnoteHeader)) {
          durationBySlot = slotDurations;
        } else if (extractedCells.length > 0) {
          headerRows.push(extractedCells);
        }
      }

      const parseHeaderMonth = (txt) => {
        const m = String(txt || '').match(
          /\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:t|tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\b/i
        );
        if (!m) return null;
        const k = m[1].slice(0, 3).toLowerCase();
        return (
          {
            jan: 1,
            feb: 2,
            mar: 3,
            apr: 4,
            may: 5,
            jun: 6,
            jul: 7,
            aug: 8,
            sep: 9,
            oct: 10,
            nov: 11,
            dec: 12,
          }[k] || null
        );
      };

      const isAnnualDur = (d) => /^(?:12\s*months?|5[23]\s*weeks?)\s*ended$/i.test(d || '');
      let colEntries = [];
      const yearRowIdx = headerRows.findLastIndex((r) =>
        r.some((c) => !c.isFootnoteHeader && /\b(20\d{2})\b/.test(c.text))
      );
      if (yearRowIdx !== -1) {
        const yearDataCells = headerRows[yearRowIdx].filter((c) => !c.isFootnoteHeader);
        const prevDataCells =
          yearRowIdx > 0 ? headerRows[yearRowIdx - 1].filter((c) => !c.isFootnoteHeader) : null;
        const canUsePrevRow =
          prevDataCells &&
          prevDataCells.length > 0 &&
          !prevDataCells.every((p) => p.colspan >= 4) &&
          (prevDataCells.length === yearDataCells.length || prevDataCells.length >= 2);
        const yearCellDurs = yearDataCells.map((c) => durationBySlot[c.slot] || '');
        const hasMixedDurations =
          stmtType !== 'BS' &&
          yearDataCells.every((c, idx) => !/\b(20\d{2})\b/.test(c.text) || Boolean(yearCellDurs[idx])) &&
          yearCellDurs.some(isAnnualDur) &&
          yearCellDurs.some((d) => d && !isAnnualDur(d));

        for (let c = 0; c < yearDataCells.length; c++) {
          const cell = yearDataCells[c];
          if (hasMixedDurations && yearCellDurs[c] && !isAnnualDur(yearCellDurs[c])) continue;
          let prevText = '';
          if (canUsePrevRow) {
            if (prevDataCells.length === yearDataCells.length) {
              prevText = prevDataCells[c].text;
            } else {
              const slotMatch = prevDataCells.find(
                (p) =>
                  p.slot <= cell.slot &&
                  cell.slot < p.slot + p.colspan &&
                  p.colspan <= cell.colspan * 2
              );
              if (slotMatch) prevText = slotMatch.text;
            }
          }
          const fullDateText = prevText ? `${prevText} ${cell.text}` : cell.text;
          const yrMatch = fullDateText.match(/\b(20\d{2})\b/);
          if (yrMatch) {
            let yr = parseInt(yrMatch[1], 10);
            if (/\bJan(?:uary)?\.?\s*(?:0?[1-9]|1[0-5])\b/i.test(fullDateText)) {
              yr -= 1;
            }
            colEntries.push({ colIdx: c, yr, month: parseHeaderMonth(fullDateText) });
          }
        }
      }

      if (stmtType === 'BS' && colEntries.length >= 5) {
        const monthsByYr = new Map();
        for (const e of colEntries) {
          if (e.month === null) continue;
          if (!monthsByYr.has(e.yr)) monthsByYr.set(e.yr, new Set());
          monthsByYr.get(e.yr).add(e.month);
        }
        const hasMultiQuarterYr = [...monthsByYr.values()].some((s) => s.size >= 3);
        const fyeAnchorMonth = colEntries[0]?.month ?? null;
        if (hasMultiQuarterYr && fyeAnchorMonth !== null) {
          const filtered = colEntries.filter((e) => e.month === fyeAnchorMonth);
          if (filtered.length >= 2) colEntries = filtered;
        }
      }

      for (let i = 1; i < colEntries.length; i++) {
        if (colEntries[i].yr === colEntries[i - 1].yr) {
          colEntries[i].yr = colEntries[i - 1].yr - 1;
        }
      }
      const colYears = colEntries.map((e) => e.yr);

      let rows = [];
      let currentSection = '';
      let cfCashBalanceCount = 0;
      let inRestatementBlock = false;
      let inOtherEntityBlock = false;
      let afterMainLse = false;
      const postLseDebtRows = [];

      // 다차원 복합 지주사(ProductOrServiceAxis 부문 손익·BS·CF가 실제로 존재하는 기업) 전용 플래그
      // AAPL·CME 등 매출 세분화/주식종류만 있는 단일 기업에서는 반드시 false -> 기존 5종목 회귀 0건 보장
      const firstProdRhIdx = trBlocks.findIndex(
        (tr) => /class=["']rh["']/i.test(tr) && /ProductOrServiceAxis/i.test(tr)
      );
      const afterProdTrs = firstProdRhIdx !== -1 ? trBlocks.slice(firstProdRhIdx + 1) : [];
      const hasFullSegmentPnL =
        stmtType === 'IS' &&
        firstProdRhIdx !== -1 &&
        afterProdTrs.some((tr) => /defref_us-gaap_Revenues(?:[^a-zA-Z0-9]|$)/i.test(tr)) &&
        afterProdTrs.some((tr) => /defref_us-gaap_CostsAndExpenses(?:[^a-zA-Z0-9]|$)/i.test(tr));
      const hasSegmentBS =
        stmtType === 'BS' &&
        firstProdRhIdx !== -1 &&
        afterProdTrs.some((tr) => /defref_us-gaap_Assets(?:[^a-zA-Z0-9]|$)/i.test(tr));
      const hasSegmentCF =
        stmtType === 'CF' &&
        firstProdRhIdx !== -1 &&
        afterProdTrs.some((tr) =>
          /NetCashProvidedByUsedInFinancingActivities|ProceedsFromIssuanceOfLongTermDebt|RepaymentsOfLongTermDebt/i.test(
            tr
          )
        );
      const isMultiSegmentTable = hasFullSegmentPnL || hasSegmentBS || hasSegmentCF;
      const hasCombinedRRUtilMember =
        hasFullSegmentPnL &&
        trBlocks.some(
          (tr) =>
            /class=["']rh["']/i.test(tr) &&
            /Railroad,\s*Utilities\s+and\s+Energy/i.test(tr)
        );

      let currentDimAxis = '';
      let currentDimMember = '';
      let currentBlockAbstract = '';
      let nonSegAxis = '';
      let nonSegMember = '';
      let nonSegBlockAbstract = '';
      let currentBsAbstract = '';
      const postLseDimRows = [];
      const postLseMezzRows = [];

      const classifyBsSectionAbstract = (aTag, rLbl) => {
        if (/^AssetsCurrentAbstract$/i.test(aTag) || /^current\s+assets:?$/i.test(rLbl)) {
          return 'ASSETS_CURRENT';
        }
        if (
          /^(?:AssetsNoncurrentAbstract|OtherAssetsAbstract|IntangibleAssetsAbstract)$/i.test(aTag) ||
          /^(?:other\s+assets|non-?current\s+assets|long-?term\s+assets):?$/i.test(rLbl)
        ) {
          return 'ASSETS_NONCURRENT';
        }
        if (/^LiabilitiesCurrentAbstract$/i.test(aTag) || /^current\s+liabilities:?$/i.test(rLbl)) {
          return 'LIABILITIES_CURRENT';
        }
        if (
          /^(?:LiabilitiesNoncurrentAbstract|LongTermLiabilitiesAbstract)$/i.test(aTag) ||
          /^(?:long-?term\s+liabilities|non-?current\s+liabilities):?$/i.test(rLbl)
        ) {
          return 'LIABILITIES_NONCURRENT';
        }
        return '';
      };

      for (let trIdx = 0; trIdx < trBlocks.length; trIdx++) {
        const tr = trBlocks[trIdx];
        if (/<th[\s>]/i.test(tr)) continue;

        let rawLabel = '';
        const linkMatch = tr.match(/<a[^>]*onclick=["'][^"']*showAR[^"']*["'][^>]*>([\s\S]*?)<\/a>/i);
        if (linkMatch) {
          rawLabel = linkMatch[1];
        } else {
          const plMatch = tr.match(/<td[^>]*class=["']pl[^"']*["'][^>]*>([\s\S]*?)<\/td>/i);
          if (plMatch) rawLabel = plMatch[1];
        }
        rawLabel = rawLabel
          .replace(/<[^>]+>/g, '')
          .replace(/&amp;/g, '&')
          .replace(/&#160;|&nbsp;/gi, ' ')
          .replace(/\s+/g, ' ')
          .trim();

        const hasNumericCell = /class=["']nump?["']/i.test(tr);

        if (!hasNumericCell) {
          if (/RestatementAxis/i.test(tr) || /^restated\s+balance$/i.test(rawLabel)) {
            inRestatementBlock = true;
            inOtherEntityBlock = false;
            currentBsAbstract = '';
            if (isMultiSegmentTable) {
              currentDimAxis = '';
              currentDimMember = '';
              currentBlockAbstract = '';
            } else {
              nonSegAxis = '';
              nonSegMember = '';
              nonSegBlockAbstract = '';
            }
            continue;
          } else if (/class=["']rh["']/i.test(tr) || /Axis/i.test(tr)) {
            inRestatementBlock = false;
            currentBsAbstract = '';
            // 복합공시(Combined 10-K, 예: NEE + 자회사 FPL)의 자회사 LegalEntityAxis/OperatingSegments 블록 격리
            inOtherEntityBlock =
              rows.length > 0 &&
              /(?:LegalEntityAxis|ConsolidatedEntitiesAxis|EntityNameAxis|ConsolidationItemsAxis|StatementBusinessSegmentsAxis)/i.test(
                tr
              );
            if (inOtherEntityBlock) {
              continue;
            }
            if (isMultiSegmentTable && /class=["']rh["']/i.test(tr)) {
              if (/ProductOrServiceAxis/i.test(tr)) {
                currentDimAxis = 'ProductOrServiceAxis';
              } else if (/StatementClassOfStockAxis/i.test(tr)) {
                currentDimAxis = 'StatementClassOfStockAxis';
              } else {
                currentDimAxis = 'OTHER_AXIS';
              }
              currentDimMember = canonicalizeDimensionMember(currentDimAxis, rawLabel);
              currentBlockAbstract = '';
              currentSection = '';
              continue;
            }
            if (!isMultiSegmentTable && /class=["']rh["']/i.test(tr)) {
              if (/ProductOrServiceAxis/i.test(tr)) {
                nonSegAxis = 'ProductOrServiceAxis';
              } else if (/LongtermDebtTypeAxis/i.test(tr)) {
                nonSegAxis = 'LongtermDebtTypeAxis';
              } else if (/StatementClassOfStockAxis/i.test(tr)) {
                nonSegAxis = 'StatementClassOfStockAxis';
              } else {
                nonSegAxis = 'OTHER_AXIS';
              }
              nonSegMember = rawLabel.replace(/\[member\]/gi, '').replace(/\s+/g, ' ').trim();
              nonSegBlockAbstract = '';
              if (
                stmtType === 'CF' &&
                (nonSegAxis === 'LongtermDebtTypeAxis' ||
                  nonSegAxis === 'StatementClassOfStockAxis')
              ) {
                currentSection = '';
                continue;
              }
            }
          }

          if (isMultiSegmentTable) {
            const absDefMatch = tr.match(/defref_([a-zA-Z0-9-]+?)_([a-zA-Z0-9]+)/);
            const absTag = absDefMatch ? absDefMatch[2] : '';
            if (/^RevenuesAbstract$/i.test(absTag) || /^revenues:?$/i.test(rawLabel)) {
              currentBlockAbstract = 'REVENUES';
            } else if (
              /^CostsAndExpensesAbstract$/i.test(absTag) ||
              /^costs and expenses:?$/i.test(rawLabel)
            ) {
              currentBlockAbstract = 'COSTS';
            } else if (/^AssetsAbstract$/i.test(absTag) || /^assets:?$/i.test(rawLabel)) {
              currentBlockAbstract = 'ASSETS';
            } else if (
              /^LiabilitiesAbstract$/i.test(absTag) ||
              /^liabilities:?$/i.test(rawLabel)
            ) {
              currentBlockAbstract = 'LIABILITIES';
            } else if (/FinancingActivities.*Abstract/i.test(absTag)) {
              currentBlockAbstract = 'FINANCING';
            } else if (/InvestingActivities.*Abstract/i.test(absTag)) {
              currentBlockAbstract = 'INVESTING';
            } else if (
              /OperatingActivities.*Abstract|AdjustmentsToReconcile.*Abstract/i.test(absTag)
            ) {
              currentBlockAbstract = 'OPERATING';
            } else if (/CashAndCashEquivalents.*Abstract/i.test(absTag)) {
              currentBlockAbstract = 'CASH';
            } else if (/EarningsPerShareAbstract/i.test(absTag)) {
              currentBlockAbstract = 'EPS';
            }
          } else if (
            stmtType === 'CF' &&
            (nonSegAxis === 'LongtermDebtTypeAxis' ||
              nonSegAxis === 'StatementClassOfStockAxis' ||
              nonSegAxis === 'OTHER_AXIS')
          ) {
            const absDefMatch = tr.match(/defref_([a-zA-Z0-9-]+?)_([a-zA-Z0-9]+)/);
            const absTag = absDefMatch ? absDefMatch[2] : '';
            if (
              /OperatingActivities.*Abstract|AdjustmentsToReconcile.*Abstract|LossGainAbstract/i.test(
                absTag
              ) ||
              /operating activities|adjustments to net income|\(gain\)\s*loss\s+on/i.test(rawLabel)
            ) {
              nonSegBlockAbstract = 'OPERATING';
            } else if (
              /InvestingActivities.*Abstract/i.test(absTag) ||
              /investing activities/i.test(rawLabel)
            ) {
              nonSegBlockAbstract = 'INVESTING';
            } else if (
              /FinancingActivities.*Abstract/i.test(absTag) ||
              /financing activities/i.test(rawLabel)
            ) {
              nonSegBlockAbstract = 'FINANCING';
            } else if (/CashAndCashEquivalents.*Abstract/i.test(absTag)) {
              nonSegBlockAbstract = 'CASH';
            }
          } else if (stmtType === 'BS') {
            const absDefMatch = tr.match(/defref_([a-zA-Z0-9-]+?)_([a-zA-Z0-9]+)/);
            const absTag = absDefMatch ? absDefMatch[2] : '';
            const bsAbs = classifyBsSectionAbstract(absTag, rawLabel);
            if (nonSegAxis && nonSegAxis !== 'StatementClassOfStockAxis') {
              if (bsAbs) nonSegBlockAbstract = bsAbs;
            } else if (!nonSegAxis && bsAbs) {
              currentBsAbstract = bsAbs;
            }
          }

          if (
            rawLabel &&
            !/^(operating activities|investing activities|financing activities|assets|liabilities)$/i.test(
              rawLabel
            )
          ) {
            currentSection = rawLabel;
          }
          continue;
        }

        const fullDefMatch = tr.match(/(defref_([a-zA-Z0-9-]+?)_([a-zA-Z0-9]+)(?:_([^'"\s)]+))?)/);
        const fullDefKey = fullDefMatch ? fullDefMatch[1] : '';
        const prefix = fullDefMatch ? fullDefMatch[2] : 'custom';
        let tag = fullDefMatch ? fullDefMatch[3] : '';
        const dimKey = fullDefMatch && fullDefMatch[4] ? fullDefMatch[4] : '';

        if (isMultiSegmentTable) {
          if (currentDimAxis === 'OTHER_AXIS') continue;
          if (stmtType === 'BS' && currentDimAxis === 'StatementClassOfStockAxis') continue;
          if (hasSegmentBS && tag === 'CashCashEquivalentsAndUSTreasuryBills') continue;
          if (hasSegmentCF && currentDimAxis === 'ProductOrServiceAxis') {
            const isSegFinDebt =
              tag === 'ProceedsFromIssuanceOfLongTermDebt' || tag === 'RepaymentsOfLongTermDebt';
            if (currentBlockAbstract !== 'FINANCING' && !isSegFinDebt) continue;
            if (!isSegFinDebt && !/^CashFlow|Debt|Borrow/i.test(tag)) continue;
          }
        }

        if (inOtherEntityBlock) {
          continue;
        }

        if (isExcludedGrossOrAccumRow(stmtType, tag, rawLabel)) {
          continue;
        }

        let cleanLabel = stripDynamicParentheticals(rawLabel || tag);
        if (!cleanLabel) continue;

        const isDebtAxisCF =
          !isMultiSegmentTable && stmtType === 'CF' && nonSegAxis === 'LongtermDebtTypeAxis';
        if (
          isDebtAxisCF &&
          (nonSegBlockAbstract === 'FINANCING' || nonSegBlockAbstract === 'INVESTING') &&
          tag &&
          rows.some((r) => !r.nonSegAxis && r.tag === tag)
        ) {
          continue;
        }
        const isDebtOpRow = isDebtAxisCF && nonSegBlockAbstract === 'OPERATING';
        let debtDimKey = '';
        if (isDebtOpRow) {
          if (nonSegMember && !cleanLabel.toLowerCase().includes(nonSegMember.toLowerCase())) {
            cleanLabel = `${cleanLabel} (${nonSegMember})`;
          }
          debtDimKey = 'DEBT_' + (nonSegMember || 'OTHER').toUpperCase().replace(/[^A-Z0-9]+/g, '_');
        }
        const isClassFinRow =
          !isMultiSegmentTable &&
          stmtType === 'CF' &&
          nonSegAxis === 'StatementClassOfStockAxis' &&
          nonSegBlockAbstract === 'FINANCING';
        let classDimKey = '';
        if (isClassFinRow) {
          if (nonSegMember && !cleanLabel.toLowerCase().includes(nonSegMember.toLowerCase())) {
            cleanLabel = `${cleanLabel} (${nonSegMember})`;
          }
          classDimKey =
            'CLASS_' + (nonSegMember || 'OTHER').toUpperCase().replace(/[^A-Z0-9]+/g, '_');
        }

        let effectiveDimMember = isMultiSegmentTable ? currentDimMember : '';
        let forceNonTotalInSubSegment = false;

        if (
          hasFullSegmentPnL &&
          hasCombinedRRUtilMember &&
          currentDimAxis === 'ProductOrServiceAxis'
        ) {
          if (effectiveDimMember === 'SEG_SUB_RAILROAD') {
            effectiveDimMember = 'SEG_RAILROAD_UTILITIES_ENERGY';
            if (tag === 'Revenues') {
              tag = 'RailroadTransportationRevenues';
              cleanLabel = 'Railroad transportation revenues';
              forceNonTotalInSubSegment = true;
            } else if (tag === 'OperatingExpenses') {
              cleanLabel = 'Freight rail transportation expenses';
              forceNonTotalInSubSegment = true;
            }
          } else if (effectiveDimMember === 'SEG_SUB_UTILITIES_ENERGY') {
            effectiveDimMember = 'SEG_RAILROAD_UTILITIES_ENERGY';
            if (
              tag === 'RevenueFromContractWithCustomerExcludingAssessedTax' ||
              tag === 'SalesRevenueNet'
            ) {
              tag = 'UtilityAndEnergyOperatingRevenues';
              cleanLabel = 'Utility and energy operating revenues';
              forceNonTotalInSubSegment = true;
            } else if (tag === 'CostsOfServicesAndOperatingExpenses') {
              cleanLabel = 'Utilities and energy cost of sales and other expenses';
            }
          }
        }

        let thisCashIndex = 0;
        if (
          stmtType === 'CF' &&
          (!isMultiSegmentTable || currentDimAxis !== 'ProductOrServiceAxis') &&
          (tag === 'CashAndCashEquivalentsAtCarryingValue' ||
            tag === 'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents' ||
            /^cash and cash equivalents/i.test(cleanLabel))
        ) {
          cfCashBalanceCount++;
          thisCashIndex = cfCashBalanceCount;
        }

        const authMeta =
          xbrlAuthMap[fullDefKey] || xbrlAuthMap[`defref_${prefix}_${tag}`] || {};

        if (
          /percentItemType|pureItemType/i.test(authMeta.dataType || '') ||
          /Percentage$/i.test(tag || '') ||
          (stmtType === 'BS' &&
            (/perShareItemType|sharesItemType/i.test(authMeta.dataType || '') ||
              /(?:ParOrStatedValuePerShare|SharesAuthorized|SharesIssued|SharesOutstanding|TreasuryStock(?:Common|Preferred)?Shares)$/i.test(
                tag || ''
              )))
        ) {
          continue;
        }

        const rowType = detectRowType(tag, cleanLabel, currentSection, stmtType, authMeta.dataType);
        if (stmtType === 'BS' && rowType === 'SHARES') {
          continue;
        }
        const contextualLabel =
          isDebtOpRow || isClassFinRow
            ? cleanLabel
            : buildContextualLabel(
              cleanLabel,
              currentSection,
              rowType,
              stmtType,
              thisCashIndex,
              isMultiSegmentTable
            );

        const valueCells =
          tr.match(/<td[^>]*class=["'](?:nump|num|text)["'][^>]*>[\s\S]*?<\/td>/gi) || [];

        const rawNumbersByYear = {};
        let hasAnyNumber = false;

        for (const { colIdx, yr } of colEntries) {
          if (colIdx >= valueCells.length) continue;
          const td = valueCells[colIdx];
          const rawText = td
            .replace(/<[^>]+>/g, '')
            .replace(/&#160;|&nbsp;|\$|,/gi, '')
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
            hasAnyNumber = true;
          }
        }

        if (!hasAnyNumber) continue;

        if (inRestatementBlock) {
          if (tag) {
            const existingMain = rows.find((r) => r.tag === tag && !r.dimKey);
            if (existingMain) {
              for (const [yr, signedNum] of Object.entries(rawNumbersByYear)) {
                if (existingMain.rawNumbersByYear[yr] === undefined) {
                  existingMain.rawNumbersByYear[yr] = signedNum;
                }
              }
            }
          }
          continue;
        }

        let isTotalRow = isTrueSubtotalRow(cleanLabel, tag, stmtType, rowType);
        if (hasFullSegmentPnL && currentDimAxis === 'ProductOrServiceAxis') {
          if (forceNonTotalInSubSegment) {
            isTotalRow = false;
          } else {
            isTotalRow =
              rowType === 'USD' &&
              (tag === 'Revenues' || tag === 'CostsAndExpenses' || /^total\s+/i.test(cleanLabel));
          }
        } else if (hasSegmentBS && currentDimAxis === 'ProductOrServiceAxis') {
          isTotalRow =
            rowType === 'USD' &&
            (tag === 'Assets' ||
              tag === 'Liabilities' ||
              tag === 'LiabilitiesExcludingIncomeTaxesPrincipallyDeferred' ||
              /^total\s+|,\s*total$/i.test(cleanLabel));
        }

        if (isTotalRow && rowType === 'USD') {
          currentSection = '';
          currentBsAbstract = '';
        }

        const isCashBalanceInCF = stmtType === 'CF' && thisCashIndex > 0;
        const resolvedDimKey = isCashBalanceInCF
          ? `CF_CASH_${thisCashIndex}`
          : debtDimKey || classDimKey || effectiveDimMember || dimKey;

        const parsedRowObj = {
          prefix,
          tag,
          dimAxis: isMultiSegmentTable ? currentDimAxis : '',
          dimMember: effectiveDimMember,
          blockAbstract: isMultiSegmentTable ? currentBlockAbstract : '',
          nonSegAxis: !isMultiSegmentTable ? nonSegAxis : '',
          nonSegMember: !isMultiSegmentTable ? nonSegMember : '',
          nonSegBlockAbstract: !isMultiSegmentTable ? nonSegBlockAbstract : '',
          bsSectionAbstract:
            stmtType === 'BS' && !isMultiSegmentTable && !nonSegAxis ? currentBsAbstract : '',
          dimKey: resolvedDimKey,
          isMainRow: Boolean(tag && !resolvedDimKey && !isCashBalanceInCF),
          rowType,
          isEps: rowType === 'EPS',
          companyLabel: cleanLabel,
          contextualLabel,
          normLabel: normalizeLabel(cleanLabel),
          normContextualLabel: normalizeLabel(contextualLabel),
          isTotalRow,
          rawNumbersByYear,
          rValuesByYear: {},
        };

        if (afterMainLse) {
          if (nonSegAxis === 'ProductOrServiceAxis' && /LongTermDebt/i.test(tag) && rowType === 'USD') {
            postLseDebtRows.push(parsedRowObj);
          } else if (
            stmtType === 'BS' &&
            rowType === 'USD' &&
            !isTotalRow &&
            isMezzanineOrTemporaryEquityRow(tag, cleanLabel)
          ) {
            postLseMezzRows.push(parsedRowObj);
          } else if (
            stmtType === 'BS' &&
            nonSegAxis &&
            nonSegAxis !== 'StatementClassOfStockAxis' &&
            rowType === 'USD' &&
            !isTotalRow
          ) {
            let inferredAbs = nonSegBlockAbstract;
            if (!inferredAbs) {
              if (/^AssetsOfDisposalGroup.*Current$/i.test(tag || '')) inferredAbs = 'ASSETS_CURRENT';
              else if (/^LiabilitiesOfDisposalGroup.*Current$/i.test(tag || '')) {
                inferredAbs = 'LIABILITIES_CURRENT';
              } else if (
                /^(?:DisposalGroup.*AssetsNoncurrent|AssetsOfDisposalGroup.*Noncurrent)$/i.test(
                  tag || ''
                )
              ) {
                inferredAbs = 'ASSETS_NONCURRENT';
              } else if (
                /^(?:DisposalGroup.*LiabilitiesNoncurrent|LiabilitiesOfDisposalGroup.*Noncurrent)$/i.test(
                  tag || ''
                )
              ) {
                inferredAbs = 'LIABILITIES_NONCURRENT';
              }
            }
            if (inferredAbs) {
              parsedRowObj.nonSegBlockAbstract = inferredAbs;
              postLseDimRows.push(parsedRowObj);
            }
          }
          continue;
        }

        rows.push(parsedRowObj);

        if (
          stmtType === 'BS' &&
          !hasSegmentBS &&
          (tag === 'LiabilitiesAndStockholdersEquity' ||
            /^total liabilities and (?:stockholders|shareholders)'? equity$|^total liabilities and equity$/i.test(
              cleanLabel
            ))
        ) {
          const remainingTrs = trBlocks.slice(trIdx + 1);
          const hasRemainingSegDebt =
            remainingTrs.some((t) => /ProductOrServiceAxis/i.test(t)) &&
            remainingTrs.some((t) => /LongTermDebt/i.test(t));
          const hasRemainingDimBsSection =
            remainingTrs.some((t) => /class=["']rh["']/i.test(t)) &&
            remainingTrs.some((t) =>
              /Assets(?:Current|Noncurrent)Abstract|OtherAssetsAbstract|IntangibleAssetsAbstract|Liabilities(?:Current|Noncurrent)Abstract|DisposalGroupIncludingDiscontinuedOperation|AssetsOfDisposalGroupIncludingDiscontinuedOperation|LiabilitiesOfDisposalGroupIncludingDiscontinuedOperation|TemporaryEquityCarryingAmount|RedeemableNoncontrollingInterest/i.test(
                t
              )
            );
          if (!hasRemainingSegDebt && !hasRemainingDimBsSection) {
            break;
          }
          afterMainLse = true;
        }
      }

      const verifiedDollarScale = determineExactTableScaleFromFacts(
        rows,
        colYears,
        factsObj,
        accessionNumber,
        headerDefaultScale
      );

      const cfTableHasNegatives =
        stmtType === 'CF' &&
        rows.some(
          (rr) =>
            rr.rowType === 'USD' &&
            Object.values(rr.rawNumbersByYear || {}).some((n) => typeof n === 'number' && n < 0)
        );

      for (const r of rows) {
        const polarity =
          stmtType === 'CF'
            ? classifyCashFlowRowPolarity(r.tag, r.contextualLabel, '')
            : 'NEUTRAL';
        const rawValsArr = Object.values(r.rawNumbersByYear || {}).filter(
          (n) => typeof n === 'number' && n !== 0
        );
        const rowHasMixedSignsInTable =
          stmtType === 'CF' &&
          cfTableHasNegatives &&
          rawValsArr.some((n) => n > 0) &&
          rawValsArr.some((n) => n < 0);

        for (const [yr, signedNum] of Object.entries(r.rawNumbersByYear)) {
          if (r.rowType === 'EPS') {
            r.rValuesByYear[yr] = Number(signedNum.toFixed(2));
          } else if (r.rowType === 'SHARES') {
            if (
              isMultiSegmentTable &&
              r.dimMember === 'CLASS_A' &&
              Math.abs(signedNum) >= 100000 &&
              Math.abs(signedNum) < 50000000
            ) {
              r.shareUnit = 'SHARES';
              r.rValuesByYear[yr] = Math.round(signedNum);
            } else {
              let scaledShares = null;
              if (r.tag && !r.dimMember && !r.dimKey) {
                const concept = factsObj?.[r.prefix]?.[r.tag] || factsObj?.['us-gaap']?.[r.tag];
                const shareUnitKey = concept?.units
                  ? Object.keys(concept.units).find(
                      (k) => k.toLowerCase().includes('share') && !k.includes('USD')
                    )
                  : null;
                const shareRecords = shareUnitKey ? concept.units[shareUnitKey] : null;
                if (Array.isArray(shareRecords)) {
                  const yrNum = Number(yr);
                  let matchedRec = shareRecords.find((rec) => {
                    if (rec.accn !== accessionNumber) return false;
                    if (rec.start && rec.end) {
                      const days = (new Date(rec.end) - new Date(rec.start)) / (1000 * 3600 * 24);
                      if (days < 250 || days > 400) return false;
                    }
                    let endYr = rec.end ? parseInt(rec.end.slice(0, 4), 10) : rec.fy;
                    if (rec.end && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(rec.end)) endYr -= 1;
                    return endYr === yrNum;
                  });
                  if (!matchedRec) {
                    matchedRec = shareRecords.find((rec) => {
                      if (rec.form !== '10-K' && rec.form !== '10-K/A') return false;
                      if (rec.start && rec.end) {
                        const days = (new Date(rec.end) - new Date(rec.start)) / (1000 * 3600 * 24);
                        if (days < 250 || days > 400) return false;
                      }
                      const isEarlyFye = Boolean(
                        rec.end && /^20\d{2}-(?:01-(?:1[6-9]|2\d|3[01])|0[2-5]-\d{2})$/.test(rec.end)
                      );
                      if (!isEarlyFye && rec.frame && new RegExp(`^CY${yrNum}(?:Q4)?I?$`).test(rec.frame)) {
                        return true;
                      }
                      let endYr = rec.end ? parseInt(rec.end.slice(0, 4), 10) : rec.fy;
                      if (rec.end && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(rec.end)) endYr -= 1;
                      return endYr === yrNum;
                    });
                  }
                  if (
                    matchedRec &&
                    typeof matchedRec.val === 'number' &&
                    Math.abs(matchedRec.val) > 0 &&
                    Math.abs(signedNum) > 0
                  ) {
                    const targetMillions = Math.abs(matchedRec.val) / 1e6;
                    if (targetMillions >= 0.05) {
                      const ratio = Math.abs(signedNum) / targetMillions;
                      if (ratio > 0.4 && ratio < 2.5) scaledShares = signedNum;
                      else if (ratio > 400 && ratio < 2500) scaledShares = signedNum * 0.001;
                      else if (ratio > 400000 && ratio < 2500000) scaledShares = signedNum * 0.000001;
                    }
                  }
                }
              }
              if (scaledShares === null) {
                scaledShares = signedNum * shareScaleToMillions;
                if (Math.abs(scaledShares) >= 40000 && Math.abs(scaledShares) < 50000000) {
                  scaledShares = scaledShares * 0.001;
                } else if (Math.abs(scaledShares) >= 50000000) {
                  scaledShares = scaledShares * 0.000001;
                }
              }
              r.shareUnit = 'MILLIONS';
              r.rValuesByYear[yr] = Math.round(scaledShares);
            }
          } else {
            let val = Math.round(signedNum * verifiedDollarScale);
            const isNetAcqInflow =
              stmtType === 'CF' &&
              cfTableHasNegatives &&
              r.tag === 'PaymentsToAcquireBusinessesNetOfCashAcquired' &&
              signedNum > 0;
            const isNetProceedsOutflow =
              stmtType === 'CF' &&
              cfTableHasNegatives &&
              polarity === 'INFLOW' &&
              signedNum < 0 &&
              (rowHasMixedSignsInTable ||
                /^(?:ProceedsFromDivestitureOfBusinessesNetOfCashDivested|ProceedsFromIssuanceOrSaleOfEquity)$|NetOfCashDivested/i.test(
                  r.tag
                ) ||
                (r.prefix && r.prefix !== 'us-gaap' && /^ProceedsFrom/i.test(r.tag || '')));
            if (isNetAcqInflow) {
              val = Math.abs(val);
              if (!r.netAcquisitionInflowYears) r.netAcquisitionInflowYears = new Set();
              r.netAcquisitionInflowYears.add(String(yr));
            } else if (isNetProceedsOutflow) {
              val = -Math.abs(val);
              if (!r.netProceedsOutflowYears) r.netProceedsOutflowYears = new Set();
              r.netProceedsOutflowYears.add(String(yr));
            } else if (val !== 0 && polarity === 'OUTFLOW') {
              val = -Math.abs(val);
            } else if (val !== 0 && polarity === 'INFLOW') {
              val = Math.abs(val);
            }
            r.rValuesByYear[yr] = val;
          }
        }
        delete r.rawNumbersByYear;
      }

      // BS 3개년 이상 컬럼 중 자본변동 롤포워드용 희소 컬럼(Assets & LSE 모두 공백인 과거 연도) 제거 (ASC 606 혼재 방지)
      if (stmtType === 'BS' && colYears.length >= 3) {
        const assetsRefRow = rows.find((r) => r.tag === 'Assets' && !r.dimKey);
        const lseRefRow = rows.find(
          (r) =>
            (r.tag === 'LiabilitiesAndStockholdersEquity' ||
              /^total liabilities and (?:stockholders|shareholders)'? equity$|^total liabilities and equity$/i.test(
                r.companyLabel || ''
              )) &&
            !r.dimKey
        );
        if (assetsRefRow || lseRefRow) {
          const fullBsYears = new Set(
            colYears.filter(
              (yr) =>
                assetsRefRow?.rValuesByYear?.[yr] !== undefined ||
                lseRefRow?.rValuesByYear?.[yr] !== undefined
            )
          );
          if (fullBsYears.size >= 2 && fullBsYears.size < colYears.length) {
            for (const yr of colYears) {
              if (fullBsYears.has(yr)) continue;
              for (const r of rows) delete r.rValuesByYear[yr];
              for (const pr of postLseDimRows) delete pr.rawNumbersByYear?.[yr];
              for (const pr of postLseMezzRows) delete pr.rawNumbersByYear?.[yr];
              for (const pr of postLseDebtRows) delete pr.rawNumbersByYear?.[yr];
            }
          }
        }
      }

      const latestColYr = colYears[0];

      // BS 총부채·자본합계(LSE) 이후 ProductOrServiceAxis로 분리 공시된 유동·비유동 장기차입금 복원 (CAT 등)
      if (stmtType === 'BS' && postLseDebtRows.length > 0 && latestColYr !== undefined) {
        for (const pr of postLseDebtRows) {
          for (const [yr, signedNum] of Object.entries(pr.rawNumbersByYear)) {
            pr.rValuesByYear[yr] = Math.round(signedNum * verifiedDollarScale);
          }
          delete pr.rawNumbersByYear;
        }

        const groupedByTag = new Map();
        for (let pi = 0; pi < postLseDebtRows.length; pi++) {
          const pr = postLseDebtRows[pi];
          if (!pr.tag) continue;
          if (!groupedByTag.has(pr.tag)) {
            groupedByTag.set(pr.tag, { sample: pr, byYear: {}, seenMembersByYear: {} });
          }
          const g = groupedByTag.get(pr.tag);
          const memKey = pr.nonSegMember || pr.dimKey || String(pi);
          for (const [yr, val] of Object.entries(pr.rValuesByYear)) {
            if (typeof val !== 'number') continue;
            if (!g.seenMembersByYear[yr]) g.seenMembersByYear[yr] = new Set();
            if (g.seenMembersByYear[yr].has(memKey)) continue;
            g.seenMembersByYear[yr].add(memKey);
            g.byYear[yr] = (g.byYear[yr] || 0) + val;
          }
        }

        const currentDebtTags = ['LongTermDebtAndCapitalLeaseObligationsCurrent', 'LongTermDebtCurrent'];
        const liabCurrIdx = rows.findIndex((r) => r.tag === 'LiabilitiesCurrent' && !r.dimKey);
        if (
          liabCurrIdx !== -1 &&
          !rows.some((r) => !r.dimKey && currentDebtTags.includes(r.tag))
        ) {
          for (const cTag of currentDebtTags) {
            const g = groupedByTag.get(cTag);
            if (!g || g.byYear[latestColYr] === undefined) continue;
            let clStart = 0;
            for (let i = liabCurrIdx - 1; i >= 0; i--) {
              if (rows[i].isTotalRow && rows[i].rowType === 'USD') {
                clStart = i + 1;
                break;
              }
            }
            let compSum = 0;
            for (let i = clStart; i < liabCurrIdx; i++) {
              if (rows[i].rowType === 'USD' && !rows[i].isTotalRow && !rows[i].dimKey) {
                compSum += rows[i].rValuesByYear[latestColYr] || 0;
              }
            }
            const clVal = rows[liabCurrIdx].rValuesByYear[latestColYr];
            if (clVal !== undefined) {
              const gapBefore = Math.abs(clVal - compSum);
              const gapAfter = Math.abs(clVal - (compSum + g.byYear[latestColYr]));
              if (gapBefore > 3 && gapAfter <= 3) {
                const s = g.sample;
                rows.splice(liabCurrIdx, 0, {
                  prefix: s.prefix,
                  tag: s.tag,
                  dimAxis: '',
                  dimMember: '',
                  blockAbstract: '',
                  nonSegAxis: '',
                  nonSegMember: '',
                  nonSegBlockAbstract: '',
                  dimKey: '',
                  isMainRow: true,
                  rowType: 'USD',
                  isEps: false,
                  companyLabel: s.companyLabel,
                  contextualLabel: s.companyLabel,
                  normLabel: s.normLabel,
                  normContextualLabel: s.normLabel,
                  isTotalRow: false,
                  rValuesByYear: g.byYear,
                });
                break;
              }
            }
          }
        }

        const nonCurrDebtTags = [
          'LongTermDebtAndCapitalLeaseObligations',
          'LongTermDebtNoncurrent',
          'LongTermDebt',
        ];
        const updatedLiabCurrIdx = rows.findIndex((r) => r.tag === 'LiabilitiesCurrent' && !r.dimKey);
        const updatedLiabIdx = rows.findIndex((r) => r.tag === 'Liabilities' && !r.dimKey);
        if (
          updatedLiabCurrIdx !== -1 &&
          updatedLiabIdx > updatedLiabCurrIdx &&
          !rows.some((r) => !r.dimKey && nonCurrDebtTags.includes(r.tag))
        ) {
          for (const ncTag of nonCurrDebtTags) {
            const g = groupedByTag.get(ncTag);
            if (!g || g.byYear[latestColYr] === undefined) continue;
            const clVal = rows[updatedLiabCurrIdx].rValuesByYear[latestColYr];
            const liabVal = rows[updatedLiabIdx].rValuesByYear[latestColYr];
            if (clVal === undefined || liabVal === undefined) continue;
            let nonCurrSum = 0;
            for (let i = updatedLiabCurrIdx + 1; i < updatedLiabIdx; i++) {
              if (rows[i].rowType === 'USD' && !rows[i].isTotalRow && !rows[i].dimKey) {
                nonCurrSum += rows[i].rValuesByYear[latestColYr] || 0;
              }
            }
            const gapBefore = Math.abs(liabVal - (clVal + nonCurrSum));
            const gapAfter = Math.abs(liabVal - (clVal + nonCurrSum + g.byYear[latestColYr]));
            if (gapBefore > 3 && gapAfter <= 3) {
              const s = g.sample;
              rows.splice(updatedLiabIdx, 0, {
                prefix: s.prefix,
                tag: s.tag,
                dimAxis: '',
                dimMember: '',
                blockAbstract: '',
                nonSegAxis: '',
                nonSegMember: '',
                nonSegBlockAbstract: '',
                dimKey: '',
                isMainRow: true,
                rowType: 'USD',
                isEps: false,
                companyLabel: s.companyLabel,
                contextualLabel: s.companyLabel,
                normLabel: s.normLabel,
                normContextualLabel: s.normLabel,
                isTotalRow: false,
                rValuesByYear: g.byYear,
              });
              break;
            }
          }
        }
      }

      if (stmtType === 'BS' && !hasSegmentBS && postLseDimRows.length > 0) {
        for (const pr of postLseDimRows) {
          for (const [yr, signedNum] of Object.entries(pr.rawNumbersByYear || {})) {
            pr.rValuesByYear[yr] = Math.round(signedNum * verifiedDollarScale);
          }
          delete pr.rawNumbersByYear;
        }
        const bsDimTargets = [
          { absKey: 'ASSETS_CURRENT', subTag: 'AssetsCurrent' },
          { absKey: 'ASSETS_NONCURRENT', subTag: 'AssetsNoncurrent' },
          { absKey: 'LIABILITIES_CURRENT', subTag: 'LiabilitiesCurrent' },
          { absKey: 'LIABILITIES_NONCURRENT', subTag: 'LiabilitiesNoncurrent' },
        ];
        for (const tgt of bsDimTargets) {
          const candRows = postLseDimRows.filter((r) => r.nonSegBlockAbstract === tgt.absKey);
          if (candRows.length === 0) continue;
          const subIdx = rows.findIndex(
            (r) => !r.dimKey && r.isTotalRow && r.rowType === 'USD' && r.tag === tgt.subTag
          );
          if (subIdx === -1) continue;
          let prevTotalIdx = -1;
          for (let i = subIdx - 1; i >= 0; i--) {
            if (rows[i].isTotalRow && rows[i].rowType === 'USD' && !rows[i].dimMember) {
              prevTotalIdx = i;
              break;
            }
          }
          const prevSubStart = prevTotalIdx + 1;
          const absStart = findTightenedBsSectionStart(rows, prevSubStart, subIdx, tgt.absKey);
          const candidateStarts = absStart > prevSubStart ? [absStart, prevSubStart] : [prevSubStart];

          let matchedStart = false;
          for (const compStart of candidateStarts) {
            let improvedYearCount = 0;
            let worsenedYearCount = 0;
            for (const yr of colYears) {
              const subVal = rows[subIdx].rValuesByYear?.[yr];
              if (subVal === undefined || subVal === null) continue;
              const addSum = candRows.reduce((acc, cr) => acc + (cr.rValuesByYear?.[yr] ?? 0), 0);
              if (addSum === 0) continue;
              let baseSum = 0;
              for (let i = compStart; i < subIdx; i++) {
                if (
                  rows[i].rowType !== 'USD' ||
                  rows[i].isTotalRow ||
                  isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].contextualLabel)
                ) {
                  continue;
                }
                baseSum += rows[i].rValuesByYear?.[yr] ?? 0;
              }
              const errBefore = Math.abs(baseSum - subVal);
              const errAfter = Math.abs(baseSum + addSum - subVal);
              if (errBefore > 3 && errAfter <= 3) improvedYearCount++;
              else if (errBefore <= 3 && errAfter > 3) worsenedYearCount++;
            }
            if (improvedYearCount > 0 && worsenedYearCount === 0) {
              matchedStart = true;
              break;
            }
          }
          if (matchedStart) {
            for (const cr of candRows) {
              cr.dimKey =
                'DIM_' + (cr.nonSegMember || 'OTHER').toUpperCase().replace(/[^A-Z0-9]+/g, '_');
              cr.isMainRow = false;
              cr.bsSectionAbstract = tgt.absKey;
            }
            rows.splice(subIdx, 0, ...candRows);
          }
        }
      }

      // BS 총부채·자본합계(LSE) 이후 클래스별/별도 메자닌(임시자본) 행 복원 (HAYW Class A 등)
      if (stmtType === 'BS' && !hasSegmentBS && postLseMezzRows.length > 0) {
        for (const pr of postLseMezzRows) {
          for (const [yr, signedNum] of Object.entries(pr.rawNumbersByYear || {})) {
            pr.rValuesByYear[yr] = Math.round(signedNum * verifiedDollarScale);
          }
          delete pr.rawNumbersByYear;
        }
        const nonZeroMezz = postLseMezzRows.filter((pr) =>
          Object.values(pr.rValuesByYear || {}).some((v) => v !== 0)
        );
        if (nonZeroMezz.length > 0) {
          const assetsRow = rows.find((r) => !r.dimKey && r.isTotalRow && r.tag === 'Assets');
          const liabRow = rows.find(
            (r) =>
              !r.dimKey &&
              r.isTotalRow &&
              ['Liabilities', 'LiabilitiesIncludingDeferredIncomeTaxes'].includes(r.tag)
          );
          const clRow = rows.find((r) => !r.dimKey && r.isTotalRow && r.tag === 'LiabilitiesCurrent');
          const nclRow = rows.find(
            (r) => !r.dimKey && r.isTotalRow && r.tag === 'LiabilitiesNoncurrent'
          );
          let eqIdx = -1;
          for (let i = rows.length - 1; i >= 0; i--) {
            if (
              !rows[i].dimKey &&
              rows[i].isTotalRow &&
              [
                'StockholdersEquity',
                'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
                'PartnersCapital',
                'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest',
              ].includes(rows[i].tag)
            ) {
              eqIdx = i;
              break;
            }
          }
          if (assetsRow && eqIdx !== -1 && (liabRow || (clRow && nclRow))) {
            const existMezz = rows.filter((r) =>
              isMezzanineOrTemporaryEquityRow(r.tag, r.contextualLabel)
            );
            let improved = 0;
            let worsened = 0;
            for (const yr of colYears) {
              const aVal = assetsRow.rValuesByYear?.[yr];
              const lVal =
                liabRow?.rValuesByYear?.[yr] ??
                (clRow?.rValuesByYear?.[yr] !== undefined && nclRow?.rValuesByYear?.[yr] !== undefined
                  ? clRow.rValuesByYear[yr] + nclRow.rValuesByYear[yr]
                  : undefined);
              const eqVal = rows[eqIdx].rValuesByYear?.[yr];
              if (aVal === undefined || lVal === undefined || eqVal === undefined) continue;
              const addM = nonZeroMezz.reduce((acc, mr) => acc + (mr.rValuesByYear?.[yr] ?? 0), 0);
              if (addM === 0) continue;
              const exM = existMezz.reduce((acc, mr) => acc + (mr.rValuesByYear?.[yr] ?? 0), 0);
              const errBefore = Math.abs(lVal + exM + eqVal - aVal);
              const errAfter = Math.abs(lVal + exM + addM + eqVal - aVal);
              if (errBefore > 3 && errAfter <= 3) improved++;
              else if (errBefore <= 3 && errAfter > 3) worsened++;
            }
            if (improved > 0 && worsened === 0) {
              const liabBoundaryIdx = rows.findIndex(
                (r) =>
                  !r.dimKey &&
                  r.isTotalRow &&
                  ['Liabilities', 'LiabilitiesIncludingDeferredIncomeTaxes'].includes(r.tag)
              );
              const nclBoundaryIdx = rows.findIndex(
                (r) => !r.dimKey && r.isTotalRow && r.tag === 'LiabilitiesNoncurrent'
              );
              const clBoundaryIdx = rows.findIndex(
                (r) => !r.dimKey && r.isTotalRow && r.tag === 'LiabilitiesCurrent'
              );
              const searchFrom = Math.max(liabBoundaryIdx, nclBoundaryIdx, clBoundaryIdx);
              let insertMezzIdx = eqIdx;
              for (let i = searchFrom + 1; i < eqIdx; i++) {
                if (rows[i].rowType !== 'USD') continue;
                if (
                  isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].contextualLabel) ||
                  /CommonStock|PreferredStock|AdditionalPaidInCapital|RetainedEarnings|TreasuryStock|AccumulatedOtherComprehensive|StockholdersEquity|PartnersCapital|MinorityInterest/i.test(
                    rows[i].tag || ''
                  ) ||
                  /^(?:common stock|preferred stock|additional paid-?in|retained earnings|accumulated deficit|accumulated other comprehensive|treasury stock)/i.test(
                    rows[i].companyLabel || ''
                  )
                ) {
                  insertMezzIdx = i;
                  break;
                }
              }
              for (const mr of nonZeroMezz) {
                if (nonZeroMezz.length === 1) {
                  mr.dimKey = '';
                  mr.isMainRow = true;
                  mr.nonSegAxis = '';
                  mr.nonSegMember = '';
                } else {
                  mr.dimKey =
                    'DIM_' + (mr.nonSegMember || 'MEZZ').toUpperCase().replace(/[^A-Z0-9]+/g, '_');
                  mr.isMainRow = false;
                }
              }
              rows.splice(insertMezzIdx, 0, ...nonZeroMezz);
            }
          }
        }
      }

      // BS 재고자산 세부 구성항목(원재료/재공품/완제품) + 합계(InventoryNet) 동시 공시 중복 제거 (PG, SHW 2016 등)
      if (stmtType === 'BS' && !hasSegmentBS && rows.length > 2) {
        const invNetIdx = rows.findIndex((r) => r.tag === 'InventoryNet' && !r.dimKey);
        if (invNetIdx >= 2) {
          const subIndices = [];
          for (let k = invNetIdx - 1; k >= 0; k--) {
            if (rows[k].dimKey || rows[k].isTotalRow || rows[k].rowType !== 'USD') break;
            if (
              /^Inventory(?:RawMaterials|WorkInProcess|FinishedGoods|Parts|Supplies|LIFO|Reserve|Valuation)/i.test(
                rows[k].tag || ''
              )
            ) {
              subIndices.push(k);
            } else {
              break;
            }
          }
          if (subIndices.length >= 2) {
            const invNetRow = rows[invNetIdx];
            let overlapCount = 0;
            let allMatch = true;
            for (const yr of colYears) {
              const totVal = invNetRow.rValuesByYear[yr];
              if (totVal === undefined) continue;
              const subVals = subIndices
                .map((si) => rows[si].rValuesByYear[yr])
                .filter((v) => typeof v === 'number');
              if (subVals.length < 2) continue;
              overlapCount++;
              const subSum = subVals.reduce((a, b) => a + b, 0);
              if (Math.abs(subSum - totVal) > 3) {
                allMatch = false;
                break;
              }
            }
            if (overlapCount > 0 && allMatch) {
              const removeSet = new Set(subIndices);
              rows = rows.filter((_, idx) => !removeSet.has(idx));
            }
          }
        }

        // BS 매출채권 총액(AccountsReceivableGrossCurrent)과 순액(AccountsReceivableNetCurrent) 동시 공시 시 총액 중복 제거 (SPSC 등)
        const grossArIdx = rows.findIndex(
          (r) =>
            !r.dimKey &&
            r.rowType === 'USD' &&
            !r.isTotalRow &&
            r.tag === 'AccountsReceivableGrossCurrent'
        );
        const netArRow = rows.find(
          (r) =>
            !r.dimKey &&
            r.rowType === 'USD' &&
            !r.isTotalRow &&
            r.tag === 'AccountsReceivableNetCurrent'
        );
        if (grossArIdx !== -1 && netArRow) {
          const grossArRow = rows[grossArIdx];
          let arOverlap = 0;
          let allGrossGteNet = true;
          for (const yr of colYears) {
            const gVal = grossArRow.rValuesByYear?.[yr];
            const nVal = netArRow.rValuesByYear?.[yr];
            if (typeof gVal !== 'number' || typeof nVal !== 'number') continue;
            arOverlap++;
            if (gVal < nVal - 3) {
              allGrossGteNet = false;
              break;
            }
          }
          if (arOverlap > 0 && allGrossGteNet) {
            rows.splice(grossArIdx, 1);
          }
        }
      }

      // 다차원 복합 지주사(BRK-A 등) 단일 10-K 구조 재배치 (수치 확정 후 / 소계 검산 전)
      if (hasSegmentCF && latestColYr !== undefined) {
        const segFinRows = rows.filter((r) => r.dimAxis === 'ProductOrServiceAxis');
        const mainRows = rows.filter((r) => r.dimAxis !== 'ProductOrServiceAxis');
        const invSubIdx = mainRows.findIndex((r) =>
          [
            'NetCashProvidedByUsedInInvestingActivities',
            'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
          ].includes(r.tag)
        );
        const finSubIdx = mainRows.findIndex((r) =>
          [
            'NetCashProvidedByUsedInFinancingActivities',
            'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
          ].includes(r.tag)
        );
        if (segFinRows.length > 0 && finSubIdx !== -1) {
          const insertAt = invSubIdx !== -1 && invSubIdx < finSubIdx ? invSubIdx + 1 : finSubIdx;
          mainRows.splice(insertAt, 0, ...segFinRows);
          rows = mainRows;
        }
      }

      if (stmtType === 'CF' && !isMultiSegmentTable) {
        const debtOpRows = rows.filter(
          (r) => r.nonSegAxis === 'LongtermDebtTypeAxis' && r.nonSegBlockAbstract === 'OPERATING'
        );
        if (debtOpRows.length > 0) {
          const otherRows = rows.filter(
            (r) => !(r.nonSegAxis === 'LongtermDebtTypeAxis' && r.nonSegBlockAbstract === 'OPERATING')
          );
          const opSubIdx = otherRows.findIndex((r) =>
            [
              'NetCashProvidedByUsedInOperatingActivities',
              'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
            ].includes(r.tag)
          );
          if (opSubIdx !== -1) {
            otherRows.splice(opSubIdx, 0, ...debtOpRows);
            rows = otherRows;
          }
        }
        const classFinRows = rows.filter(
          (r) =>
            r.nonSegAxis === 'StatementClassOfStockAxis' &&
            r.nonSegBlockAbstract === 'FINANCING'
        );
        if (classFinRows.length > 0) {
          const otherRows = rows.filter(
            (r) =>
              !(
                r.nonSegAxis === 'StatementClassOfStockAxis' &&
                r.nonSegBlockAbstract === 'FINANCING'
              )
          );
          const finSubIdx = otherRows.findIndex((r) =>
            [
              'NetCashProvidedByUsedInFinancingActivities',
              'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
            ].includes(r.tag)
          );
          if (finSubIdx !== -1) {
            let prevSubIdx = -1;
            for (let i = finSubIdx - 1; i >= 0; i--) {
              if (otherRows[i].isTotalRow && otherRows[i].rowType === 'USD') {
                prevSubIdx = i;
                break;
              }
            }
            let improvedYearCount = 0;
            let worsenedYearCount = 0;
            for (const yr of colYears) {
              const finVal = otherRows[finSubIdx].rValuesByYear?.[yr];
              if (finVal === undefined || finVal === null) continue;
              const addSum = classFinRows.reduce(
                (acc, cr) => acc + (cr.rValuesByYear?.[yr] ?? 0),
                0
              );
              if (addSum === 0) continue;
              let baseSum = 0;
              for (let i = prevSubIdx + 1; i < finSubIdx; i++) {
                if (otherRows[i].rowType === 'USD' && !otherRows[i].isTotalRow) {
                  baseSum += otherRows[i].rValuesByYear?.[yr] ?? 0;
                }
              }
              const errBefore = Math.abs(baseSum - finVal);
              const errAfter = Math.abs(baseSum + addSum - finVal);
              if (errBefore > 3 && errAfter <= 3) improvedYearCount++;
              else if (errBefore <= 3 && errAfter > 3) worsenedYearCount++;
            }
            if (improvedYearCount > 0 && worsenedYearCount === 0) {
              otherRows.splice(finSubIdx, 0, ...classFinRows);
              rows = otherRows;
            }
          }
        }
        const cfActivitySections = [
          {
            kind: 'OPERATING',
            tags: [
              'NetCashProvidedByUsedInOperatingActivities',
              'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
            ],
          },
          {
            kind: 'INVESTING',
            tags: [
              'NetCashProvidedByUsedInInvestingActivities',
              'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
            ],
          },
          {
            kind: 'FINANCING',
            tags: [
              'NetCashProvidedByUsedInFinancingActivities',
              'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
            ],
          },
        ];
        for (const sec of cfActivitySections) {
          const candRows = rows.filter(
            (r) =>
              r.nonSegAxis === 'OTHER_AXIS' &&
              r.nonSegBlockAbstract === sec.kind &&
              r.rowType === 'USD' &&
              !r.isTotalRow
          );
          if (candRows.length === 0) continue;
          const candSet = new Set(candRows);
          const otherRows = rows.filter((r) => !candSet.has(r));
          const subIdx = otherRows.findIndex((r) => sec.tags.includes(r.tag));
          if (subIdx === -1) continue;

          let prevSubIdx = -1;
          for (let i = subIdx - 1; i >= 0; i--) {
            if (otherRows[i].isTotalRow && otherRows[i].rowType === 'USD') {
              prevSubIdx = i;
              break;
            }
          }
          let compStart = prevSubIdx + 1;
          if (sec.kind === 'OPERATING' && prevSubIdx === -1) {
            compStart = findOperatingCfStartIdx(otherRows, subIdx);
          }

          let improvedYearCount = 0;
          let worsenedYearCount = 0;
          for (const yr of colYears) {
            const subVal = otherRows[subIdx].rValuesByYear?.[yr];
            if (subVal === undefined || subVal === null) continue;
            const addSum = candRows.reduce((acc, cr) => acc + (cr.rValuesByYear?.[yr] ?? 0), 0);
            if (addSum === 0) continue;
            let baseSum = 0;
            for (let i = compStart; i < subIdx; i++) {
              if (
                otherRows[i].rowType !== 'USD' ||
                otherRows[i].isTotalRow ||
                isIntermediateCfAdjustmentSubtotal(otherRows[i].tag, otherRows[i].companyLabel)
              ) {
                continue;
              }
              baseSum += otherRows[i].rValuesByYear?.[yr] ?? 0;
            }
            const errBefore = Math.abs(baseSum - subVal);
            const errAfter = Math.abs(baseSum + addSum - subVal);
            if (errBefore > 3 && errAfter <= 3) improvedYearCount++;
            else if (errBefore <= 3 && errAfter > 3) worsenedYearCount++;
          }

          if (improvedYearCount > 0 && worsenedYearCount === 0) {
            for (const cr of candRows) {
              cr.dimKey =
                'DIM_' + (cr.nonSegMember || 'OTHER').toUpperCase().replace(/[^A-Z0-9]+/g, '_');
              cr.isMainRow = false;
            }
            let insertAt = subIdx;
            if (
              sec.kind === 'OPERATING' &&
              subIdx > 0 &&
              isIntermediateCfAdjustmentSubtotal(
                otherRows[subIdx - 1].tag,
                otherRows[subIdx - 1].companyLabel
              )
            ) {
              insertAt = subIdx - 1;
            }
            otherRows.splice(insertAt, 0, ...candRows);
            rows = otherRows;
          }
        }
      }

      if (hasSegmentBS && latestColYr !== undefined) {
        const block0Rows = rows.filter((r) => !r.dimMember);
        const segMap = new Map();
        for (const r of rows) {
          if (r.dimAxis !== 'ProductOrServiceAxis' || !r.dimMember) continue;
          if (!segMap.has(r.dimMember)) {
            segMap.set(r.dimMember, { assetRows: [], liabRows: [], seenAssetsTotal: false });
          }
          const grp = segMap.get(r.dimMember);
          const isLiab =
            r.blockAbstract === 'LIABILITIES' ||
            (grp.seenAssetsTotal && r.blockAbstract !== 'ASSETS');
          if (isLiab) {
            grp.liabRows.push(r);
          } else {
            grp.assetRows.push(r);
            if (r.isTotalRow && r.tag === 'Assets') grp.seenAssetsTotal = true;
          }
        }

        const b0AssetsIdx = block0Rows.findIndex((r) => r.tag === 'Assets' && r.isTotalRow);
        const b0LiabIdx = block0Rows.findIndex(
          (r, idx) =>
            idx > b0AssetsIdx &&
            (r.tag === 'Liabilities' || r.tag === 'LiabilitiesIncludingDeferredIncomeTaxes') &&
            r.isTotalRow
        );

        if (b0AssetsIdx !== -1 && b0LiabIdx !== -1 && segMap.size > 0) {
          const b0AssetsTotal = block0Rows[b0AssetsIdx];
          const b0LiabDetail = block0Rows.slice(b0AssetsIdx + 1, b0LiabIdx);
          const b0LiabTotal = block0Rows[b0LiabIdx];
          const b0EquityAfter = block0Rows.slice(b0LiabIdx + 1);

          let segAssetSum = 0;
          let hasAllSegAssets = true;
          for (const grp of segMap.values()) {
            const st = grp.assetRows.find((r) => r.isTotalRow && r.tag === 'Assets');
            const v = st?.rValuesByYear?.[latestColYr];
            if (v === undefined) hasAllSegAssets = false;
            else segAssetSum += v;
          }

          const b0AssetVal = b0AssetsTotal.rValuesByYear?.[latestColYr];
          if (hasAllSegAssets && b0AssetVal !== undefined && Math.abs(segAssetSum - b0AssetVal) <= 5) {
            const segLiabTagSet = new Set();
            for (const grp of segMap.values()) {
              for (const lr of grp.liabRows) {
                if (!lr.isTotalRow && lr.tag) segLiabTagSet.add(lr.tag);
              }
            }
            const consolidatedOnlyLiabRows = b0LiabDetail.filter(
              (r) => !r.tag || !segLiabTagSet.has(r.tag)
            );
            const reassembled = [];
            for (const grp of segMap.values()) reassembled.push(...grp.assetRows);
            reassembled.push(b0AssetsTotal);
            for (const grp of segMap.values()) reassembled.push(...grp.liabRows);
            reassembled.push(...consolidatedOnlyLiabRows);
            reassembled.push(b0LiabTotal);
            reassembled.push(...b0EquityAfter);
            rows = reassembled;
          }
        }
      }

      if (hasFullSegmentPnL && latestColYr !== undefined) {
        const block0Rows = rows.filter((r) => !r.dimMember);
        const classRows = rows.filter((r) => r.dimAxis === 'StatementClassOfStockAxis');
        const segMap = new Map();
        for (const r of rows) {
          if (r.dimAxis !== 'ProductOrServiceAxis' || !r.dimMember) continue;
          if (!segMap.has(r.dimMember)) {
            segMap.set(r.dimMember, { revRows: [], costRows: [], seenRevTotal: false });
          }
          const grp = segMap.get(r.dimMember);
          const isCost =
            r.blockAbstract === 'COSTS' || (grp.seenRevTotal && r.blockAbstract !== 'REVENUES');
          if (isCost) {
            grp.costRows.push(r);
          } else {
            grp.revRows.push(r);
            if (r.isTotalRow && r.tag === 'Revenues') grp.seenRevTotal = true;
          }
        }

        const b0RevIdx = block0Rows.findIndex((r) => r.tag === 'Revenues' && r.isTotalRow);
        const b0CostIdx = block0Rows.findIndex(
          (r, idx) => idx > b0RevIdx && r.tag === 'CostsAndExpenses' && r.isTotalRow
        );

        if (b0RevIdx !== -1 && b0CostIdx !== -1 && segMap.size > 0) {
          const b0RevTotal = block0Rows[b0RevIdx];
          const b0CostTotal = block0Rows[b0CostIdx];

          let segRevSum = 0;
          let segCostSum = 0;
          let validSegTotals = true;
          for (const grp of segMap.values()) {
            const sr = grp.revRows.find((r) => r.isTotalRow && r.tag === 'Revenues');
            const sc = grp.costRows.find((r) => r.isTotalRow && r.tag === 'CostsAndExpenses');
            if (sr?.rValuesByYear?.[latestColYr] === undefined || sc?.rValuesByYear?.[latestColYr] === undefined) {
              validSegTotals = false;
              break;
            }
            segRevSum += sr.rValuesByYear[latestColYr];
            segCostSum += sc.rValuesByYear[latestColYr];
          }

          const b0RevVal = b0RevTotal.rValuesByYear?.[latestColYr];
          const b0CostVal = b0CostTotal.rValuesByYear?.[latestColYr];
          if (
            validSegTotals &&
            b0RevVal !== undefined &&
            b0CostVal !== undefined &&
            Math.abs(segRevSum - b0RevVal) <= 5 &&
            Math.abs(segCostSum - b0CostVal) <= 5
          ) {
            const segDetailTags = new Set();
            for (const grp of segMap.values()) {
              for (const sr of [...grp.revRows, ...grp.costRows]) {
                if (!sr.isTotalRow && sr.tag) segDetailTags.add(sr.tag);
              }
            }

            const nonopKeepTags = new Set([
              'GainLossOnInvestments',
              'GainLossOnDerivativeInstrumentsNetPretax',
              'NonoperatingGainsLosses',
              'NonoperatingIncomeExpense',
            ]);
            const rawBetween = block0Rows.slice(b0RevIdx + 1, b0CostIdx);
            const hasSplitInvAndDeriv =
              rawBetween.some((r) => r.tag === 'GainLossOnInvestments') &&
              rawBetween.some((r) => r.tag === 'GainLossOnDerivativeInstrumentsNetPretax');

            const betweenRevAndCost = rawBetween.filter((r) => {
              if (hasSplitInvAndDeriv && r.tag === 'NonoperatingIncomeExpense') return false;
              if (nonopKeepTags.has(r.tag)) return true;
              if (r.tag && segDetailTags.has(r.tag)) return false;
              return true;
            });

            const b0BelowCosts = block0Rows.slice(b0CostIdx + 1);
            const reassembled = [];
            for (const grp of segMap.values()) reassembled.push(...grp.revRows);
            reassembled.push(b0RevTotal);
            reassembled.push(...betweenRevAndCost);
            for (const grp of segMap.values()) reassembled.push(...grp.costRows);
            reassembled.push(b0CostTotal);
            reassembled.push(...b0BelowCosts);
            reassembled.push(...classRows);
            rows = reassembled;
          }
        }
      }

      if (stmtType === 'CF' && !hasSegmentCF) {
        const opSubIdx = rows.findIndex(
          (r) =>
            !r.dimKey &&
            r.isTotalRow &&
            r.rowType === 'USD' &&
            [
              'NetCashProvidedByUsedInOperatingActivities',
              'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
            ].includes(r.tag)
        );
        const actSubIndices = rows
          .map((r, idx) =>
            !r.dimKey &&
            r.isTotalRow &&
            r.rowType === 'USD' &&
            [
              'NetCashProvidedByUsedInOperatingActivities',
              'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
              'NetCashProvidedByUsedInInvestingActivities',
              'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
              'NetCashProvidedByUsedInFinancingActivities',
              'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
            ].includes(r.tag)
              ? idx
              : -1
          )
          .filter((x) => x >= 0);
        const lastActIdx = actSubIndices.length > 0 ? Math.max(...actSubIndices) : -1;
        if (opSubIdx !== -1 && lastActIdx > opSubIdx) {
          const compStart = findOperatingCfStartIdx(rows, opSubIdx);
          const baseSumByYr = {};
          const failingYears = [];
          for (const yr of colYears) {
            const subVal = rows[opSubIdx].rValuesByYear?.[yr];
            if (subVal === undefined || subVal === null) continue;
            let baseSum = 0;
            let hasComp = false;
            for (let i = compStart; i < opSubIdx; i++) {
              if (
                rows[i].rowType !== 'USD' ||
                rows[i].isTotalRow ||
                isIntermediateCfAdjustmentSubtotal(rows[i].tag, rows[i].companyLabel)
              ) {
                continue;
              }
              const v = rows[i].rValuesByYear?.[yr];
              if (typeof v === 'number') {
                baseSum += v;
                hasComp = true;
              }
            }
            if (!hasComp) continue;
            baseSumByYr[yr] = baseSum;
            if (Math.abs(baseSum - subVal) > 3) failingYears.push(yr);
          }

          if (failingYears.length > 0) {
            for (let rIdx = lastActIdx + 1; rIdx < rows.length; rIdx++) {
              const cand = rows[rIdx];
              if (cand.rowType !== 'USD' || cand.isTotalRow || cand.dimKey) continue;
              if (
                /Cash|RestrictedCash|ExchangeRate|IncomeTaxesPaid|InterestPaid|RightOfUse|LeaseLiabilities|CapitalExpenditure|Supplemental|Noncash/i.test(
                  `${cand.tag || ''} ${cand.companyLabel || ''}`
                )
              ) {
                continue;
              }
              let matchedSign = 0;
              for (const sign of [1, -1]) {
                let improved = 0;
                let worsened = 0;
                for (const yr of colYears) {
                  if (!(yr in baseSumByYr)) continue;
                  const subVal = rows[opSubIdx].rValuesByYear[yr];
                  const candVal = (cand.rValuesByYear?.[yr] ?? 0) * sign;
                  const errBefore = Math.abs(baseSumByYr[yr] - subVal);
                  const errAfter = Math.abs(baseSumByYr[yr] + candVal - subVal);
                  if (errBefore > 3 && candVal !== 0 && errAfter <= 3) improved++;
                  else if (errBefore <= 3 && errAfter > 3) worsened++;
                }
                const allNonZeroFailingFixed = failingYears.every((yr) => {
                  const candVal = (cand.rValuesByYear?.[yr] ?? 0) * sign;
                  if (candVal === 0) return true;
                  return (
                    Math.abs(baseSumByYr[yr] + candVal - rows[opSubIdx].rValuesByYear[yr]) <= 3
                  );
                });
                if (
                  improved >= Math.min(2, failingYears.length) &&
                  worsened === 0 &&
                  allNonZeroFailingFixed
                ) {
                  matchedSign = sign;
                  break;
                }
              }
              if (matchedSign !== 0) {
                for (const [yr, val] of Object.entries(cand.rValuesByYear || {})) {
                  if (typeof val === 'number' && val !== 0) {
                    cand.rValuesByYear[yr] = val * matchedSign;
                  }
                }
                rows.splice(rIdx, 1);
                rows.splice(opSubIdx, 0, cand);
                break;
              }
            }
          }
        }
      }

      if (stmtType === 'CF') {
        autoFixRowSignsBySubtotalReconciliation(rows, colYears, 'CF');
      }

      if (stmtType === 'BS') {
        const hasBothEquityTotals =
          rows.some((r) => r.tag === 'StockholdersEquity') &&
          rows.some(
            (r) =>
              r.tag === 'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest'
          );
        if (hasBothEquityTotals) {
          const parentEqRow = rows.find((r) => r.tag === 'StockholdersEquity');
          if (parentEqRow) {
            parentEqRow.contextualLabel = 'Total parent shareholders equity';
          }
        }
        const hasBothPartnersTotals =
          rows.some((r) => r.tag === 'PartnersCapital') &&
          rows.some(
            (r) => r.tag === 'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest'
          );
        if (hasBothPartnersTotals) {
          const parentPartnersRow = rows.find((r) => r.tag === 'PartnersCapital');
          if (parentPartnersRow) {
            parentPartnersRow.contextualLabel = 'Total parent partners capital';
          }
        }
      }

      return rows;
    };

    const [IS, BS, CF] = await Promise.all([
      parseRHtmlTable(statementFiles.IS, 'IS'),
      parseRHtmlTable(statementFiles.BS, 'BS'),
      parseRHtmlTable(statementFiles.CF, 'CF'),
    ]);

    return { IS, BS, CF };
  } catch (e) {
    return null;
  }
}

// 16. 소계 확정 잠금(Reconciled-Year Lock) 및 단일 연도 소계 중복 판별 헬퍼
function findTightenedBsSectionStart(rows, compStartIdx, subIdx, targetAbstract) {
  let firstMatchIdx = -1;
  let hasLeadingNonAbstract = false;
  for (let i = compStartIdx; i < subIdx; i++) {
    if (rows[i].rowType !== 'USD' || rows[i].isTotalRow || rows[i].dimMember) continue;
    if (rows[i].bsSectionAbstract === targetAbstract) {
      if (firstMatchIdx === -1) firstMatchIdx = i;
    } else if (firstMatchIdx === -1) {
      hasLeadingNonAbstract = true;
    }
  }
  return firstMatchIdx !== -1 && hasLeadingNonAbstract ? firstMatchIdx : compStartIdx;
}

function getVirtualNonCurrentLiabContext(rows, stmtType) {
  if (stmtType !== 'BS') return null;
  if (
    rows.some(
      (r) =>
        r.isTotalRow &&
        !r.dimMember &&
        ['Liabilities', 'LiabilitiesIncludingDeferredIncomeTaxes', 'LiabilitiesNoncurrent'].includes(r.tag)
    )
  ) {
    return null;
  }

  const clIdx = rows.findIndex((r) => r.isTotalRow && !r.dimMember && r.tag === 'LiabilitiesCurrent');
  if (clIdx === -1) return null;

  let eqIdx = -1;
  for (let i = clIdx + 1; i < rows.length; i++) {
    if (
      rows[i].isTotalRow &&
      !rows[i].dimMember &&
      [
        'StockholdersEquity',
        'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
        'PartnersCapital',
        'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest',
      ].includes(rows[i].tag)
    ) {
      eqIdx = i;
    }
  }
  if (eqIdx === -1) return null;

  const lseIdx = rows.findIndex(
    (r, idx) => idx > eqIdx && r.isTotalRow && !r.dimMember && r.tag === 'LiabilitiesAndStockholdersEquity'
  );
  if (lseIdx === -1) return null;

  let eqStartIdx = eqIdx;
  for (let i = clIdx + 1; i < eqIdx; i++) {
    if (rows[i].rowType !== 'USD') continue;
    if (
      isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].contextualLabel) ||
      /CommonStock|PreferredStock|AdditionalPaidInCapital|RetainedEarnings|TreasuryStock|AccumulatedOtherComprehensive|StockholdersEquity|PartnersCapital|MinorityInterest/i.test(
        rows[i].tag || ''
      ) ||
      /^(?:common stock|preferred stock|additional paid-?in|retained earnings|accumulated deficit|accumulated other comprehensive|treasury stock)/i.test(
        rows[i].companyLabel || ''
      )
    ) {
      eqStartIdx = i;
      break;
    }
  }

  const nonCurrLiabIndices = [];
  for (let i = clIdx + 1; i < eqStartIdx; i++) {
    if (
      rows[i].rowType === 'USD' &&
      !rows[i].isTotalRow &&
      !isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].contextualLabel) &&
      !isSpecialStandaloneBsLiabilityRow(rows[i].tag, rows[i].contextualLabel)
    ) {
      nonCurrLiabIndices.push(i);
    }
  }

  const mezzIndices = [];
  for (let i = clIdx + 1; i < lseIdx; i++) {
    if (rows[i].rowType === 'USD' && isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].contextualLabel)) {
      mezzIndices.push(i);
    }
  }

  const expectedForYear = (yr) => {
    const lseVal = rows[lseIdx].rValuesByYear?.[yr];
    const eqVal = rows[eqIdx].rValuesByYear?.[yr];
    const clVal = rows[clIdx].rValuesByYear?.[yr];
    if (
      lseVal === undefined ||
      lseVal === null ||
      eqVal === undefined ||
      eqVal === null ||
      clVal === undefined ||
      clVal === null
    ) {
      return null;
    }
    const mezzSum = mezzIndices.reduce((acc, mi) => acc + (rows[mi].rValuesByYear?.[yr] ?? 0), 0);
    return lseVal - eqVal - clVal - mezzSum;
  };

  return { clIdx, eqStartIdx, eqIdx, lseIdx, nonCurrLiabIndices, expectedForYear };
}

function findOwningSubtotalRow(
  rows,
  targetPos,
  isTargetAlreadyInRows = false,
  targetDimMember = '',
  stmtType = ''
) {
  const startSearchIdx = isTargetAlreadyInRows ? targetPos + 1 : targetPos;
  const effDim =
    targetDimMember || (isTargetAlreadyInRows ? rows[targetPos]?.dimMember : '') || '';
  if (stmtType === 'BS' && !effDim) {
    const vCtx = getVirtualNonCurrentLiabContext(rows, stmtType);
    if (
      vCtx &&
      targetPos > vCtx.clIdx &&
      (isTargetAlreadyInRows ? targetPos < vCtx.eqStartIdx : targetPos <= vCtx.eqStartIdx)
    ) {
      return { lockedYears: rows[vCtx.lseIdx].virtualNonCurrLiabLockedYears || new Set() };
    }
  }
  for (let i = startSearchIdx; i < rows.length; i++) {
    if (!rows[i].isTotalRow || rows[i].rowType !== 'USD') continue;
    if (stmtType !== 'CF') {
      if (effDim) {
        if ((rows[i].dimMember || '') !== effDim) continue;
      } else {
        if (rows[i].dimAxis === 'ProductOrServiceAxis') continue;
      }
    } else if (rows[i].dimAxis === 'ProductOrServiceAxis') {
      continue;
    }
    return rows[i];
  }
  return null;
}

function updateReconciledYearLocks(rows, stmtType) {
  const hasSegPnL =
    stmtType === 'IS' && rows.some((r) => r.dimAxis === 'ProductOrServiceAxis' && r.isTotalRow);
  if (stmtType !== 'BS' && stmtType !== 'CF' && !hasSegPnL) return;

  let prevTotalIdx = -1;
  for (let sIdx = 0; sIdx < rows.length; sIdx++) {
    if (!rows[sIdx].isTotalRow || rows[sIdx].rowType !== 'USD') continue;

    const subTotalRow = rows[sIdx];
    const prevTotalRow = prevTotalIdx !== -1 ? rows[prevTotalIdx] : null;
    const prevIdxSnapshot = prevTotalIdx;
    prevTotalIdx = sIdx;

    if (stmtType === 'IS') {
      if (subTotalRow.dimAxis !== 'ProductOrServiceAxis') continue;
    }

    if (stmtType === 'CF') {
      const isActivitySubtotal =
        [
          'NetCashProvidedByUsedInOperatingActivities',
          'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations',
          'NetCashProvidedByUsedInInvestingActivities',
          'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations',
          'NetCashProvidedByUsedInFinancingActivities',
          'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations',
        ].includes(subTotalRow.tag) ||
        (/^net cash (?:provided by|used in)|^cash (?:flows? from|generated by|used in)/i.test(
          subTotalRow.companyLabel || ''
        ) &&
          !/PeriodIncreaseDecrease/i.test(subTotalRow.tag || ''));
      if (!isActivitySubtotal) continue;
    }

    if (stmtType === 'BS') {
      if (
        subTotalRow.tag === 'LiabilitiesAndStockholdersEquity' ||
        /^total liabilities and (?:stockholders|shareholders)'? equity$|^total liabilities and equity$/i.test(
          subTotalRow.companyLabel || ''
        )
      ) {
        continue;
      }
    }

    let compStartIdx = prevIdxSnapshot + 1;
    if (stmtType === 'CF' && prevIdxSnapshot === -1) {
      compStartIdx = findOperatingCfStartIdx(rows, sIdx);
    }
    if (
      stmtType === 'BS' &&
      (subTotalRow.tag === 'AssetsNoncurrent' || subTotalRow.tag === 'LiabilitiesNoncurrent')
    ) {
      const targetAbs =
        subTotalRow.tag === 'AssetsNoncurrent' ? 'ASSETS_NONCURRENT' : 'LIABILITIES_NONCURRENT';
      const absStart = findTightenedBsSectionStart(rows, compStartIdx, sIdx, targetAbs);
      if (absStart > compStartIdx) compStartIdx = absStart;
    }

    const isSegSubtotal = subTotalRow.dimAxis === 'ProductOrServiceAxis' && Boolean(subTotalRow.dimMember);
    const compIndices = [];
    for (let i = compStartIdx; i < sIdx; i++) {
      if (rows[i].rowType !== 'USD' || rows[i].isTotalRow) continue;
      if (stmtType === 'CF' && isIntermediateCfAdjustmentSubtotal(rows[i].tag, rows[i].companyLabel)) {
        continue;
      }
      if (
        stmtType === 'BS' &&
        (isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].companyLabel) ||
          isSpecialStandaloneBsLiabilityRow(rows[i].tag, rows[i].companyLabel))
      ) {
        continue;
      }
      if (stmtType !== 'CF') {
        if (isSegSubtotal) {
          if ((rows[i].dimMember || '') !== subTotalRow.dimMember) continue;
        } else if (rows[i].dimAxis === 'ProductOrServiceAxis') {
          continue;
        }
      }
      compIndices.push(i);
    }

    if (compIndices.length === 0) continue;

    let sameSidePrevRow = null;
    if (stmtType === 'BS' && !isSegSubtotal && prevTotalRow) {
      const isAssetsGrand =
        (subTotalRow.tag === 'Assets' || /^total assets$/i.test(subTotalRow.companyLabel || '')) &&
        ['AssetsCurrent', 'AssetsNoncurrent'].includes(prevTotalRow.tag);
      const isLiabGrand =
        (subTotalRow.tag === 'Liabilities' ||
          /^total liabilities$/i.test(subTotalRow.companyLabel || '')) &&
        ['LiabilitiesCurrent', 'LiabilitiesNoncurrent', 'Deposits'].includes(prevTotalRow.tag);
      const isTotalEqFromParent =
        (subTotalRow.tag ===
          'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest' &&
          prevTotalRow.tag === 'StockholdersEquity') ||
        (subTotalRow.tag ===
          'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest' &&
          prevTotalRow.tag === 'PartnersCapital');
      if (isAssetsGrand || isLiabGrand || isTotalEqFromParent) {
        sameSidePrevRow = prevTotalRow;
      }
    }

    if (!subTotalRow.lockedYears) subTotalRow.lockedYears = new Set();

    for (const [yr, subVal] of Object.entries(subTotalRow.rValuesByYear || {})) {
      if (subVal === undefined || subVal === null) continue;
      const compVals = compIndices
        .map((idx) => rows[idx].rValuesByYear?.[yr])
        .filter((v) => v !== undefined && v !== null);
      if (compVals.length === 0) continue;

      let sum = compVals.reduce((a, b) => a + b, 0);
      if (sameSidePrevRow) {
        const prevVal = sameSidePrevRow.rValuesByYear?.[yr];
        if (prevVal === undefined || prevVal === null) continue;
        sum += prevVal;
      }

      if (Math.abs(sum - subVal) <= 3) {
        subTotalRow.lockedYears.add(String(yr));
      }
    }
  }

  const vCtx = getVirtualNonCurrentLiabContext(rows, stmtType);
  if (vCtx) {
    const lseRow = rows[vCtx.lseIdx];
    lseRow.virtualNonCurrLiabLockedYears = lseRow.virtualNonCurrLiabLockedYears || new Set();
    for (const yr of Object.keys(lseRow.rValuesByYear || {})) {
      const exp = vCtx.expectedForYear(yr);
      if (exp === null) continue;
      const vals = vCtx.nonCurrLiabIndices
        .map((i) => rows[i].rValuesByYear?.[yr])
        .filter((v) => v !== undefined && v !== null);
      if (vals.length === 0) continue;
      const sum = vals.reduce((a, b) => a + b, 0);
      if (Math.abs(sum - exp) <= 3) {
        lseRow.virtualNonCurrLiabLockedYears.add(String(yr));
      }
    }
  }
}

// 단일 연도(yr) 기준 소계 중복 판별 (Phase A 메자닌 보호 + Phase B 다차원 격리 포함)
function isSingleYearRedundantByReconciliation(
  rows,
  targetPos,
  yr,
  candidateVal,
  stmtType,
  isTargetAlreadyInRows = false,
  targetDimMember = ''
) {
  if (!['IS', 'BS', 'CF'].includes(stmtType)) return false;
  if (candidateVal === undefined || candidateVal === null || candidateVal === 0) return false;

  const effDim =
    targetDimMember || (isTargetAlreadyInRows ? rows[targetPos]?.dimMember : '') || '';

  if (stmtType === 'BS' && !effDim) {
    const vCtx = getVirtualNonCurrentLiabContext(rows, stmtType);
    if (
      vCtx &&
      targetPos > vCtx.clIdx &&
      (isTargetAlreadyInRows ? targetPos < vCtx.eqStartIdx : targetPos <= vCtx.eqStartIdx)
    ) {
      const exp = vCtx.expectedForYear(yr);
      if (exp === null) return false;
      const compVals = [];
      for (const i of vCtx.nonCurrLiabIndices) {
        if (isTargetAlreadyInRows && i === targetPos) continue;
        const v = rows[i].rValuesByYear?.[yr];
        if (v !== undefined && v !== null) compVals.push(v);
      }
      if (compVals.length === 0) return false;
      const sumWithout = compVals.reduce((a, b) => a + b, 0);
      const errWithout = Math.abs(sumWithout - exp);
      const errWith = Math.abs(sumWithout + candidateVal - exp);
      return errWithout <= 4 && errWith > errWithout + 2;
    }
  }

  let subTotalIdx = -1;
  const startSearchIdx = isTargetAlreadyInRows ? targetPos + 1 : targetPos;
  for (let i = startSearchIdx; i < rows.length; i++) {
    if (!rows[i].isTotalRow || rows[i].rowType !== 'USD') continue;
    if (stmtType !== 'CF') {
      if (effDim) {
        if ((rows[i].dimMember || '') !== effDim) continue;
      } else if (rows[i].dimAxis === 'ProductOrServiceAxis') {
        continue;
      }
    } else if (rows[i].dimAxis === 'ProductOrServiceAxis') {
      continue;
    }
    subTotalIdx = i;
    break;
  }
  if (subTotalIdx === -1) return false;

  let prevTotalIdx = -1;
  for (let i = targetPos - 1; i >= 0; i--) {
    if (rows[i].isTotalRow && rows[i].rowType === 'USD') {
      prevTotalIdx = i;
      break;
    }
  }

  const subTotalRow = rows[subTotalIdx];
  const prevTotalRow = prevTotalIdx !== -1 ? rows[prevTotalIdx] : null;
  const subVal = subTotalRow.rValuesByYear?.[yr];
  if (subVal === undefined || subVal === null) return false;

  const isSegSubtotal = subTotalRow.dimAxis === 'ProductOrServiceAxis' && Boolean(subTotalRow.dimMember);

  let compStartIdx = prevTotalIdx + 1;
  if (stmtType === 'CF' && prevTotalIdx === -1) {
    compStartIdx = findOperatingCfStartIdx(rows, subTotalIdx);
  }
  if (
    stmtType === 'BS' &&
    !isSegSubtotal &&
    (subTotalRow.tag === 'AssetsNoncurrent' || subTotalRow.tag === 'LiabilitiesNoncurrent')
  ) {
    const targetAbs =
      subTotalRow.tag === 'AssetsNoncurrent' ? 'ASSETS_NONCURRENT' : 'LIABILITIES_NONCURRENT';
    const absStart = findTightenedBsSectionStart(rows, compStartIdx, subTotalIdx, targetAbs);
    if (absStart > compStartIdx) compStartIdx = absStart;
  }

  const otherComponentRows = [];
  for (let i = compStartIdx; i < subTotalIdx; i++) {
    if (isTargetAlreadyInRows && i === targetPos) continue;
    if (stmtType === 'CF' && isIntermediateCfAdjustmentSubtotal(rows[i].tag, rows[i].companyLabel)) {
      continue;
    }
    if (
      stmtType === 'BS' &&
      (isMezzanineOrTemporaryEquityRow(rows[i].tag, rows[i].companyLabel) ||
        isSpecialStandaloneBsLiabilityRow(rows[i].tag, rows[i].companyLabel))
    ) {
      continue;
    }
    if (stmtType !== 'CF') {
      if (isSegSubtotal && (rows[i].dimMember || '') !== subTotalRow.dimMember) continue;
      if (!isSegSubtotal && rows[i].dimAxis === 'ProductOrServiceAxis') continue;
    }
    otherComponentRows.push(rows[i]);
  }

  const compVals = otherComponentRows
    .map((r) => r.rValuesByYear?.[yr])
    .filter((v) => v !== undefined && v !== null);
  if (compVals.length === 0) return false;

  const sumWithout = compVals.reduce((a, b) => a + b, 0);
  const sumWith = sumWithout + candidateVal;

  if (stmtType === 'CF') {
    const errWithout = Math.abs(sumWithout - subVal);
    const errWith = Math.abs(sumWith - subVal);
    return errWithout <= 4 && errWith > errWithout + 2;
  }

  if (stmtType === 'BS') {
    const isGrandTotalBS =
      !isSegSubtotal &&
      (subTotalRow.tag === 'Assets' ||
        subTotalRow.tag === 'Liabilities' ||
        /^total assets$|^total liabilities$/i.test(subTotalRow.companyLabel));

    if (isGrandTotalBS && prevTotalRow && prevTotalRow.rValuesByYear?.[yr] !== undefined) {
      const prevVal = prevTotalRow.rValuesByYear[yr];
      const errCombWithout = Math.abs(prevVal + sumWithout - subVal);
      const errCombWith = Math.abs(prevVal + sumWith - subVal);
      if (errCombWithout <= 4 && errCombWith > errCombWithout + 2) return true;
    }

    const errWithout = Math.abs(sumWithout - subVal);
    const errWith = Math.abs(sumWith - subVal);
    return errWithout <= 4 && errWith > errWithout + 2;
  }

  if (isSegSubtotal) {
    const errWithout = Math.abs(sumWithout - subVal);
    const errWith = Math.abs(sumWith - subVal);
    return errWithout <= 4 && errWith > errWithout + 2;
  }

  if (stmtType === 'IS' && !isSegSubtotal && isTargetAlreadyInRows) {
    const targetRow = rows[targetPos];
    const isTopLineRevenue = [
      'Revenues',
      'SalesRevenueNet',
      'RevenueFromContractWithCustomerExcludingAssessedTax',
      'RevenueFromContractWithCustomerIncludingAssessedTax',
      'RegulatedAndUnregulatedOperatingRevenue',
    ].includes(targetRow?.tag);
    const isCostTotal =
      subTotalRow.tag === 'CostsAndExpenses' ||
      subTotalRow.tag === 'OperatingExpenses' ||
      /costs?|expenses?/i.test(subTotalRow.companyLabel || '');
    if (isTopLineRevenue && isCostTotal) {
      return false;
    }
  }

  const candidatesWithout = [Math.abs(sumWithout - subVal)];
  const candidatesWith = [Math.abs(sumWith - subVal)];

  if (compVals.length >= 2) {
    const firstMinusRestWithout = compVals[0] - compVals.slice(1).reduce((a, b) => a + b, 0);
    const firstMinusRestWith =
      compVals[0] - (compVals.slice(1).reduce((a, b) => a + b, 0) + candidateVal);
    candidatesWithout.push(Math.abs(firstMinusRestWithout - subVal));
    candidatesWith.push(Math.abs(firstMinusRestWith - subVal));
  }

  if (prevTotalRow && prevTotalRow.rValuesByYear?.[yr] !== undefined) {
    const prevVal = prevTotalRow.rValuesByYear[yr];
    candidatesWithout.push(Math.abs(prevVal - sumWithout - subVal));
    candidatesWithout.push(Math.abs(prevVal + sumWithout - subVal));
    candidatesWith.push(Math.abs(prevVal - sumWith - subVal));
    candidatesWith.push(Math.abs(prevVal + sumWith - subVal));
  }

  const minErrWithout = Math.min(...candidatesWithout);
  const minErrWith = Math.min(...candidatesWith);

  return minErrWithout <= 4 && minErrWith > minErrWithout + 2;
}

// 17. 5단계 멀티패스 매칭 + 연도별 소계 항등식 필터링 엔진 (MSA Safety 버전과 100% 동일하게 고정!)
async function fetchMultiYear10KStructure(cik, factsObj, headers, options = {}) {
  try {
    const all10Ks = await discoverAll10KFilings(
      cik,
      factsObj,
      headers,
      options.preloadedSubData || null
    );
    if (all10Ks.length === 0) return null;

    const cikUnpadded = String(parseInt(cik, 10));
    const parsedFilings = [];

    for (let i = 0; i < all10Ks.length; i += 4) {
      if (i > 0) await delay(120);
      const batch = all10Ks.slice(i, i + 4);
      const results = await Promise.all(
        batch.map((f) => fetchSingle10KRStatements(cikUnpadded, f.accn, factsObj, headers))
      );
      for (const res of results) {
        if (res && (res.IS?.length || res.BS?.length || res.CF?.length)) {
          parsedFilings.push(res);
        }
      }
    }

    // 지주사 재편 등으로 IS보다 BS 공시 시작연도가 늦고 10년 미만인 경우(예: DIS 2017 BS),
    // 엄격히 검증된 승계 전 법인 CIK의 과거 10-K를 추가 병합 (기존 10년 완료/단독 IPO 기업은 미실행)
    if (all10Ks.length < 10 && parsedFilings.length > 0 && options.targetSubData) {
      const isYears = [];
      const bsYears = [];
      for (const pf of parsedFilings) {
        for (const r of pf.IS || []) {
          for (const y of Object.keys(r.rValuesByYear || {})) isYears.push(Number(y));
        }
        for (const r of pf.BS || []) {
          if (r.tag === 'Assets' || r.tag === 'LiabilitiesAndStockholdersEquity') {
            for (const y of Object.keys(r.rValuesByYear || {})) bsYears.push(Number(y));
          }
        }
      }
      const minIsYr = isYears.length ? Math.min(...isYears) : Infinity;
      const minBsYr = bsYears.length ? Math.min(...bsYears) : Infinity;
      if (minIsYr < minBsYr && minBsYr < Infinity) {
        const pred = await findVerifiedPredecessor10KCik({
          targetCik: cik,
          targetEntityName: options.targetEntityName || '',
          targetTicker: options.targetTicker || '',
          targetSubData: options.targetSubData,
          headers,
          requiredMaxReportYear: minBsYr - 1,
        });
        if (pred) {
          let predFactsObj = factsObj;
          try {
            const pfRes = await fetch(
              makeUrl(['data', 'sec', 'gov'], `api/xbrl/companyfacts/CIK${pred.cik}.json`),
              { headers }
            );
            if (pfRes.ok) {
              const pfData = await pfRes.json();
              if (pfData?.facts) predFactsObj = pfData.facts;
            }
          } catch {}
          const seenAccns = new Set(all10Ks.map((f) => f.accn));
          const extra10Ks = pred.filings10K
            .filter((f) => !seenAccns.has(f.accn) && f.reportYear <= minBsYr)
            .slice(0, Math.max(1, 10 - all10Ks.length));
          const predCikUnpadded = String(parseInt(pred.cik, 10));
          for (let i = 0; i < extra10Ks.length; i += 4) {
            if (i > 0) await delay(120);
            const batch = extra10Ks.slice(i, i + 4);
            const results = await Promise.all(
              batch.map((f) => fetchSingle10KRStatements(predCikUnpadded, f.accn, predFactsObj, headers))
            );
            for (const res of results) {
              if (res && (res.IS?.length || res.BS?.length || res.CF?.length)) {
                parsedFilings.push(res);
              }
            }
          }
        }
      }
    }

    if (parsedFilings.length === 0) return null;

    for (const f of parsedFilings) {
      for (const stmtType of ['IS', 'BS', 'CF']) {
        for (const r of f[stmtType] || []) {
          r.exactKor = r.dimMember
            ? null
            : matchExactKoreanAccountName(r.contextualLabel, r.tag, stmtType, r.rowType) ||
              (r.isMainRow && STANDARD_KOR_DICT[r.tag]?.kor) ||
              null;
        }
      }
    }

    const baseStructure = parsedFilings[0];

    for (const stmtType of ['IS', 'BS', 'CF']) {
      const anchorRows = baseStructure[stmtType] || [];
      anchorRows.forEach((ar) => {
        ar.isRecentAnchor = true;
      });
      updateReconciledYearLocks(anchorRows, stmtType);

      for (let fIdx = 1; fIdx < parsedFilings.length; fIdx++) {
        const olderRows = parsedFilings[fIdx][stmtType] || [];
        const usedAnchorIndices = new Set();
        const olderToAnchorMap = new Map();

        const lockMatch = (oIdx, aIdx) => {
          usedAnchorIndices.add(aIdx);
          olderToAnchorMap.set(oIdx, aIdx);
          const targetAnchor = anchorRows[aIdx];
          const sourceOlder = olderRows[oIdx];
          const isSegIS = stmtType === 'IS' && targetAnchor.dimAxis === 'ProductOrServiceAxis';
          const isIntermCfSub =
            stmtType === 'CF' &&
            isIntermediateCfAdjustmentSubtotal(targetAnchor.tag, targetAnchor.companyLabel);
          const ownerSub =
            (stmtType === 'BS' || stmtType === 'CF' || isSegIS) &&
            targetAnchor.rowType === 'USD' &&
            !targetAnchor.isTotalRow &&
            !isIntermCfSub &&
            !isMezzanineOrTemporaryEquityRow(targetAnchor.tag, targetAnchor.contextualLabel) &&
            !isSpecialStandaloneBsLiabilityRow(targetAnchor.tag, targetAnchor.contextualLabel)
              ? findOwningSubtotalRow(anchorRows, aIdx, true, targetAnchor.dimMember || '', stmtType)
              : null;

          let isSignFactor = 1;
          if (stmtType === 'IS' && targetAnchor.rowType === 'USD') {
            let oppCount = 0;
            let sameCount = 0;
            for (const [yr, oVal] of Object.entries(sourceOlder.rValuesByYear || {})) {
              const aVal = targetAnchor.rValuesByYear?.[yr];
              if (
                typeof aVal === 'number' &&
                typeof oVal === 'number' &&
                aVal !== 0 &&
                oVal !== 0 &&
                Math.abs(Math.abs(aVal) - Math.abs(oVal)) <= 2
              ) {
                if (Math.sign(aVal) === -Math.sign(oVal)) oppCount++;
                else sameCount++;
              }
            }
            if (oppCount > 0 && sameCount === 0) {
              isSignFactor = -1;
            }
          }

          for (const [yr, val] of Object.entries(sourceOlder.rValuesByYear)) {
            if (ownerSub?.lockedYears?.has(String(yr))) continue;
            if (targetAnchor.rValuesByYear[yr] === undefined) {
              targetAnchor.rValuesByYear[yr] =
                typeof val === 'number' && val !== 0 && isSignFactor === -1 ? -val : val;
              if (sourceOlder.netAcquisitionInflowYears?.has(String(yr))) {
                if (!targetAnchor.netAcquisitionInflowYears) {
                  targetAnchor.netAcquisitionInflowYears = new Set();
                }
                targetAnchor.netAcquisitionInflowYears.add(String(yr));
              }
              if (sourceOlder.netProceedsOutflowYears?.has(String(yr))) {
                if (!targetAnchor.netProceedsOutflowYears) {
                  targetAnchor.netProceedsOutflowYears = new Set();
                }
                targetAnchor.netProceedsOutflowYears.add(String(yr));
              }
            }
          }
        };

        const getCfActivitySection = (rowsList, rowIdx) => {
          if (stmtType !== 'CF') return 'OTHER';
          if (rowsList[rowIdx]?.isTotalRow) return 'OTHER';
          for (let i = rowIdx; i < rowsList.length; i++) {
            const r = rowsList[i];
            if (r.dimAxis === 'ProductOrServiceAxis' && r.dimMember) continue;
            if (!r.isTotalRow || r.rowType !== 'USD') continue;
            const t = r.tag || '';
            const l = r.companyLabel || '';
            if (/OperatingActivities/i.test(t) || /operating activities/i.test(l)) return 'OPERATING';
            if (/InvestingActivities/i.test(t) || /investing activities/i.test(l)) return 'INVESTING';
            if (/FinancingActivities/i.test(t) || /financing activities/i.test(l)) return 'FINANCING';
            return 'OTHER';
          }
          for (let i = rowIdx - 1; i >= 0; i--) {
            const r = rowsList[i];
            if (r.dimAxis === 'ProductOrServiceAxis' && r.dimMember) continue;
            if (!r.isTotalRow || r.rowType !== 'USD') continue;
            const t = r.tag || '';
            const l = r.companyLabel || '';
            if (/FinancingActivities/i.test(t) || /financing activities/i.test(l)) {
              return 'POST_FINANCING';
            }
          }
          return 'OTHER';
        };
        const olderCfSectionCache = olderRows.map((_, i) => getCfActivitySection(olderRows, i));
        const anchorCfSectionCache = anchorRows.map((_, i) => getCfActivitySection(anchorRows, i));
        const isSameCfSection = (oIdx, aIdx) => {
          if (stmtType !== 'CF') return true;
          if (olderRows[oIdx]?.isTotalRow || anchorRows[aIdx]?.isTotalRow) return true;
          const oSec = olderCfSectionCache[oIdx] || 'OTHER';
          const aSec = anchorCfSectionCache[aIdx] || 'OTHER';
          const oDim = olderRows[oIdx]?.dimKey || '';
          const aDim = anchorRows[aIdx]?.dimKey || '';
          if (
            (oSec !== 'OTHER' && aSec === 'OTHER' && /^(?:DEBT_|CLASS_|DIM_)/.test(oDim)) ||
            (aSec !== 'OTHER' && oSec === 'OTHER' && /^(?:DEBT_|CLASS_|DIM_)/.test(aDim))
          ) {
            return false;
          }
          if (
            (oSec === 'POST_FINANCING' && aSec !== 'POST_FINANCING' && aSec !== 'OTHER') ||
            (aSec === 'POST_FINANCING' && oSec !== 'POST_FINANCING' && oSec !== 'OTHER')
          ) {
            return false;
          }
          if (oSec === 'OTHER' || aSec === 'OTHER') return true;
          return oSec === aSec;
        };

        // [Pass 1] 태그 + 차원(dimKey/dimMember) + 문맥라벨 + rowType 완전 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.tag) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              !(or.nonSegAxis && !ar.nonSegAxis) &&
              ar.rowType === or.rowType &&
              ar.tag === or.tag &&
              ar.dimKey === or.dimKey &&
              (ar.dimMember || '') === (or.dimMember || '') &&
              ar.normContextualLabel === or.normContextualLabel
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 1b] 다차원(dimMember) 동일 + 태그 + isTotalRow 완전 일치 (표현 라벨 변화 대응)
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.dimMember || !or.tag) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              ar.rowType === or.rowType &&
              (ar.dimMember || '') === or.dimMember &&
              ar.tag === or.tag &&
              ar.isTotalRow === or.isTotalRow
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 1c] 비다차원 부채·주식종류·관계기업 등 보조 축(DEBT_*, CLASS_*, DIM_*) 전용 매칭
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (
            !or.dimKey ||
            (!or.dimKey.startsWith('DEBT_') &&
              !or.dimKey.startsWith('CLASS_') &&
              !or.dimKey.startsWith('DIM_')) ||
            !or.tag
          ) {
            continue;
          }
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              ar.rowType === or.rowType &&
              ar.tag === or.tag &&
              ar.dimKey === or.dimKey
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 2] 본문 메인 태그(isMainRow) + rowType 완전 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.isMainRow || !or.tag) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              !(or.nonSegAxis && !ar.nonSegAxis) &&
              ar.rowType === or.rowType &&
              ar.isMainRow &&
              ar.tag === or.tag
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 3] 문맥 포함 정규화 영문 라벨(normContextualLabel) + 차원 + rowType 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.normContextualLabel) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              !(or.nonSegAxis && !ar.nonSegAxis) &&
              ar.rowType === or.rowType &&
              (ar.dimMember || '') === (or.dimMember || '') &&
              (!or.dimMember || ar.isTotalRow === or.isTotalRow) &&
              ar.normContextualLabel === or.normContextualLabel
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 4] 확정 한글 계정명(exactKor) + 차원 + rowType 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.exactKor) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              !(or.nonSegAxis && !ar.nonSegAxis) &&
              ar.rowType === or.rowType &&
              (ar.dimMember || '') === (or.dimMember || '') &&
              ar.exactKor === or.exactKor
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 5] 과거 동의어 태그 그룹(HISTORICAL_SYNONYM_GROUPS) + rowType 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.isMainRow || !or.tag) continue;
          let synGroup = HISTORICAL_SYNONYM_GROUPS.find((g) => g.includes(or.tag));
          if (
            !synGroup &&
            stmtType === 'BS' &&
            (or.tag === 'DeferredTaxLiabilitiesNoncurrent' || or.tag === 'DeferredIncomeTaxLiabilitiesNet')
          ) {
            if (getVirtualNonCurrentLiabContext(anchorRows, 'BS') !== null) {
              synGroup = ['DeferredIncomeTaxLiabilitiesNet', 'DeferredTaxLiabilitiesNoncurrent'];
            } else {
              const otherDtlTag =
                or.tag === 'DeferredTaxLiabilitiesNoncurrent'
                  ? 'DeferredIncomeTaxLiabilitiesNet'
                  : 'DeferredTaxLiabilitiesNoncurrent';
              const clIdxAnchor = anchorRows.findIndex(
                (r) => !r.dimMember && r.isTotalRow && r.tag === 'LiabilitiesCurrent'
              );
              const eqIdxAnchor = anchorRows.findIndex(
                (r) =>
                  !r.dimMember &&
                  r.isTotalRow &&
                  [
                    'StockholdersEquity',
                    'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
                    'PartnersCapital',
                    'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest',
                  ].includes(r.tag)
              );
              const targetIdxAnchor = anchorRows.findIndex(
                (ar, idx) =>
                  !usedAnchorIndices.has(idx) &&
                  ar.rowType === or.rowType &&
                  ar.isMainRow &&
                  !ar.dimMember &&
                  ar.tag === otherDtlTag
              );
              const canMerge1to1Dtl =
                targetIdxAnchor !== -1 &&
                !anchorRows.some((ar) => !ar.dimMember && ar.tag === or.tag) &&
                !olderRows.some((r) => !r.dimMember && r.tag === otherDtlTag) &&
                (clIdxAnchor === -1 || targetIdxAnchor > clIdxAnchor) &&
                (eqIdxAnchor === -1 || targetIdxAnchor < eqIdxAnchor) &&
                Object.keys(or.rValuesByYear || {}).some(
                  (yr) => anchorRows[targetIdxAnchor].rValuesByYear?.[yr] === undefined
                );
              if (canMerge1to1Dtl) {
                synGroup = ['DeferredIncomeTaxLiabilitiesNet', 'DeferredTaxLiabilitiesNoncurrent'];
              }
            }
          }
          if (!synGroup) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              !(or.nonSegAxis && !ar.nonSegAxis) &&
              ar.rowType === or.rowType &&
              ar.isMainRow &&
              (ar.dimMember || '') === (or.dimMember || '') &&
              synGroup.includes(ar.tag)
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 5b] 다차원(dimMember) 전용 동의어 태그 그룹 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.dimMember || !or.tag) continue;
          const dimGroups = [...DIM_SYNONYM_GROUPS, ...HISTORICAL_SYNONYM_GROUPS];
          const synGroup = dimGroups.find((g) => g.includes(or.tag));
          if (!synGroup) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              isSameCfSection(oIdx, idx) &&
              ar.rowType === or.rowType &&
              (ar.dimMember || '') === or.dimMember &&
              ar.isTotalRow === or.isTotalRow &&
              synGroup.includes(ar.tag)
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // 과거 소멸 부문(예: 2016 Finance & Financial Products) 삽입 전, 해당 연도 부문 합계가 이미 연결 합계와 일치하거나 10개년 범위 밖인지 확인
        const anchorYearsNumeric = [];
        for (const ar of anchorRows) {
          for (const k of Object.keys(ar.rValuesByYear || {})) {
            const n = parseInt(k, 10);
            if (!isNaN(n)) anchorYearsNumeric.push(n);
          }
        }
        const minAnchorWindowYear =
          anchorYearsNumeric.length > 0 ? Math.max(...anchorYearsNumeric) - 9 : 0;

        const preExistingDimMembers = new Set(
          anchorRows.map((ar) => ar.dimMember).filter(Boolean)
        );
        const isLegacySegmentAlreadyPartitionedInYear = (or, yr) => {
          if (or.dimAxis !== 'ProductOrServiceAxis' || !or.dimMember) return false;
          if (preExistingDimMembers.has(or.dimMember)) return false;
          if (minAnchorWindowYear && parseInt(yr, 10) < minAnchorWindowYear) return true;

          if (stmtType === 'IS') {
            const consolRev = anchorRows.find((ar) => !ar.dimMember && ar.tag === 'Revenues');
            const cVal = consolRev?.rValuesByYear?.[yr];
            if (cVal === undefined) return false;
            const segRevSum = anchorRows
              .filter(
                (ar) =>
                  ar.dimAxis === 'ProductOrServiceAxis' &&
                  preExistingDimMembers.has(ar.dimMember) &&
                  ar.isTotalRow &&
                  ar.tag === 'Revenues'
              )
              .reduce((acc, ar) => acc + (ar.rValuesByYear?.[yr] ?? 0), 0);
            return segRevSum !== 0 && Math.abs(segRevSum - cVal) <= 5;
          }

          if (stmtType === 'BS') {
            if (or.blockAbstract === 'ASSETS' || (!or.blockAbstract && or.tag === 'Assets')) {
              const consolAssets = anchorRows.find((ar) => !ar.dimMember && ar.tag === 'Assets');
              const cVal = consolAssets?.rValuesByYear?.[yr];
              if (cVal === undefined) return false;
              const segAssetSum = anchorRows
                .filter(
                  (ar) =>
                    ar.dimAxis === 'ProductOrServiceAxis' &&
                    preExistingDimMembers.has(ar.dimMember) &&
                    ar.isTotalRow &&
                    ar.tag === 'Assets'
                )
                .reduce((acc, ar) => acc + (ar.rValuesByYear?.[yr] ?? 0), 0);
              return segAssetSum !== 0 && Math.abs(segAssetSum - cVal) <= 5;
            }
            const consolLiab = anchorRows.find(
              (ar) =>
                !ar.dimMember &&
                (ar.tag === 'Liabilities' || ar.tag === 'LiabilitiesIncludingDeferredIncomeTaxes')
            );
            const cVal = consolLiab?.rValuesByYear?.[yr];
            if (cVal === undefined) return false;
            const segLiabSum = anchorRows
              .filter(
                (ar) =>
                  ar.dimAxis === 'ProductOrServiceAxis' &&
                  preExistingDimMembers.has(ar.dimMember) &&
                  ar.isTotalRow &&
                  (ar.tag === 'Liabilities' || ar.tag === 'LiabilitiesExcludingIncomeTaxesPrincipallyDeferred')
              )
              .reduce((acc, ar) => acc + (ar.rValuesByYear?.[yr] ?? 0), 0);
            const defTax = anchorRows.find((ar) => !ar.dimMember && ar.tag === 'IncomeTaxesPrincipallyDeferred');
            const dtVal = defTax?.rValuesByYear?.[yr] ?? 0;
            return segLiabSum !== 0 && Math.abs(segLiabSum + dtVal - cVal) <= 5;
          }

          return false;
        };

        // [Insert Pass] Pass 1~5에서 매칭되지 않은 과거 행 처리
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];

          if (
            stmtType === 'IS' &&
            ['카드수수료수익', '서비스수수료수익', '투자및중개수수료수익', '투자은행(IB)수수료수익'].includes(
              or.exactKor
            ) &&
            anchorRows.some((ar) => ar.exactKor === '수수료및커미션수익')
          ) {
            continue;
          }

          let insertPos = -1;
          for (let prev = oIdx - 1; prev >= 0; prev--) {
            if (olderToAnchorMap.has(prev)) {
              insertPos = olderToAnchorMap.get(prev) + 1;
              break;
            }
          }
          if (insertPos === -1) {
            for (let next = oIdx + 1; next < olderRows.length; next++) {
              if (olderToAnchorMap.has(next)) {
                insertPos = olderToAnchorMap.get(next);
                break;
              }
            }
          }
          if (insertPos === -1) insertPos = anchorRows.length;

          // BS 과거 소멸 부문(SEG_FINANCE_PRODUCTS) 자산/부채 섹션 위치 보정
          if (
            stmtType === 'BS' &&
            or.dimAxis === 'ProductOrServiceAxis' &&
            or.dimMember &&
            !anchorRows.some((ar) => ar.dimMember === or.dimMember)
          ) {
            if (or.blockAbstract === 'ASSETS' || or.tag === 'Assets') {
              const consolAssetsIdx = anchorRows.findIndex((ar) => !ar.dimMember && ar.tag === 'Assets');
              if (consolAssetsIdx !== -1) insertPos = consolAssetsIdx;
            } else if (or.blockAbstract === 'LIABILITIES' || /Liabilities/i.test(or.tag || '')) {
              const defTaxIdx = anchorRows.findIndex(
                (ar) => !ar.dimMember && ar.tag === 'IncomeTaxesPrincipallyDeferred'
              );
              const consolLiabIdx = anchorRows.findIndex(
                (ar) =>
                  !ar.dimMember &&
                  (ar.tag === 'Liabilities' || ar.tag === 'LiabilitiesIncludingDeferredIncomeTaxes')
              );
              if (defTaxIdx !== -1) insertPos = defTaxIdx;
              else if (consolLiabIdx !== -1) insertPos = consolLiabIdx;
            }
          }

          const isMezz =
            stmtType === 'BS' &&
            (isMezzanineOrTemporaryEquityRow(or.tag, or.contextualLabel) ||
              isSpecialStandaloneBsLiabilityRow(or.tag, or.contextualLabel));
          const isIntermCfSub =
            stmtType === 'CF' && isIntermediateCfAdjustmentSubtotal(or.tag, or.companyLabel);
          const isSegIS = stmtType === 'IS' && or.dimAxis === 'ProductOrServiceAxis';
          const ownerSubForInsert =
            !isMezz &&
            !isIntermCfSub &&
            (stmtType === 'BS' || stmtType === 'CF' || isSegIS) &&
            or.rowType === 'USD' &&
            !or.isTotalRow
              ? findOwningSubtotalRow(anchorRows, insertPos, false, or.dimMember || '', stmtType)
              : null;

          const filteredValuesByYear = {};
          for (const [yr, val] of Object.entries(or.rValuesByYear)) {
            if (isLegacySegmentAlreadyPartitionedInYear(or, yr)) {
              continue;
            }
            if (!isMezz && !isIntermCfSub && or.rowType === 'USD' && !or.isTotalRow) {
              if (ownerSubForInsert?.lockedYears?.has(String(yr))) {
                continue;
              }
              if (
                isSingleYearRedundantByReconciliation(
                  anchorRows,
                  insertPos,
                  yr,
                  val,
                  stmtType,
                  false,
                  or.dimMember || ''
                )
              ) {
                continue;
              }
            }
            filteredValuesByYear[yr] = val;
          }

          if (Object.keys(filteredValuesByYear).length === 0) {
            continue;
          }

          const newHistoricalRow = {
            ...or,
            isRecentAnchor: false,
            rValuesByYear: filteredValuesByYear,
          };

          anchorRows.splice(insertPos, 0, newHistoricalRow);

          for (const [k, v] of olderToAnchorMap.entries()) {
            if (v >= insertPos) {
              olderToAnchorMap.set(k, v + 1);
            }
          }
          olderToAnchorMap.set(oIdx, insertPos);
        }

        // [특수 보완] 과거 세부 수수료 -> 최신 통합 '수수료및커미션수익' 합산 보완
        if (stmtType === 'IS') {
          const feeRow = anchorRows.find((r) => r.exactKor === '수수료및커미션수익');
          if (feeRow) {
            const feeSubRows = olderRows.filter((or) =>
              ['카드수수료수익', '서비스수수료수익', '투자및중개수수료수익', '투자은행(IB)수수료수익'].includes(
                or.exactKor
              )
            );
            if (feeSubRows.length >= 2) {
              const olderYears = new Set();
              feeSubRows.forEach((sr) =>
                Object.keys(sr.rValuesByYear).forEach((y) => olderYears.add(y))
              );
              for (const yr of olderYears) {
                if (feeRow.rValuesByYear[yr] === undefined) {
                  const sumFee = feeSubRows.reduce(
                    (acc, sr) => acc + (sr.rValuesByYear[yr] || 0),
                    0
                  );
                  if (sumFee !== 0) feeRow.rValuesByYear[yr] = sumFee;
                }
              }
            }
          }
        }

        // 공시(fIdx) 1건 병합이 완료된 시점에 소계 오차 <= 3 달성 연도를 잠금
        updateReconciledYearLocks(anchorRows, stmtType);
      }
    }

    return baseStructure;
  } catch (e) {
    return null;
  }
}

// 18. 최근 보고된 10-K 기준 과거 10개년도 범위 산출 (수치 로직 100% 동일)
function determineTenYearWindow(usGaap, actualStructure) {
  const yearsSet = new Set();

  if (actualStructure) {
    for (const stmtType of ['IS', 'BS', 'CF']) {
      for (const row of actualStructure[stmtType] || []) {
        for (const yr of Object.keys(row.rValuesByYear || {})) {
          yearsSet.add(parseInt(yr, 10));
        }
      }
    }
  }

  const anchorTags = [
    'Assets',
    'Revenues',
    'RevenueFromContractWithCustomerExcludingAssessedTax',
    'NetIncomeLoss',
    'StockholdersEquity',
    'InterestIncomeExpenseNet',
  ];

  for (const tag of anchorTags) {
    const usdList = usGaap?.[tag]?.units?.USD;
    if (!Array.isArray(usdList)) continue;
    for (const r of usdList) {
      if (r.form === '10-K' || r.form === '10-K/A') {
        const isEarlyFye = Boolean(
          r.end && /^20\d{2}-(?:01-(?:1[6-9]|2\d|3[01])|0[2-5]-\d{2})$/.test(r.end)
        );
        if (isEarlyFye) {
          const fm = r.frame ? r.frame.match(/^CY\d{4}(Q\d)?(I)?$/) : null;
          if (!fm || !fm[1] || fm[1] === 'Q4') {
            yearsSet.add(parseInt(r.end.slice(0, 4), 10));
          }
        } else if (r.frame) {
          const m = r.frame.match(/^CY(\d{4})/);
          if (m) yearsSet.add(parseInt(m[1], 10));
        } else if (r.fy) {
          yearsSet.add(Number(r.fy));
        }
      }
    }
  }

  const sorted = Array.from(yearsSet)
    .filter((y) => y >= 2005 && y <= new Date().getFullYear() + 1)
    .sort((a, b) => a - b);

  if (sorted.length === 0) {
    throw new Error('해당 기업의 10-K 연간 재무 데이터를 찾을 수 없습니다.');
  }

  const latestYear = sorted[sorted.length - 1];
  const years = [];
  for (let y = latestYear - 9; y <= latestYear; y++) {
    years.push(y);
  }
  return years;
}

// 19. CompanyFacts JSON에서 10개년 수치 추출 (수치 로직 100% 동일)
function extract10YearSeriesForTag(
  factsObj,
  prefix,
  tag,
  targetYears,
  stmtType,
  rowType,
  rValuesByYear = {},
  excludedSynonymTags = null
) {
  if (!tag) return null;
  const usGaap = factsObj?.['us-gaap'] || {};
  const customTax = factsObj?.[prefix] || {};

  const synonymGroup = HISTORICAL_SYNONYM_GROUPS.find((grp) => grp.includes(tag));
  const candidateTags = synonymGroup
    ? [
        tag,
        ...synonymGroup.filter(
          (t) => t !== tag && !(excludedSynonymTags && excludedSynonymTags.has(t))
        ),
      ]
    : [tag];

  const getValueFromConcept = (conceptObj, year) => {
    if (!conceptObj || !conceptObj.units) return null;

    let unitKey;
    if (rowType === 'EPS') {
      unitKey = Object.keys(conceptObj.units).find((k) => k.includes('USD/shares') || k.includes('USD / shares'));
    } else if (rowType === 'SHARES') {
      unitKey = Object.keys(conceptObj.units).find(
        (k) => k.toLowerCase().includes('share') && !k.includes('USD')
      );
    } else {
      unitKey = Object.keys(conceptObj.units).find((k) => k === 'USD');
    }
    if (!unitKey) return null;

    const records = conceptObj.units[unitKey];
    if (!Array.isArray(records)) return null;

    const matches = records.filter((r) => {
      if (r.form !== '10-K' && r.form !== '10-K/A') return false;

      if (stmtType === 'BS' && r.start) return false;
      if (stmtType !== 'BS' && r.start && r.end) {
        const days = (new Date(r.end) - new Date(r.start)) / (1000 * 3600 * 24);
        if (days < 250 || days > 400) return false;
      }

      const isEarlyFye = Boolean(
        r.end && /^20\d{2}-(?:01-(?:1[6-9]|2\d|3[01])|0[2-5]-\d{2})$/.test(r.end)
      );
      if (r.frame) {
        const m = r.frame.match(/^CY(\d{4})(Q\d)?(I)?$/);
        if (m) {
          const frameYear = parseInt(m[1], 10);
          const qStr = m[2];
          if (qStr && qStr !== 'Q4') return false;
          if (!isEarlyFye) {
            return frameYear === year;
          }
        }
      }

      let endYr = r.end ? parseInt(r.end.slice(0, 4), 10) : r.fy;
      if (r.end && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(r.end)) endYr -= 1;
      return endYr === year;
    });

    if (matches.length === 0) return null;
    matches.sort((a, b) => (a.filed || '').localeCompare(b.filed || ''));
    const latest = matches[matches.length - 1];
    return typeof latest.val === 'number' ? { val: latest.val, unitKey } : null;
  };

  let groupInv = 0;
  let groupKeep = 0;
  const tagSignMap = {};

  for (const candTag of candidateTags) {
    const conceptObj = customTax[candTag] || usGaap[candTag];
    if (!conceptObj) continue;
    let inv = 0;
    let keep = 0;
    for (const y of targetYears) {
      const rVal = rValuesByYear?.[y];
      const fObj = getValueFromConcept(conceptObj, y);
      if (rVal && fObj && fObj.val) {
        const fMil = rowType === 'EPS' ? Number(fObj.val.toFixed(2)) : Math.round(fObj.val / 1e6);
        if (fMil !== 0 && Math.abs(Math.abs(rVal) - Math.abs(fMil)) <= 2) {
          if (Math.sign(rVal) !== Math.sign(fMil)) {
            inv++;
            groupInv++;
          } else {
            keep++;
            groupKeep++;
          }
        }
      }
    }
    if (inv > 0 || keep > 0) {
      tagSignMap[candTag] = inv > keep ? -1 : 1;
    }
  }

  const defaultGroupSign = groupInv > groupKeep ? -1 : 1;

  const valuesByYear = {};
  let detectedUnitKey = 'USD';

  for (const y of targetYears) {
    let found = null;
    let appliedSign = defaultGroupSign;
    for (const candTag of candidateTags) {
      const conceptObj = customTax[candTag] || usGaap[candTag];
      found = getValueFromConcept(conceptObj, y);
      if (found !== null) {
        appliedSign = tagSignMap[candTag] ?? defaultGroupSign;
        break;
      }
    }

    if (found !== null) {
      valuesByYear[y] = found.val * appliedSign;
      detectedUnitKey = found.unitKey;
    } else {
      valuesByYear[y] = null;
    }
  }

  return { valuesByYear, unitKey: detectedUnitKey };
}

// =========================================================================================
// ★ 20. 수치 확정 이후 마지막 단계(Presentation Layer)에서만 작동하는 한글 계정명 정밀 변환 함수들
// (수치나 행 구조에는 단 1%도 관여하지 않고 오직 엑셀 A열의 텍스트와 굵기만 교정함)
// =========================================================================================

function translateSpecialAndInsuranceLabels(contextualLabel, tag = '', stmtType = '') {
  const raw = stripDynamicParentheticals(contextualLabel).replace(/\((?:in\s+dollars\s+per\s+share|per\s+share)\)/gi, '').trim();
  if (!raw) return null;

  const specialRules = [
    // RLI Corp 등 보험·특수업종 전용 사전 (수치 추출 완료 후 A열 이름만 변환)
    [/^cost\s+of\s+hardware$/i, '하드웨어원가'],
    [/^customer\s+relationships?(?:,?\s*net(?:\s+of\s+amortization)?)?$/i, '고객관계무형자산'],
    [/^product\s+sales$/i, '제품매출액'],
    [/^net premiums earned$|^premiums earned,?\s*net$|^premiums earned$/i, '순경과보험료(보험료수익)'],
    [/^net investment income$|^investment income,?\s*net$/i, '순투자수익'],
    [/^net realized (?:investment )?gains?(?: \(losses\))?$|^realized gains? \(losses\),?\s*net$/i,
      stmtType === 'CF' ? '실현투자손익(순) 조정' : '실현투자손익(순)'],
    [/^other-?than-?temporary-?impairment losses on investments.*$/i, '투자자산손상차손(OTTI)'],
    [/^net unrealized gains? (?:\(losses\) )?on equity securities.*$|^unrealized gains? (?:\(losses\) )?on equity securities.*$/i,
      stmtType === 'CF' ? '지분증권미실현평가손익(순) 조정' : '지분증권미실현평가손익(순)'],
    [/^losses and settlement expenses?$|^losses and loss adjustment expenses?$|^policyholder benefits and claims.*$/i,
      '손해액및손해조사비(보험금지급)'],
    [/^policy acquisition costs?$|^amortization of deferred policy acquisition costs?$/i,
      '보험계약취득비(신계약비)'],
    [/^insurance operating expenses?$/i, '보험영업비용'],
    [/^interest expense on debt$/i, '차입금이자비용'],
    [/^general corporate expenses?$/i, '일반법인관리비'],
    [/^total expenses$|^total benefits,?\s*losses and expenses$/i, '총비용합계'],
    [/^current$/i, stmtType === 'IS' ? '당기법인세비용' : '유동항목'],
    [/^other comprehensive (?:earnings|income) \(loss\),?\s*net of tax$|^other comprehensive (?:earnings|income),?\s*net of tax$/i,
      '기타포괄손익(세후)'],
    [/^comprehensive (?:earnings|income)$|^total comprehensive (?:earnings|income)$/i, '총포괄손익'],
    [/^comprehensive (?:earnings|income) per share$/i, '주당포괄손익(달러)'],
    [/^cash dividends paid per common share$|^dividends declared per common share$/i, '보통주 주당현금배당금(달러)'],
    [/^available-?for-?sale,?\s*at fair value$|^fixed maturities available-?for-?sale.*$/i,
      '매도가능채무증권(공정가치)'],
    [/^held-?to-?maturity,?\s*at amortized cost$|^fixed maturities held-?to-?maturity.*$/i,
      '만기보유채무증권(상각후원가)'],
    [/^equity securities available-?for-?sale,?\s*at fair value$/i, '매도가능지분증권(공정가치)'],
    [/^equity securities,?\s*at fair value$|^equity securities$/i, '지분증권(공정가치)'],
    [/^short-?term investments,?\s*at cost.*$|^short-?term investments$/i, '단기금융상품'],
    [/^other invested assets$/i, '기타투자자산'],
    [/^cash$/i, '현금'],
    [/^total investments and cash$/i, '[투자자산및현금 합계]'],
    [/^total investments$/i, '[투자자산 합계]'],
    [/^premiums and reinsurance balances receivable.*$|^premiums receivable.*$/i,
      stmtType === 'CF' ? '보험료및재보험수취채권의 증감' : '보험료및재보험수취채권(순액)'],
    [/^ceded unearned premiums?$|^prepaid reinsurance premiums?$/i,
      stmtType === 'CF' ? '출재미경과보험료의 증감' : '출재미경과보험료(선급재보험료)'],
    [/^reinsurance balances recoverable on unpaid losses.*$|^reinsurance recoverables?.*$/i,
      stmtType === 'CF' ? '재보험미수금의 증감' : '재보험미수금(지급준비금관련)'],
    [/^deferred policy acquisition costs?,?\s*net$|^deferred policy acquisition costs?$/i,
      stmtType === 'CF' ? '이연신계약비의 증감' : '이연신계약비(이연보험계약취득비)'],
    [/^goodwill and intangibles?$|^goodwill and intangible assets?,?\s*net$/i, '영업권및무형자산(순액)'],
    [/^unpaid losses and settlement expenses?$|^reserve for losses and loss adjustment expenses?$/i,
      stmtType === 'CF' ? '지급준비금(미지급손해액)의 증감' : '지급준비금(미지급손해액및손해조사비)'],
    [/^unearned premiums?$/i,
      stmtType === 'CF' ? '미경과보험료의 증감' : '미경과보험료(보험계약부채)'],
    [/^reinsurance balances payable$|^reinsurance payable$/i,
      stmtType === 'CF' ? '재보험미지급금의 증감' : '재보험미지급금'],
    [/^funds held$|^funds held under reinsurance treaties$/i,
      stmtType === 'CF' ? '재보험예수금의 증감' : '재보험예수금(보유자금부채)'],
    [/^bonds payable,?\s*long-?term debt$/i, '사채및장기차입금'],
    // 복합 지주사(BRK-A 등) 비차원 특수 계정 (수치 확정 후 Presentation Layer 전용)
    [/^investment gains?\s*(?:\(|\/)?\s*loss(?:es)?\)?:?$/i,
      stmtType === 'CF' ? '투자손익 조정' : '투자손익'],
    [/^derivative (?:contract )?gains?\s*(?:\(|\/)?\s*loss(?:es)?\)?:?$/i, '파생상품 계약 손익'],
    [/^investment and derivative (?:contract )?gains?(?: \(losses\)|\/losses)?:?$/i, '투자및파생상품 계약 손익(통합)'],
    [/^income taxes,?\s*principally deferred$/i, '이연법인세부채(법인세부채)'],
    [/^common stock at par value$/i, '보통주자본금'],
    [/^capital in excess of par value$/i, '자본잉여금(보통주)'],
    [/^berkshire (?:hathaway )?shareholders'? equity$/i, '지배기업주주지분'],
    [/^discount accretion on investments.*$/i, '채무증권할증·할인차금상각(순)'],
    [/^equity method investment impairment losses?$/i, '지분법투자주식손상차손'],
    [/^deferred charges\s*[-—–]\s*retroactive reinsurance$|^deferred charges reinsurance assumed$/i, '소급재보험이연자산의 증감'],
    [/^unearned insurance premiums$/i, '미경과보험료의 증감'],
    [/^receivables and originated loans$/i, '수취채권및대출채권의 증감'],
    [/^purchases of equity securities$/i, '지분증권의 취득'],
    [/^sales(?: and redemptions)? of equity securities$/i, '지분증권의 처분'],
    [/^purchases of u\.?s\.?\s*treasury bills and fixed maturity securities$/i, '미국국채및채무증권의 취득'],
    [/^sales of u\.?s\.?\s*treasury bills and fixed maturity securities$/i, '미국국채및채무증권의 처분'],
    [/^redemptions and maturities of u\.?s\.?\s*treasury bills and fixed maturity securities$/i, '미국국채및채무증권의 만기상환'],
    [/^purchases of loans and finance receivables$/i, '대출및금융채권의 취득'],
    [/^collections of loans and finance receivables$/i, '대출및금융채권의 회수'],
    [/^restricted cash,?\s*included in other assets at end of year$/i, '기말 사용제한현금'],
  ];

  for (const [regex, kor] of specialRules) {
    if (regex.test(raw)) return kor;
  }
  return null;
}

const SEGMENT_PREFIX_KOR_MAP = {
  SEG_INSURANCE_OTHER: '보험및기타',
  SEG_RAILROAD_UTILITIES_ENERGY: '철도·유틸리티·에너지',
  SEG_FINANCE_PRODUCTS: '금융·금융상품',
  SEG_SUB_RAILROAD: '철도운송',
  SEG_SUB_UTILITIES_ENERGY: '유틸리티·에너지',
};

function translateDimensionalRowKorName(item, stmtType) {
  const dim = String(item.dimMember || '');
  const raw = stripDynamicParentheticals(item.contextualLabel || item.companyLabel || '')
    .replace(/:$/, '')
    .trim();
  const tag = String(item.tag || '');

  // 주식 종류(Class A / Class B) 주당순이익·발행주식수 구분
  if (dim === 'CLASS_A' || dim === 'CLASS_B' || dim.startsWith('CLASS_')) {
    const classLabel = dim === 'CLASS_A' ? 'Class A' : dim === 'CLASS_B' ? 'Class B' : dim.replace(/^CLASS_/, 'Class ');
    const isDiluted = /diluted/i.test(`${raw} ${tag}`);
    const kind = isDiluted ? '희석' : '기본';
    if (item.rowType === 'EPS') {
      return `${classLabel} ${kind} 주당순이익(달러)`;
    }
    if (item.rowType === 'SHARES') {
      const unitStr = item.shareUnit === 'SHARES' ? '주' : '백만주';
      return `${classLabel} ${kind} 가중평균발행주식수(${unitStr})`;
    }
  }

  const mappedSegKor = SEGMENT_PREFIX_KOR_MAP[dim] || null;
  let baseKor = null;
  let isTotalLabel = false;

  // 부문 소계 행
  if (item.isTotalRow) {
    if (stmtType === 'IS') {
      if (tag === 'Revenues' || /revenues?$/i.test(raw)) {
        baseKor = '매출액 합계';
        isTotalLabel = true;
      } else if (tag === 'CostsAndExpenses' || /costs? and expenses?$/i.test(raw)) {
        baseKor = '원가및영업비용 합계';
        isTotalLabel = true;
      }
    }
    if (stmtType === 'BS') {
      if (tag === 'Assets' || /assets$/i.test(raw)) {
        baseKor = '자산 합계';
        isTotalLabel = true;
      } else if (
        tag === 'Liabilities' ||
        tag === 'LiabilitiesExcludingIncomeTaxesPrincipallyDeferred' ||
        /liabilities/i.test(raw)
      ) {
        baseKor = '부채 합계';
        isTotalLabel = true;
      }
    }
  }

  // 부문 세부 행 매핑 (Presentation Layer 전용)
  const segTagMap = {
    // IS 부문 수익
    PremiumsEarnedNet: '보험료수익',
    SalesAndServiceRevenue: '판매및서비스매출액',
    OperatingLeaseLeaseIncome: '리스수익',
    OperatingLeasesIncomeStatementLeaseRevenue: '리스수익',
    InvestmentIncomeInterestDividendAndOther: '이자·배당및기타투자수익',
    RailroadTransportationRevenues: '철도운송매출액',
    FreightRailTransportationRevenues: '철도운송매출액',
    UtilityAndEnergyOperatingRevenues: '유틸리티및에너지영업수익',
    EnergyOperatingRevenues: '유틸리티및에너지영업수익',
    ServiceRevenuesAndOtherIncome: '서비스수익및기타수익',
    // IS 부문 비용
    LiabilityForUnpaidClaimsAndClaimsAdjustmentExpenseIncurredClaims1: '보험손해액및손해조사비',
    IncurredClaimsPropertyCasualtyAndLiability: '보험손해액및손해조사비',
    PolicyholderBenefitsAndClaimsIncurredNet: '생명·연금·건강보험급부금',
    PolicyholderBenefitsAndClaimsIncurredLifeAnnuityAndHealth: '생명·연금·건강보험급부금',
    PolicyholderBenefitsAndClaimsIncurredLifeAndAnnuity: '생명·연금·건강보험급부금',
    InsuranceUnderwritingExpenses: '보험인수비용',
    CostOfGoodsAndServicesSold: '판매및서비스원가',
    CostOfLeasing: '리스원가',
    SellingGeneralAndAdministrativeExpense: '판매비와관리비',
    GoodwillAndIntangibleAssetImpairment: '영업권및무형자산손상차손',
    InterestExpense: '이자비용',
    OperatingExpenses: '철도운송비용',
    CostsOfServicesAndOperatingExpenses: '유틸리티및에너지영업비용',
    OtherExpenses: '기타영업비용',
    // BS 부문 자산
    CashAndCashEquivalentsAtCarryingValue: '현금및현금성자산',
    USTreasuryBills: '미국 국채 단기투자',
    AvailableForSaleSecuritiesDebtSecurities: '채무증권투자',
    EquitySecuritiesFvNi: '지분증권투자',
    AvailableForSaleSecuritiesEquitySecurities: '지분증권투자',
    EquityMethodInvestments: '지분법적용투자주식',
    NotesReceivableNet: '대출및금융수취채권',
    LoansAndLeasesReceivableNetReportedAmount: '대출및금융수취채권',
    PropertyPlantAndEquipmentAndAssetsHeldForLeaseNet: '유형자산및리스제공자산',
    PropertySubjectToOrAvailableForOperatingLeaseNet: '리스제공장비자산',
    PropertyPlantAndEquipmentNet: '유형자산',
    InventoryNet: '재고자산',
    Goodwill: '영업권',
    IntangibleAssetsNetExcludingGoodwill: '무형자산',
    UnamortizedDeferredChargesOnReinsuranceAssumed: '소급재보험이연자산',
    RegulatoryAssets: '규제자산',
    OtherInvestments: '기타투자자산',
    OtherAssets: '기타자산',
    // BS 부문 부채
    LiabilityForClaimsAndClaimsAdjustmentExpense: '지급준비금(미지급손해액)',
    LiabilityForClaimsAndClaimsAdjustmentExpensePropertyCasualtyLiability: '지급준비금(미지급손해액)',
    LiabilityForUnpaidClaimsAndClaimAdjustmentExpenseRetroactiveReinsurance: '소급재보험 지급준비금',
    LiabilityForUnpaidClaimsAndClaimAdjustmentExpenseRetroactiveReinsuranceContracts: '소급재보험 지급준비금',
    UnearnedPremiums: '미경과보험료',
    LiabilityForFuturePolicyBenefits: '생명·연금·건강보험급부부채',
    PolicyholderFunds: '기타보험계약자부채',
    AccountsPayableAndAccruedLiabilitiesCurrentAndNoncurrent: '매입채무·미지급금및기타부채',
    OtherPayablesToBrokerDealersAndClearingOrganizations: '미국 국채 매입미지급금',
    PayableForPurchaseOfUSTreasuryBills: '미국 국채 매입미지급금',
    AircraftRepurchaseLiabilitiesAndDeferredRevenueLeasesNet: '항공기재매입부채및선수리스수익',
    DerivativeLiabilities: '파생상품계약부채',
    RegulatoryLiabilities: '규제부채',
    DebtAndCapitalLeaseObligations: '지급어음및차입금',
    // CF 부문 재무활동
    ProceedsFromIssuanceOfLongTermDebt: '장기차입금의 차입',
    RepaymentsOfLongTermDebt: '장기차입금의 상환',
  };

  if (!baseKor) {
    baseKor = segTagMap[tag] || null;
    if (!baseKor) {
      if (tag === 'PremiumsAndOtherReceivablesNet' || tag === 'AccountsAndOtherReceivablesNet') {
        baseKor = /^other receivables$/i.test(raw) ? '기타수취채권' : '수취채권';
      } else if (tag === 'SalesRevenueNet' || tag === 'RevenueFromContractWithCustomerExcludingAssessedTax') {
        baseKor = '판매및서비스매출액';
      } else if (tag === 'GainLossOnInvestments') {
        baseKor = '투자손익';
      } else if (tag === 'GainLossOnDerivativeInstrumentsNetPretax') {
        baseKor = '파생상품 계약 손익';
      }
    }

    if (!baseKor) {
      baseKor =
        matchExactKoreanAccountName(raw, tag, stmtType, item.rowType) ||
        translateSpecialAndInsuranceLabels(raw, tag, stmtType) ||
        STANDARD_KOR_DICT[tag]?.kor ||
        null;
    }
  }

  const validBaseKor = baseKor && !/[a-zA-Z]{3,}/.test(baseKor) ? baseKor : null;

  if (mappedSegKor) {
    if (validBaseKor) {
      return isTotalLabel ? `[${mappedSegKor} ${validBaseKor}]` : `[${mappedSegKor}] ${validBaseKor}`;
    }
    return null;
  }

  return {
    pendingSegDim: dim,
    baseKor: validBaseKor,
    isTotalLabel,
  };
}

const PRESERVED_SHORT_ACRONYMS = new Set([
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
  'GAAP',
  'IFRS',
  'SEC',
  'IRS',
]);

function sanitizeAiKoreanLabel(aiKor) {
  if (!aiKor || typeof aiKor !== 'string') return null;
  let s = aiKor.trim();
  if (!/[가-힣]/.test(s)) return null;

  // 허용된 표준 약어가 아닌 영문 단독 괄호 병기(예: (MSRs), (ATM), (Mortgage Servicing Rights)) 제거
  s = s.replace(/\(([A-Za-z0-9.\s,&/'-]+)\)/g, (full, inner) => {
    const token = inner.trim();
    if (!/[A-Za-z]/.test(token)) return full;
    if (PRESERVED_SHORT_ACRONYMS.has(token.toUpperCase())) return full;
    return '';
  });

  s = s.replace(/\s+/g, ' ').trim();
  if (!/[가-힣]/.test(s)) return null;

  const residualTokens = s.match(/[A-Za-z]+/g) || [];
  const hasBadEng = residualTokens.some((tok) => !PRESERVED_SHORT_ACRONYMS.has(tok.toUpperCase()));
  if (hasBadEng) {
    s = s
      .replace(/[A-Za-z]+(?:[-/&.\s]+[A-Za-z]+)*/g, (m) => {
        const clean = m.trim().toUpperCase();
        return PRESERVED_SHORT_ACRONYMS.has(clean) ? m : ' ';
      })
      .replace(/\(\s*\)/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  if (!/[가-힣]/.test(s)) return null;
  const finalCheck = s.match(/[A-Za-z]+/g) || [];
  if (finalCheck.some((tok) => !PRESERVED_SHORT_ACRONYMS.has(tok.toUpperCase()))) {
    return null;
  }

  // K-IFRS 표준 용어 보정 및 AI 흔한 오타·단어 중복 결정적 정규화
  s = s
    .replace(/사업결정/g, '사업결합')
    .replace(/제한용현금성자산/g, '사용제한현금성자산')
    .replace(/사용제한예금|(?<!사용)(?:제한용현금|제한용예금|제한예금)/g, '사용제한현금')
    .replace(/사용사용제한현금/g, '사용제한현금')
    .replace(/차무상환/g, '채무상환')
    .replace(/받어음/g, '받을어음')
    .replace(/처분처분/g, '처분')
    .replace(/조건부대가는부(?:채)?/g, '조건부대가부채')
    .replace(/미지급비용\s*(?:및|과|와)\s*미지급비용/g, '매입채무 및 미지급비용')
    .replace(/운영리스/g, '영업리스');

  return s;
}

// Gemini 응답 지연/실패 시에도 영문이 남지 않도록 보장하는 최종 한글 구문 조합기
function fallbackKoreanLabelBuilder(engLabel, tag = '', stmtType = '') {
  const src = stripDynamicParentheticals(engLabel || tag);
  if (!src) return '기타재무항목';
  if (!/[a-zA-Z]{2,}/.test(src)) return src;

  const lower = src.toLowerCase();
  if (
    /^payment of preferred stock dividends\s*\(([^)]+)\)$/i.test(src)
  ) {
    const inner = RegExp.$1
      .replace(/series\s+([a-z0-9]+)\s+preferred stock/i, '시리즈$1 우선주')
      .replace(/preferred stock/i, '우선주')
      .trim();
    return `우선주 배당금의 지급 (${inner})`;
  }
  const phraseDict = [
    ['customer relationships', '고객관계무형자산'],
    ['cost of hardware', '하드웨어원가'],
    ['mortgage servicing rights', '모기지관리권자산'],
    ['servicing rights', '관리권자산'],
    ['business combination', '사업결합'],
    ['asset retirement', '자산복구충당부채'],
    ['senior notes', '선순위사채'],
    ['financial products', '금융상품'],
    ['financial services', '금융서비스'],
    ['restricted cash', '사용제한현금'],
    ['contingent consideration', '조건부대가'],
    ['rental equipment', '렌탈장비'],
    ['deferred gain', '이연이익'],
    ['pension settlement', '연금정산'],
    ['commercial paper', '기업어음'],
    ['unconsolidated entities', '비연결법인'],
    ['real estate held for investment', '투자부동산'],
    ['real estate held for sale', '매각예정부동산'],
    ['revolving credit', '회전신용한도'],
    ['line of credit', '신용한도'],
    ['mortgages payable', '담보차입금'],
    ['bargain purchase gain', '염가매수차익'],
    ['securities borrowed', '차입유가증권'],
    ['trading liabilities', '단기매매금융부채'],
    ['beneficial interests', '수익증권'],
    ['escrow deposits', '에스크로예치금'],
    ['available-for-sale', '매도가능금융자산'],
    ['held-to-maturity', '만기보유금융자산'],
    ['fixed maturities', '채무증권'],
    ['equity securities', '지분증권'],
    ['debt securities', '채무증권'],
    ['marketable securities', '유가증권'],
    ['short-term investments', '단기금융상품'],
    ['invested assets', '투자자산'],
    ['investment income', '투자수익'],
    ['reinsurance', '재보험'],
    ['premiums', '보험료'],
    ['unpaid losses', '미지급손해액'],
    ['settlement expenses', '손해조사비'],
    ['policy acquisition costs', '보험계약취득비'],
    ['goodwill and intangibles', '영업권및무형자산'],
    ['goodwill', '영업권'],
    ['intangible assets', '무형자산'],
    ['property, plant and equipment', '유형자산'],
    ['property and equipment', '유형자산'],
    ['operating lease', '영업리스'],
    ['finance lease', '금융리스'],
    ['accounts receivable', '매출채권'],
    ['receivables', '수취채권'],
    ['inventories', '재고자산'],
    ['prepaid expenses', '선급비용'],
    ['accounts payable', '매입채무'],
    ['accrued expenses', '미지급비용'],
    ['accrued liabilities', '미지급부채'],
    ['income taxes', '법인세'],
    ['deferred tax', '이연법인세'],
    ['short-term debt', '단기차입금'],
    ['long-term debt', '장기차입금'],
    ['bonds payable', '사채'],
    ['interest expense', '이자비용'],
    ['interest income', '이자수익'],
    ['dividends', '배당금'],
    ['treasury stock', '자기주식'],
    ['common stock', '보통주'],
    ['preferred stock', '우선주'],
    ['retained earnings', '이익잉여금'],
    ['comprehensive earnings', '포괄손익'],
    ['comprehensive income', '포괄손익'],
    ['depreciation', '감가상각비'],
    ['amortization', '상각비'],
    ['impairment', '손상차손'],
    ['operating expenses', '영업비용'],
    ['corporate expenses', '법인관리비'],
    ['capitalization', '총자본화금액'],
    ['machinery', '기계장비'],
    ['transportation', '운송'],
    ['industrial', '산업'],
    ['healthcare', '헬스케어'],
    ['equipment', '장비'],
    ['energy', '에너지'],
    ['power', '전력'],
    ['products', '제품'],
    ['services', '서비스'],
    ['cash', '현금'],
  ];

  let working = lower;
  const matched = [];
  for (const [eng, kor] of phraseDict) {
    if (working.includes(eng)) {
      if (!matched.includes(kor)) matched.push(kor);
      working = working.replace(new RegExp(eng, 'g'), ' ');
    }
  }

  if (matched.length > 0) {
    if (stmtType === 'SEGMENT') {
      return matched.slice(0, 3).join('·');
    }
    let base = matched.slice(0, 2).join('및');
    if (/^total\s+/i.test(lower)) return `${base}합계`;
    if (stmtType === 'CF') {
      if (/^(?:purchases|additions|payments|repayments|repurchases)/i.test(lower)) return `${base} 관련 현금유출`;
      if (/^proceeds/i.test(lower)) return `${base} 관련 현금유입`;
      if (/change|increase|decrease/i.test(lower)) return `${base}의 증감`;
    }
    if (/other/i.test(lower) && !base.startsWith('기타')) base = `기타${base}`;
    return base;
  }

  if (stmtType === 'SEGMENT') return '기타사업부문';
  return src;
}

const PREFERRED_FLASH_MODELS = [
  'gemini-flash-latest',
  'gemini-flash-lite-latest',
  'gemini-3.1-flash-lite',
  'gemini-3.8-flash',
];

async function translateGeminiChunk(batchItems, modelCandidates, deadModels, apiKey) {
  const indexedPayload = {};
  batchItems.forEach((item, idx) => {
    indexedPayload[String(idx)] = {
      stmt: item.stmtType,
      label: item.contextualLabel,
      tag: item.tag,
    };
  });

  const prompt = `당신은 미국 SEC 10-K 재무제표 전문 한국어 공인회계사입니다.
아래 JSON 객체의 영문 재무제표 계정명(stmt, label, tag)들을 K-IFRS 한국 실무 회계 기준에 맞는 정확하고 자연스러운 **순수 한글 계정명**으로 번역하세요.
- 숫자 Key("0", "1"...)를 빠짐없이 그대로 유지하여 JSON 객체로만 응답하세요.
- 괄호 안에 영문 원어나 약어(MSRs, ATM 등)를 병기하지 말고, 달러($)·(in dollars) 표기도 제거하세요.
- 합계(Total)가 아닌 개별 영업비용은 절대 '영업비용합계'로 번역하지 말고 구체적으로 번역하세요.
- K-IFRS 표준 용어 준수: Senior notes→선순위사채, 부채성 Provisions/Accruals/Reserves→충당부채, Business combination→사업결합, Debt extinguishment→채무상환, Restricted cash→사용제한현금.
- 현금흐름표(stmt: "CF")의 운전자본 변동(tag에 IncreaseDecrease/NetChange 포함 또는 자산·부채 항목)은 반드시 '~의 증감'으로 표기하세요.
- label이 사업부문명 단독(예: "Financial Products")이면 tag의 회계 성격(단기차입금·장기차입금·차입·상환 등)을 함께 반영해 구분하세요.
- stmt가 "SEGMENT"인 항목은 사업부문 접두사(예: "기계·에너지·운송", "금융상품")로 간결한 한글 명사구만 출력하세요.
- 오타나 단어 중복 없이 자연스러운 한국어 회계 용어로만 출력하세요.

응답 예시:
{"0": "시장조성및유사활동수익", "1": "매도가능채무증권의 취득"}

[번역할 계정 목록]
${JSON.stringify(indexedPayload)}`;

  const chunkResult = {};
  for (const modelName of modelCandidates) {
    if (deadModels.has(modelName)) continue;
    const url = makeUrl(
      ['generativelanguage', 'googleapis', 'com'],
      `v1beta/models/${modelName}:generateContent?key=${apiKey}`
    );

    for (let attempt = 0; attempt < 2; attempt++) {
      try {
        const response = await fetch(url, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            contents: [{ parts: [{ text: prompt }] }],
            generationConfig: {
              temperature: 0.1,
              responseMimeType: 'application/json',
            },
          }),
          signal: AbortSignal.timeout(12000),
        });

        if (response.status === 404 || response.status === 403 || response.status === 429) {
          deadModels.add(modelName);
          break;
        }
        if (response.status === 400 || response.status === 401) {
          return chunkResult;
        }
        if (response.status === 503 || response.status >= 500) {
          if (attempt === 0) {
            await delay(400);
            continue;
          }
          break;
        }
        if (!response.ok) break;

        const data = await response.json();
        const rawText = data?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
        let parsed = null;
        try {
          parsed = JSON.parse(rawText);
        } catch {
          const jsonMatch = rawText.match(/\{[\s\S]*\}/);
          if (jsonMatch) {
            try {
              parsed = JSON.parse(jsonMatch[0]);
            } catch {}
          }
        }
        if (!parsed || typeof parsed !== 'object') break;

        batchItems.forEach((item, idx) => {
          const val = parsed[String(idx)];
          const kor = typeof val === 'string' ? val : val?.kor;
          if (kor && typeof kor === 'string' && kor.trim()) {
            chunkResult[item.rowKey] = kor.trim();
          }
        });

        if (Object.keys(chunkResult).length > 0) return chunkResult;
        break;
      } catch (e) {
        if (attempt === 0) {
          await delay(400);
          continue;
        }
      }
    }
  }

  return chunkResult;
}

// Gemini AI 한글 번역 함수 (20개 단위 청크 + 최대 4동시 병렬 + 누락키 2차 보완)
async function alignAndTranslateWithGemini(itemsToTranslate, allSheetsRows, apiKey) {
  if (!itemsToTranslate || itemsToTranslate.length === 0) return {};

  let candidateModels = [...PREFERRED_FLASH_MODELS];
  try {
    const listUrl = makeUrl(
      ['generativelanguage', 'googleapis', 'com'],
      `v1beta/models?key=${apiKey}`
    );
    const listRes = await fetch(listUrl, { signal: AbortSignal.timeout(4000) });
    if (listRes.ok) {
      const listData = await listRes.json();
      const available = (listData.models || [])
        .filter(
          (m) =>
            Array.isArray(m.supportedGenerationMethods) &&
            m.supportedGenerationMethods.includes('generateContent') &&
            /gemini/i.test(m.name)
        )
        .map((m) => m.name.replace(/^models\//, ''));

      const availSet = new Set(available);
      const preferredAvail = PREFERRED_FLASH_MODELS.filter((m) => availSet.has(m));
      if (preferredAvail.length > 0) {
        candidateModels = preferredAvail;
      }
    }
  } catch (e) {}

  const deadModels = new Set();
  const result = {};

  const runChunksInPool = async (targetItems, chunkSize = 20, maxConcurrency = 4) => {
    const chunks = [];
    for (let i = 0; i < targetItems.length; i += chunkSize) {
      chunks.push(targetItems.slice(i, i + chunkSize));
    }
    let nextIdx = 0;
    const workerCount = Math.min(maxConcurrency, chunks.length);
    const workers = Array.from({ length: workerCount }, async () => {
      while (nextIdx < chunks.length) {
        const myIdx = nextIdx++;
        const chunkMap = await translateGeminiChunk(
          chunks[myIdx],
          candidateModels,
          deadModels,
          apiKey
        );
        Object.assign(result, chunkMap);
      }
    });
    await Promise.all(workers);
  };

  await runChunksInPool(itemsToTranslate, 20, 4);

  const missingItems = itemsToTranslate.filter((it) => !result[it.rowKey]);
  if (missingItems.length > 0 && missingItems.length < itemsToTranslate.length) {
    await runChunksInPool(missingItems, 20, 4);
  }

  return result;
}

export async function POST(request) {
  try {
    const { ticker, apiKey, userEmail } = await request.json();

    if (!ticker || !apiKey) {
      return NextResponse.json(
        { error: '티커와 Gemini API Key를 모두 입력해 주세요.' },
        { status: 400 }
      );
    }

    const safeEmail =
      userEmail && userEmail.includes('@') ? userEmail.trim() : 'sec_user@gmail.com';
    const headers = {
      'User-Agent': `SecExcelWebApp ${safeEmail}`,
      'Accept-Encoding': 'gzip, deflate',
      Accept: 'application/json',
    };

    // 1. 티커로 CIK 번호 조회
    const tickerUrl = makeUrl(['www', 'sec', 'gov'], 'files/company_tickers.json');
    const tickerRes = await fetch(tickerUrl, { headers });
    if (!tickerRes.ok) {
      throw new Error('SEC 기업 목록 서버에 연결할 수 없습니다.');
    }
    const tickerData = await tickerRes.json();
    const query = ticker.trim().toUpperCase();
    const matched = Object.values(tickerData).find(
      (item) =>
        item.ticker.toUpperCase() === query || item.title.toUpperCase() === query
    );

    if (!matched) {
      return NextResponse.json(
        { error: `입력하신 '${ticker}'에 해당하는 미국 상장기업을 찾을 수 없습니다.` },
        { status: 404 }
      );
    }

    let cik = String(matched.cik_str).padStart(10, '0');
    const entityName = matched.title || matched.ticker;

    // 1-1. Submissions 사전 확인: 해외기업 20-F/40-F 조기 차단 및 지주사 재편 승계 전 법인 CIK 폴백(XOM 등)
    let subData = null;
    try {
      const subUrl = makeUrl(['data', 'sec', 'gov'], `submissions/CIK${cik}.json`);
      const subRes = await fetch(subUrl, { headers });
      if (subRes.ok) {
        subData = await subRes.json();
      }
    } catch {}

    const recentForms = subData?.filings?.recent?.form || [];
    const has10KInRecent = recentForms.some((f) => /^10-K(?:\/A)?$/i.test(f));
    const has20FInRecent = recentForms.some((f) => /^(?:20-F|40-F)(?:\/A)?$/i.test(f));
    if (!has10KInRecent && has20FInRecent) {
      return NextResponse.json(
        {
          error: `'${entityName}'(${query})은(는) 미국 국내기업(10-K)이 아닌 해외 상장기업(SEC 20-F/40-F 공시 대상)으로, 본 10-K 재무제표 자동 추출 대상에서 제외됩니다.`,
        },
        { status: 422 }
      );
    }

    const hasFundFormInRecent = recentForms.some((f) =>
      /^(?:N-CSR|N-CSRS|N-CEN|NPORT-P|N-Q)(?:\/A)?$/i.test(f)
    );
    if (!has10KInRecent && hasFundFormInRecent && subData) {
      const existing10KsForFundCheck = await extract10KsFromSubmissionsData(subData, headers, 1);
      if (existing10KsForFundCheck.length === 0) {
        return NextResponse.json(
          {
            error: `'${entityName}'(${query})은(는) 일반 상장기업(10-K)이 아닌 미국 투자회사법(1940 Act) 등록 펀드(SEC N-CSR/N-CSRS 공시 대상)로, 본 10-K 재무제표 자동 추출 대상에서 제외됩니다.`,
          },
          { status: 422 }
        );
      }
    }

    if (!has10KInRecent && subData) {
      const existing10Ks = await extract10KsFromSubmissionsData(subData, headers, 1);
      if (existing10Ks.length === 0) {
        const pred = await findVerifiedPredecessor10KCik({
          targetCik: cik,
          targetEntityName: entityName,
          targetTicker: query,
          targetSubData: subData,
          headers,
          requiredMaxReportYear: null,
        });
        if (pred) {
          cik = pred.cik;
          subData = pred.subData;
        } else {
          const hasPre10KIpoOrQuarterly = recentForms.some((f) =>
            /^(?:S-1|S-1\/A|S-4|S-4\/A|424B[1-4]|10-Q|10-Q\/A|8-A12B|F-1|F-1\/A)$/i.test(f)
          );
          if (hasPre10KIpoOrQuarterly) {
            return NextResponse.json(
              {
                error: `'${entityName}'(${query})은(는) 최근 IPO(신규 상장) 등으로 아직 첫 연간 사업보고서(SEC 10-K)가 공시되지 않은 기업으로, 본 10-K 재무제표 자동 추출 대상에서 제외됩니다.`,
              },
              { status: 422 }
            );
          }
        }
      }
    }

    // 2. CompanyFacts JSON 로드 후 10년 치 10-K R*.htm 전체 병합 (수치 로직 100% 동일)
    const factsUrl = makeUrl(['data', 'sec', 'gov'], `api/xbrl/companyfacts/CIK${cik}.json`);
    const factsRes = await fetch(factsUrl, { headers });
    if (!factsRes.ok) {
      throw new Error(`SEC 재무 데이터 조회 실패 (상태 코드: ${factsRes.status})`);
    }

    const factsData = await factsRes.json();
    const factsObj = factsData?.facts || {};
    const usGaap = factsObj['us-gaap'] || {};

    const actualStructure = await fetchMultiYear10KStructure(cik, factsObj, headers, {
      preloadedSubData: subData,
      targetSubData: subData,
      targetEntityName: entityName,
      targetTicker: query,
    });

    // 3. 최근 보고된 10-K 기준 10개년도 범위 산출 (수치 로직 100% 동일)
    const targetYears = determineTenYearWindow(usGaap, actualStructure);
    const latestYear = targetYears[targetYears.length - 1];

    // 4. 다중 10-K R*.htm 직접 추출 수치를 기본으로 하고, 연도별 오차 검증을 거쳐 CompanyFacts 보완 (수치 로직 100% 동일)
    const buildFinalSeriesForSheet = (stmtType) => {
      const rRows = actualStructure?.[stmtType] || [];
      const resultRows = [];
      updateReconciledYearLocks(rRows, stmtType);

      const hasMultiSegSheet = rRows.some((r) => r.dimAxis === 'ProductOrServiceAxis');
      const sheetOwnedMainUsdTags = new Set(
        rRows.filter((r) => !r.dimMember && r.rowType === 'USD' && r.tag).map((r) => r.tag)
      );
      const IS_DISAGGREGATED_REV_COGS_TAGS = new Set([
        'SalesRevenueServicesNet',
        'SalesRevenueGoodsNet',
        'SalesRevenueServicesGross',
        'SalesRevenueGoodsGross',
        'LicensesRevenue',
        'LicenseAndServicesRevenue',
        'HardwareRevenue',
        'SoftwareRevenue',
        'MaintenanceRevenue',
        'CostOfGoodsSold',
        'CostOfServices',
        'CostOfGoodsSoldExcludingDepreciationDepletionAndAmortization',
        'CostOfServicesExcludingDepreciationDepletionAndAmortization',
        'LicenseCosts',
      ]);
      const IS_TOTAL_REV_TAGS = new Set([
        'RevenueFromContractWithCustomerExcludingAssessedTax',
        'RevenueFromContractWithCustomerIncludingAssessedTax',
        'Revenues',
        'SalesRevenueNet',
        'SalesAndServiceRevenue',
      ]);
      const IS_TOTAL_COGS_TAGS = new Set([
        'CostOfRevenue',
        'CostOfGoodsAndServicesSold',
        'CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization',
      ]);
      const coveredTotalRevYears = new Set();
      const coveredTotalCogsYears = new Set();
      if (stmtType === 'IS') {
        for (const r of rRows) {
          if (!r.isRecentAnchor || r.dimMember || r.rowType !== 'USD') continue;
          const targetSet = IS_TOTAL_REV_TAGS.has(r.tag)
            ? coveredTotalRevYears
            : IS_TOTAL_COGS_TAGS.has(r.tag)
              ? coveredTotalCogsYears
              : null;
          if (!targetSet) continue;
          for (const yrStr of Object.keys(r.rValuesByYear || {})) {
            const yn = Number(yrStr);
            if (Number.isFinite(yn)) targetSet.add(yn);
          }
        }
      }
      const combinedNonopInRR =
        stmtType === 'IS' && hasMultiSegSheet
          ? rRows.filter(
              (r) =>
                !r.dimMember &&
                (r.tag === 'NonoperatingGainsLosses' || r.tag === 'NonoperatingIncomeExpense')
            )
          : [];
      const lastCfFinTotalIdx =
        stmtType === 'CF'
          ? rRows.reduce(
              (last, r, i) =>
                !r.dimMember &&
                r.isTotalRow &&
                r.rowType === 'USD' &&
                (/FinancingActivities/i.test(r.tag || '') ||
                  /financing activities/i.test(r.companyLabel || ''))
                  ? i
                  : last,
              -1
            )
          : -1;
      const lastCfCashChangeTotalIdx =
        stmtType === 'CF'
          ? rRows.reduce(
              (last, r, i) =>
                !r.dimMember &&
                r.isTotalRow &&
                r.rowType === 'USD' &&
                r.tag &&
                /PeriodIncreaseDecrease/i.test(r.tag)
                  ? i
                  : last,
              -1
            )
          : -1;
      const postFinancingStartIdx =
        lastCfFinTotalIdx !== -1
          ? Math.max(lastCfFinTotalIdx, lastCfCashChangeTotalIdx)
          : -1;

      const enrichedRows = rRows.map((rRow, idx) => {
        const rowKey = `${stmtType}_${idx}_${rRow.dimMember || ''}_${rRow.tag}_${rRow.normContextualLabel}`;
        const stdRule = STANDARD_KOR_DICT[rRow.tag] || null;
        const polarity =
          stmtType === 'CF'
            ? classifyCashFlowRowPolarity(rRow.tag, rRow.contextualLabel, rRow.exactKor || stdRule?.kor || '')
            : 'NEUTRAL';
        const isMezzRow =
          stmtType === 'BS' &&
          (isMezzanineOrTemporaryEquityRow(rRow.tag, rRow.contextualLabel) ||
            isSpecialStandaloneBsLiabilityRow(rRow.tag, rRow.contextualLabel));
        const isIntermCfSub =
          stmtType === 'CF' && isIntermediateCfAdjustmentSubtotal(rRow.tag, rRow.companyLabel);
        const isSegIS = stmtType === 'IS' && rRow.dimAxis === 'ProductOrServiceAxis';
        const ownerSub =
          !isMezzRow &&
          !isIntermCfSub &&
          (stmtType === 'BS' || stmtType === 'CF' || isSegIS) &&
          rRow.rowType === 'USD' &&
          !rRow.isTotalRow
            ? findOwningSubtotalRow(rRows, idx, true, rRow.dimMember || '', stmtType)
            : null;

        const skipCompanyFactsForBranch =
          stmtType === 'IS' &&
          hasMultiSegSheet &&
          !rRow.dimMember &&
          rRow.tag === 'GainLossOnDerivativeInstrumentsNetPretax';

        const skipCompanyFactsForPostFinActivity =
          stmtType === 'CF' &&
          postFinancingStartIdx !== -1 &&
          idx > postFinancingStartIdx &&
          !rRow.isTotalRow &&
          rRow.rowType === 'USD' &&
          /^(?:ProceedsFrom|PaymentsToAcquire|PaymentsFor|RepaymentsOf)/i.test(rRow.tag || '');

        const skipCompanyFactsForDimIsSubline =
          stmtType === 'IS' &&
          rRow.rowType === 'USD' &&
          !rRow.isTotalRow &&
          !rRow.dimMember &&
          Boolean(rRow.nonSegAxis) &&
          rRow.nonSegAxis !== 'StatementClassOfStockAxis' &&
          Boolean(rRow.tag) &&
          (() => {
            const synGrp = HISTORICAL_SYNONYM_GROUPS.find((g) => g.includes(rRow.tag));
            return rRows.some(
              (other) =>
                other !== rRow &&
                !other.dimMember &&
                !other.nonSegAxis &&
                other.rowType === 'USD' &&
                other.tag &&
                (other.tag === rRow.tag || (synGrp && synGrp.includes(other.tag)))
            );
          })();

        const cfSeries =
          rRow.isMainRow &&
          !skipCompanyFactsForBranch &&
          !skipCompanyFactsForDimIsSubline &&
          !skipCompanyFactsForPostFinActivity
            ? extract10YearSeriesForTag(
                factsObj,
                rRow.prefix,
                rRow.tag,
                targetYears,
                stmtType,
                rRow.rowType,
                rRow.rValuesByYear,
                !rRow.dimMember && rRow.rowType === 'USD' && !rRow.isTotalRow
                  ? sheetOwnedMainUsdTags
                  : null
              )
            : null;

        const rRowNumericYears = Object.keys(rRow.rValuesByYear || {})
          .map((k) => Number(k))
          .filter((n) => Number.isFinite(n));
        const maxRYearForRow = rRowNumericYears.length > 0 ? Math.max(...rRowNumericYears) : null;

        const mergedYearMap = { ...(rRow.rValuesByYear || {}) };
        const backfilledYears = new Set();
        const netProceedsOutflowYears = new Set(rRow.netProceedsOutflowYears || []);
        const anchorNonZeroVals = Object.values(mergedYearMap).filter(
          (v) => typeof v === 'number' && v !== 0
        );
        const isConsistentlyNegativeAnchorIS =
          stmtType === 'IS' &&
          anchorNonZeroVals.length > 0 &&
          anchorNonZeroVals.every((v) => v < 0);

        for (const y of targetYears) {
          if (mergedYearMap[y] !== undefined) continue;

          if (
            stmtType === 'IS' &&
            rRow.rowType === 'USD' &&
            !rRow.isRecentAnchor &&
            !rRow.dimMember &&
            !rRow.isTotalRow &&
            maxRYearForRow !== null &&
            y > maxRYearForRow &&
            IS_DISAGGREGATED_REV_COGS_TAGS.has(rRow.tag)
          ) {
            const totalCovered = /Cost/i.test(rRow.tag) ? coveredTotalCogsYears : coveredTotalRevYears;
            if (totalCovered.has(y)) continue;
          }

          if (
            stmtType === 'IS' &&
            hasMultiSegSheet &&
            !rRow.dimMember &&
            rRow.tag === 'GainLossOnInvestments' &&
            combinedNonopInRR.some((cr) => cr.rValuesByYear?.[y] !== undefined)
          ) {
            continue;
          }

          const rawFactVal = cfSeries?.valuesByYear?.[y];
          if (rawFactVal === null || rawFactVal === undefined) continue;

          let candidateVal;
          if (rRow.rowType === 'EPS') {
            candidateVal = Number(rawFactVal.toFixed(2));
          } else if (rRow.rowType === 'SHARES') {
            let shareDivisor = rRow.shareUnit === 'SHARES' ? 1 : 1e6;
            if (rRow.shareUnit !== 'SHARES') {
              const votes = { 1: 0, 1e3: 0, 1e6: 0 };
              for (const [oy, anchorVal] of Object.entries(mergedYearMap)) {
                const factVal = cfSeries?.valuesByYear?.[oy];
                if (
                  typeof anchorVal === 'number' &&
                  anchorVal !== 0 &&
                  typeof factVal === 'number' &&
                  factVal !== 0
                ) {
                  const ratio = Math.abs(factVal) / Math.abs(anchorVal);
                  if (ratio > 400000 && ratio < 2500000) votes[1e6]++;
                  else if (ratio > 400 && ratio < 2500) votes[1e3]++;
                  else if (ratio > 0.4 && ratio < 2.5) votes[1]++;
                }
              }
              const maxV = Math.max(votes[1e6], votes[1e3], votes[1]);
              if (maxV > 0) {
                shareDivisor = votes[1e6] === maxV ? 1e6 : votes[1e3] === maxV ? 1e3 : 1;
              }
            }
            candidateVal = Math.round(rawFactVal / shareDivisor);
          } else {
            const scaled = Math.round(rawFactVal / 1e6);
            const isBackfillNetProceedsOutflow =
              stmtType === 'CF' &&
              scaled < 0 &&
              /^(?:ProceedsFromDivestitureOfBusinessesNetOfCashDivested|ProceedsFromIssuanceOrSaleOfEquity)$|NetOfCashDivested/i.test(
                rRow.tag
              );
            if (scaled === 0) {
              candidateVal = 0;
            } else if (isBackfillNetProceedsOutflow) {
              candidateVal = scaled;
              netProceedsOutflowYears.add(String(y));
            } else if (stdRule?.sign === 'neg' || polarity === 'OUTFLOW') {
              candidateVal = -Math.abs(scaled);
            } else if (
              (stdRule?.sign === 'pos' || polarity === 'INFLOW') &&
              !(isConsistentlyNegativeAnchorIS && scaled < 0)
            ) {
              candidateVal = Math.abs(scaled);
            } else {
              candidateVal = scaled;
            }
          }

          if (!isMezzRow && !isIntermCfSub && rRow.rowType === 'USD' && !rRow.isTotalRow) {
            if (ownerSub?.lockedYears?.has(String(y))) {
              continue;
            }
            if (
              isSingleYearRedundantByReconciliation(
                rRows,
                idx,
                String(y),
                candidateVal,
                stmtType,
                true,
                rRow.dimMember || ''
              )
            ) {
              continue;
            }
          }

          mergedYearMap[y] = candidateVal;
          backfilledYears.add(String(y));
        }

        if (stmtType === 'CF' && rRow.rowType === 'USD') {
          for (const y of targetYears) {
            const v = mergedYearMap[y];
            if (v === undefined || v === null || v === 0) continue;
            if (stdRule?.sign === 'neg' || polarity === 'OUTFLOW') {
              if (!rRow.netAcquisitionInflowYears?.has(String(y))) {
                mergedYearMap[y] = -Math.abs(v);
              }
            } else if (stdRule?.sign === 'pos' || polarity === 'INFLOW') {
              if (!netProceedsOutflowYears.has(String(y))) {
                mergedYearMap[y] = Math.abs(v);
              }
            }
          }
        }

        return {
          ...rRow,
          rowKey,
          stdRule,
          backfilledYears,
          netProceedsOutflowYears,
          rValuesByYear: mergedYearMap,
        };
      });

      if (stmtType === 'IS' && hasMultiSegSheet) {
        // 시계열 분기 행 정리: 투자손익이 개별 공시된 연도(2016~2019, 2022~2025)는 통합행에서 제거하여 중복 방지
        const invGainRow = enrichedRows.find(
          (r) => !r.dimMember && r.tag === 'GainLossOnInvestments'
        );
        const combinedRows = enrichedRows.filter(
          (r) =>
            !r.dimMember &&
            (r.tag === 'NonoperatingGainsLosses' || r.tag === 'NonoperatingIncomeExpense')
        );
        if (invGainRow && combinedRows.length > 0) {
          for (const y of targetYears) {
            if (invGainRow.rValuesByYear?.[y] !== undefined) {
              for (const cr of combinedRows) {
                delete cr.rValuesByYear[y];
              }
            }
          }
        }
      }

      for (let idx = 0; idx < enrichedRows.length; idx++) {
        const rRow = enrichedRows[idx];
        if (
          stmtType === 'BS' &&
          (isMezzanineOrTemporaryEquityRow(rRow.tag, rRow.contextualLabel) ||
            isSpecialStandaloneBsLiabilityRow(rRow.tag, rRow.contextualLabel))
        ) {
          continue;
        }
        if (stmtType === 'IS' && hasMultiSegSheet && !rRow.dimMember) {
          continue;
        }
        const hasRecentPresence = targetYears
          .slice(-2)
          .some((y) => rRow.rValuesByYear[y] !== undefined);

        if (!rRow.isTotalRow && rRow.rowType === 'USD') {
          for (const y of targetYears) {
            const isDisconnectedOlderCfYear =
              stmtType === 'CF' &&
              y < targetYears[targetYears.length - 2] &&
              rRow.rValuesByYear[y + 1] === undefined;
            if (
              hasRecentPresence &&
              !rRow.backfilledYears?.has(String(y)) &&
              !isDisconnectedOlderCfYear
            ) {
              continue;
            }
            const val = rRow.rValuesByYear[y];
            if (val === undefined || val === null) continue;
            if (
              isSingleYearRedundantByReconciliation(
                enrichedRows,
                idx,
                String(y),
                val,
                stmtType,
                true,
                rRow.dimMember || ''
              )
            ) {
              delete rRow.rValuesByYear[y];
            }
          }
        }
      }

      autoFixRowSignsBySubtotalReconciliation(enrichedRows, targetYears, stmtType);

      if (stmtType === 'CF') {
        for (let idx = 0; idx < enrichedRows.length; idx++) {
          const rRow = enrichedRows[idx];
          if (rRow.isTotalRow || rRow.rowType !== 'USD') continue;
          const hasRecentPresence = targetYears
            .slice(-2)
            .some((y) => rRow.rValuesByYear[y] !== undefined);
          for (const y of targetYears) {
            const isDisconnectedOlderCfYear =
              y < targetYears[targetYears.length - 2] &&
              rRow.rValuesByYear[y + 1] === undefined;
            if (
              hasRecentPresence &&
              !rRow.backfilledYears?.has(String(y)) &&
              !isDisconnectedOlderCfYear
            ) {
              continue;
            }
            const val = rRow.rValuesByYear[y];
            if (val === undefined || val === null) continue;
            if (
              isSingleYearRedundantByReconciliation(
                enrichedRows,
                idx,
                String(y),
                val,
                stmtType,
                true,
                rRow.dimMember || ''
              )
            ) {
              delete rRow.rValuesByYear[y];
            }
          }
        }
      }

      if (stmtType === 'BS') {
        const assetsRow = enrichedRows.find((r) => r.tag === 'Assets' && !r.dimMember);
        const lseRow = enrichedRows.find(
          (r) =>
            (r.tag === 'LiabilitiesAndStockholdersEquity' ||
              /^total liabilities and (?:stockholders|shareholders)'? equity$|^total liabilities and equity$/i.test(
                r.companyLabel || ''
              )) &&
            !r.dimMember
        );
        // 상장 전(Pre-IPO) 연도 등 BS 본표 자체가 없는 연도에 CompanyFacts 부주석 잔여 셀만 유입되는 현상 차단 (DUOL 2019 등)
        if (assetsRow || lseRow) {
          for (const y of targetYears) {
            if (
              assetsRow?.rValuesByYear?.[y] === undefined &&
              lseRow?.rValuesByYear?.[y] === undefined
            ) {
              for (const r of enrichedRows) {
                if (r.rValuesByYear) delete r.rValuesByYear[y];
              }
            }
          }
        }

        const liabRow = enrichedRows.find(
          (r) =>
            (r.tag === 'Liabilities' || r.tag === 'LiabilitiesIncludingDeferredIncomeTaxes') &&
            !r.dimMember
        );
        const totalEqRow = enrichedRows.find(
          (r) =>
            r.tag === 'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest' ||
            r.tag === 'PartnersCapitalIncludingPortionAttributableToNoncontrollingInterest'
        );
        const parentEqRow = enrichedRows.find(
          (r) => r.tag === 'StockholdersEquity' || r.tag === 'PartnersCapital'
        );
        const nciRow = enrichedRows.find((r) => r.tag === 'MinorityInterest');
        const mezzRows = enrichedRows.filter((r) =>
          isMezzanineOrTemporaryEquityRow(r.tag, r.contextualLabel)
        );
        const capRow = enrichedRows.find(
          (r) =>
            (r.tag === 'CapitalizationLongtermDebtAndEquity' ||
              /^total\s+capitalization$/i.test(r.companyLabel || '')) &&
            !r.dimMember
        );
        const liabCurrRow = enrichedRows.find(
          (r) => r.tag === 'LiabilitiesCurrent' && !r.dimMember
        );

        // SPAC/역합병 전후 등 해당 연도에 메자닌 없이 자산=부채+자본이 이미 성립하는 고아 메자닌 수치 차단
        if (mezzRows.length > 0 && assetsRow && liabRow && (totalEqRow || parentEqRow)) {
          for (const y of targetYears) {
            const aVal = assetsRow.rValuesByYear?.[y];
            const lVal = liabRow.rValuesByYear?.[y];
            const eqVal =
              totalEqRow?.rValuesByYear?.[y] ??
              (parentEqRow?.rValuesByYear?.[y] !== undefined && (nciRow?.rValuesByYear?.[y] ?? 0) === 0
                ? parentEqRow.rValuesByYear[y]
                : undefined);
            const lseVal = lseRow?.rValuesByYear?.[y];
            const mezzVal = mezzRows.reduce((acc, mr) => acc + (mr.rValuesByYear?.[y] ?? 0), 0);
            if (
              aVal !== undefined &&
              lVal !== undefined &&
              eqVal !== undefined &&
              mezzVal !== 0 &&
              Math.abs(lVal + eqVal - aVal) <= 3 &&
              (lseVal === undefined || Math.abs(lVal + eqVal - lseVal) <= 3) &&
              Math.abs(lVal + eqVal + mezzVal - aVal) > 3
            ) {
              for (const mr of mezzRows) {
                if (mr.rValuesByYear?.[y] !== undefined) delete mr.rValuesByYear[y];
              }
            }
          }
        }

        if (totalEqRow && parentEqRow) {
          for (const y of targetYears) {
            if (
              totalEqRow.rValuesByYear[y] === undefined &&
              parentEqRow.rValuesByYear[y] !== undefined
            ) {
              const nciVal = nciRow?.rValuesByYear?.[y] ?? 0;
              const mezzVal = mezzRows.reduce((acc, mr) => acc + (mr.rValuesByYear?.[y] ?? 0), 0);
              const aVal = assetsRow?.rValuesByYear?.[y];
              const lVal = liabRow?.rValuesByYear?.[y];
              if (
                nciVal === 0 &&
                (aVal === undefined ||
                  lVal === undefined ||
                  Math.abs(lVal + mezzVal + parentEqRow.rValuesByYear[y] - aVal) <= 3)
              ) {
                totalEqRow.rValuesByYear[y] = parentEqRow.rValuesByYear[y];
              }
            }
          }
        }

        // 유틸리티(Capitalization 포맷 등) 과거 공시에서 부채총계 행이 생략된 연도 보완 (NEE 2016~2018 등)
        for (const y of targetYears) {
          const aVal = assetsRow?.rValuesByYear?.[y];
          const lseVal = lseRow?.rValuesByYear?.[y];
          const hasUtilityOrFullBs =
            capRow?.rValuesByYear?.[y] !== undefined ||
            (aVal !== undefined &&
              lseVal !== undefined &&
              Math.abs(lseVal - aVal) <= 3 &&
              liabCurrRow?.rValuesByYear?.[y] !== undefined);

          if (
            hasUtilityOrFullBs &&
            totalEqRow &&
            parentEqRow &&
            nciRow &&
            totalEqRow.rValuesByYear[y] === undefined &&
            parentEqRow.rValuesByYear[y] !== undefined &&
            nciRow.rValuesByYear[y] !== undefined
          ) {
            totalEqRow.rValuesByYear[y] =
              parentEqRow.rValuesByYear[y] + nciRow.rValuesByYear[y];
          }

          if (liabRow && liabRow.rValuesByYear[y] === undefined && aVal !== undefined && hasUtilityOrFullBs) {
            const eqVal =
              totalEqRow?.rValuesByYear?.[y] ??
              (parentEqRow?.rValuesByYear?.[y] !== undefined
                ? parentEqRow.rValuesByYear[y] + (nciRow?.rValuesByYear?.[y] ?? 0)
                : undefined);
            const mezzVal = mezzRows.reduce((acc, mr) => acc + (mr.rValuesByYear?.[y] ?? 0), 0);
            if (eqVal !== undefined && aVal - mezzVal - eqVal > 0) {
              liabRow.rValuesByYear[y] = aVal - mezzVal - eqVal;
            }
          }
        }
      }

      for (let idx = 0; idx < enrichedRows.length; idx++) {
        const rRow = enrichedRows[idx];
        const finalValues = targetYears.map((y) =>
          rRow.rValuesByYear[y] !== undefined ? rRow.rValuesByYear[y] : null
        );

        if (finalValues.every((v) => v === null || v === 0)) continue;

        resultRows.push({
          rowKey: rRow.rowKey,
          stmtType,
          rowType: rRow.rowType,
          tag: rRow.tag,
          dimAxis: rRow.dimAxis || '',
          dimMember: rRow.dimMember || '',
          blockAbstract: rRow.blockAbstract || '',
          shareUnit: rRow.shareUnit || '',
          companyLabel: rRow.companyLabel,
          contextualLabel: rRow.contextualLabel,
          exactKor: rRow.exactKor,
          isMainRow: rRow.isMainRow,
          isTotalRow: rRow.isTotalRow,
          isEps: rRow.rowType === 'EPS',
          stdRule: rRow.stdRule,
          values: finalValues,
        });
      }

      return resultRows;
    };

    const rawSheets = {
      IS: buildFinalSeriesForSheet('IS'),
      BS: buildFinalSeriesForSheet('BS'),
      CF: buildFinalSeriesForSheet('CF'),
    };

    // =========================================================================================
    // 5. ★ 수치 확정 이후 한글 계정명(korName) 결정 단계 (수치 데이터는 이미 위 4단계에서 100% 확정됨!)
    // =========================================================================================
    const REVENUE_TOTAL_TAGS = new Set([
      'RevenueFromContractWithCustomerExcludingAssessedTax',
      'RevenueFromContractWithCustomerIncludingAssessedTax',
      'Revenues',
      'SalesRevenueNet',
    ]);
    const COGS_TOTAL_TAGS = new Set([
      'CostOfRevenue',
      'CostOfGoodsAndServicesSold',
      'CostOfGoodsAndServiceExcludingDepreciationDepletionAndAmortization',
    ]);
    const needsAiList = [];
    const segNeedsAiMap = new Map();
    const pendingSegOnlyItems = [];

    for (const key of ['IS', 'BS', 'CF']) {
      const hasMultiSegInSheet = rawSheets[key].some(
        (r) => r.dimAxis === 'ProductOrServiceAxis'
      );
      const hasSeparateTotalRevenueRow =
        key === 'IS' &&
        rawSheets.IS.some(
          (r) =>
            !r.dimMember &&
            REVENUE_TOTAL_TAGS.has(r.tag) &&
            (r.isTotalRow ||
              /^total\s+(?:operating\s+|net\s+|consolidated\s+)?(?:revenues?|sales)/i.test(
                stripDynamicParentheticals(r.contextualLabel || '')
              ))
        );
      const hasSeparateTotalCogsRow =
        key === 'IS' &&
        rawSheets.IS.some((r) => !r.dimMember && COGS_TOTAL_TAGS.has(r.tag));
      const hasOpTotal =
        key === 'CF' && rawSheets.CF.some((r) => r.tag === 'NetCashProvidedByUsedInOperatingActivities');
      const hasInvTotal =
        key === 'CF' && rawSheets.CF.some((r) => r.tag === 'NetCashProvidedByUsedInInvestingActivities');
      const hasFinTotal =
        key === 'CF' && rawSheets.CF.some((r) => r.tag === 'NetCashProvidedByUsedInFinancingActivities');
      const hasTotalOrDiscEps =
        key === 'IS' &&
        rawSheets.IS.some(
          (r) =>
            r.rowType === 'EPS' &&
            !r.dimMember &&
            (r.tag === 'EarningsPerShareBasic' ||
              r.tag === 'EarningsPerShareDiluted' ||
              /DiscontinuedOperation/i.test(r.tag || '') ||
              /\bdiscontinued\b/i.test(stripDynamicParentheticals(r.contextualLabel || '')))
        );
      for (const item of rawSheets[key]) {
        // (0) 다차원(dimMember) 전용 정밀 한글 변환 (부문 접두사 + Class A/B 구분)
        if (item.dimMember) {
          const dimRes = translateDimensionalRowKorName(item, key);
          if (typeof dimRes === 'string') {
            item.korName = dimRes;
          } else if (dimRes && typeof dimRes === 'object' && dimRes.pendingSegDim) {
            const rawSegEng = dimRes.pendingSegDim
              .replace(/^SEG_/, '')
              .replace(/_/g, ' ')
              .toLowerCase()
              .replace(/\b\w/g, (c) => c.toUpperCase());
            const segKey = `__SEG__${dimRes.pendingSegDim}`;
            if (!segNeedsAiMap.has(segKey)) {
              segNeedsAiMap.set(segKey, {
                rowKey: segKey,
                stmtType: 'SEGMENT',
                contextualLabel: rawSegEng,
                tag: dimRes.pendingSegDim,
                rawSegEng,
              });
            }
            item.pendingSegKey = segKey;
            item.isSegTotalLabel = Boolean(dimRes.isTotalLabel);
            if (dimRes.baseKor) {
              item.resolvedBaseKor = dimRes.baseKor;
              pendingSegOnlyItems.push(item);
            } else {
              needsAiList.push(item);
            }
          } else {
            item.pendingSegPrefix = SEGMENT_PREFIX_KOR_MAP[item.dimMember];
            needsAiList.push(item);
          }
          continue;
        }

        // (0b) 계속영업·중단영업 EPS 라벨 분리 (Step 5 Presentation Layer 전용 — Pass 4 불변 보장)
        if (key === 'IS' && item.rowType === 'EPS') {
          const rawLbl = stripDynamicParentheticals(item.contextualLabel || '');
          const epsTag = String(item.tag || '');
          if (!/comprehensive/i.test(`${rawLbl} ${epsTag}`)) {
            const isDiscontinued =
              /DiscontinuedOperation/i.test(epsTag) || /\bdiscontinued\b/i.test(rawLbl);
            const isContinuing =
              !isDiscontinued &&
              ((hasTotalOrDiscEps && /IncomeLossFromContinuingOperationsPer/i.test(epsTag)) ||
                /\bcontinuing\b/i.test(rawLbl));
            const isDiluted = /diluted/i.test(rawLbl) || /Diluted/i.test(epsTag);
            const isBasic =
              (!isDiluted && (/basic/i.test(rawLbl) || /Basic/i.test(epsTag))) ||
              (/basic/i.test(rawLbl) && !/diluted/i.test(rawLbl));
            if (isDiscontinued) {
              if (isBasic) {
                item.korName = '중단영업 기본 주당순이익(달러)';
                continue;
              }
              if (isDiluted) {
                item.korName = '중단영업 희석 주당순이익(달러)';
                continue;
              }
            }
            if (isContinuing) {
              if (isBasic) {
                item.korName = '계속영업 기본 주당순이익(달러)';
                continue;
              }
              if (isDiluted) {
                item.korName = '계속영업 희석 주당순이익(달러)';
                continue;
              }
            }
          }
        }

        // 다차원 복합 지주사 비차원 특수 라벨 보정 (5개 기준 종목에는 해당 문구 미존재)
        if (key === 'IS' && hasMultiSegInSheet) {
          if (/before income taxes and equity/i.test(item.contextualLabel)) {
            item.korName = '지분법손익반영전 세전이익';
            continue;
          }
          if (
            /^investment and derivative/i.test(
              stripDynamicParentheticals(item.contextualLabel)
            )
          ) {
            item.korName = '투자및파생상품 계약 손익(통합)';
            continue;
          }
          if (item.tag === 'ProfitLoss') {
            item.korName = '연결당기순이익';
            continue;
          }
          if (item.tag === 'NetIncomeLoss') {
            item.korName = '당기순이익';
            continue;
          }
        }
        if (key === 'CF' && hasMultiSegInSheet) {
          if (
            item.exactKor === '법인세비용' ||
            item.tag === 'IncomeTaxExpenseBenefit' ||
            /^income taxes$/i.test(stripDynamicParentheticals(item.contextualLabel))
          ) {
            item.korName = '법인세(이연법인세 등) 증감';
            continue;
          }
        }
        if (key === 'CF' && !item.dimMember) {
          if (isIntermediateCfAdjustmentSubtotal(item.tag, item.contextualLabel)) {
            item.korName = '[조정항목합계]';
            continue;
          }
          if (item.tag === 'NetCashProvidedByUsedInOperatingActivitiesContinuingOperations' && hasOpTotal) {
            item.korName = '[계속영업 영업활동현금흐름]';
            continue;
          }
          if (item.tag === 'NetCashProvidedByUsedInInvestingActivitiesContinuingOperations' && hasInvTotal) {
            item.korName = '[계속영업 투자활동현금흐름]';
            continue;
          }
          if (item.tag === 'NetCashProvidedByUsedInFinancingActivitiesContinuingOperations' && hasFinTotal) {
            item.korName = '[계속영업 재무활동현금흐름]';
            continue;
          }
        }

        if (key === 'IS' && !item.dimMember && item.tag === 'InterestIncomeExpenseNet') {
          item.korName = '순이자손익';
          continue;
        }

        // (1) OperatingExpenses 태그인데 실제 영문명이 'Insurance operating expenses' 등 구체적 개별 비용인 경우
        // '영업비용합계'로 잘못 찍히지 않도록 실제 영문 라벨 기준으로 번역하고 굵은 글씨(bold)를 해제함!
        const isSpecificOperatingExpense =
          (item.tag === 'OperatingExpenses' || item.tag === 'CostsAndExpenses') &&
          !/^(?:total\s+)?(?:operating\s+expenses?|costs?\s+and\s+expenses?)$/i.test(
            item.contextualLabel
          );

        if (isSpecificOperatingExpense) {
          item.isNonTotalOperatingExpense = true;
          const specialKor = translateSpecialAndInsuranceLabels(
            item.contextualLabel,
            item.tag,
            key
          );
          if (specialKor) {
            item.korName = specialKor;
          } else {
            needsAiList.push(item);
          }
          continue;
        }

        // (1b) 총매출 행이 별도로 존재하는 손익계산서에서 'Passenger' 등 세부 매출 행이 RevenueFromContractWithCustomer*
        // 태그를 공유하는 경우 '매출액' 고정 매핑을 막고 원문 기준으로 개별 번역 + bold 해제
        const rawRevLbl = stripDynamicParentheticals(item.contextualLabel || '').trim();
        const isSpecificRevenueSubline =
          key === 'IS' &&
          hasSeparateTotalRevenueRow &&
          !item.isTotalRow &&
          (item.tag === 'RevenueFromContractWithCustomerExcludingAssessedTax' ||
            item.tag === 'RevenueFromContractWithCustomerIncludingAssessedTax') &&
          Boolean(rawRevLbl) &&
          !/^(?:total\s+)?(?:operating\s+|net\s+|consolidated\s+)?(?:revenues?|sales|net\s+sales|revenue\s+from\s+contracts?\s+with\s+customers?)$/i.test(
            rawRevLbl
          );

        if (isSpecificRevenueSubline) {
          item.isNonTotalRevenue = true;
          const directRevKor =
            matchExactKoreanAccountName(item.contextualLabel, '', key, item.rowType) ||
            translateSpecialAndInsuranceLabels(item.contextualLabel, item.tag, key);
          if (directRevKor && directRevKor !== '매출액') {
            item.korName = directRevKor;
          } else {
            needsAiList.push(item);
          }
          continue;
        }

        // (1c) 총매출원가 행이 별도로 존재하는 손익계산서에서 'Cost of hardware' 등 세부 원가 행이
        // CostOfGoodsSold/CostOfServices 태그로 '매출원가' 중복 표기되는 현상 방지 (Step 5 전용)
        const rawCogsLbl = stripDynamicParentheticals(item.contextualLabel || '').trim();
        const isSpecificCogsSubline =
          key === 'IS' &&
          hasSeparateTotalCogsRow &&
          !item.isTotalRow &&
          (item.tag === 'CostOfGoodsSold' || item.tag === 'CostOfServices') &&
          Boolean(rawCogsLbl) &&
          !/^(?:total\s+)?(?:cost\s+of\s+(?:revenue|revenues|sales|goods\s+sold|goods\s+and\s+services\s+sold))$/i.test(
            rawCogsLbl
          );

        if (isSpecificCogsSubline) {
          const directCogsKor =
            translateSpecialAndInsuranceLabels(item.contextualLabel, item.tag, key) ||
            matchExactKoreanAccountName(item.contextualLabel, '', key, item.rowType);
          if (directCogsKor && directCogsKor !== '매출원가') {
            item.korName = directCogsKor;
          } else {
            needsAiList.push(item);
          }
          continue;
        }

        // (2) 기존 정밀 사전 또는 보험·특수업종 사전에 매칭되는 경우
        const specialKor = translateSpecialAndInsuranceLabels(
          item.contextualLabel,
          item.tag,
          key
        );
        if (item.exactKor && !/[a-zA-Z]{3,}/.test(item.exactKor)) {
          item.korName = item.exactKor;
        } else if (specialKor) {
          item.korName = specialKor;
        } else if (item.stdRule?.kor) {
          item.korName = item.stdRule.kor;
        } else {
          needsAiList.push(item);
        }
      }
    }

    if (needsAiList.length > 0 || segNeedsAiMap.size > 0) {
      const allAiItems = [...needsAiList, ...segNeedsAiMap.values()];
      const aiMap = await alignAndTranslateWithGemini(allAiItems, rawSheets, apiKey.trim());
      const segKorMap = {};
      for (const [segKey, segItem] of segNeedsAiMap.entries()) {
        const aiSeg = sanitizeAiKoreanLabel(aiMap[segKey]);
        segKorMap[segKey] =
          aiSeg || fallbackKoreanLabelBuilder(segItem.rawSegEng, segItem.tag, 'SEGMENT');
      }
      for (const item of needsAiList) {
        const aiKor = sanitizeAiKoreanLabel(aiMap[item.rowKey]);
        let baseKor;
        if (aiKor) {
          baseKor = aiKor;
        } else {
          baseKor = fallbackKoreanLabelBuilder(
            item.contextualLabel,
            item.tag,
            item.stmtType
          );
        }
        const segPrefix = item.pendingSegKey ? segKorMap[item.pendingSegKey] : item.pendingSegPrefix;
        item.korName = segPrefix
          ? item.isSegTotalLabel
            ? `[${segPrefix} ${baseKor}]`
            : `[${segPrefix}] ${baseKor}`
          : baseKor;
      }
      for (const item of pendingSegOnlyItems) {
        const segPrefix = segKorMap[item.pendingSegKey] || '기타사업부문';
        item.korName = item.isSegTotalLabel
          ? `[${segPrefix} ${item.resolvedBaseKor}]`
          : `[${segPrefix}] ${item.resolvedBaseKor}`;
      }
    }

    // (3) Step 5 번역 완료 후 시트 문맥 기반 동일 라벨 세분화 (수치·행 구조 불변, A열 표시 텍스트만 보정)
    {
      let cfSec = 0;
      for (const item of rawSheets.CF) {
        if (!item.dimMember) {
          const rawLbl = stripDynamicParentheticals(item.contextualLabel || '');
          if (
            /^other(?:[\s,\-—–]+net)?$/i.test(rawLbl) &&
            (item.korName === '기타손익(순)' || item.korName === '기타')
          ) {
            item.korName =
              cfSec === 0
                ? '기타영업활동현금흐름(순)'
                : cfSec === 1
                ? '기타투자활동현금흐름(순)'
                : '기타재무활동현금흐름(순)';
          }
        }
        if (
          !item.dimMember &&
          item.isTotalRow &&
          /(?:Operating|Investing|Financing)Activities/i.test(item.tag || '')
        ) {
          cfSec++;
        }
      }
    }

    if (rawSheets.BS.length > 0) {
      const hasNonCurrLiabTotal = rawSheets.BS.some(
        (r) => !r.dimMember && r.korName === '[비유동부채 합계]'
      );
      const assetsRowBs = rawSheets.BS.find((r) => !r.dimMember && r.korName === '[[자산총계]]');
      const clIdxBs = rawSheets.BS.findIndex((r) => !r.dimMember && r.korName === '[유동부채 합계]');
      let eqIdxBs = -1;
      for (let i = rawSheets.BS.length - 1; i >= 0; i--) {
        if (!rawSheets.BS[i].dimMember && rawSheets.BS[i].korName === '[[자본총계]]') {
          eqIdxBs = i;
          break;
        }
      }
      if (!hasNonCurrLiabTotal && assetsRowBs && clIdxBs !== -1 && eqIdxBs > clIdxBs + 1) {
        const clRowBs = rawSheets.BS[clIdxBs];
        const eqRowBs = rawSheets.BS[eqIdxBs];
        const mezzRowsBs = rawSheets.BS.filter(
          (r) => !r.dimMember && isMezzanineOrTemporaryEquityRow(r.tag, r.contextualLabel)
        );
        for (let i = clIdxBs + 1; i < eqIdxBs; i++) {
          const cand = rawSheets.BS[i];
          if (cand.dimMember || cand.rowType !== 'USD') continue;
          const rawCandLbl = stripDynamicParentheticals(cand.contextualLabel || '').trim();
          if (!/^total\s+other\s+liabilities$/i.test(rawCandLbl)) continue;

          let checked = 0;
          let matched = 0;
          for (let yi = 0; yi < targetYears.length; yi++) {
            const aVal = assetsRowBs.values[yi];
            const clVal = clRowBs.values[yi];
            const eqVal = eqRowBs.values[yi];
            const cVal = cand.values[yi];
            if (aVal === null || clVal === null || eqVal === null || cVal === null) continue;
            checked++;
            const mVal = mezzRowsBs.reduce((acc, mr) => acc + (mr.values[yi] ?? 0), 0);
            if (Math.abs(clVal + cVal + mVal + eqVal - aVal) <= 3) {
              matched++;
            }
          }
          if (checked > 0 && matched === checked) {
            cand.korName = '[비유동부채 합계]';
            break;
          }
        }
      }

      if (clIdxBs !== -1 && eqIdxBs > clIdxBs + 1) {
        for (let i = clIdxBs + 1; i < eqIdxBs; i++) {
          const cand = rawSheets.BS[i];
          if (cand.dimMember || cand.rowType !== 'USD' || !cand.isTotalRow) continue;
          if (cand.korName === '[비유동부채 합계]' || cand.korName === '[[부채총계]]') continue;
          const rawCandLbl = stripDynamicParentheticals(cand.contextualLabel || '').trim();
          const isOtherNclTotal =
            cand.tag === 'LiabilitiesOtherThanLongtermDebtNoncurrent' ||
            /^total\s+other\s+(?:long-?term|non-?current)\s+liabilities$/i.test(rawCandLbl);
          if (!isOtherNclTotal) continue;

          for (let startIdx = i - 1; startIdx > clIdxBs; startIdx--) {
            const topRow = rawSheets.BS[startIdx];
            if (
              topRow.dimMember ||
              topRow.rowType !== 'USD' ||
              topRow.isTotalRow ||
              isMezzanineOrTemporaryEquityRow(topRow.tag, topRow.contextualLabel)
            ) {
              break;
            }
            const slice = rawSheets.BS.slice(startIdx, i);
            const tol = Math.max(3, Math.ceil((slice.length + 1) * 0.5));
            let checked = 0;
            let matched = 0;
            for (let yi = 0; yi < targetYears.length; yi++) {
              const cVal = cand.values[yi];
              if (cVal === null) continue;
              const sum = slice.reduce((acc, r) => acc + (r.values[yi] ?? 0), 0);
              checked++;
              if (Math.abs(sum - cVal) <= tol) matched++;
            }
            if (checked > 0 && matched === checked) {
              cand.korName = '[기타비유동부채 합계]';
              break;
            }
          }
        }
      }

      const caIdxBs = rawSheets.BS.findIndex(
        (r) => !r.dimMember && r.korName === '[유동자산 합계]'
      );
      const assetsIdxBs = rawSheets.BS.findIndex(
        (r) => !r.dimMember && r.korName === '[[자산총계]]'
      );
      if (caIdxBs !== -1 && assetsIdxBs > caIdxBs + 1) {
        for (let i = caIdxBs + 1; i < assetsIdxBs; i++) {
          const cand = rawSheets.BS[i];
          if (cand.dimMember || cand.rowType !== 'USD' || !cand.isTotalRow) continue;
          const rawCandLbl = stripDynamicParentheticals(cand.contextualLabel || '').trim();
          const isOtherAssetsCand =
            cand.korName === '[비유동자산 합계]' ||
            cand.tag === 'RegulatedEntityOtherAssetsNoncurrent' ||
            /^total\s+other\s+(?:non-?current\s+)?assets$/i.test(rawCandLbl);
          if (!isOtherAssetsCand) continue;
          const hasOtherSignal =
            /other/i.test(rawCandLbl) ||
            cand.tag === 'OtherAssetsNoncurrent' ||
            cand.tag === 'RegulatedEntityOtherAssetsNoncurrent' ||
            cand.tag === 'OtherAssets';
          if (!hasOtherSignal) continue;

          const caRow = rawSheets.BS[caIdxBs];
          const assetsRow = rawSheets.BS[assetsIdxBs];
          let checked = 0;
          let fullMatch = 0;
          for (let yi = 0; yi < targetYears.length; yi++) {
            const caV = caRow.values[yi];
            const cV = cand.values[yi];
            const aV = assetsRow.values[yi];
            if (caV === null || cV === null || aV === null) continue;
            checked++;
            if (Math.abs(caV + cV - aV) <= 3) fullMatch++;
          }
          if (checked > 0 && fullMatch === 0) {
            cand.korName = '[기타비유동자산 합계]';
          }
        }
      }
    }

    if (rawSheets.BS.length > 0) {
      const bsGroups = new Map();
      const hasCaTotal = rawSheets.BS.some((r) => !r.dimMember && r.korName === '[유동자산 합계]');
      const hasClTotal = rawSheets.BS.some((r) => !r.dimMember && r.korName === '[유동부채 합계]');
      const assetsIdx = rawSheets.BS.findIndex((r) => !r.dimMember && r.korName === '[[자산총계]]');
      const caIdx = rawSheets.BS.findIndex((r) => !r.dimMember && r.korName === '[유동자산 합계]');
      const clIdx = rawSheets.BS.findIndex((r) => !r.dimMember && r.korName === '[유동부채 합계]');
      const liabIdx = rawSheets.BS.findIndex((r) => !r.dimMember && r.korName === '[[부채총계]]');

      rawSheets.BS.forEach((item, idx) => {
        if (item.dimMember || item.isTotalRow || !item.korName) return;
        const tag = String(item.tag || '');
        const isTagCurr = /Current$/i.test(tag) && !/Noncurrent$|CurrentAndNoncurrent$/i.test(tag);
        const isTagNoncurr = /Noncurrent$/i.test(tag) && !/CurrentAndNoncurrent$/i.test(tag);
        const inCurrSec =
          (hasCaTotal && caIdx !== -1 && idx < caIdx) ||
          (hasClTotal && clIdx !== -1 && (assetsIdx === -1 || idx > assetsIdx) && idx < clIdx);
        const inNoncurrSec =
          (hasCaTotal && caIdx !== -1 && idx > caIdx && (assetsIdx === -1 || idx < assetsIdx)) ||
          (hasClTotal && clIdx !== -1 && idx > clIdx && (liabIdx === -1 || idx < liabIdx));
        if (!bsGroups.has(item.korName)) bsGroups.set(item.korName, []);
        bsGroups.get(item.korName).push({
          item,
          isCurr: isTagCurr || (!isTagNoncurr && inCurrSec),
          isNoncurr: isTagNoncurr || (!isTagCurr && inNoncurrSec),
        });
      });

      for (const entries of bsGroups.values()) {
        if (entries.length < 2) continue;
        const hasCurr = entries.some((e) => e.isCurr && !e.isNoncurr);
        const hasNoncurr = entries.some((e) => e.isNoncurr && !e.isCurr);
        if (!hasCurr || !hasNoncurr) continue;
        for (const { item, isCurr, isNoncurr } of entries) {
          if (isCurr && !isNoncurr && !/유동|단기/.test(item.korName)) {
            item.korName = `${item.korName}(유동)`;
          } else if (isNoncurr && !isCurr && !/비유동|장기/.test(item.korName)) {
            item.korName = /^단기/.test(item.korName)
              ? item.korName.replace(/^단기/, '장기')
              : `${item.korName}(비유동)`;
          }
        }
      }
    }

    // 6. 시트별 최종 행 조립 (MSA Safety 버전과 100% 동일)
    const sheetsData = { IS: [], BS: [], CF: [] };

    for (const key of ['IS', 'BS', 'CF']) {
      const mergedRows = [];
      let currentCfSection = 0;

      for (const item of rawSheets[key]) {
        const itemCfSection = key === 'CF' ? currentCfSection : 0;
        const isBold =
          item.isNonTotalOperatingExpense ||
          item.isNonTotalRevenue ||
          (item.dimMember && !item.isTotalRow)
            ? false
            : Boolean(item.stdRule?.bold || item.isTotalRow || /^\[.*\]$/.test(item.korName));

        const existingRow = mergedRows.find(
          (r) =>
            r.korName === item.korName &&
            r.rowType === item.rowType &&
            (r.dimMember || '') === (item.dimMember || '') &&
            (key !== 'CF' || r.cfSection === itemCfSection) &&
            r.values.some((v, i) => v === null && item.values[i] !== null) &&
            !r.values.some(
              (v, i) => v !== null && item.values[i] !== null && Math.abs(v - item.values[i]) > 5
            )
        );

        if (existingRow) {
          existingRow.values = existingRow.values.map((v, i) =>
            v !== null && v !== 0 ? v : item.values[i] !== null ? item.values[i] : v
          );
          existingRow.bold = existingRow.bold || isBold;
        } else {
          const sig = `${item.rowType}_${item.dimMember || ''}_${itemCfSection}_${item.korName}_${item.values.join(',')}`;
          if (
            !mergedRows.some(
              (r) =>
                `${r.rowType}_${r.dimMember || ''}_${r.cfSection ?? 0}_${r.korName}_${r.values.join(',')}` === sig
            )
          ) {
            mergedRows.push({
              rowType: item.rowType,
              dimMember: item.dimMember || '',
              cfSection: itemCfSection,
              korName: item.korName,
              isEps: item.isEps,
              bold: isBold,
              values: [...item.values],
            });
          }
        }

        if (
          key === 'CF' &&
          !item.dimMember &&
          item.isTotalRow &&
          /(?:Operating|Investing|Financing)Activities/i.test(item.tag || '')
        ) {
          currentCfSection++;
        }
      }

      sheetsData[key] = mergedRows;
    }

    // 7. ExcelJS 워크북 생성
    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'SEC 10-K Analyzer';
    workbook.created = new Date();

    const sheetConfigs = [
      { key: 'IS', title: '손익계산서' },
      { key: 'BS', title: '재무상태표' },
      { key: 'CF', title: '현금흐름표' },
    ];

    for (const cfg of sheetConfigs) {
      const ws = workbook.addWorksheet(cfg.title);
      ws.views = [{ state: 'frozen', xSplit: 1, ySplit: 1 }];

      const headerRow = ws.addRow(['구분', ...targetYears.map(String)]);
      headerRow.height = 24;
      headerRow.font = { bold: true, color: { argb: 'FF1E3A8A' } };
      headerRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFEFF6FF' },
      };
      headerRow.alignment = { horizontal: 'center', vertical: 'middle' };

      for (const item of sheetsData[cfg.key]) {
        const row = ws.addRow([item.korName, ...item.values]);
        if (item.bold) {
          row.font = { bold: true };
          row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: { argb: 'FFF8FAFC' },
          };
        }

        row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
          if (colNumber === 1) return;
          const v = cell.value;
          if (typeof v === 'number') {
            cell.numFmt = item.isEps
              ? '#,##0.00;(#,##0.00);"-"'
              : '#,##0;(#,##0);"-"';
            if (v < 0) {
              cell.font = {
                ...(item.bold ? { bold: true } : {}),
                color: { argb: 'FFDC2626' },
              };
            }
          }
        });
      }

      ws.columns.forEach((col, idx) => {
        let maxLen = idx === 0 ? 32 : 13;
        col.eachCell({ includeEmpty: true }, (cell) => {
          const text = cell.value ? String(cell.value) : '';
          const korCount = (text.match(/[가-힣]/g) || []).length;
          const visualLen = text.length + korCount * 0.8;
          if (visualLen > maxLen) maxLen = Math.min(visualLen + 4, 52);
        });
        col.width = Math.ceil(maxLen);
      });
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const startShort = String(targetYears[0]).slice(-2);
    const endShort = String(latestYear).slice(-2);
    const safeName = entityName.replace(/[\\/:*?"<>|]/g, '').trim();
    const fileName = `${safeName}_10-K_${startShort}_${endShort}.xlsx`;

    return new NextResponse(buffer, {
      status: 200,
      headers: {
        'Content-Type':
          'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        'X-File-Name': encodeURIComponent(fileName),
      },
    });
  } catch (error) {
    console.error('분석 오류:', error);
    return NextResponse.json(
      { error: error.message || '재무제표 생성 중 오류가 발생했습니다.' },
      { status: 500 }
    );
  }
}
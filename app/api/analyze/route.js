// app/api/analyze/route.js
import { NextResponse } from 'next/server';
import ExcelJS from 'exceljs';

export const maxDuration = 60;

// 복사 시 마크다운 URL 변환 깨짐 방지 헬퍼
function makeUrl(hostParts, pathPart) {
  return ['https:', '', hostParts.join('.'), pathPart].join('/');
}

const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

// 1. 재무상태표(BS)에서 제외할 순수 취득원가(Gross) 및 단독 감가상각누계액 태그
const BS_GROSS_AND_ACCUM_EXCLUDE = new Set([
  'PropertyPlantAndEquipmentGross',
  'AccumulatedDepreciationDepletionAndAmortizationPropertyPlantAndEquipment',
  'FiniteLivedIntangibleAssetsGross',
  'FiniteLivedIntangibleAssetsAccumulatedAmortization',
  'IntangibleAssetsGrossExcludingGoodwill',
  'AllowanceForDoubtfulAccountsReceivableCurrent',
]);

// 2. XBRL 1:1 동의어 태그 그룹 (연도별 XBRL 택소노미 개정 연결용)
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
  ['AccruedLiabilitiesCurrent', 'OtherAccruedLiabilitiesCurrent'],
  ['LongTermDebtCurrent', 'LongTermDebtAndCapitalLeaseObligationsCurrent'],
  ['LongTermDebtNoncurrent', 'LongTermDebtAndCapitalLeaseObligations', 'LongTermDebt'],
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

// 3. 표준 US-GAAP 태그 사전
const STANDARD_KOR_DICT = {
  // === 손익계산서 (IS) ===
  RevenueFromContractWithCustomerExcludingAssessedTax: { kor: '매출액', sign: 'raw', bold: true },
  RevenueFromContractWithCustomerIncludingAssessedTax: { kor: '매출액', sign: 'raw', bold: true },
  Revenues: { kor: '매출액', sign: 'raw', bold: true },
  SalesRevenueNet: { kor: '매출액', sign: 'raw', bold: true },
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
  RestructuringCharges: { kor: '매장폐쇄및기타비용(순액)', sign: 'raw' },
  RestructuringSettlementAndImpairmentProvisions: { kor: '매장폐쇄및기타비용(순액)', sign: 'raw' },
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
  StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest: {
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

// 4. [방법 2 XBRL 원리] R*.htm 하단에 내장된 XBRL 정의 테이블(authRefData)에서 정확한 DataType, Balance, PeriodType 추출
function parseEmbeddedXbrlAuthRefs(html) {
  const authMap = {};
  const blocks = html.split(/<table[^>]*class=["']authRefData["'][^>]*>/i);
  for (let i = 1; i < blocks.length; i++) {
    const chunk = blocks[i].slice(0, 2500);
    const idMatch =
      blocks[i - 1].slice(-300).match(/id=["'](defref_[^"']+)["']/i) ||
      chunk.match(/id=["'](defref_[^"']+)["']/i);
    const nameMatch = chunk.match(/<tr>\s*<td[^>]*>\s*Name:\s*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i);
    const typeMatch = chunk.match(/<tr>\s*<td[^>]*>\s*(?:Data\s+)?Type:\s*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i);
    const balMatch = chunk.match(/<tr>\s*<td[^>]*>\s*Balance(?:\s+Type)?:\s*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i);
    const perMatch = chunk.match(/<tr>\s*<td[^>]*>\s*Period(?:\s+Type)?:\s*<\/td>\s*<td[^>]*>([^<]+)<\/td>/i);

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

// 5. 라벨 내 가변 주석 문구 및 'net of accumulated depreciation of $...' 정제 함수
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

// 6. XBRL 내장 DataType(monetaryItemType | perShareItemType | sharesItemType) + 태그/문맥 2중 검증으로 행 속성 판별
function detectRowType(tag, cleanLabel, currentSection, stmtType, xbrlDataType = '') {
  if (/perShareItemType/i.test(xbrlDataType)) return 'EPS';
  if (/sharesItemType/i.test(xbrlDataType)) return 'SHARES';
  if (/monetaryItemType/i.test(xbrlDataType)) return 'USD';

  const t = String(tag || '');
  const lbl = String(cleanLabel || '').trim();
  const sec = String(currentSection || '').trim();

  if (
    /WeightedAverageNumberOf|SharesOutstanding|SharesIssued|AverageCommonShares|CommonStockShares/i.test(
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
      /weighted[- ]average.*shares|shares\s+used\s+in|average\s+common\s+shares|number\s+of\s+shares/i.test(
        lbl
      )
    ) {
      return 'SHARES';
    }
    if (/\bper\s+(?:common\s+|diluted\s+|basic\s+)?share\b/i.test(lbl)) {
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

// 7. 단순 밑줄(class="reu")이 아니라 실제 회계적 소계/총계(Subtotal/Total) 행인지 정확히 판별
function isTrueSubtotalRow(cleanLabel, tag, stmtType, rowType) {
  if (rowType !== 'USD') return false;
  const lbl = String(cleanLabel || '').trim();
  const t = String(tag || '');

  if (stmtType === 'BS') {
    if (/^total\s+/i.test(lbl)) return true;
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
      'StockholdersEquity',
      'StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest',
      'LiabilitiesAndStockholdersEquity',
    ].includes(t);
  }

  if (stmtType === 'CF') {
    if (
      /^net cash (?:provided by|used in)|^cash (?:flows? from|generated by|used in)|^(?:net )?increase \(decrease\) in cash/i.test(
        lbl
      )
    ) {
      return true;
    }
    return [
      'NetCashProvidedByUsedInOperatingActivities',
      'NetCashProvidedByUsedInInvestingActivities',
      'NetCashProvidedByUsedInFinancingActivities',
      'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalentsPeriodIncreaseDecreaseIncludingExchangeRateEffect',
      'CashAndCashEquivalentsPeriodIncreaseDecrease',
    ].includes(t);
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

// 8. 재무상태표에서 순액(Net) 자산은 보호하고 순수 취득원가(Gross)·단독 감가상각누계액 행만 제외
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

// 9. 상위 섹션 제목(sectionHeader)과 행 속성(rowType)을 결합하여 모호함 없는 문맥 라벨 생성
function buildContextualLabel(cleanLabel, sectionHeader, rowType, stmtType = '', cashFlowCashIndex = 0) {
  const lbl = stripDynamicParentheticals(cleanLabel).replace(/:$/, '').trim();
  const sec = stripDynamicParentheticals(sectionHeader || '').replace(/:$/, '').trim();

  if (rowType === 'EPS' && /^basic$/i.test(lbl)) return 'Earnings per share - Basic';
  if (rowType === 'EPS' && /^diluted$/i.test(lbl)) return 'Earnings per share - Diluted';
  if (rowType === 'SHARES' && /^basic$/i.test(lbl)) return 'Weighted average shares - Basic';
  if (rowType === 'SHARES' && /^diluted$/i.test(lbl)) return 'Weighted average shares - Diluted';

  if (stmtType === 'CF') {
    if (/beginning of (?:period|year)|at beginning of/i.test(lbl) || cashFlowCashIndex === 1) {
      return 'Cash and cash equivalents, beginning of period';
    }
    if (/end of (?:period|year)|at end of/i.test(lbl) || cashFlowCashIndex === 2) {
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

// 10. 확정적(Deterministic) 1:1 한글 계정명 매핑 함수
function matchExactKoreanAccountName(contextualLabel, tag = '', stmtType = '', rowType = 'USD') {
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

  const exactRules = [
    // 은행·금융·거래소(CME 등) 특화 항목
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

    // 재무상태표(BS - 일반 및 금융업 공통)
    [/^cash and due from banks$/i, '현금및타행예치금'],
    [/^interest-?bearing deposits with the federal reserve.*|^interest-?bearing deposits in banks$/i,
      '연방준비은행및타행이자부예치금'],
    [/^cash and cash equivalents,?\s*beginning of (?:period|year)$/i, '기초 현금및현금성자산'],
    [/^cash and cash equivalents,?\s*end of (?:period|year)$/i, '기말 현금및현금성자산'],
    [/^cash and cash equivalents$/i, '현금및현금성자산'],
    [/^restricted cash$/i, '사용제한현금'],
    [/^marketable securities$/i, '단기유가증권'],
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
    [/^deferred income tax liabilities,?\s*net$|^deferred income tax liabilities$|^deferred income taxes$/i,
      stmtType === 'BS' ? '이연법인세부채' : '이연법인세비용(혜택)'],
    [/^common stock$/i, '보통주자본금'],
    [/^preferred stock$/i, '우선주자본금'],
    [/^additional paid-?in capital$/i, '자본잉여금(보통주)'],
    [/^common stock and additional paid-?in capital$/i, '보통주자본금및자본잉여금'],
    [/^retained earnings$|^retained earnings \(accumulated deficit\)$|^accumulated deficit$/i,
      '이익잉여금(결손금)'],
    [/^treasury stock$|^treasury stock,?\s*at cost$/i, '자기주식'],
    [/^accumulated other comprehensive income \(loss\)$|^accumulated other comprehensive loss$/i,
      '기타포괄손익누계액'],

    // 현금흐름표(CF)
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
    [/^dividends paid to noncontrolling interests?$|^distributions to noncontrolling interests?$/i,
      '비지배지분 배당금의 지급'],
    [/^other financing activities,?\s*net$/i, '기타재무활동현금흐름(순)'],

    // 일반 제조·유통·IT 공통
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
    [/^other income \(loss\)$|^other income,?\s*net$|^other income \(expense\),?\s*net$|^other,?\s*net$/i,
      '기타손익(순)'],
    [/^income before income taxes$|^income \(loss\) before income taxes$|^income before provision for income taxes$|^earnings before income taxes$/i,
      '법인세비용차감전순이익'],
    [/^provision for income taxes$|^income tax provision$|^income tax expense$|^income taxes$/i, '법인세비용'],
    [/^net income$|^net earnings$|^net income \(loss\)$/i, '당기순이익'],
    [/^net increase \(decrease\) in cash and cash equivalents.*$/i, '현금및현금성자산의 순증감'],
    [/^effect of exchange rate changes on cash.*/i, '현금에 대한 환율변동효과'],
    [/^total current assets$/i, '[유동자산 합계]'],
    [/^total non-?current assets$|^total other assets$/i, '[비유동자산 합계]'],
    [/^total assets$/i, '[[자산총계]]'],
    [/^total current liabilities$/i, '[유동부채 합계]'],
    [/^total non-?current liabilities$|^total long-?term liabilities$/i, '[비유동부채 합계]'],
    [/^total liabilities$/i, '[[부채총계]]'],
    [/^total stockholders'? equity$|^total shareholders'? equity$|^total equity$/i, '[[자본총계]]'],
    [/^total .*?(?:stockholders|shareholders)'? equity$/i, '지배기업주주지분'],
    [/^total liabilities and (?:stockholders|shareholders)'? equity$|^total liabilities and equity$/i,
      '[[부채와자본총계]]'],
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

// 11. ★ 현금흐름표(CF) 내 명백한 현금 유출(Outflow: '-') / 현금 유입(Inflow: '+') 행 판별 함수
function classifyCashFlowRowPolarity(tag = '', contextualLabel = '', korName = '') {
  const t = String(tag || '');
  const lbl = String(contextualLabel || '').toLowerCase();
  const kor = String(korName || '');

  // 'ProceedsFromPaymentsFor...' 또는 'ProceedsFromRepaymentsOf...'처럼 유입·유출이 결합된 순증감 태그는 제외
  if (/ProceedsFrom(?:Payments|Repayments)|PaymentsForProceeds|IncreaseDecrease|NetChange/i.test(t)) {
    return 'NEUTRAL';
  }
  if (/\bnet\s+change\b|\bincrease\s*\(decrease\)|순증감|증감|기타.*순/i.test(`${lbl} ${kor}`)) {
    return 'NEUTRAL';
  }

  // 1) 명백한 현금 유출(Outflow -> 반드시 음수 '-')
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

  // 2) 명백한 현금 유입(Inflow -> 반드시 양수 '+')
  if (/^ProceedsFrom/i.test(t)) {
    return 'INFLOW';
  }
  if (/^(?:proceeds from|borrowings under|issuance of)/i.test(lbl)) {
    return 'INFLOW';
  }
  if (/현금유입|차입금의\s*(?:차입|유입)|행사에\s*따른\s*유입|매입제도로\s*인한\s*유입|만기회수|만기도래\s*및\s*조기상환/.test(kor)) {
    return 'INFLOW';
  }

  return 'NEUTRAL';
}

// 12. ★ 소계 항등식 역산을 통한 자동 부호 교정 엔진 (Reconciliation Sign Solver)
// 각 연도·섹션(예: 2025년 투자활동현금흐름)에서 구성 항목들의 합이 소계(예: -258)와 불일치할 때,
// 특정 항목의 부호(+ <-> -)를 반전시키면 소계와 정확히 일치(오차 <= 3)하는 경우 자동으로 부호를 교정!
function autoFixRowSignsBySubtotalReconciliation(rows, targetYears, stmtType) {
  if (stmtType !== 'CF' && stmtType !== 'BS') return;

  // 소계(Subtotal) 구간 인덱스 수집
  let prevTotalIdx = -1;
  for (let sIdx = 0; sIdx < rows.length; sIdx++) {
    if (!rows[sIdx].isTotalRow || rows[sIdx].rowType !== 'USD') continue;

    const subTotalRow = rows[sIdx];
    let compStartIdx = prevTotalIdx + 1;

    if (stmtType === 'CF' && prevTotalIdx === -1) {
      const netIncIdx = rows.findIndex(
        (r, idx) =>
          idx < sIdx &&
          (r.tag === 'NetIncomeLoss' ||
            r.tag === 'ProfitLoss' ||
            /^net income|^net earnings/i.test(r.companyLabel))
      );
      if (netIncIdx !== -1) compStartIdx = netIncIdx;
    }

    const compIndices = [];
    for (let i = compStartIdx; i < sIdx; i++) {
      if (rows[i].rowType === 'USD' && !rows[i].isTotalRow) {
        compIndices.push(i);
      }
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

        // 이미 소계가 일치하면 패스
        if (currentErr <= 3) continue;

        // 단일 항목(idx)의 부호를 반전(val -> -val)시켰을 때 소계가 완벽히 맞아떨어지는지 검사
        let bestFlipIdx = -1;
        let bestNewErr = currentErr;

        for (const idx of activeIndices) {
          const val = rows[idx].rValuesByYear[y];
          const flippedSum = currentSum - 2 * val;
          const flippedErr = Math.abs(flippedSum - subVal);

          if (flippedErr <= 3 && flippedErr < bestNewErr) {
            bestNewErr = flippedErr;
            bestFlipIdx = idx;
          }
        }

        if (bestFlipIdx !== -1) {
          rows[bestFlipIdx].rValuesByYear[y] = -rows[bestFlipIdx].rValuesByYear[y];
        }
      }
    }

    prevTotalIdx = sIdx;
  }
}

// 13. 동일 XBRL 태그·동일 연도 1:1 전수 대조 투표로 표 전체의 단일 화폐 스케일 확정
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
          if (rec.frame && new RegExp(`^CY${yr}(?:Q4)?I?$`).test(rec.frame)) return true;
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

// 14. CompanyFacts + Submissions(과거 아카이브 포함)에서 최근 10개년 10-K 공시 목록(accn) 완벽 추출
async function discoverAll10KFilings(cik, factsObj, headers) {
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
    const subUrl = makeUrl(['data', 'sec', 'gov'], `submissions/CIK${cik}.json`);
    const subRes = await fetch(subUrl, { headers });
    if (subRes.ok) {
      const subData = await subRes.json();
      const collectFromRecentObj = (recentObj) => {
        if (!recentObj || !Array.isArray(recentObj.form)) return;
        for (let i = 0; i < recentObj.form.length; i++) {
          if (recentObj.form[i] === '10-K') {
            const accn = recentObj.accessionNumber[i];
            const filed = recentObj.filingDate?.[i] || '';
            const repDate = recentObj.reportDate?.[i] || '';
            let repYear = repDate ? parseInt(repDate.slice(0, 4), 10) : parseInt(filed.slice(0, 4), 10) - 1;
            if (repDate && /^20\d{2}-01-(?:0[1-9]|1[0-5])$/.test(repDate)) repYear -= 1;
            if (accn && !accMap.has(accn)) {
              accMap.set(accn, { accn, filed, reportYear: repYear });
            }
          }
        }
      };

      collectFromRecentObj(subData?.filings?.recent);

      const archiveFiles = subData?.filings?.files || [];
      for (let f = 0; f < archiveFiles.length && accMap.size < 11; f++) {
        const fileName = archiveFiles[f]?.name;
        if (!fileName) continue;
        const archUrl = makeUrl(['data', 'sec', 'gov'], `submissions/${fileName}`);
        const archRes = await fetch(archUrl, { headers });
        if (archRes.ok) {
          const archData = await archRes.json();
          collectFromRecentObj(archData);
        }
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

// 15. 단일 10-K 공시의 FilingSummary.xml 및 R*.htm(내장 XBRL 메타데이터 포함)에서 3대 재무제표 표 데이터 추출
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
        !/Operations|Statements of Income and|Statements of Earnings and/i.test(shortName)
      ) {
        continue;
      }

      if (!statementFiles.IS && /Operations|Income|Earnings/i.test(shortName)) {
        statementFiles.IS = htmlFile;
      } else if (!statementFiles.BS && /Balance Sheet|Financial Position|Financial Condition/i.test(shortName)) {
        statementFiles.BS = htmlFile;
      } else if (!statementFiles.CF && /Cash Flow/i.test(shortName)) {
        statementFiles.CF = htmlFile;
      }
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
      for (const tr of trBlocks) {
        if (/class=["']nump?["']/i.test(tr)) break;
        if (!/<th[\s>]/i.test(tr)) continue;
        const thCells = tr.match(/<th[\s\S]*?<\/th>/gi) || [];
        const extractedCells = [];

        for (const th of thCells) {
          if (/class=["']tl["']/i.test(th)) continue;
          const text = th
            .replace(/<[^>]+>/g, ' ')
            .replace(/&#160;|&nbsp;/gi, ' ')
            .replace(/\s+/g, ' ')
            .trim();
          if (/^\d+\s*months?\s*ended$/i.test(text)) continue;
          extractedCells.push(text);
        }
        if (extractedCells.length > 0) {
          headerRows.push(extractedCells);
        }
      }

      let colYears = [];
      const yearRowIdx = headerRows.findLastIndex((r) => r.some((c) => /\b(20\d{2})\b/.test(c)));
      if (yearRowIdx !== -1) {
        const yearRow = headerRows[yearRowIdx];
        const prevRow =
          yearRowIdx > 0 && headerRows[yearRowIdx - 1].length === yearRow.length
            ? headerRows[yearRowIdx - 1]
            : null;

        for (let c = 0; c < yearRow.length; c++) {
          const fullDateText = prevRow ? `${prevRow[c]} ${yearRow[c]}` : yearRow[c];
          const yrMatch = fullDateText.match(/\b(20\d{2})\b/);
          if (yrMatch) {
            let yr = parseInt(yrMatch[1], 10);
            if (/\bJan(?:uary)?\.?\s*(?:0?[1-9]|1[0-5])\b/i.test(fullDateText)) {
              yr -= 1;
            }
            colYears.push(yr);
          }
        }
      }

      for (let i = 1; i < colYears.length; i++) {
        if (colYears[i] === colYears[i - 1]) {
          colYears[i] = colYears[i - 1] - 1;
        }
      }

      const rows = [];
      let currentSection = '';
      let cfCashBalanceCount = 0;

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
        rawLabel = rawLabel
          .replace(/<[^>]+>/g, '')
          .replace(/&amp;/g, '&')
          .replace(/&#160;|&nbsp;/gi, ' ')
          .replace(/\s+/g, ' ')
          .trim();

        const hasNumericCell = /class=["']nump?["']/i.test(tr);

        if (!hasNumericCell) {
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
        const tag = fullDefMatch ? fullDefMatch[3] : '';
        const dimKey = fullDefMatch && fullDefMatch[4] ? fullDefMatch[4] : '';

        if (isExcludedGrossOrAccumRow(stmtType, tag, rawLabel)) {
          continue;
        }

        const cleanLabel = stripDynamicParentheticals(rawLabel || tag);
        if (!cleanLabel) continue;

        let thisCashIndex = 0;
        if (
          stmtType === 'CF' &&
          (tag === 'CashAndCashEquivalentsAtCarryingValue' ||
            tag === 'CashCashEquivalentsRestrictedCashAndRestrictedCashEquivalents' ||
            /^cash and cash equivalents/i.test(cleanLabel))
        ) {
          cfCashBalanceCount++;
          thisCashIndex = cfCashBalanceCount;
        }

        const authMeta =
          xbrlAuthMap[fullDefKey] || xbrlAuthMap[`defref_${prefix}_${tag}`] || {};

        const rowType = detectRowType(tag, cleanLabel, currentSection, stmtType, authMeta.dataType);
        const contextualLabel = buildContextualLabel(
          cleanLabel,
          currentSection,
          rowType,
          stmtType,
          thisCashIndex
        );

        const valueCells =
          tr.match(/<td[^>]*class=["'](?:nump|num|text)["'][^>]*>[\s\S]*?<\/td>/gi) || [];

        const rawNumbersByYear = {};
        let hasAnyNumber = false;

        for (let i = 0; i < Math.min(colYears.length, valueCells.length); i++) {
          const yr = colYears[i];
          const td = valueCells[i];
          const rawText = td
            .replace(/<[^>]+>/g, '')
            .replace(/&#160;|&nbsp;|\$|,/gi, '')
            .trim();

          if (!rawText || rawText === '-' || rawText === '—' || rawText === '–') continue;

          // ★ 1차 방어: 닫는 괄호 ')'가 옆 각주 셀(<td class="fn">)로 분리되어 있어도 여는 괄호 '('만 있으면 100% 음수로 인식!
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

        const isTotalRow = isTrueSubtotalRow(cleanLabel, tag, stmtType, rowType);

        if (isTotalRow && rowType === 'USD') {
          currentSection = '';
        }

        const isCashBalanceInCF = stmtType === 'CF' && thisCashIndex > 0;

        rows.push({
          prefix,
          tag,
          dimKey: isCashBalanceInCF ? `CF_CASH_${thisCashIndex}` : dimKey,
          isMainRow: Boolean(tag && !dimKey && !isCashBalanceInCF),
          rowType,
          isEps: rowType === 'EPS',
          companyLabel: cleanLabel,
          contextualLabel,
          normLabel: normalizeLabel(cleanLabel),
          normContextualLabel: normalizeLabel(contextualLabel),
          isTotalRow,
          rawNumbersByYear,
          rValuesByYear: {},
        });

        if (
          stmtType === 'BS' &&
          (tag === 'LiabilitiesAndStockholdersEquity' ||
            /^total liabilities and (?:stockholders|shareholders)'? equity$|^total liabilities and equity$/i.test(
              cleanLabel
            ))
        ) {
          break;
        }
      }

      const verifiedDollarScale = determineExactTableScaleFromFacts(
        rows,
        colYears,
        factsObj,
        accessionNumber,
        headerDefaultScale
      );

      for (const r of rows) {
        const polarity =
          stmtType === 'CF'
            ? classifyCashFlowRowPolarity(r.tag, r.contextualLabel, '')
            : 'NEUTRAL';

        for (const [yr, signedNum] of Object.entries(r.rawNumbersByYear)) {
          if (r.rowType === 'EPS') {
            r.rValuesByYear[yr] = Number(signedNum.toFixed(2));
          } else if (r.rowType === 'SHARES') {
            let scaledShares = signedNum * shareScaleToMillions;
            if (Math.abs(scaledShares) >= 40000 && Math.abs(scaledShares) < 50000000) {
              scaledShares = scaledShares * 0.001;
            } else if (Math.abs(scaledShares) >= 50000000) {
              scaledShares = scaledShares * 0.000001;
            }
            r.rValuesByYear[yr] = Math.round(scaledShares);
          } else {
            let val = Math.round(signedNum * verifiedDollarScale);
            if (val !== 0 && polarity === 'OUTFLOW') val = -Math.abs(val);
            else if (val !== 0 && polarity === 'INFLOW') val = Math.abs(val);
            r.rValuesByYear[yr] = val;
          }
        }
        delete r.rawNumbersByYear;
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

// 16. 단일 연도(yr) 기준으로 특정 행의 값(candidateVal)이 이미 다른 행들에 의해 소계가 완성되어 중복인지 판별
function isSingleYearRedundantByReconciliation(
  rows,
  targetPos,
  yr,
  candidateVal,
  stmtType,
  isTargetAlreadyInRows = false
) {
  if (!['IS', 'BS', 'CF'].includes(stmtType)) return false;
  if (candidateVal === undefined || candidateVal === null || candidateVal === 0) return false;

  let subTotalIdx = -1;
  const startSearchIdx = isTargetAlreadyInRows ? targetPos + 1 : targetPos;
  for (let i = startSearchIdx; i < rows.length; i++) {
    if (rows[i].isTotalRow && rows[i].rowType === 'USD') {
      subTotalIdx = i;
      break;
    }
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

  let compStartIdx = prevTotalIdx + 1;
  if (stmtType === 'CF' && prevTotalIdx === -1) {
    const netIncIdx = rows.findIndex(
      (r, idx) =>
        idx < subTotalIdx &&
        (r.tag === 'NetIncomeLoss' ||
          r.tag === 'ProfitLoss' ||
          /^net income|^net earnings/i.test(r.companyLabel))
    );
    if (netIncIdx !== -1) compStartIdx = netIncIdx;
  }

  const otherComponentRows = [];
  for (let i = compStartIdx; i < subTotalIdx; i++) {
    if (isTargetAlreadyInRows && i === targetPos) continue;
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
      subTotalRow.tag === 'Assets' ||
      subTotalRow.tag === 'Liabilities' ||
      /^total assets$|^total liabilities$/i.test(subTotalRow.companyLabel);

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

// 17. 5단계 멀티패스 매칭 + 연도별(Per-Year) 소계 항등식 필터링 엔진
async function fetchMultiYear10KStructure(cik, factsObj, headers) {
  try {
    const all10Ks = await discoverAll10KFilings(cik, factsObj, headers);
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

    if (parsedFilings.length === 0) return null;

    for (const f of parsedFilings) {
      for (const stmtType of ['IS', 'BS', 'CF']) {
        for (const r of f[stmtType] || []) {
          r.exactKor =
            matchExactKoreanAccountName(r.contextualLabel, r.tag, stmtType, r.rowType) ||
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

      for (let fIdx = 1; fIdx < parsedFilings.length; fIdx++) {
        const olderRows = parsedFilings[fIdx][stmtType] || [];
        const usedAnchorIndices = new Set();
        const olderToAnchorMap = new Map();

        const lockMatch = (oIdx, aIdx) => {
          usedAnchorIndices.add(aIdx);
          olderToAnchorMap.set(oIdx, aIdx);
          const targetAnchor = anchorRows[aIdx];
          const sourceOlder = olderRows[oIdx];
          for (const [yr, val] of Object.entries(sourceOlder.rValuesByYear)) {
            if (targetAnchor.rValuesByYear[yr] === undefined) {
              targetAnchor.rValuesByYear[yr] = val;
            }
          }
        };

        // [Pass 1] 태그 + 차원(dimKey) + 문맥라벨 + rowType 완전 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.tag) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              ar.rowType === or.rowType &&
              ar.tag === or.tag &&
              ar.dimKey === or.dimKey &&
              ar.normContextualLabel === or.normContextualLabel
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
              ar.rowType === or.rowType &&
              ar.isMainRow &&
              ar.tag === or.tag
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 3] 문맥 포함 정규화 영문 라벨(normContextualLabel) + rowType 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.normContextualLabel) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              ar.rowType === or.rowType &&
              ar.normContextualLabel === or.normContextualLabel
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 4] 확정 한글 계정명(exactKor) + rowType 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.exactKor) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              ar.rowType === or.rowType &&
              ar.exactKor === or.exactKor
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Pass 5] 과거 동의어 태그 그룹(HISTORICAL_SYNONYM_GROUPS) + rowType 일치
        for (let oIdx = 0; oIdx < olderRows.length; oIdx++) {
          if (olderToAnchorMap.has(oIdx)) continue;
          const or = olderRows[oIdx];
          if (!or.isMainRow || !or.tag) continue;
          const synGroup = HISTORICAL_SYNONYM_GROUPS.find((g) => g.includes(or.tag));
          if (!synGroup) continue;
          const aIdx = anchorRows.findIndex(
            (ar, idx) =>
              !usedAnchorIndices.has(idx) &&
              ar.rowType === or.rowType &&
              ar.isMainRow &&
              synGroup.includes(ar.tag)
          );
          if (aIdx !== -1) lockMatch(oIdx, aIdx);
        }

        // [Insert Pass] Pass 1~5에서 매칭되지 않은 과거 행 처리 (연도별 오차 비교 후 필요한 연도만 정밀 삽입!)
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

          const filteredValuesByYear = {};
          for (const [yr, val] of Object.entries(or.rValuesByYear)) {
            if (
              or.rowType === 'USD' &&
              !or.isTotalRow &&
              isSingleYearRedundantByReconciliation(anchorRows, insertPos, yr, val, stmtType, false)
            ) {
              continue;
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
      }
    }

    return baseStructure;
  } catch (e) {
    return null;
  }
}

// 18. 최근 보고된 10-K 기준 과거 10개년도 범위(과거 -> 최신순) 산출
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
        if (r.frame) {
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

// 19. CompanyFacts JSON에서 10개년 수치 추출
function extract10YearSeriesForTag(factsObj, prefix, tag, targetYears, stmtType, rowType, rValuesByYear = {}) {
  if (!tag) return null;
  const usGaap = factsObj?.['us-gaap'] || {};
  const customTax = factsObj?.[prefix] || {};

  const synonymGroup = HISTORICAL_SYNONYM_GROUPS.find((grp) => grp.includes(tag));
  const candidateTags = synonymGroup
    ? [tag, ...synonymGroup.filter((t) => t !== tag)]
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

      if (r.frame) {
        const m = r.frame.match(/^CY(\d{4})(Q\d)?(I)?$/);
        if (m) {
          const frameYear = parseInt(m[1], 10);
          const qStr = m[2];
          if (qStr && qStr !== 'Q4') return false;
          return frameYear === year;
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

  // 전체 동의어 그룹 통합 부호 투표 (개별 태그에 겹치는 연도가 없더라도 그룹 전체의 R*.htm 부호를 공유)
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

// 20. 사전에 없는 계정들의 한글 번역 및 구형->신형 동일 계정 매핑
async function alignAndTranslateWithGemini(itemsToTranslate, allSheetsRows, apiKey) {
  if (!itemsToTranslate || itemsToTranslate.length === 0) return {};

  const indexedPayload = {};
  itemsToTranslate.forEach((item, idx) => {
    indexedPayload[String(idx)] = {
      stmt: item.stmtType,
      label: item.contextualLabel,
      tag: item.tag,
    };
  });

  const prompt = `당신은 미국 SEC 10-K 재무제표 전문 한국어 공인회계사입니다.
아래 JSON 객체의 영문 재무제표 계정명(label, tag)들을 한국 실무 회계 기준(K-IFRS / 금융·제조·유통 회계기준)에 맞는 정확하고 자연스러운 **순수 한글 계정명**으로 번역하세요.
- 숫자 Key("0", "1"...)를 그대로 유지하세요.
- 영어 알파벳이나 괄호 안의 달러($) 숫자를 남기지 말고, 원래 영문 회계 계정의 의미를 정확히 살린 한글 계정명(문자열)으로 응답하세요.
- 연도별로 표현만 살짝 바뀐 동일 성격의 계정(예: Long-term debt and capital lease obligations 와 Long-term debt and other financing obligations)은 동일한 표준 한글 계정명('장기차입금및금융리스부채')으로 통일하세요.

응답 예시:
{"0": "시장조성및유사활동수익", "1": "매도가능채무증권의 취득"}

[번역할 계정 목록]
${JSON.stringify(indexedPayload)}`;

  let candidateModels = ['gemini-3.0-flash', 'gemini-2.5-flash', 'gemini-2.0-flash'];
  try {
    const listUrl = makeUrl(
      ['generativelanguage', 'googleapis', 'com'],
      `v1beta/models?key=${apiKey}`
    );
    const listRes = await fetch(listUrl);
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

      const flashModels = available.filter((n) => /flash/i.test(n) && !/8b|image|tts|audio/i.test(n));
      if (flashModels.length > 0) {
        candidateModels = [...new Set([...flashModels, ...candidateModels])];
      }
    }
  } catch (e) {}

  for (const modelName of candidateModels) {
    const url = makeUrl(
      ['generativelanguage', 'googleapis', 'com'],
      `v1beta/models/${modelName}:generateContent?key=${apiKey}`
    );

    try {
      const response = await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          contents: [{ parts: [{ text: prompt }] }],
          generationConfig: { temperature: 0.1 },
        }),
      });

      if (!response.ok) continue;

      const data = await response.json();
      const rawText = data?.candidates?.[0]?.content?.parts?.[0]?.text || '{}';
      const jsonMatch = rawText.match(/\{[\s\S]*\}/);
      if (!jsonMatch) continue;

      const parsed = JSON.parse(jsonMatch[0]);
      const result = {};
      itemsToTranslate.forEach((item, idx) => {
        const val = parsed[String(idx)];
        const kor = typeof val === 'string' ? val : val?.kor;
        if (kor && typeof kor === 'string') {
          result[item.rowKey] = kor.trim();
        }
      });

      if (Object.keys(result).length > 0) return result;
    } catch (e) {
      continue;
    }
  }

  return {};
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

    const cik = String(matched.cik_str).padStart(10, '0');
    const entityName = matched.title || matched.ticker;

    // 2. CompanyFacts JSON 로드 후 10년 치 10-K R*.htm 전체 병합
    const factsUrl = makeUrl(['data', 'sec', 'gov'], `api/xbrl/companyfacts/CIK${cik}.json`);
    const factsRes = await fetch(factsUrl, { headers });
    if (!factsRes.ok) {
      throw new Error(`SEC 재무 데이터 조회 실패 (상태 코드: ${factsRes.status})`);
    }

    const factsData = await factsRes.json();
    const factsObj = factsData?.facts || {};
    const usGaap = factsObj['us-gaap'] || {};

    const actualStructure = await fetchMultiYear10KStructure(cik, factsObj, headers);

    // 3. 최근 보고된 10-K 기준 10개년도 범위 산출 (과거 -> 최신순)
    const targetYears = determineTenYearWindow(usGaap, actualStructure);
    const latestYear = targetYears[targetYears.length - 1];

    // 4. 다중 10-K R*.htm 직접 추출 수치를 기본으로 하고, 연도별 오차 검증을 거쳐 CompanyFacts 보완
    const buildFinalSeriesForSheet = (stmtType) => {
      const rRows = actualStructure?.[stmtType] || [];
      const resultRows = [];

      const enrichedRows = rRows.map((rRow, idx) => {
        const rowKey = `${stmtType}_${idx}_${rRow.tag}_${rRow.normContextualLabel}`;
        const stdRule = STANDARD_KOR_DICT[rRow.tag] || null;
        const polarity =
          stmtType === 'CF'
            ? classifyCashFlowRowPolarity(rRow.tag, rRow.contextualLabel, rRow.exactKor || stdRule?.kor || '')
            : 'NEUTRAL';

        const cfSeries = rRow.isMainRow
          ? extract10YearSeriesForTag(
              factsObj,
              rRow.prefix,
              rRow.tag,
              targetYears,
              stmtType,
              rRow.rowType,
              rRow.rValuesByYear
            )
          : null;

        const mergedYearMap = { ...(rRow.rValuesByYear || {}) };

        for (const y of targetYears) {
          if (mergedYearMap[y] !== undefined) continue;

          const rawFactVal = cfSeries?.valuesByYear?.[y];
          if (rawFactVal === null || rawFactVal === undefined) continue;

          let candidateVal;
          if (rRow.rowType === 'EPS') {
            candidateVal = Number(rawFactVal.toFixed(2));
          } else if (rRow.rowType === 'SHARES') {
            let s = rawFactVal;
            if (Math.abs(s) >= 50000000) s = s / 1e6;
            else if (Math.abs(s) >= 40000) s = s / 1e3;
            candidateVal = Math.round(s);
          } else {
            const scaled = Math.round(rawFactVal / 1e6);
            if (scaled === 0) {
              candidateVal = 0;
            } else if (stdRule?.sign === 'neg' || polarity === 'OUTFLOW') {
              candidateVal = -Math.abs(scaled);
            } else if (stdRule?.sign === 'pos' || polarity === 'INFLOW') {
              candidateVal = Math.abs(scaled);
            } else {
              candidateVal = scaled;
            }
          }

          if (
            rRow.rowType === 'USD' &&
            !rRow.isTotalRow &&
            isSingleYearRedundantByReconciliation(rRows, idx, String(y), candidateVal, stmtType, true)
          ) {
            continue;
          }

          mergedYearMap[y] = candidateVal;
        }

        // ★ 2차 방어: 현금흐름표의 명백한 유출(OUTFLOW) 또는 유입(INFLOW) 행은 모든 연도에 대해 부호 일관성 강제!
        if (stmtType === 'CF' && rRow.rowType === 'USD') {
          for (const y of targetYears) {
            const v = mergedYearMap[y];
            if (v === undefined || v === null || v === 0) continue;
            if (stdRule?.sign === 'neg' || polarity === 'OUTFLOW') {
              mergedYearMap[y] = -Math.abs(v);
            } else if (stdRule?.sign === 'pos' || polarity === 'INFLOW') {
              mergedYearMap[y] = Math.abs(v);
            }
          }
        }

        return {
          ...rRow,
          rowKey,
          stdRule,
          rValuesByYear: mergedYearMap,
        };
      });

      // 과거 전용 행의 연도별 중복 셀 정리
      for (let idx = 0; idx < enrichedRows.length; idx++) {
        const rRow = enrichedRows[idx];
        const hasRecentPresence = targetYears
          .slice(-2)
          .some((y) => rRow.rValuesByYear[y] !== undefined);

        if (!hasRecentPresence && !rRow.isTotalRow && rRow.rowType === 'USD') {
          for (const y of targetYears) {
            const val = rRow.rValuesByYear[y];
            if (val === undefined || val === null) continue;
            if (
              isSingleYearRedundantByReconciliation(
                enrichedRows,
                idx,
                String(y),
                val,
                stmtType,
                true
              )
            ) {
              delete rRow.rValuesByYear[y];
            }
          }
        }
      }

      // ★ 3차 방어: 소계 항등식 역산을 통한 자동 부호 교정 (예: MSA 2025년 투자활동현금흐름 +189 -> -189 자동 교정!)
      autoFixRowSignsBySubtotalReconciliation(enrichedRows, targetYears, stmtType);

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

    // 5. 한글 계정명 결정 (AI 번역 완료 후 한글 계정명 기준으로도 유출/유입 부호 일관성 재확인)
    const needsAiList = [];

    for (const key of ['IS', 'BS', 'CF']) {
      for (const item of rawSheets[key]) {
        if (item.exactKor) {
          item.korName = item.exactKor;
        } else if (item.stdRule?.kor) {
          item.korName = item.stdRule.kor;
        } else {
          needsAiList.push(item);
        }
      }
    }

    if (needsAiList.length > 0) {
      const aiMap = await alignAndTranslateWithGemini(needsAiList, rawSheets, apiKey.trim());
      for (const item of needsAiList) {
        item.korName = aiMap[item.rowKey] || item.stdRule?.kor || item.contextualLabel;
      }
    }

    // AI 번역된 한글 계정명 기준으로 현금흐름표 유출/유입 부호 최종 확인
    for (const item of rawSheets.CF) {
      if (item.rowType !== 'USD' || item.isTotalRow) continue;
      const pol = classifyCashFlowRowPolarity(item.tag, item.contextualLabel, item.korName);
      if (pol === 'OUTFLOW') {
        item.values = item.values.map((v) => (v === null || v === 0 ? v : -Math.abs(v)));
      } else if (pol === 'INFLOW') {
        item.values = item.values.map((v) => (v === null || v === 0 ? v : Math.abs(v)));
      }
    }

    // 6. 시트별 최종 행 조립 (동일 rowType 및 동일 한글명을 가진 상호보완 연도 행 자동 병합)
    const sheetsData = { IS: [], BS: [], CF: [] };

    for (const key of ['IS', 'BS', 'CF']) {
      const mergedRows = [];

      for (const item of rawSheets[key]) {
        const isBold = Boolean(item.stdRule?.bold || item.isTotalRow);

        const existingRow = mergedRows.find(
          (r) =>
            r.korName === item.korName &&
            r.rowType === item.rowType &&
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
          const sig = `${item.rowType}_${item.korName}_${item.values.join(',')}`;
          if (!mergedRows.some((r) => `${r.rowType}_${r.korName}_${r.values.join(',')}` === sig)) {
            mergedRows.push({
              rowType: item.rowType,
              korName: item.korName,
              isEps: item.isEps,
              bold: isBold,
              values: [...item.values],
            });
          }
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
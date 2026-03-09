import * as XLSX from 'xlsx';
import './style.css';

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Target columns for formula fill Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
const FORMULA_COLUMNS: string[] = [
  'Total RePaid', 'MonthlyShare', 'TotalAdvance', 'Shares C/F', 'Loans C/F', 'TotalCash',
];

const PAYMENT_COLS = [
  'Cash',
  'Paybill',
  'Bank',
  'LoanRepayment',
  'AdvanceRepayment',
  'RiskFund',
  'Fine',
  'FineDeduction',
  'ShareDeduction',
  'AdvanceDeduction',
  'RiskFundOut',
] as const;
type PaymentCol = typeof PAYMENT_COLS[number];
const LIVE_RESULT_COLS = ['MonthlyShare', 'Shares C/F', 'Loans C/F', 'TotalCash', 'TotalAdvance', 'Total RePaid'] as const;
type LiveResultCol = typeof LIVE_RESULT_COLS[number];
type RecordTab = 'savings' | 'loan';

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ DOM refs Ã¢â‚¬â€ Step 1 Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
const dropZone = document.getElementById('dropZone')! as HTMLDivElement;
const fileInput = document.getElementById('fileInput')! as HTMLInputElement;
const browseBtn = document.getElementById('browseBtn')! as HTMLButtonElement;
const fileInfo = document.getElementById('fileInfo')! as HTMLDivElement;
const fileNameEl = document.getElementById('fileName')! as HTMLParagraphElement;
const fileSizeEl = document.getElementById('fileSize')! as HTMLParagraphElement;
const clearBtn = document.getElementById('clearBtn')! as HTMLButtonElement;
const processBtn = document.getElementById('processBtn')! as HTMLButtonElement;
const btnText = processBtn.querySelector('.btn-text')! as HTMLSpanElement;
const spinner = document.getElementById('spinner')! as HTMLSpanElement;
const logCard = document.getElementById('logCard')! as HTMLDivElement;
const logIcon = document.getElementById('logIcon')! as HTMLSpanElement;
const logTitle = document.getElementById('logTitle')! as HTMLSpanElement;
const logList = document.getElementById('logList')! as HTMLUListElement;
const formulaBanner = document.getElementById('formulaBanner')! as HTMLDivElement;

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ DOM refs Ã¢â‚¬â€ Step 2 Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
const step2Section = document.getElementById('step2Section')! as HTMLElement;
const memberTableBody = document.getElementById('memberTableBody')! as HTMLTableSectionElement;
const applyBtn = document.getElementById('applyBtn')! as HTMLButtonElement;
const startOverBtn = document.getElementById('startOverBtn')! as HTMLButtonElement;
const contribSummary = document.getElementById('contribSummary')! as HTMLDivElement;
const memberCountEl = document.getElementById('memberCount')! as HTMLSpanElement;
const colTotalExpected = document.getElementById('colTotalExpected')!;
const colTotalCash = document.getElementById('colTotalCash')!;
const colTotalPaybill = document.getElementById('colTotalPaybill')!;
const colTotalBank = document.getElementById('colTotalBank')!;

// Step indicator elements
const step1Ind = document.getElementById('step1Indicator')! as HTMLDivElement;
const step2Ind = document.getElementById('step2Indicator')! as HTMLDivElement;
const step3Ind = document.getElementById('step3Indicator')! as HTMLDivElement;

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Detail Popup refs Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
const detailOverlay = document.getElementById('detailOverlay')! as HTMLDivElement;
const popupAvatar = document.getElementById('popupAvatar')! as HTMLDivElement;
const popupMemberNo = document.getElementById('popupMemberNo')! as HTMLParagraphElement;
const popupMemberName = document.getElementById('popupMemberName')! as HTMLParagraphElement;
const popupExpected = document.getElementById('popupExpected')! as HTMLSpanElement;
const popupPrincipal = document.getElementById('popupPrincipal')! as HTMLSpanElement;
const popupInstallment = document.getElementById('popupInstallment')! as HTMLSpanElement;
const popupTotalShares = document.getElementById('popupTotalShares')! as HTMLSpanElement;
const popupLoanBalance = document.getElementById('popupLoanBalance')! as HTMLSpanElement;
const popupShareLoanDiff = document.getElementById('popupShareLoanDiff')! as HTMLSpanElement;
const popupLoanInterest = document.getElementById('popupLoanInterest')! as HTMLSpanElement;
const popupAdvBalance = document.getElementById('popupAdvBalance')! as HTMLSpanElement;
const popupAdvInterest = document.getElementById('popupAdvInterest')! as HTMLSpanElement;
const popupMShare = document.getElementById('popupMShare')! as HTMLSpanElement;
const popupLoanRepayment = document.getElementById('popupLoanRepayment')! as HTMLInputElement;
const popupAdvRepayment = document.getElementById('popupAdvRepayment')! as HTMLInputElement;
const popupRiskFund = document.getElementById('popupRiskFund')! as HTMLInputElement;
const popupFine = document.getElementById('popupFine')! as HTMLInputElement;
const popupFineDeduction = document.getElementById('popupFineDeduction')! as HTMLInputElement;
const popupShareDeduction = document.getElementById('popupShareDeduction')! as HTMLInputElement;
const popupAdvanceDeduction = document.getElementById('popupAdvanceDeduction')! as HTMLInputElement;
const popupRiskFundOut = document.getElementById('popupRiskFundOut')! as HTMLInputElement;
const popupTotalCash = document.getElementById('popupTotalCash')! as HTMLSpanElement;
const popupTotalAdvance = document.getElementById('popupTotalAdvance')! as HTMLSpanElement;
const popupTotalRepaid = document.getElementById('popupTotalRepaid')! as HTMLSpanElement;
const popupMShareResult = document.getElementById('popupMShareResult')! as HTMLSpanElement;
const popupClose = document.getElementById('popupClose')! as HTMLButtonElement;
const popupDismiss = document.getElementById('popupDismiss')! as HTMLButtonElement;

// Record popup refs
const recordOverlay = document.getElementById('recordOverlay')! as HTMLDivElement;
const recordAvatar = document.getElementById('recordAvatar')! as HTMLDivElement;
const recordMemberNo = document.getElementById('recordMemberNo')! as HTMLParagraphElement;
const recordMemberName = document.getElementById('recordMemberName')! as HTMLParagraphElement;
const savingsTabBtn = document.getElementById('savingsTabBtn')! as HTMLButtonElement;
const loanRecordTabBtn = document.getElementById('loanRecordTabBtn')! as HTMLButtonElement;
const savingsPanel = document.getElementById('savingsPanel')! as HTMLDivElement;
const loanRecordPanel = document.getElementById('loanRecordPanel')! as HTMLDivElement;
const recordSavingsValue = document.getElementById('recordSavingsValue')! as HTMLSpanElement;
const recordSavingsBalance = document.getElementById('recordSavingsBalance')! as HTMLSpanElement;
const recordLoanPrincipal = document.getElementById('recordLoanPrincipal')! as HTMLSpanElement;
const recordLoanInterest = document.getElementById('recordLoanInterest')! as HTMLSpanElement;
const recordLoanBalance = document.getElementById('recordLoanBalance')! as HTMLSpanElement;
const recordClose = document.getElementById('recordClose')! as HTMLButtonElement;
const recordDismiss = document.getElementById('recordDismiss')! as HTMLButtonElement;

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ State Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
let selectedFile: File | null = null;
let processedWorkbook: XLSX.WorkBook | null = null;
let processedSheetName = '';
let outputName = 'modified.xlsx';

interface MemberRow {
  memberNo: string | number;
  memberName: string;
  memberStatus: string;
  isInactive: boolean;
  rowIdx: number;
  hasAlertColor: boolean;
  expected: number;
  principal: number;
  installment: number;
  loanBalance: number;
  loanInterest: number;
  advanceBalance: number;
  advanceInterest: number;
  monthlyShare: number;      // from Excel (for display)
  totalShares: number;
  shareTransfer: number;
  fineDeduction: number;
  shareDeduction: number;
  riskFundOut: number;
  shareOut: number;
  nonCashOut: number;
  sharesBalanceBase: number;
  sharesBalance: number;
  loansBalanceBase: number;
  loansBalance: number;
  loanRepayment: number;
  advRepayment: number;
  // Occasional fields from Excel
  advInterestPaid: number;
  advDeduction: number;
  loanInterestPaid: number;
  registrationFee: number;
  passBook: number;
  fine: number;
}

/** Live map of entered amounts Ã¢â‚¬â€ includes riskFund per member */
interface PaymentEntry {
  cash: number; paybill: number; bank: number;
  loanRepayment: number; advRepayment: number; riskFund: number;
  fine: number;
  fineDeduction: number;
  shareDeduction: number;
  advanceDeduction: number;
  riskFundOut: number;
  reducingInterest: number;
  monthlyShare: number;
  sharesBalance: number;
  loansBalance: number;
}

const paymentMap = new Map<number, PaymentEntry>();
const memberRowMap = new Map<number, MemberRow>();
let activeRecordMember: MemberRow | null = null;

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Helpers Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1048576).toFixed(2)} MB`;
}

function fmt(n: number): string {
  return n.toLocaleString('en-KE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

const MONEY_PLACEHOLDER = 'KSh 0.00';
const popupEditableInputs = [
  popupRiskFund,
  popupLoanRepayment,
  popupAdvRepayment,
  popupFine,
  popupFineDeduction,
  popupShareDeduction,
  popupAdvanceDeduction,
  popupRiskFundOut,
];

function syncBodyScrollLock() {
  const overlayOpen = !detailOverlay.classList.contains('hidden') || !recordOverlay.classList.contains('hidden');
  document.body.style.overflow = overlayOpen ? 'hidden' : '';
}

function getHeaderIndex(colMap: Record<string, number>, headers: string[]): number | undefined {
  return headers.find((header) => colMap[header] !== undefined)
    ? colMap[headers.find((header) => colMap[header] !== undefined)!]
    : undefined;
}

function normalizeColorHex(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const hex = value.trim().replace(/^#/, '').toUpperCase();
  if (/^[0-9A-F]{8}$/.test(hex)) return hex.slice(2);
  if (/^[0-9A-F]{6}$/.test(hex)) return hex;
  return null;
}

function isAlertHex(hex: string): boolean {
  const red = parseInt(hex.slice(0, 2), 16);
  const green = parseInt(hex.slice(2, 4), 16);
  const blue = parseInt(hex.slice(4, 6), 16);
  return red >= 150 && red >= green + 35 && red >= blue + 35;
}

function cellHasAlertColor(cell: XLSX.CellObject | undefined): boolean {
  const style = (cell as XLSX.CellObject & {
    s?: {
      font?: { color?: { rgb?: string } };
      fill?: {
        fgColor?: { rgb?: string };
        bgColor?: { rgb?: string };
      };
      color?: { rgb?: string };
    };
  })?.s;
  if (!style) return false;

  const candidates = [
    style.font?.color?.rgb,
    style.fill?.fgColor?.rgb,
    style.fill?.bgColor?.rgb,
    style.color?.rgb,
  ];

  return candidates
    .map(normalizeColorHex)
    .some((hex): hex is string => !!hex && isAlertHex(hex));
}

function rowHasAlertColor(
  sheet: XLSX.WorkSheet,
  rowIdx: number,
  range: XLSX.Range,
  excludedCols: number[] = [],
): boolean {
  for (let c = range.s.c; c <= range.e.c; c++) {
    if (excludedCols.includes(c)) continue;
    const cell = sheet[XLSX.utils.encode_cell({ r: rowIdx, c })];
    if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
    if (cellHasAlertColor(cell)) return true;
  }
  return false;
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Live formula engine Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function deriveLiveValues(m: MemberRow) {
  const pay = paymentMap.get(m.rowIdx);
  if (!pay) {
    const sharesBalance = (m.totalShares + m.monthlyShare + m.shareTransfer)
      - (m.fineDeduction + m.shareDeduction + m.advDeduction + m.riskFundOut + m.shareOut + m.nonCashOut);

    return {
      totalCash: 0,
      totalAdvance: 0,
      totalRepaid: 0,
      monthlyShare: m.monthlyShare,
      sharesBalance: m.sharesBalance || Math.max(0, sharesBalance),
      loansBalance: m.loansBalance || Math.max(0, m.loanBalance - m.loanRepayment),
    };
  }

  const totalCash = pay.cash + pay.paybill + pay.bank;
  const totalAdvance = pay.advRepayment + m.advInterestPaid - pay.advanceDeduction;
  const totalRepaid = totalCash - (m.passBook + pay.riskFund + totalAdvance + pay.fine);
  const loanBase = pay.loanRepayment + pay.reducingInterest + m.registrationFee;
  const monthlyShare = totalRepaid > loanBase
    ? totalRepaid - loanBase
    : totalRepaid - m.registrationFee;
  const sharesBalance = Math.max(
    0,
    (m.totalShares + monthlyShare + m.shareTransfer)
    - (pay.fineDeduction + pay.shareDeduction + pay.advanceDeduction + pay.riskFundOut + m.shareOut + m.nonCashOut)
  );
  const loansBalance = Math.max(0, m.loanBalance - pay.loanRepayment);

  return {
    totalCash,
    totalAdvance,
    totalRepaid,
    monthlyShare,
    sharesBalance,
    loansBalance,
  };
}

function syncDerivedPaymentState(m: MemberRow) {
  const pay = paymentMap.get(m.rowIdx);
  const live = deriveLiveValues(m);
  if (pay) {
    pay.monthlyShare = live.monthlyShare;
    pay.sharesBalance = live.sharesBalance;
    pay.loansBalance = live.loansBalance;
  }
  return live;
}

function calcLiveValues(m: MemberRow) {
  const live = syncDerivedPaymentState(m);

  popupTotalCash.textContent = live.totalCash >= 0 ? fmt(live.totalCash) : MONEY_PLACEHOLDER;
  popupTotalAdvance.textContent = live.totalAdvance >= 0 ? fmt(live.totalAdvance) : MONEY_PLACEHOLDER;
  popupTotalRepaid.textContent = fmt(live.totalRepaid);
  popupMShare.textContent = fmt(live.monthlyShare);
  popupMShareResult.textContent = fmt(live.monthlyShare);
}

function setDetailPopupEditableState(isEditable: boolean) {
  detailOverlay.classList.toggle('readonly-member', !isEditable);
  popupEditableInputs.forEach((input) => {
    input.disabled = !isEditable;
  });
}

function log(msg: string, type: 'default' | 'success' | 'warn' | 'error' = 'default') {
  const li = document.createElement('li');
  li.textContent = msg;
  if (type !== 'default') li.classList.add(type);
  logList.appendChild(li);
  logCard.scrollTop = logCard.scrollHeight;
}

function resetLog() {
  logList.innerHTML = '';
  logIcon.textContent = '...';
  logTitle.textContent = 'Processing...';
  logCard.classList.remove('hidden');
}

function setLoading(loading: boolean) {
  processBtn.disabled = loading || !selectedFile;
  btnText.textContent = loading ? 'Processing...' : 'Process File';
  spinner.classList.toggle('hidden', !loading);
}

function setStepActive(n: 1 | 2 | 3) {
  [step1Ind, step2Ind, step3Ind].forEach((el, i) => {
    el.classList.remove('active', 'done');
    if (i + 1 < n) el.classList.add('done');
    if (i + 1 === n) el.classList.add('active');
  });
}

function shiftFormula(formula: string, srcRow: number, tgtRow: number): string {
  return formula.replace(
    /([A-Za-z_][\w]*!)?(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g,
    (match, sheet, colDollar, col, rowDollar, row) => {
      if (rowDollar === '$') return match;
      if (parseInt(row, 10) === srcRow) return `${sheet ?? ''}${colDollar}${col}${rowDollar}${tgtRow}`;
      return match;
    }
  );
}

function getInstallment(principal: number): number {
  if (principal <= 0) return 0;
  if (principal <= 5000) return 335;
  if (principal <= 10000) return 500;
  if (principal <= 15000) return 750;
  if (principal <= 20000) return 1000;
  if (principal <= 25000) return 1250;
  if (principal <= 30000) return 1500;
  if (principal <= 35000) return 1500;
  if (principal <= 40000) return 1600;
  if (principal <= 45000) return 1800;
  if (principal <= 50000) return 2000;
  if (principal <= 55000) return 2200;
  if (principal <= 60000) return 2400;
  if (principal <= 65000) return 2600;
  if (principal <= 70000) return 2800;
  if (principal <= 75000) return 3000;
  if (principal <= 80000) return 3200;
  if (principal <= 85000) return 3400;
  if (principal <= 90000) return 3600;
  if (principal <= 95000) return 3800;
  if (principal <= 100000) return 4000;
  if (principal <= 120000) return 4800;
  if (principal <= 140000) return 5600;
  if (principal <= 160000) return 6400;
  if (principal <= 180000) return 7200;
  if (principal <= 200000) return 8000;
  if (principal <= 250000) return 10000;
  if (principal <= 300000) return 12000;
  if (principal <= 350000) return 14000;
  if (principal <= 400000) return 16000;
  if (principal <= 450000) return 18000;
  if (principal <= 500000) return 20000;
  if (principal <= 550000) return 22000;
  if (principal <= 600000) return 24000;
  if (principal <= 650000) return 26000;
  if (principal <= 700000) return 28000;
  if (principal <= 750000) return 30000;
  if (principal <= 800000) return 32000;
  if (principal <= 850000) return 34000;
  if (principal <= 900000) return 36000;
  return 36000;
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ File selection Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function acceptFile(file: File) {
  if (!file.name.match(/\.xlsx?$/i)) { alert('Please upload an Excel file (.xlsx or .xls).'); return; }
  selectedFile = file;
  processedWorkbook = null;
  formulaBanner.classList.add('hidden');
  step2Section.classList.add('hidden');
  logCard.classList.add('hidden');
  setStepActive(1);
  fileNameEl.textContent = file.name;
  fileSizeEl.textContent = formatBytes(file.size);
  fileInfo.classList.remove('hidden');
  processBtn.disabled = false;
}

function clearFile() {
  selectedFile = null;
  processedWorkbook = null;
  paymentMap.clear();
  memberRowMap.clear();
  activeRecordMember = null;
  fileInput.value = '';
  fileInfo.classList.add('hidden');
  processBtn.disabled = true;
  formulaBanner.classList.add('hidden');
  logCard.classList.add('hidden');
  step2Section.classList.add('hidden');
  memberTableBody.innerHTML = '';
  closeDetailPopup();
  closeRecordPopup();
  setStepActive(1);
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Drag and drop Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
dropZone.addEventListener('dragover', (e) => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault(); dropZone.classList.remove('drag-over');
  const file = e.dataTransfer?.files[0];
  if (file) acceptFile(file);
});
dropZone.addEventListener('click', () => fileInput.click());
browseBtn.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
fileInput.addEventListener('change', () => { if (fileInput.files?.[0]) acceptFile(fileInput.files[0]); });
clearBtn.addEventListener('click', (e) => { e.stopPropagation(); clearFile(); });
startOverBtn.addEventListener('click', () => clearFile());

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Core: formula fill & processing Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
async function processFile(file: File) {
  setLoading(true);
  resetLog();
  formulaBanner.classList.add('hidden');
  step2Section.classList.add('hidden');
  closeDetailPopup();
  closeRecordPopup();

  try {
    log('Reading file...');
    const arrayBuffer = await file.arrayBuffer();
    const data = new Uint8Array(arrayBuffer);
    const workbook = XLSX.read(data, { type: 'array', cellFormula: true, cellNF: true, cellStyles: true });

    const sheetName = workbook.SheetNames[0];
    if (!sheetName) throw new Error('No sheets found in the workbook.');
    const sheet = workbook.Sheets[sheetName];
    log(`Sheet: "${sheetName}"`, 'success');

    const ref = sheet['!ref'];
    if (!ref) throw new Error('Sheet appears to be empty.');
    const range = XLSX.utils.decode_range(ref);

    // Build colMap for every header
    const colMap: Record<string, number> = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
      if (!cell) continue;
      const header = (cell.v as string)?.toString().trim() ?? '';
      colMap[header] = c;
    }

    // Ã¢â€â‚¬Ã¢â€â‚¬ Formula fill Ã¢â€â‚¬Ã¢â€â‚¬
    const found = FORMULA_COLUMNS.filter(col => colMap[col] !== undefined);
    const missing = FORMULA_COLUMNS.filter(col => colMap[col] === undefined);
    if (found.length === 0) throw new Error('None of the target formula column headers found in row 1.');
    log(`Found ${found.length} formula column(s): ${found.join(', ')}`, 'success');
    if (missing.length > 0) log(`Skipped (not found): ${missing.join(', ')}`, 'warn');

    let lastRow = 1;
    for (let r = 1; r <= range.e.r; r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c })];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
          if (r > lastRow) lastRow = r;
          break;
        }
      }
    }
    log(`Last data row: ${lastRow + 1}`, 'success');
    if (lastRow < 2) throw new Error('No data rows detected below the header row.');

    let totalCopied = 0;
    for (const colName of FORMULA_COLUMNS) {
      const colIdx = colMap[colName];
      if (colIdx === undefined) continue;
      const srcAddr = XLSX.utils.encode_cell({ r: 1, c: colIdx });
      const srcCell = sheet[srcAddr];
      if (!srcCell || (srcCell.v === undefined && !srcCell.f)) { log(`"${colName}" - row 2 empty, skipping.`, 'warn'); continue; }
      const hasFormula = !!srcCell.f;
      const srcFormula = srcCell.f ?? null;
      let copiedInCol = 0;
      for (let tgtRow = 2; tgtRow <= lastRow; tgtRow++) {
        const tgtAddr = XLSX.utils.encode_cell({ r: tgtRow, c: colIdx });
        if (hasFormula && srcFormula) {
          sheet[tgtAddr] = { t: srcCell.t, f: shiftFormula(srcFormula, 2, tgtRow + 1) };
        } else {
          sheet[tgtAddr] = { t: srcCell.t, v: srcCell.v, w: srcCell.w };
        }
        if (tgtRow > range.e.r) range.e.r = tgtRow;
        if (colIdx > range.e.c) range.e.c = colIdx;
        copiedInCol++;
      }
      totalCopied += copiedInCol;
      log(`"${colName}" - ${hasFormula ? 'formula' : 'value'} copied to ${copiedInCol} row(s).`, 'success');
    }
    sheet['!ref'] = XLSX.utils.encode_range(range);
    log(`Done! ${totalCopied} cells updated.`, 'success');

    // Ã¢â€â‚¬Ã¢â€â‚¬ Collect member rows with breakdown Ã¢â€â‚¬Ã¢â€â‚¬
    const memberNoColIdx = colMap['MemberNo'];
    const memberNameIdx = colMap['MemberName'];
    const pLoanCol = colMap['PrincipalLoan'];
    const lBalCol = colMap['LoanBalance'];
    const aBalCol = colMap['AdvanceBalance'];
    const mShareCol = colMap['MonthlyShare'];
    const totalSharesCol = colMap['TotalShares'];
    const shareTransferCol = colMap['ShareTransfer'];
    const fineDeductionCol = colMap['FineDeduction'];
    const shareDeductionCol = colMap['ShareDeduction'];
    const riskFundOutCol = colMap['RiskFundOut'];
    const shareOutCol = colMap['ShareOut'];
    const nonCashOutCol = colMap['NonCashOut'];
    const memberStatusCol = colMap['MemberStatus'];
    const sharesBaseCol = getHeaderIndex(colMap, ['Shares B/F', 'Share B/F', 'Shares BF', 'Share BF']);
    const sharesBalanceCol = colMap['Shares C/F'];
    const loansBaseCol = getHeaderIndex(colMap, ['Loans B/F', 'Loan B/F', 'Loans BF', 'Loan BF']);
    const loansBalanceCol = colMap['Loans C/F'];
    const collected: MemberRow[] = [];

    if (memberNoColIdx === undefined) {
      log('Warning: "MemberNo" column not found - contribution table empty.', 'warn');
    } else {
      for (let r = 1; r <= lastRow; r++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c: memberNoColIdx })];
        if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
        const memberNo = cell.v as string | number;

        const nameCell = memberNameIdx !== undefined ? sheet[XLSX.utils.encode_cell({ r, c: memberNameIdx })] : null;
        const memberName = nameCell && nameCell.v ? String(nameCell.v).trim() : 'Unknown';
        const statusCell = memberStatusCol !== undefined ? sheet[XLSX.utils.encode_cell({ r, c: memberStatusCol })] : null;
        const memberStatus = statusCell && statusCell.v ? String(statusCell.v).trim() : 'Active';
        const isInactive = memberStatus.toLowerCase() === 'inactive';
        const excludedColorCols = [
          memberNoColIdx,
          memberNameIdx,
        ].filter((idx): idx is number => idx !== undefined);
        const hasAlertColor = rowHasAlertColor(sheet, r, range, excludedColorCols);

        const getNum = (cIdx?: number): number => {
          if (cIdx === undefined) return 0;
          const c = sheet[XLSX.utils.encode_cell({ r, c: cIdx })];
          const v = parseFloat(c?.v as string);
          return isNaN(v) ? 0 : v;
        };

        const principal = getNum(pLoanCol);
        const installment = getInstallment(principal);
        const loanBalance = getNum(lBalCol);
        const loanInterest = Math.round(loanBalance * 0.015);
        const advanceBalance = getNum(aBalCol);
        const advanceInterest = Math.round(advanceBalance * 0.10);
        const monthlyShare = getNum(mShareCol); // Extracted from Excel
        const totalShares = getNum(totalSharesCol);
        const shareTransfer = getNum(shareTransferCol);
        const fineDeduction = getNum(fineDeductionCol);
        const shareDeduction = getNum(shareDeductionCol);
        const riskFundOut = getNum(riskFundOutCol);
        const shareOut = getNum(shareOutCol);
        const nonCashOut = getNum(nonCashOutCol);
        const sharesBalanceBase = getNum(sharesBaseCol);
        const sharesBalance = getNum(sharesBalanceCol);
        const loansBalanceBase = getNum(loansBaseCol);
        const loansBalance = getNum(loansBalanceCol);
        const loanRepayment = 0;
        const advRepayment = 0;

        const advInterestPaid = getNum(colMap['AdvanceInterestPaid']);
        const advDeduction = getNum(colMap['AdvanceDeduction']);
        const loanInterestPaid = Math.round(loanBalance * 0.015);
        const registrationFee = getNum(colMap['RegistrationFee']);
        const passBook = getNum(colMap['PassBook']);
        const fine = getNum(colMap['Fine']);

        // The expected calculation uses constant 500 for share and 50 for risk fund.
        const expected = installment + loanInterest + advanceBalance + advanceInterest + 500 + 50;

        collected.push({
          memberNo,
          memberName,
          memberStatus,
          isInactive,
          rowIdx: r,
          hasAlertColor,
          expected,
          principal,
          installment,
          loanBalance,
          loanInterest,
          advanceBalance,
          advanceInterest,
          monthlyShare,
          totalShares,
          shareTransfer,
          fineDeduction,
          shareDeduction,
          riskFundOut,
          shareOut,
          nonCashOut,
          sharesBalanceBase,
          sharesBalance,
          loansBalanceBase,
          loansBalance,
          loanRepayment,
          advRepayment,
          advInterestPaid,
          advDeduction,
          loanInterestPaid,
          registrationFee,
          passBook,
          fine
        });
      }
      log(`Loaded ${collected.length} member(s).`, 'success');
    }

    processedWorkbook = workbook;
    processedSheetName = sheetName;
    outputName = `${file.name.replace(/\.xlsx?$/i, '')}_filled.xlsx`;
    paymentMap.clear();

    logIcon.textContent = 'OK';
    logTitle.textContent = 'Formulas filled!';
    formulaBanner.classList.remove('hidden');
    setStepActive(2);
    showContributionSection(collected);

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    log(`Error: ${msg}`, 'error');
    logIcon.textContent = 'ERR';
    logTitle.textContent = 'Processing failed';
    console.error(err);
  } finally {
    setLoading(false);
  }
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Build contribution table Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function showContributionSection(members: MemberRow[]) {
  memberTableBody.innerHTML = '';
  paymentMap.clear(); // Reset map on re-process
  memberRowMap.clear();
  activeRecordMember = null;

  members.forEach((m) => {
    memberRowMap.set(m.rowIdx, m);
    paymentMap.set(m.rowIdx, {
      cash: 0,
      paybill: 0,
      bank: 0,
      loanRepayment: 0,
      advRepayment: 0,
      riskFund: 0,
      fine: m.fine,
      fineDeduction: m.fineDeduction,
      shareDeduction: m.shareDeduction,
      advanceDeduction: m.advDeduction,
      riskFundOut: m.riskFundOut,
      reducingInterest: m.loanInterest,
      monthlyShare: m.monthlyShare,
      sharesBalance: m.sharesBalance,
      loansBalance: m.loansBalance,
    });

    const tr = document.createElement('tr');
    tr.dataset.rowIdx = String(m.rowIdx);
    if (m.hasAlertColor) tr.classList.add('alert-row');
    if (m.isInactive) tr.classList.add('inactive-row');

    const tdMno = document.createElement('td');
    tdMno.className = 'td-mno';
    const memberNoWrap = document.createElement('span');
    memberNoWrap.className = 'member-no-wrap';
    const memberNoText = document.createElement('span');
    memberNoText.className = 'member-no-text';
    memberNoText.textContent = String(m.memberNo);
    memberNoWrap.appendChild(memberNoText);
    if (m.isInactive) {
      const badge = document.createElement('span');
      badge.className = 'member-status-badge';
      badge.textContent = 'InActive';
      memberNoWrap.appendChild(badge);
      tdMno.title = 'This member is inactive and cannot be edited.';
    }
    tdMno.appendChild(memberNoWrap);
    tdMno.addEventListener('click', () => openDetailPopup(m));

    const tdExp = document.createElement('td');
    tdExp.className = 'td-exp';
    tdExp.textContent = fmt(m.expected);
    tdExp.addEventListener('click', () => openRecordPopup(m));

    const createInputCol = (type: 'cash' | 'paybill' | 'bank') => {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.step = '0.01';
      input.className = `amt-inp ${type}-inp`;
      if (m.hasAlertColor) input.classList.add('alert-inp');
      input.disabled = m.isInactive;
      input.placeholder = '0';
      if (m.isInactive) input.title = 'Inactive members cannot be edited.';
      input.addEventListener('input', () => {
        const val = parseFloat(input.value) || 0;
        const pay = paymentMap.get(m.rowIdx)!;
        if (type === 'cash') pay.cash = val;
        if (type === 'paybill') pay.paybill = val;
        if (type === 'bank') pay.bank = val;

        const totalContrib = pay.cash + pay.paybill + pay.bank;
        if (totalContrib >= 50) {
          pay.riskFund = 50;
        } else {
          pay.riskFund = 0; // or leave it user customized? Re-evaluate to 0 if under 50.
        }

        updateFooterTotals();
        if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
        autoSaveToSheet(m.rowIdx);
      });
      td.appendChild(input);
      return td;
    };

    tr.appendChild(tdMno);
    tr.appendChild(tdExp);
    tr.appendChild(createInputCol('cash'));
    tr.appendChild(createInputCol('paybill'));
    tr.appendChild(createInputCol('bank'));

    memberTableBody.appendChild(tr);
  });

  const sumExpected = members.reduce((s, m) => s + m.expected, 0);
  colTotalExpected.textContent = sumExpected > 0 ? fmt(sumExpected) : MONEY_PLACEHOLDER;

  memberCountEl.textContent = String(members.length);
  contribSummary.classList.remove('hidden');

  updateFooterTotals();

  step2Section.classList.remove('hidden');
  setTimeout(() => step2Section.scrollIntoView({ behavior: 'smooth', block: 'start' }), 150);
}

function updateFooterTotals() {
  let cashTotal = 0;
  let paybillTotal = 0;
  let bankTotal = 0;

  for (const pay of paymentMap.values()) {
    cashTotal += pay.cash;
    paybillTotal += pay.paybill;
    bankTotal += pay.bank;
  }

  colTotalCash.textContent = cashTotal > 0 ? fmt(cashTotal) : MONEY_PLACEHOLDER;
  colTotalPaybill.textContent = paybillTotal > 0 ? fmt(paybillTotal) : MONEY_PLACEHOLDER;
  colTotalBank.textContent = bankTotal > 0 ? fmt(bankTotal) : MONEY_PLACEHOLDER;
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Detail Popup Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function setRecordTab(tab: RecordTab) {
  const savingsActive = tab === 'savings';
  savingsTabBtn.classList.toggle('active', savingsActive);
  loanRecordTabBtn.classList.toggle('active', !savingsActive);
  savingsPanel.classList.toggle('hidden', !savingsActive);
  loanRecordPanel.classList.toggle('hidden', savingsActive);
}

function refreshRecordPopup(m: MemberRow) {
  const payment = paymentMap.get(m.rowIdx);
  const live = syncDerivedPaymentState(m);
  recordSavingsValue.textContent = fmt(payment?.monthlyShare ?? live.monthlyShare);
  recordSavingsBalance.textContent = fmt(payment?.sharesBalance ?? live.sharesBalance);
  recordLoanPrincipal.textContent = fmt(payment?.loanRepayment ?? 0);
  recordLoanInterest.textContent = fmt(payment?.reducingInterest ?? m.loanInterest);
  recordLoanBalance.textContent = fmt(payment?.loansBalance ?? live.loansBalance);
}

function openRecordPopup(m: MemberRow, tab: RecordTab = 'savings') {
  const initialsSrc = m.memberName && m.memberName !== 'Unknown' ? String(m.memberName) : String(m.memberNo);
  recordAvatar.textContent = initialsSrc.slice(0, 2).toUpperCase();
  recordMemberNo.textContent = String(m.memberNo);
  recordMemberName.textContent = m.isInactive ? `${m.memberName} (InActive)` : m.memberName;
  activeRecordMember = m;
  refreshRecordPopup(m);
  setRecordTab(tab);
  recordOverlay.classList.remove('hidden');
  syncBodyScrollLock();
}

function closeRecordPopup() {
  recordOverlay.classList.add('hidden');
  activeRecordMember = null;
  syncBodyScrollLock();
}

function openDetailPopup(m: MemberRow) {
  // Use memberName for initials if possible, fallback to memberNo
  const initialsSrc = m.memberName && m.memberName !== 'Unknown' ? String(m.memberName) : String(m.memberNo);
  const initials = initialsSrc.slice(0, 2).toUpperCase();
  popupAvatar.textContent = initials;
  popupMemberNo.textContent = String(m.memberNo);
  popupMemberName.textContent = m.isInactive ? `${m.memberName} (InActive)` : m.memberName;

  popupExpected.textContent = `KSh ${fmt(m.expected)}`;
  popupPrincipal.textContent = m.principal > 0 ? fmt(m.principal) : MONEY_PLACEHOLDER;
  popupInstallment.textContent = m.installment > 0 ? fmt(m.installment) : MONEY_PLACEHOLDER;
  popupTotalShares.textContent = m.totalShares > 0 ? fmt(m.totalShares) : MONEY_PLACEHOLDER;
  popupLoanBalance.textContent = m.loanBalance > 0 ? fmt(m.loanBalance) : MONEY_PLACEHOLDER;
  popupShareLoanDiff.textContent = fmt(m.totalShares - m.loanBalance);
  popupLoanInterest.textContent = m.loanInterest > 0 ? fmt(m.loanInterest) : MONEY_PLACEHOLDER;
  popupAdvBalance.textContent = m.advanceBalance > 0 ? fmt(m.advanceBalance) : MONEY_PLACEHOLDER;
  popupAdvInterest.textContent = m.advanceInterest > 0 ? fmt(m.advanceInterest) : MONEY_PLACEHOLDER;
  popupMShare.textContent = fmt(syncDerivedPaymentState(m).monthlyShare);

  const payData = paymentMap.get(m.rowIdx)!;
  popupRiskFund.value = String(payData.riskFund);
  popupLoanRepayment.value = payData.loanRepayment ? payData.loanRepayment.toString() : '';
  popupAdvRepayment.value = payData.advRepayment ? payData.advRepayment.toString() : '';
  popupFine.value = payData.fine ? payData.fine.toString() : '';
  popupFineDeduction.value = payData.fineDeduction ? payData.fineDeduction.toString() : '';
  popupShareDeduction.value = payData.shareDeduction ? payData.shareDeduction.toString() : '';
  popupAdvanceDeduction.value = payData.advanceDeduction ? payData.advanceDeduction.toString() : '';
  popupRiskFundOut.value = payData.riskFundOut ? payData.riskFundOut.toString() : '';
  setDetailPopupEditableState(!m.isInactive);

  popupRiskFund.oninput = () => {
    let val = parseInt(popupRiskFund.value) || 0;
    if (val > 50) { val = 50; popupRiskFund.value = '50'; }
    payData.riskFund = val;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };

  popupLoanRepayment.oninput = () => {
    payData.loanRepayment = parseFloat(popupLoanRepayment.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };
  popupAdvRepayment.oninput = () => {
    payData.advRepayment = parseFloat(popupAdvRepayment.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };
  popupFine.oninput = () => {
    payData.fine = parseFloat(popupFine.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };
  popupFineDeduction.oninput = () => {
    payData.fineDeduction = parseFloat(popupFineDeduction.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };
  popupShareDeduction.oninput = () => {
    payData.shareDeduction = parseFloat(popupShareDeduction.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };
  popupAdvanceDeduction.oninput = () => {
    payData.advanceDeduction = parseFloat(popupAdvanceDeduction.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };
  popupRiskFundOut.oninput = () => {
    payData.riskFundOut = parseFloat(popupRiskFundOut.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
  };

  calcLiveValues(m); // initial paint for computed popup rows

  detailOverlay.classList.remove('hidden');
  syncBodyScrollLock();
}

function closeDetailPopup() {
  detailOverlay.classList.add('hidden');
  syncBodyScrollLock();
}

[popupClose, popupDismiss].forEach(btn => btn.addEventListener('click', closeDetailPopup));
detailOverlay.addEventListener('click', (e) => { if (e.target === detailOverlay) closeDetailPopup(); });
[recordClose, recordDismiss].forEach(btn => btn.addEventListener('click', closeRecordPopup));
recordOverlay.addEventListener('click', (e) => { if (e.target === recordOverlay) closeRecordPopup(); });
savingsTabBtn.addEventListener('click', () => setRecordTab('savings'));
loanRecordTabBtn.addEventListener('click', () => setRecordTab('loan'));
document.addEventListener('keydown', (e) => {
  if (e.key !== 'Escape') return;
  if (!recordOverlay.classList.contains('hidden')) {
    closeRecordPopup();
    return;
  }
  closeDetailPopup();
});

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Auto-save: write current row data to in-memory workbook Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function autoSaveToSheet(rowIdx: number) {
  if (!processedWorkbook) return;
  const sheet = processedWorkbook.Sheets[processedSheetName];
  // Build column map if not cached
  const ref = sheet['!ref'];
  if (!ref) return;
  const range = XLSX.utils.decode_range(ref);
  const colCache: Partial<Record<PaymentCol | LiveResultCol, number>> = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (!cell) continue;
    const hdr = (cell.v as string)?.toString().trim() as PaymentCol | LiveResultCol;
    if (PAYMENT_COLS.includes(hdr as PaymentCol)) colCache[hdr] = c;
    if (LIVE_RESULT_COLS.includes(hdr as LiveResultCol)) colCache[hdr] = c;
  }
  const payment = paymentMap.get(rowIdx);
  if (!payment) return;
  const {
    cash,
    paybill,
    bank,
    loanRepayment,
    advRepayment,
    riskFund,
    fine,
    fineDeduction,
    shareDeduction,
    advanceDeduction,
    riskFundOut,
  } = payment;
  const vals: Record<PaymentCol, number> = {
    Cash: cash, Paybill: paybill, Bank: bank,
    LoanRepayment: loanRepayment, AdvanceRepayment: advRepayment,
    RiskFund: riskFund,
    Fine: fine,
    FineDeduction: fineDeduction,
    ShareDeduction: shareDeduction,
    AdvanceDeduction: advanceDeduction,
    RiskFundOut: riskFundOut,
  };
  for (const col of PAYMENT_COLS) {
    const cIdx = colCache[col];
    if (cIdx === undefined) continue;
    const addr = XLSX.utils.encode_cell({ r: rowIdx, c: cIdx });
    sheet[addr] = { t: 'n', v: vals[col] || 0 };
  }

  const member = memberRowMap.get(rowIdx);
  if (!member) return;

  const live = syncDerivedPaymentState(member);
  const liveVals: Record<LiveResultCol, number> = {
    MonthlyShare: live.monthlyShare,
    'Shares C/F': live.sharesBalance,
    'Loans C/F': live.loansBalance,
    TotalCash: live.totalCash,
    TotalAdvance: live.totalAdvance,
    'Total RePaid': live.totalRepaid,
  };

  for (const col of LIVE_RESULT_COLS) {
    const cIdx = colCache[col];
    if (cIdx === undefined) continue;
    const addr = XLSX.utils.encode_cell({ r: rowIdx, c: cIdx });
    sheet[addr] = { t: 'n', v: liveVals[col] || 0 };
  }
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Apply contributions & download Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function applyContributions() {
  if (!processedWorkbook) return;

  const sheet = processedWorkbook.Sheets[processedSheetName];
  const ref = sheet['!ref']!;
  const range = XLSX.utils.decode_range(ref);

  const payColIdx: Partial<Record<PaymentCol, number>> = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (!cell) continue;
    const hdr = (cell.v as string)?.toString().trim() as PaymentCol;
    if (PAYMENT_COLS.includes(hdr)) payColIdx[hdr] = c;
  }

  const missingCols = PAYMENT_COLS.filter(p => payColIdx[p] === undefined);
  if (missingCols.length > 0) {
    alert(`Note: Columns not found and will be skipped: ${missingCols.join(', ')}`);
  }

  let writes = 0;

  for (const [rowIdx, payment] of paymentMap) {
    const {
      cash,
      paybill,
      bank,
      loanRepayment,
      advRepayment,
      riskFund,
      fine,
      fineDeduction,
      shareDeduction,
      advanceDeduction,
      riskFundOut,
    } = payment;

    const vals: Record<PaymentCol, number> = {
      Cash: cash, Paybill: paybill, Bank: bank,
      LoanRepayment: loanRepayment, AdvanceRepayment: advRepayment,
      RiskFund: riskFund,
      Fine: fine,
      FineDeduction: fineDeduction,
      ShareDeduction: shareDeduction,
      AdvanceDeduction: advanceDeduction,
      RiskFundOut: riskFundOut,
    };

    for (const col of PAYMENT_COLS) {
      const cIdx = payColIdx[col];
      if (cIdx === undefined) continue;
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: cIdx });
      sheet[addr] = { t: 'n', v: vals[col] || 0 }; // write 0 if blank/unpaid
      if (vals[col] > 0) writes++;
    }

    autoSaveToSheet(rowIdx);
  }

  const buf = XLSX.write(processedWorkbook, { bookType: 'xlsx', type: 'buffer' }) as ArrayBuffer;
  const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = outputName; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);

  setStepActive(3);
  step3Ind.classList.remove('active');
  step3Ind.classList.add('done');

  console.log(`Written: ${writes} payment value(s).`);
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Event wiring Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
processBtn.addEventListener('click', () => { if (selectedFile) processFile(selectedFile); });
applyBtn.addEventListener('click', applyContributions);

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Page init: enforce correct state on load Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
// Prevents spinner / formula-filled banner from showing before any file action.
(function initPageState() {
  spinner.classList.add('hidden');
  formulaBanner.classList.add('hidden');
  logCard.classList.add('hidden');
  step2Section.classList.add('hidden');
  fileInfo.classList.add('hidden');
  processBtn.disabled = true;
  btnText.textContent = 'Process File';
})();

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ Warning before exit Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
window.addEventListener('beforeunload', (e) => {
  if (selectedFile) {
    e.preventDefault();
    e.returnValue = ''; // Standard way to trigger the confirmation dialog
  }
});


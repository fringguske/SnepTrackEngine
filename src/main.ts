import * as XLSX from 'xlsx';
import { inject } from '@vercel/analytics';
import './style.css';

inject();

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
let activeDraftKey: string | null = null;

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

type DraftPaymentFields = Pick<PaymentEntry,
  'cash' | 'paybill' | 'bank' |
  'loanRepayment' | 'advRepayment' | 'riskFund' |
  'fine' | 'fineDeduction' | 'shareDeduction' | 'advanceDeduction' | 'riskFundOut'
>;
type DraftPayloadV1 = {
  version: 1;
  file: { name: string; size: number; lastModified: number };
  savedAt: number;
  entries: Record<string, DraftPaymentFields>;
};

const DRAFT_STORAGE_PREFIX = 'snepbotv1:draft:v1:';
let draftSaveTimer: number | null = null;

function getFileFingerprint(file: File): string {
  return `${file.name}|${file.size}|${file.lastModified}`;
}

function getDraftStorageKeyForFile(file: File): string {
  return `${DRAFT_STORAGE_PREFIX}${getFileFingerprint(file)}`;
}

function safeParseDraft(raw: string | null): DraftPayloadV1 | null {
  if (!raw) return null;
  try {
    const parsed = JSON.parse(raw) as DraftPayloadV1;
    if (!parsed || parsed.version !== 1 || typeof parsed.entries !== 'object') return null;
    return parsed;
  } catch {
    return null;
  }
}

function readActiveDraft(): DraftPayloadV1 | null {
  if (!activeDraftKey) return null;
  return safeParseDraft(localStorage.getItem(activeDraftKey));
}

function schedulePersistDraft() {
  if (!activeDraftKey || !selectedFile) return;
  if (draftSaveTimer) window.clearTimeout(draftSaveTimer);
  draftSaveTimer = window.setTimeout(() => {
    draftSaveTimer = null;
    persistDraftNow();
  }, 250);
}

function persistDraftNow() {
  if (!activeDraftKey || !selectedFile) return;

  const entries: Record<string, DraftPaymentFields> = {};
  for (const [rowIdx, pay] of paymentMap.entries()) {
    const member = memberRowMap.get(rowIdx);
    if (!member) continue;
    const key = String(member.memberNo).trim();

    // Store only meaningful edits to keep localStorage small.
    const hasAny =
      pay.cash !== 0 || pay.paybill !== 0 || pay.bank !== 0 ||
      pay.loanRepayment !== 0 || pay.advRepayment !== 0 || pay.riskFund !== 0 ||
      pay.fine !== member.fine ||
      pay.fineDeduction !== member.fineDeduction ||
      pay.shareDeduction !== member.shareDeduction ||
      pay.advanceDeduction !== member.advDeduction ||
      pay.riskFundOut !== member.riskFundOut;

    if (!hasAny) continue;

    entries[key] = {
      cash: pay.cash,
      paybill: pay.paybill,
      bank: pay.bank,
      loanRepayment: pay.loanRepayment,
      advRepayment: pay.advRepayment,
      riskFund: pay.riskFund,
      fine: pay.fine,
      fineDeduction: pay.fineDeduction,
      shareDeduction: pay.shareDeduction,
      advanceDeduction: pay.advanceDeduction,
      riskFundOut: pay.riskFundOut,
    };
  }

  const payload: DraftPayloadV1 = {
    version: 1,
    file: { name: selectedFile.name, size: selectedFile.size, lastModified: selectedFile.lastModified },
    savedAt: Date.now(),
    entries,
  };

  try {
    localStorage.setItem(activeDraftKey, JSON.stringify(payload));
  } catch {
    // Best-effort; ignore quota errors.
  }
}

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

type LoanRepaymentBracket = {
  minPrincipal: number;
  maxPrincipal: number;
  period: number;
  installment: number;
  interest: number;
  savings: number;
  total: number;
};

// Loan repayment schedule (scripted directly from the provided CSV values).
const LOAN_REPAYMENT_SCHEDULE: LoanRepaymentBracket[] = [
  { minPrincipal: 1, maxPrincipal: 5000, period: 15, installment: 335, interest: 75, savings: 550, total: 960 },
  { minPrincipal: 5001, maxPrincipal: 10000, period: 20, installment: 500, interest: 150, savings: 550, total: 1200 },
  { minPrincipal: 10001, maxPrincipal: 15000, period: 20, installment: 750, interest: 225, savings: 550, total: 1525 },
  { minPrincipal: 15001, maxPrincipal: 20000, period: 20, installment: 1000, interest: 300, savings: 550, total: 1850 },
  { minPrincipal: 20001, maxPrincipal: 25000, period: 20, installment: 1250, interest: 375, savings: 550, total: 2175 },
  { minPrincipal: 25001, maxPrincipal: 30000, period: 20, installment: 1500, interest: 450, savings: 550, total: 2500 },
  { minPrincipal: 30001, maxPrincipal: 35000, period: 24, installment: 1500, interest: 525, savings: 550, total: 2575 },
  { minPrincipal: 35001, maxPrincipal: 40000, period: 25, installment: 1600, interest: 600, savings: 550, total: 2750 },
  { minPrincipal: 40001, maxPrincipal: 45000, period: 25, installment: 1800, interest: 675, savings: 550, total: 3025 },
  { minPrincipal: 45001, maxPrincipal: 50000, period: 25, installment: 2000, interest: 750, savings: 550, total: 3300 },
  { minPrincipal: 50001, maxPrincipal: 55000, period: 25, installment: 2200, interest: 825, savings: 550, total: 3575 },
  { minPrincipal: 55001, maxPrincipal: 60000, period: 25, installment: 2400, interest: 900, savings: 550, total: 3850 },
  { minPrincipal: 60001, maxPrincipal: 65000, period: 25, installment: 2600, interest: 975, savings: 550, total: 4125 },
  { minPrincipal: 65001, maxPrincipal: 70000, period: 25, installment: 2800, interest: 1050, savings: 550, total: 4400 },
  { minPrincipal: 70001, maxPrincipal: 75000, period: 25, installment: 3000, interest: 1125, savings: 550, total: 4675 },
  { minPrincipal: 75001, maxPrincipal: 80000, period: 25, installment: 3200, interest: 1200, savings: 550, total: 4950 },
  { minPrincipal: 80001, maxPrincipal: 85000, period: 25, installment: 3400, interest: 1275, savings: 550, total: 5225 },
  { minPrincipal: 85001, maxPrincipal: 90000, period: 25, installment: 3600, interest: 1350, savings: 550, total: 5500 },
  { minPrincipal: 90001, maxPrincipal: 95000, period: 25, installment: 3800, interest: 1425, savings: 550, total: 5775 },
  { minPrincipal: 95001, maxPrincipal: 100000, period: 25, installment: 4000, interest: 1500, savings: 550, total: 6050 },
  { minPrincipal: 100001, maxPrincipal: 120000, period: 25, installment: 4800, interest: 1800, savings: 550, total: 7150 },
  { minPrincipal: 120001, maxPrincipal: 140000, period: 25, installment: 5600, interest: 2100, savings: 550, total: 8250 },
  { minPrincipal: 140001, maxPrincipal: 160000, period: 25, installment: 6400, interest: 2400, savings: 550, total: 9350 },
  { minPrincipal: 160001, maxPrincipal: 180000, period: 25, installment: 7200, interest: 2700, savings: 550, total: 10450 },
  { minPrincipal: 180001, maxPrincipal: 200000, period: 25, installment: 8000, interest: 3000, savings: 550, total: 11550 },
  { minPrincipal: 200001, maxPrincipal: 250000, period: 25, installment: 10000, interest: 3750, savings: 1050, total: 14800 },
  { minPrincipal: 250001, maxPrincipal: 300000, period: 25, installment: 12000, interest: 4500, savings: 1050, total: 17550 },
  { minPrincipal: 300001, maxPrincipal: 350000, period: 25, installment: 14000, interest: 5250, savings: 1050, total: 20300 },
  { minPrincipal: 350001, maxPrincipal: 400000, period: 30, installment: 16000, interest: 6000, savings: 1050, total: 23050 },
  { minPrincipal: 400001, maxPrincipal: 450000, period: 25, installment: 18000, interest: 6750, savings: 1050, total: 25800 },
  { minPrincipal: 450001, maxPrincipal: 480000, period: 25, installment: 20000, interest: 7500, savings: 1050, total: 28550 },
  { minPrincipal: 500001, maxPrincipal: 550000, period: 25, installment: 22000, interest: 8250, savings: 1050, total: 31300 },
  { minPrincipal: 550001, maxPrincipal: 600000, period: 25, installment: 24000, interest: 9000, savings: 1050, total: 34050 },
  { minPrincipal: 600001, maxPrincipal: 650000, period: 25, installment: 26000, interest: 9750, savings: 1050, total: 36800 },
  { minPrincipal: 650001, maxPrincipal: 700000, period: 25, installment: 28000, interest: 10500, savings: 1050, total: 39550 },
  { minPrincipal: 700001, maxPrincipal: 750000, period: 25, installment: 30000, interest: 11250, savings: 1050, total: 42300 },
  { minPrincipal: 750001, maxPrincipal: 800000, period: 25, installment: 32000, interest: 12000, savings: 1050, total: 45050 },
  { minPrincipal: 800001, maxPrincipal: 850000, period: 25, installment: 34000, interest: 12750, savings: 1050, total: 47800 },
  { minPrincipal: 850001, maxPrincipal: 900000, period: 25, installment: 36000, interest: 13500, savings: 1050, total: 50550 },
];

const NO_LOAN_SAVINGS = 500;
const NO_LOAN_RISK_FUND = 50;
const NO_LOAN_SAVINGS_RISK = NO_LOAN_SAVINGS + NO_LOAN_RISK_FUND;

function getLoanRepaymentBracket(principal: number): LoanRepaymentBracket | null {
  if (principal <= 0) return null;
  const p = Math.round(principal);
  let bestLower: LoanRepaymentBracket | null = null;
  for (const bracket of LOAN_REPAYMENT_SCHEDULE) {
    if (p >= bracket.minPrincipal && p <= bracket.maxPrincipal) return bracket;
    if (bracket.maxPrincipal < p) bestLower = bracket;
  }
  return bestLower ?? LOAN_REPAYMENT_SCHEDULE[0];
}

// Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬ File selection Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬Ã¢â€â‚¬
function acceptFile(file: File) {
  if (!file.name.match(/\.xlsx?$/i)) { alert('Please upload an Excel file (.xlsx or .xls).'); return; }
  selectedFile = file;
  activeDraftKey = getDraftStorageKeyForFile(file);
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
  activeDraftKey = null;
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
        const loanBalance = getNum(lBalCol);
        const advanceBalance = getNum(aBalCol);
        const advanceInterest = Math.round(advanceBalance * 0.10);
        const hasLoan = loanBalance > 0;
        const loanBracket = hasLoan ? getLoanRepaymentBracket(principal) : null;
        const installment = loanBracket?.installment ?? 0;
        const reducingInterest = hasLoan ? Math.round(loanBalance * 0.015) : 0;
        const loanInterest = reducingInterest; // UI label is "Reducing Int. (1.5%)"
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
        const loanInterestPaid = reducingInterest;
        const registrationFee = getNum(colMap['RegistrationFee']);
        const passBook = getNum(colMap['PassBook']);
        const fine = getNum(colMap['Fine']);

        // Expected:
        // - If member has a loan:
        //   - If remaining loan balance is below the scheduled installment, charge:
        //     loanBalance + 1.5% (reducing interest) + 500 savings + 50 risk fund.
        //   - Otherwise, charge the schedule total (already includes savings/risk fund component).
        //   - Then add advance + 10% advance interest (if any).
        // - If member has no loan: advance + 10% advance interest (if any) + 500 savings + 50 risk fund.
        let loanBaseExpected = NO_LOAN_SAVINGS_RISK;
        if (hasLoan) {
          if (loanBracket && loanBalance >= installment) {
            loanBaseExpected = loanBracket.total;
          } else {
            loanBaseExpected = NO_LOAN_SAVINGS_RISK + loanBalance + reducingInterest;
          }
        }
        const expected = loanBaseExpected + advanceBalance + advanceInterest;

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

  const draft = readActiveDraft();
  let restoredRows = 0;

  members.forEach((m) => {
    memberRowMap.set(m.rowIdx, m);
    const paySeed: PaymentEntry = {
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
    };

    const draftEntry = draft?.entries?.[String(m.memberNo).trim()];
    if (draftEntry) {
      paySeed.cash = draftEntry.cash ?? paySeed.cash;
      paySeed.paybill = draftEntry.paybill ?? paySeed.paybill;
      paySeed.bank = draftEntry.bank ?? paySeed.bank;
      paySeed.loanRepayment = draftEntry.loanRepayment ?? paySeed.loanRepayment;
      paySeed.advRepayment = draftEntry.advRepayment ?? paySeed.advRepayment;
      paySeed.riskFund = draftEntry.riskFund ?? paySeed.riskFund;
      paySeed.fine = draftEntry.fine ?? paySeed.fine;
      paySeed.fineDeduction = draftEntry.fineDeduction ?? paySeed.fineDeduction;
      paySeed.shareDeduction = draftEntry.shareDeduction ?? paySeed.shareDeduction;
      paySeed.advanceDeduction = draftEntry.advanceDeduction ?? paySeed.advanceDeduction;
      paySeed.riskFundOut = draftEntry.riskFundOut ?? paySeed.riskFundOut;
      restoredRows++;
    }

    paymentMap.set(m.rowIdx, paySeed);

    if (draftEntry) autoSaveToSheet(m.rowIdx);

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

      const pay = paymentMap.get(m.rowIdx)!;
      const seeded = pay[type];
      input.value = seeded ? String(seeded) : '';

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
        schedulePersistDraft();
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
  if (restoredRows > 0) schedulePersistDraft();

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
  // Always show these balances even when 0 (per UI rules).
  popupPrincipal.textContent = fmt(m.principal);
  popupInstallment.textContent = m.installment > 0 ? fmt(m.installment) : MONEY_PLACEHOLDER;
  popupTotalShares.textContent = fmt(m.totalShares);
  popupLoanBalance.textContent = fmt(m.loanBalance);
  popupShareLoanDiff.textContent = fmt(m.totalShares - m.loanBalance);
  popupLoanInterest.textContent = m.loanInterest > 0 ? fmt(m.loanInterest) : MONEY_PLACEHOLDER;
  popupAdvBalance.textContent = fmt(m.advanceBalance);
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

  // Conditionally hide chips that don't apply, to reduce clutter on mobile.
  // Always show: advance balance, principal loan, loan balance (even when 0).
  const setChipHidden = (el: Element | null, hidden: boolean) => {
    const chip = el?.closest('.popup-chip') as HTMLElement | null;
    if (!chip) return;
    chip.classList.toggle('hidden', hidden);
  };

  const advanceIsZero = m.advanceBalance <= 0;
  setChipHidden(popupAdvRepayment, advanceIsZero);
  setChipHidden(popupAdvInterest, advanceIsZero);
  setChipHidden(popupAdvanceDeduction, advanceIsZero);

  const principalIsZero = m.principal <= 0;
  setChipHidden(popupInstallment, principalIsZero);
  setChipHidden(popupLoanRepayment, principalIsZero);
  setChipHidden(popupShareDeduction, principalIsZero);
  setChipHidden(popupLoanInterest, principalIsZero);

  popupRiskFund.oninput = () => {
    let val = parseInt(popupRiskFund.value) || 0;
    if (val > 50) { val = 50; popupRiskFund.value = '50'; }
    payData.riskFund = val;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
  };

  popupLoanRepayment.oninput = () => {
    payData.loanRepayment = parseFloat(popupLoanRepayment.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
  };
  popupAdvRepayment.oninput = () => {
    payData.advRepayment = parseFloat(popupAdvRepayment.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
  };
  popupFine.oninput = () => {
    payData.fine = parseFloat(popupFine.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
  };
  popupFineDeduction.oninput = () => {
    payData.fineDeduction = parseFloat(popupFineDeduction.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
  };
  popupShareDeduction.oninput = () => {
    payData.shareDeduction = parseFloat(popupShareDeduction.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
  };
  popupAdvanceDeduction.oninput = () => {
    payData.advanceDeduction = parseFloat(popupAdvanceDeduction.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
  };
  popupRiskFundOut.oninput = () => {
    payData.riskFundOut = parseFloat(popupRiskFundOut.value) || 0;
    calcLiveValues(m);
    if (activeRecordMember?.rowIdx === m.rowIdx) refreshRecordPopup(m);
    autoSaveToSheet(m.rowIdx);
    schedulePersistDraft();
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

  // Normalize used range for better compatibility with mobile spreadsheet viewers.
  // Some apps open at the bottom/right of the used range and/or restrict scrolling to that range.
  // We keep: header row + all header columns + all member rows.
  const exportMaxRow = (() => {
    let maxRow = 1; // at least include header + first data row
    for (const r of memberRowMap.keys()) if (r > maxRow) maxRow = r;
    return maxRow;
  })();

  const exportMaxCol = (() => {
    let lastUsedCol = 0;
    let emptyRun = 0;
    const maxC = range.e.c;
    for (let c = 0; c <= maxC; c++) {
      const headerCell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
      const row2Cell = sheet[XLSX.utils.encode_cell({ r: 1, c })];
      const headerUsed = !!(headerCell && ((headerCell.v !== undefined && headerCell.v !== null && headerCell.v !== '') || headerCell.f));
      const row2Used = !!(row2Cell && ((row2Cell.v !== undefined && row2Cell.v !== null && row2Cell.v !== '') || row2Cell.f));
      if (headerUsed || row2Used) {
        lastUsedCol = c;
        emptyRun = 0;
        continue;
      }
      if (c > lastUsedCol) {
        emptyRun++;
        if (emptyRun >= 64) break;
      }
    }

    for (const cIdx of Object.values(payColIdx)) {
      if (typeof cIdx === 'number' && cIdx > lastUsedCol) lastUsedCol = cIdx;
    }

    return lastUsedCol;
  })();

  sheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: exportMaxRow, c: exportMaxCol } });

  const buf = XLSX.write(processedWorkbook, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer;
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

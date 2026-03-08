import * as XLSX from 'xlsx';
import './style.css';

// ─── Target columns for formula fill ──────────────────────────────────────────
const FORMULA_COLUMNS: string[] = [
  'Total RePaid', 'MonthlyShare', 'TotalAdvance', 'Shares C/F', 'Loans C/F', 'TotalCash',
];

const PAYMENT_COLS = ['Cash', 'Paybill', 'Bank', 'LoanRepayment', 'AdvanceRepayment', 'RiskFund'] as const;
type PaymentCol = typeof PAYMENT_COLS[number];

// ─── DOM refs — Step 1 ─────────────────────────────────────────────────────────
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

// ─── DOM refs — Step 2 ────────────────────────────────────────────────────────
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

// ─── Detail Popup refs ────────────────────────────────────────────────────────
const detailOverlay = document.getElementById('detailOverlay')! as HTMLDivElement;
const popupAvatar = document.getElementById('popupAvatar')! as HTMLDivElement;
const popupMemberNo = document.getElementById('popupMemberNo')! as HTMLParagraphElement;
const popupMemberName = document.getElementById('popupMemberName')! as HTMLParagraphElement;
const popupExpected = document.getElementById('popupExpected')! as HTMLSpanElement;
const popupPrincipal = document.getElementById('popupPrincipal')! as HTMLSpanElement;
const popupInstallment = document.getElementById('popupInstallment')! as HTMLSpanElement;
const popupLoanBalance = document.getElementById('popupLoanBalance')! as HTMLSpanElement;
const popupLoanInterest = document.getElementById('popupLoanInterest')! as HTMLSpanElement;
const popupAdvBalance = document.getElementById('popupAdvBalance')! as HTMLSpanElement;
const popupAdvInterest = document.getElementById('popupAdvInterest')! as HTMLSpanElement;
const popupMShare = document.getElementById('popupMShare')! as HTMLSpanElement;
const popupLoanRepayment = document.getElementById('popupLoanRepayment')! as HTMLInputElement;
const popupAdvRepayment = document.getElementById('popupAdvRepayment')! as HTMLInputElement;
const popupRiskFund = document.getElementById('popupRiskFund')! as HTMLInputElement;
const popupTotalCash = document.getElementById('popupTotalCash')! as HTMLSpanElement;
const popupTotalAdvance = document.getElementById('popupTotalAdvance')! as HTMLSpanElement;
const popupTotalRepaid = document.getElementById('popupTotalRepaid')! as HTMLSpanElement;
const popupMShareResult = document.getElementById('popupMShareResult')! as HTMLSpanElement;
const popupClose = document.getElementById('popupClose')! as HTMLButtonElement;
const popupDismiss = document.getElementById('popupDismiss')! as HTMLButtonElement;

// ─── State ────────────────────────────────────────────────────────────────────
let selectedFile: File | null = null;
let processedWorkbook: XLSX.WorkBook | null = null;
let processedSheetName = '';
let outputName = 'modified.xlsx';

interface MemberRow {
  memberNo: string | number;
  memberName: string;
  rowIdx: number;
  expected: number;
  principal: number;
  installment: number;
  loanBalance: number;
  loanInterest: number;
  advanceBalance: number;
  advanceInterest: number;
  monthlyShare: number;      // from Excel (for display)
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

/** Live map of entered amounts — includes riskFund per member */
const paymentMap = new Map<number, {
  cash: number; paybill: number; bank: number;
  loanRepayment: number; advRepayment: number; riskFund: number;
}>();

// ─── Helpers ──────────────────────────────────────────────────────────────────
function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1048576).toFixed(2)} MB`;
}

function fmt(n: number): string {
  return n.toLocaleString('en-KE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

// ─── Live formula engine ─────────────────────────────────────────────────
function calcLiveValues(m: MemberRow) {
  const pay = paymentMap.get(m.rowIdx);
  if (!pay) return;

  const { cash, paybill, bank, loanRepayment, advRepayment, riskFund } = pay;

  // Step 1: TotalCash
  const totalCash = cash + paybill + bank;

  // Step 2: TotalAdvance = AdvanceRepayment + AdvanceInterestPaid - AdvanceDeduction
  const totalAdvance = advRepayment + m.advInterestPaid - m.advDeduction;

  // Step 3: TotalRepaid = TotalCash - (PassBook + RiskFund + TotalAdvance + Fine)
  const totalRepaid = totalCash - (m.passBook + riskFund + totalAdvance + m.fine);

  // Step 4: MonthlyShare formula
  const loanBase = loanRepayment + m.loanInterestPaid + m.registrationFee;
  let monthlyShareCalc: number;
  if (totalRepaid > loanBase) {
    monthlyShareCalc = totalRepaid - loanBase;
  } else {
    monthlyShareCalc = totalRepaid - m.registrationFee;
  }

  // Update popup displays
  popupTotalCash.textContent = totalCash >= 0 ? fmt(totalCash) : '—';
  popupTotalAdvance.textContent = totalAdvance >= 0 ? fmt(totalAdvance) : '—';
  popupTotalRepaid.textContent = fmt(totalRepaid);
  popupMShareResult.textContent = fmt(monthlyShareCalc);
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
  logIcon.textContent = '⏳';
  logTitle.textContent = 'Processing…';
  logCard.classList.remove('hidden');
}

function setLoading(loading: boolean) {
  processBtn.disabled = loading || !selectedFile;
  btnText.textContent = loading ? 'Processing…' : 'Process File';
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

// ─── File selection ───────────────────────────────────────────────────────────
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
  fileInput.value = '';
  fileInfo.classList.add('hidden');
  processBtn.disabled = true;
  formulaBanner.classList.add('hidden');
  logCard.classList.add('hidden');
  step2Section.classList.add('hidden');
  memberTableBody.innerHTML = '';
  setStepActive(1);
}

// ─── Drag and drop ────────────────────────────────────────────────────────────
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

// ─── Core: formula fill & processing ─────────────────────────────────────────
async function processFile(file: File) {
  setLoading(true);
  resetLog();
  formulaBanner.classList.add('hidden');
  step2Section.classList.add('hidden');

  try {
    log('Reading file…');
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

    // ── Formula fill ──
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
      if (!srcCell || (srcCell.v === undefined && !srcCell.f)) { log(`"${colName}" — row 2 empty, skipping.`, 'warn'); continue; }
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
      log(`"${colName}" — ${hasFormula ? 'formula' : 'value'} copied to ${copiedInCol} row(s).`, 'success');
    }
    sheet['!ref'] = XLSX.utils.encode_range(range);
    log(`Done! ${totalCopied} cells updated.`, 'success');

    // ── Collect member rows with breakdown ──
    const memberNoColIdx = colMap['MemberNo'];
    const memberNameIdx = colMap['MemberName'];
    const pLoanCol = colMap['PrincipalLoan'];
    const lBalCol = colMap['LoanBalance'];
    const aBalCol = colMap['AdvanceBalance'];
    const mShareCol = colMap['MonthlyShare'];
    const collected: MemberRow[] = [];

    if (memberNoColIdx === undefined) {
      log('Warning: "MemberNo" column not found — contribution table empty.', 'warn');
    } else {
      for (let r = 1; r <= lastRow; r++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c: memberNoColIdx })];
        if (!cell || cell.v === undefined || cell.v === null || cell.v === '') continue;
        const memberNo = cell.v as string | number;

        const nameCell = memberNameIdx !== undefined ? sheet[XLSX.utils.encode_cell({ r, c: memberNameIdx })] : null;
        const memberName = nameCell && nameCell.v ? String(nameCell.v).trim() : 'Unknown';

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

        collected.push({ memberNo, memberName, rowIdx: r, expected, principal, installment, loanBalance, loanInterest, advanceBalance, advanceInterest, monthlyShare, loanRepayment, advRepayment, advInterestPaid, advDeduction, loanInterestPaid, registrationFee, passBook, fine });
      }
      log(`Loaded ${collected.length} member(s).`, 'success');
    }

    processedWorkbook = workbook;
    processedSheetName = sheetName;
    outputName = `${file.name.replace(/\.xlsx?$/i, '')}_filled.xlsx`;
    paymentMap.clear();

    logIcon.textContent = '✅';
    logTitle.textContent = 'Formulas filled!';
    formulaBanner.classList.remove('hidden');
    setStepActive(2);
    showContributionSection(collected);

  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    log(`Error: ${msg}`, 'error');
    logIcon.textContent = '❌';
    logTitle.textContent = 'Processing failed';
    console.error(err);
  } finally {
    setLoading(false);
  }
}

// ─── Build contribution table ─────────────────────────────────────────────────
function showContributionSection(members: MemberRow[]) {
  memberTableBody.innerHTML = '';
  paymentMap.clear(); // Reset map on re-process

  members.forEach((m) => {
    paymentMap.set(m.rowIdx, { cash: 0, paybill: 0, bank: 0, loanRepayment: 0, advRepayment: 0, riskFund: 0 }); // Init to 0

    const tr = document.createElement('tr');
    tr.dataset.rowIdx = String(m.rowIdx);

    const tdMno = document.createElement('td');
    tdMno.className = 'td-mno';
    tdMno.textContent = String(m.memberNo);
    tdMno.addEventListener('click', () => openDetailPopup(m));

    const tdExp = document.createElement('td');
    tdExp.className = 'td-exp';
    tdExp.textContent = fmt(m.expected);

    const createInputCol = (type: 'cash' | 'paybill' | 'bank') => {
      const td = document.createElement('td');
      const input = document.createElement('input');
      input.type = 'number';
      input.min = '0';
      input.step = '0.01';
      input.className = `amt-inp ${type}-inp`;
      input.placeholder = '0';
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
  colTotalExpected.textContent = sumExpected > 0 ? fmt(sumExpected) : '—';

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

  colTotalCash.textContent = cashTotal > 0 ? fmt(cashTotal) : '—';
  colTotalPaybill.textContent = paybillTotal > 0 ? fmt(paybillTotal) : '—';
  colTotalBank.textContent = bankTotal > 0 ? fmt(bankTotal) : '—';
}

// ─── Detail Popup ─────────────────────────────────────────────────────────────
function openDetailPopup(m: MemberRow) {
  // Use memberName for initials if possible, fallback to memberNo
  const initialsSrc = m.memberName && m.memberName !== 'Unknown' ? String(m.memberName) : String(m.memberNo);
  const initials = initialsSrc.slice(0, 2).toUpperCase();
  popupAvatar.textContent = initials;
  popupMemberNo.textContent = String(m.memberNo);
  popupMemberName.textContent = m.memberName;

  popupExpected.textContent = `KES ${fmt(m.expected)}`;
  popupPrincipal.textContent = m.principal > 0 ? fmt(m.principal) : '—';
  popupInstallment.textContent = m.installment > 0 ? fmt(m.installment) : '—';
  popupLoanBalance.textContent = m.loanBalance > 0 ? fmt(m.loanBalance) : '—';
  popupLoanInterest.textContent = m.loanInterest > 0 ? fmt(m.loanInterest) : '—';
  popupAdvBalance.textContent = m.advanceBalance > 0 ? fmt(m.advanceBalance) : '—';
  popupAdvInterest.textContent = m.advanceInterest > 0 ? fmt(m.advanceInterest) : '—';
  popupMShare.textContent = fmt(m.monthlyShare);

  const payData = paymentMap.get(m.rowIdx)!;
  popupRiskFund.value = String(payData.riskFund);
  popupLoanRepayment.value = payData.loanRepayment ? payData.loanRepayment.toString() : '';
  popupAdvRepayment.value = payData.advRepayment ? payData.advRepayment.toString() : '';

  popupRiskFund.oninput = () => {
    let val = parseInt(popupRiskFund.value) || 0;
    if (val > 50) { val = 50; popupRiskFund.value = '50'; }
    payData.riskFund = val;
    calcLiveValues(m);
    autoSaveToSheet(m.rowIdx);
  };

  popupLoanRepayment.oninput = () => {
    payData.loanRepayment = parseFloat(popupLoanRepayment.value) || 0;
    calcLiveValues(m);
    autoSaveToSheet(m.rowIdx);
  };
  popupAdvRepayment.oninput = () => {
    payData.advRepayment = parseFloat(popupAdvRepayment.value) || 0;
    calcLiveValues(m);
    autoSaveToSheet(m.rowIdx);
  };

  calcLiveValues(m); // initial paint for computed popup rows

  detailOverlay.classList.remove('hidden');
  document.body.style.overflow = 'hidden'; // prevent scrolling behind popup
}

function closeDetailPopup() {
  detailOverlay.classList.add('hidden');
  document.body.style.overflow = '';
}

[popupClose, popupDismiss].forEach(btn => btn.addEventListener('click', closeDetailPopup));
detailOverlay.addEventListener('click', (e) => { if (e.target === detailOverlay) closeDetailPopup(); });
document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeDetailPopup(); });

// ─── Auto-save: write current row data to in-memory workbook ──────────────────
function autoSaveToSheet(rowIdx: number) {
  if (!processedWorkbook) return;
  const sheet = processedWorkbook.Sheets[processedSheetName];
  // Build column map if not cached
  const ref = sheet['!ref'];
  if (!ref) return;
  const range = XLSX.utils.decode_range(ref);
  const colCache: Partial<Record<PaymentCol, number>> = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (!cell) continue;
    const hdr = (cell.v as string)?.toString().trim() as PaymentCol;
    if (PAYMENT_COLS.includes(hdr)) colCache[hdr] = c;
  }
  const payment = paymentMap.get(rowIdx);
  if (!payment) return;
  const { cash, paybill, bank, loanRepayment, advRepayment, riskFund } = payment;
  const vals: Record<PaymentCol, number> = {
    Cash: cash, Paybill: paybill, Bank: bank,
    LoanRepayment: loanRepayment, AdvanceRepayment: advRepayment,
    RiskFund: riskFund
  };
  for (const col of PAYMENT_COLS) {
    const cIdx = colCache[col];
    if (cIdx === undefined) continue;
    const addr = XLSX.utils.encode_cell({ r: rowIdx, c: cIdx });
    sheet[addr] = { t: 'n', v: vals[col] || 0 };
  }
}

// ─── Apply contributions & download ──────────────────────────────────────────
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
    const { cash, paybill, bank, loanRepayment, advRepayment, riskFund } = payment;

    const vals: Record<PaymentCol, number> = {
      Cash: cash, Paybill: paybill, Bank: bank,
      LoanRepayment: loanRepayment, AdvanceRepayment: advRepayment,
      RiskFund: riskFund
    };

    for (const col of PAYMENT_COLS) {
      const cIdx = payColIdx[col];
      if (cIdx === undefined) continue;
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: cIdx });
      sheet[addr] = { t: 'n', v: vals[col] || 0 }; // write 0 if blank/unpaid
      if (vals[col] > 0) writes++;
    }
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

// ─── Event wiring ─────────────────────────────────────────────────────────────
processBtn.addEventListener('click', () => { if (selectedFile) processFile(selectedFile); });
applyBtn.addEventListener('click', applyContributions);

// ─── Page init: enforce correct state on load ─────────────────────────────────
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

// ─── Warning before exit ───────────────────────────────────────────────────────
window.addEventListener('beforeunload', (e) => {
  if (selectedFile) {
    e.preventDefault();
    e.returnValue = ''; // Standard way to trigger the confirmation dialog
  }
});

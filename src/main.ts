import * as XLSX from 'xlsx';
import './style.css';

// ─── Target columns for formula fill ──────────────────────────────────────────
const FORMULA_COLUMNS: string[] = [
  'Total RePaid',
  'MonthlyShare',
  'TotalAdvance',
  'Shares C/F',
  'Loans C/F',
  'TotalCash',
];

// Payment channel columns to write contributions into
const PAYMENT_COLS = ['Cash', 'Paybill', 'Bank', 'RiskFund'] as const;
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
const step1Section = document.getElementById('step1Section')! as HTMLElement;
const step2Section = document.getElementById('step2Section')! as HTMLElement;
const memberTableBody = document.getElementById('memberTableBody')! as HTMLTableSectionElement;
const applyBtn = document.getElementById('applyBtn')! as HTMLButtonElement;
const startOverBtn = document.getElementById('startOverBtn')! as HTMLButtonElement;
const contribSummary = document.getElementById('contribSummary')! as HTMLDivElement;
const memberCountEl = document.getElementById('memberCount')! as HTMLSpanElement;

// Step indicator elements
const step1Ind = document.getElementById('step1Indicator')! as HTMLDivElement;
const step2Ind = document.getElementById('step2Indicator')! as HTMLDivElement;
const step3Ind = document.getElementById('step3Indicator')! as HTMLDivElement;

// Footer totals
const colTotalExpected = document.getElementById('colTotalExpected')!;
const colTotalCash = document.getElementById('colTotalCash')!;
const colTotalPaybill = document.getElementById('colTotalPaybill')!;
const colTotalBank = document.getElementById('colTotalBank')!;
const colTotalAll = document.getElementById('colTotalAll')!;

// ─── State ────────────────────────────────────────────────────────────────────
let selectedFile: File | null = null;
let processedWorkbook: XLSX.WorkBook | null = null;
let processedSheetName = '';
let outputName = 'modified.xlsx';

/** Each entry represents one member row in the sheet */
interface MemberRow {
  memberNo: string | number;
  rowIdx: number;
  expected: number;
}
let memberRows: MemberRow[] = [];

// ─── Helpers ──────────────────────────────────────────────────────────────────
function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1048576) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1048576).toFixed(2)} MB`;
}

function fmt(n: number): string {
  return n.toLocaleString('en-KE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
      if (parseInt(row, 10) === srcRow) {
        return `${sheet ?? ''}${colDollar}${col}${rowDollar}${tgtRow}`;
      }
      return match;
    }
  );
}

/**
 * Returns expected loan installment based on principal brackets.
 */
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
  if (principal <= 500000) return 20000; // handles up to 500k
  if (principal <= 550000) return 22000;
  if (principal <= 600000) return 24000;
  if (principal <= 650000) return 26000;
  if (principal <= 700000) return 28000;
  if (principal <= 750000) return 30000;
  if (principal <= 800000) return 32000;
  if (principal <= 850000) return 34000;
  if (principal <= 900000) return 36000;
  return 36000; // Cap or fallback
}

// ─── File selection ───────────────────────────────────────────────────────────
function acceptFile(file: File) {
  if (!file.name.match(/\.xlsx?$/i)) {
    alert('Please upload an Excel file (.xlsx or .xls).');
    return;
  }
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
  memberRows = [];
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
dropZone.addEventListener('dragover', (e) => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const file = e.dataTransfer?.files[0];
  if (file) acceptFile(file);
});
dropZone.addEventListener('click', () => fileInput.click());
browseBtn.addEventListener('click', (e) => { e.stopPropagation(); fileInput.click(); });
fileInput.addEventListener('change', () => { if (fileInput.files?.[0]) acceptFile(fileInput.files[0]); });
clearBtn.addEventListener('click', (e) => { e.stopPropagation(); clearFile(); });
startOverBtn.addEventListener('click', () => clearFile());

// ─── Core: formula fill & processing ──────────────────────────────────────────
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

    // Map column names → col index for everything we care about
    const colMap: Record<string, number> = {};
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cellAddr = XLSX.utils.encode_cell({ r: 0, c });
      const cell = sheet[cellAddr];
      if (!cell) continue;
      const header = (cell.v as string)?.toString().trim() ?? '';
      colMap[header] = c;
    }

    // ── Formula fill ──
    const found = FORMULA_COLUMNS.filter(col => colMap[col] !== undefined);
    const missing = FORMULA_COLUMNS.filter(col => colMap[col] === undefined);

    if (found.length === 0) {
      throw new Error(
        'None of the target formula column headers were found in row 1. ' +
        'Ensure column names exactly match (case-sensitive).'
      );
    }
    log(`Found ${found.length} formula column(s): ${found.join(', ')}`, 'success');
    if (missing.length > 0) log(`Skipped (not found): ${missing.join(', ')}`, 'warn');

    // Detect last occupied row
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

      if (!srcCell || (srcCell.v === undefined && !srcCell.f)) {
        log(`"${colName}" — row 2 is empty, skipping.`, 'warn');
        continue;
      }

      const hasFormula = !!srcCell.f;
      const srcFormula = srcCell.f ?? null;
      let copiedInCol = 0;

      for (let tgtRow = 2; tgtRow <= lastRow; tgtRow++) {
        const tgtAddr = XLSX.utils.encode_cell({ r: tgtRow, c: colIdx });
        if (hasFormula && srcFormula) {
          sheet[tgtAddr] = {
            t: srcCell.t,
            f: shiftFormula(srcFormula, 2, tgtRow + 1),
          };
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

    // ── Collect MemberNo & Calculate Expected Contribution ──
    const memberNoColIdx = colMap['MemberNo'];
    const pLoanCol = colMap['PrincipalLoan'];
    const lBalCol = colMap['LoanBalance'];
    const aBalCol = colMap['AdvanceBalance'];

    const collected: MemberRow[] = [];

    if (memberNoColIdx === undefined) {
      log('Warning: "MemberNo" column not found — contribution table will be empty.', 'warn');
    } else {
      for (let r = 1; r <= lastRow; r++) {
        const cell = sheet[XLSX.utils.encode_cell({ r, c: memberNoColIdx })];
        if (cell && cell.v !== undefined && cell.v !== null && cell.v !== '') {
          const memberNo = cell.v as string | number;

          // Gather numbers for expected formula
          const getNum = (cIdx?: number) => {
            if (cIdx === undefined) return 0;
            const c = sheet[XLSX.utils.encode_cell({ r, c: cIdx })];
            const v = parseFloat(c?.v as string);
            return isNaN(v) ? 0 : v;
          };

          const pLoan = getNum(pLoanCol);
          const lBal = getNum(lBalCol);
          const aBal = getNum(aBalCol);

          const expected =
            getInstallment(pLoan) +   // Loan installment
            (lBal * 0.015) +          // 1.5% Loan Interest
            aBal +                    // AdvanceBalance
            (aBal * 0.10) +           // 10% Advance Interest
            500 +                     // Default MonthlyShare
            50;                       // Fixed RiskFund

          collected.push({ memberNo, rowIdx: r, expected });
        }
      }
      log(`Loaded ${collected.length} member(s) from "MemberNo" column.`, 'success');
    }

    // Save state
    processedWorkbook = workbook;
    processedSheetName = sheetName;
    outputName = `${file.name.replace(/\.xlsx?$/i, '')}_filled.xlsx`;
    memberRows = collected;

    // Update log UI
    logIcon.textContent = '✅';
    logTitle.textContent = 'Formulas filled!';

    // Show banner & scroll to contribution section
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

  members.forEach((m, i) => {
    const tr = document.createElement('tr');
    tr.dataset.rowIdx = String(m.rowIdx);

    tr.innerHTML = `
      <td class="row-index">${i + 1}</td>
      <td class="member-no-cell">${m.memberNo}</td>
      <td class="expected-cell">${fmt(m.expected)}</td>
      <td><input type="number" min="0" step="0.01" placeholder="0.00"
            class="amount-input cash-input" data-col="Cash" /></td>
      <td><input type="number" min="0" step="0.01" placeholder="0.00"
            class="amount-input paybill-input" data-col="Paybill" /></td>
      <td><input type="number" min="0" step="0.01" placeholder="0.00"
            class="amount-input bank-input" data-col="Bank" /></td>
      <td class="row-total">—</td>
    `;

    // Row total live update
    tr.querySelectorAll<HTMLInputElement>('.amount-input').forEach(inp => {
      inp.addEventListener('input', () => {
        updateRowTotal(tr);
        updateColumnTotals();
      });
    });

    memberTableBody.appendChild(tr);
  });

  // Show member count badge
  memberCountEl.textContent = String(members.length);
  contribSummary.classList.remove('hidden');

  // Show the section and scroll
  step2Section.classList.remove('hidden');
  setTimeout(() => {
    step2Section.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }, 150);

  // Reset totals
  updateColumnTotals();
}

function getRowInputs(tr: HTMLTableRowElement): { cash: number; paybill: number; bank: number } {
  const cash = parseFloat((tr.querySelector('[data-col="Cash"]') as HTMLInputElement).value) || 0;
  const paybill = parseFloat((tr.querySelector('[data-col="Paybill"]') as HTMLInputElement).value) || 0;
  const bank = parseFloat((tr.querySelector('[data-col="Bank"]') as HTMLInputElement).value) || 0;
  return { cash, paybill, bank };
}

function updateRowTotal(tr: HTMLTableRowElement) {
  const { cash, paybill, bank } = getRowInputs(tr);
  const total = cash + paybill + bank;
  const cell = tr.querySelector('.row-total') as HTMLTableCellElement;
  if (total > 0) {
    cell.textContent = `KES ${fmt(total)}`;
    cell.classList.add('has-value');
  } else {
    cell.textContent = '—';
    cell.classList.remove('has-value');
  }
}

function updateColumnTotals() {
  let sumExpected = 0, sumCash = 0, sumPaybill = 0, sumBank = 0;
  memberRows.forEach(m => sumExpected += m.expected);

  memberTableBody.querySelectorAll<HTMLTableRowElement>('tr').forEach(tr => {
    const { cash, paybill, bank } = getRowInputs(tr);
    sumCash += cash;
    sumPaybill += paybill;
    sumBank += bank;
  });

  colTotalExpected.textContent = sumExpected > 0 ? `KES ${fmt(sumExpected)}` : '—';
  colTotalCash.textContent = sumCash > 0 ? `KES ${fmt(sumCash)}` : '—';
  colTotalPaybill.textContent = sumPaybill > 0 ? `KES ${fmt(sumPaybill)}` : '—';
  colTotalBank.textContent = sumBank > 0 ? `KES ${fmt(sumBank)}` : '—';
  const grand = sumCash + sumPaybill + sumBank;
  colTotalAll.textContent = grand > 0 ? `KES ${fmt(grand)}` : '—';
}

// ─── Apply contributions & download ──────────────────────────────────────────
function applyContributions() {
  if (!processedWorkbook) return;

  const sheet = processedWorkbook.Sheets[processedSheetName];
  const ref = sheet['!ref']!;
  const range = XLSX.utils.decode_range(ref);

  // Build payColIdx for Cash, Paybill, Bank, RiskFund
  const payColIdx: Partial<Record<PaymentCol, number>> = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = sheet[XLSX.utils.encode_cell({ r: 0, c })];
    if (!cell) continue;
    const hdr = (cell.v as string)?.toString().trim() as PaymentCol;
    if (PAYMENT_COLS.includes(hdr)) payColIdx[hdr] = c;
  }

  const missing = PAYMENT_COLS.filter(p => payColIdx[p] === undefined);
  if (missing.length > 0) {
    alert(`Warning: The following columns were not found in your sheet and will be skipped: ${missing.join(', ')}`);
  }

  // Write each row's values
  let writes = 0;
  let riskFundWrites = 0;

  memberTableBody.querySelectorAll<HTMLTableRowElement>('tr').forEach(tr => {
    const rowIdx = parseInt(tr.dataset.rowIdx!, 10);
    const { cash, paybill, bank } = getRowInputs(tr);
    const sum = cash + paybill + bank;

    const vals: Record<Exclude<PaymentCol, 'RiskFund'>, number> = { Cash: cash, Paybill: paybill, Bank: bank };

    // 1. Write Cash/Paybill/Bank
    for (const col of ['Cash', 'Paybill', 'Bank'] as const) {
      const cIdx = payColIdx[col];
      if (cIdx === undefined) continue;
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: cIdx });
      const v = vals[col];
      if (v > 0) {
        sheet[addr] = { t: 'n', v };
        writes++;
      } else {
        // If user left blank (0), preserve existing value or write 0
        if (!sheet[addr]) {
          sheet[addr] = { t: 'n', v: 0 };
        }
      }
    }

    // 2. Autonomous RiskFund
    if (sum > 0) {
      const rfIdx = payColIdx['RiskFund'];
      if (rfIdx !== undefined) {
        const rfAddr = XLSX.utils.encode_cell({ r: rowIdx, c: rfIdx });
        sheet[rfAddr] = { t: 'n', v: 50 };
        riskFundWrites++;
      }
    }
  });

  // Serialise and download
  const buf = XLSX.write(processedWorkbook, { bookType: 'xlsx', type: 'buffer' }) as ArrayBuffer;
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = outputName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);

  // Mark step 3 done in indicator
  setStepActive(3);
  step3Ind.classList.remove('active');
  step3Ind.classList.add('done');

  console.log(`Applied ${writes} contribution value(s) and auto-filled ${riskFundWrites} RiskFund(s).`);
}

// ─── Event wiring ─────────────────────────────────────────────────────────────
processBtn.addEventListener('click', () => {
  if (selectedFile) processFile(selectedFile);
});

applyBtn.addEventListener('click', applyContributions);

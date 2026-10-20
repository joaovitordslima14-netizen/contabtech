import ExcelJS from "exceljs";
import { emptyState, getMonthSummary, mergeTaxonomy } from "./finance.js";

const MONTH_SHEETS = [
  ["Jan", 1],
  ["Fev", 2],
  ["Mar", 3],
  ["Abr", 4],
  ["Mai", 5],
  ["Jun", 6],
  ["Jul", 7],
  ["Ago", 8],
  ["Set", 9],
  ["Out", 10],
  ["Nov", 11],
  ["Dez", 12],
];

const MONTH_BY_NAME = new Map(MONTH_SHEETS);

function cellValue(sheet, address) {
  if (typeof sheet?.getCell === "function") {
    const value = sheet.getCell(address).value;
    if (value && typeof value === "object") {
      if ("result" in value) return value.result ?? "";
      if (Array.isArray(value.richText)) return value.richText.map((part) => part.text).join("");
      if ("text" in value) return value.text ?? "";
    }
    return value ?? "";
  }
  return sheet?.[address]?.v ?? "";
}

function getSheet(workbook, name) {
  return workbook.Sheets?.[name] || workbook.getWorksheet?.(name);
}

function isExample(value) {
  return typeof value === "string" && /^\s*exemplo\b/i.test(value);
}

export function parseExcelDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, "0")}-${String(value.getDate()).padStart(2, "0")}`;
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    const date = new Date(Date.UTC(1899, 11, 30) + Math.floor(value) * 86_400_000);
    if (Number.isNaN(date.getTime())) return "";
    return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`;
  }
  if (typeof value === "string") {
    const text = value.trim();
    const local = text.match(/^(\d{1,2})[/.](\d{1,2})[/.](\d{4})$/);
    if (local) return `${local[3]}-${local[2].padStart(2, "0")}-${local[1].padStart(2, "0")}`;
    if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  }
  return "";
}

function excelCents(value) {
  const amount = Number(value);
  return Number.isFinite(amount) ? Math.round(amount * 100) : 0;
}

function periodFor(year, month) {
  return `${year}-${String(month).padStart(2, "0")}`;
}

function addSourceIdentity(record, sheetName, row, field) {
  const identity = [
    record.kind || field,
    record.period,
    record.date || "",
    record.amountCents ?? record.installmentAmountCents ?? 0,
    record.category || record.brand || record.creditor || record.description || "",
    record.paymentMethod || "",
    record.description || "",
  ].join("|");
  return {
    ...record,
    source: `${sheetName}!${field}${row}`,
    sourceFingerprint: identity.toLocaleLowerCase("pt-BR"),
  };
}

function collectYearCandidates(workbook) {
  const counts = new Map();
  for (const [sheetName] of MONTH_SHEETS) {
    const sheet = getSheet(workbook, sheetName);
    if (!sheet) continue;
    const dateRanges = [
      ["B", 15, 29],
      ["L", 13, 43],
      ["S", 13, 93],
      ["E", 35, 46],
      ["O", 54, 80],
    ];
    for (const [column, firstRow, lastRow] of dateRanges) {
      for (let rowNumber = firstRow; rowNumber <= lastRow; rowNumber += 1) {
        const date = parseExcelDate(cellValue(sheet, `${column}${rowNumber}`));
        const year = Number(date.slice(0, 4));
        if (year >= 1900 && year <= 2200) counts.set(year, (counts.get(year) || 0) + 1);
      }
    }
  }
  return counts;
}

function collectTaxonomy(workbook) {
  const found = {
    categories: [],
    paymentMethods: [],
    incomeTypes: [],
    cardBrands: [],
    debtTypes: [],
  };

  for (const [sheetName] of MONTH_SHEETS) {
    const sheet = getSheet(workbook, sheetName);
    if (!sheet) continue;
    for (let row = 14; row <= 32; row += 1) found.categories.push(cellValue(sheet, `AC${row}`));
    for (let row = 39; row <= 43; row += 1) found.paymentMethods.push(cellValue(sheet, `AC${row}`));
    for (let row = 15; row <= 29; row += 1) found.incomeTypes.push(cellValue(sheet, `C${row}`));
    for (let row = 35; row <= 46; row += 1) found.cardBrands.push(cellValue(sheet, `B${row}`));
    for (let row = 54; row <= 80; row += 1) found.debtTypes.push(cellValue(sheet, `H${row}`));
  }

  return Object.fromEntries(
    Object.entries(found).map(([key, values]) => [
      key,
      [...new Set(values.filter((value) => typeof value === "string" && value.trim()).map((value) => value.trim()))],
    ]),
  );
}

const parseRawExcelDate = parseExcelDate;

function withYear(date, year) {
  if (!date) return date;
  const [, month, day] = date.split("-");
  const lastDay = new Date(Date.UTC(year, Number(month), 0)).getUTCDate();
  return `${year}-${month}-${String(Math.min(Number(day), lastDay)).padStart(2, "0")}`;
}

function readRecords(workbook, inferredYear, report, forcedYear) {
  const imported = { transactions: [], invoices: [], debts: [], goals: {} };
  const parseExcelDate = (value) => {
    const date = parseRawExcelDate(value);
    return forcedYear ? withYear(date, forcedYear) : date;
  };

  for (const [sheetName, month] of MONTH_SHEETS) {
    const sheet = getSheet(workbook, sheetName);
    if (!sheet) continue;
    const period = periodFor(inferredYear, month);

    for (let row = 15; row <= 29; row += 1) {
      const date = parseExcelDate(cellValue(sheet, `B${row}`));
      const type = String(cellValue(sheet, `C${row}`)).trim();
      const description = String(cellValue(sheet, `D${row}`)).trim();
      const amountCents = excelCents(cellValue(sheet, `E${row}`));
      if (isExample(type) || isExample(description)) {
        report.skippedExamples += 1;
        continue;
      }
      if (!amountCents) continue;
      if (!date) {
        report.undatedRecords += 1;
        imported.transactions.push(
          addSourceIdentity(
            { kind: "income", period, date: "", amountCents, incomeType: type, description },
            sheetName,
            row,
            "B",
          ),
        );
        continue;
      }
      imported.transactions.push(
        addSourceIdentity(
          { kind: "income", period: date.slice(0, 7), date, amountCents, incomeType: type, description },
          sheetName,
          row,
          "B",
        ),
      );
    }

    for (let row = 13; row <= 43; row += 1) {
      const category = String(cellValue(sheet, `K${row}`)).trim();
      const date = parseExcelDate(cellValue(sheet, `L${row}`));
      const description = String(cellValue(sheet, `M${row}`)).trim();
      const paymentMethod = String(cellValue(sheet, `N${row}`)).trim();
      const paidValue = cellValue(sheet, `O${row}`);
      const amountCents = excelCents(cellValue(sheet, `P${row}`));
      if (isExample(description)) {
        report.skippedExamples += 1;
        continue;
      }
      if (!amountCents) continue;
      const record = {
        kind: "expense",
        period: date ? date.slice(0, 7) : period,
        date,
        amountCents,
        category,
        description,
        paymentMethod,
        fixed: true,
        paid: paidValue === true || Number(paidValue) === 1,
      };
      if (!date) report.undatedRecords += 1;
      imported.transactions.push(addSourceIdentity(record, sheetName, row, "L"));
    }

    for (let row = 13; row <= 93; row += 1) {
      const category = String(cellValue(sheet, `R${row}`)).trim();
      const date = parseExcelDate(cellValue(sheet, `S${row}`));
      const description = String(cellValue(sheet, `U${row}`)).trim();
      const paymentMethod = String(cellValue(sheet, `V${row}`)).trim();
      const amountCents = excelCents(cellValue(sheet, `W${row}`));
      if (isExample(description)) {
        report.skippedExamples += 1;
        continue;
      }
      if (!amountCents) continue;
      const record = {
        kind: "expense",
        period: date ? date.slice(0, 7) : period,
        date,
        amountCents,
        category,
        description,
        paymentMethod,
        fixed: false,
        paid: true,
        installments: cellValue(sheet, `T${row}`) || "",
      };
      if (!date) report.undatedRecords += 1;
      imported.transactions.push(addSourceIdentity(record, sheetName, row, "S"));
    }

    for (let row = 37; row <= 48; row += 1) {
      const location = String(cellValue(sheet, `H${row}`)).trim();
      const amountCents = excelCents(cellValue(sheet, `I${row}`));
      if (isExample(location)) {
        report.skippedExamples += 1;
        continue;
      }
      if (!location || !amountCents) continue;
      imported.transactions.push(
        addSourceIdentity(
          { kind: "investment", period, date: "", amountCents, description: location },
          sheetName,
          row,
          "I",
        ),
      );
      report.undatedRecords += 1;
    }

    for (let row = 35; row <= 46; row += 1) {
      const brand = String(cellValue(sheet, `B${row}`)).trim();
      const paymentMethod = String(cellValue(sheet, `C${row}`)).trim();
      const amountCents = excelCents(cellValue(sheet, `D${row}`));
      const dueDate = parseExcelDate(cellValue(sheet, `E${row}`));
      const paid = cellValue(sheet, `F${row}`) === true || Number(cellValue(sheet, `F${row}`)) === 1;
      if (isExample(brand)) {
        report.skippedExamples += 1;
        continue;
      }
      if (!brand || (!dueDate && amountCents < 100 && !paid)) {
        if (brand && amountCents > 0 && !dueDate && !paid) report.skippedSmallInvoicePlaceholders += 1;
        continue;
      }
      const record = {
        period,
        brand,
        paymentMethod,
        amountCents,
        dueDate,
        paid,
      };
      imported.invoices.push(addSourceIdentity(record, sheetName, row, "D"));
    }

    for (let row = 54; row <= 80; row += 1) {
      const type = String(cellValue(sheet, `H${row}`)).trim();
      const paymentMethod = String(cellValue(sheet, `I${row}`)).trim();
      const creditor = String(cellValue(sheet, `J${row}`)).trim();
      const installmentAmountCents = excelCents(cellValue(sheet, `M${row}`));
      const totalAmountCents = excelCents(cellValue(sheet, `N${row}`));
      if (isExample(creditor)) {
        report.skippedExamples += 1;
        continue;
      }
      if (
        !creditor ||
        (!installmentAmountCents && !totalAmountCents)
      ) {
        continue;
      }
      const date = parseExcelDate(cellValue(sheet, `O${row}`));
      const record = {
        period,
        type,
        paymentMethod,
        creditor,
        currentInstallment: Number(cellValue(sheet, `K${row}`)) || 0,
        totalInstallments: Number(cellValue(sheet, `L${row}`)) || 0,
        installmentAmountCents,
        totalAmountCents,
        date,
        paid: cellValue(sheet, `P${row}`) === true || Number(cellValue(sheet, `P${row}`)) === 1,
      };
      imported.debts.push(addSourceIdentity(record, sheetName, row, "J"));
    }

    const monthIndex = month - 1;
    const goalColumn = String.fromCharCode(71 + monthIndex);
    const goalCents = excelCents(cellValue(getSheet(workbook, "Investimentos"), `${goalColumn}9`));
    if (goalCents) imported.goals[period] = goalCents;

    const rawValues = ["B10", "C10", "D10", "E10"].map((address) => {
      const value = cellValue(sheet, address);
      return value === "" || typeof value === "string" && value.startsWith("#")
        ? null
        : excelCents(value);
    });
    if (rawValues[2] === null) {
      let investido = 0;
      for (let row = 37; row <= 48; row += 1) {
        investido += excelCents(cellValue(sheet, `I${row}`));
      }
      rawValues[2] = investido;
    }
    if (rawValues.every((value) => value !== null)) {
      const calculated = getMonthSummary(
        {
          transactions: imported.transactions,
          invoices: imported.invoices,
          debts: imported.debts,
          goals: imported.goals,
        },
        period,
      );
      const actual = [
        calculated.incomeCents,
        calculated.expenseCents,
        calculated.investmentCents,
        calculated.balanceCents,
      ];
      report.comparisons.push({
        period,
        matches: rawValues.every((value, index) => value === actual[index]),
        spreadsheet: rawValues,
        imported: actual,
      });
    } else {
      report.unavailableComparisons.push(period);
    }
  }

  return imported;
}

export function parseLegacyWorkbook(workbook, { fallbackYear = new Date().getFullYear(), year } = {}) {
  const yearCounts = collectYearCandidates(workbook);
  const forcedYear = year ? Number(year) : null;
  const inferredYear =
    forcedYear || [...yearCounts].sort((left, right) => right[1] - left[1])[0]?.[0] || Number(fallbackYear);
  if (!Number.isInteger(inferredYear) || inferredYear < 1900 || inferredYear > 2200) {
    throw new Error("Não foi possível identificar um ano válido para importar a planilha.");
  }

  const report = {
    inferredYear,
    yearCounts: Object.fromEntries(yearCounts),
    undatedRecords: 0,
  skippedExamples: 0,
  skippedSmallInvoicePlaceholders: 0,
  comparisons: [],
  unavailableComparisons: [],
  };
  const imported = readRecords(workbook, inferredYear, report, forcedYear);
  const base = emptyState();
  const taxonomy = mergeTaxonomy(base.taxonomy, collectTaxonomy(workbook));

  return {
    ...imported,
    taxonomy,
    report: {
      ...report,
      transactionCount: imported.transactions.length,
      invoiceCount: imported.invoices.length,
      debtCount: imported.debts.length,
      mismatchedPeriods: report.comparisons.filter((item) => !item.matches).map((item) => item.period),
    },
  };
}

function normalizeHeader(value) {
  return String(value)
    .trim()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("pt-BR");
}

function parseXlsxAmount(value) {
  if (typeof value === "number") return Number.isFinite(value) ? Math.round(value * 100) : null;
  let text = String(value ?? "").trim().replace(/^R\$\s*/i, "").replace(/\s/g, "");
  if (text.startsWith("(") && text.endsWith(")")) text = `-${text.slice(1, -1)}`;
  if (text.includes(",")) text = text.replace(/\./g, "").replace(",", ".");
  const amount = Number(text);
  return Number.isFinite(amount) ? Math.round(amount * 100) : null;
}

function spreadsheetBoolean(value) {
  return ["sim", "true", "1", "s"].includes(normalizeHeader(value));
}

export function parseTransactionsWorksheet(sheet, { fallbackYear = new Date().getFullYear(), year } = {}) {
  const rows = [];
  sheet.eachRow({ includeEmpty: false }, (row) => {
    rows.push(row.values.slice(1).map((value) => value ?? ""));
  });
  if (rows.length < 2) throw new Error("A planilha está vazia ou não contém movimentações.");

  const headers = rows[0].map(normalizeHeader);
  const columns = {
    period: headers.indexOf("periodo"),
    date: headers.indexOf("data"),
    type: headers.indexOf("tipo"),
    category: headers.indexOf("categoria/tipo"),
    description: headers.indexOf("descricao"),
    paymentMethod: headers.indexOf("forma de pagamento"),
    fixed: headers.indexOf("fixa"),
    paid: headers.indexOf("paga"),
    amount: headers.findIndex((header) => ["valor (brl)", "valor (r$)", "valor"].includes(header)),
  };
  if (
    columns.type < 0 ||
    columns.amount < 0 ||
    (columns.period < 0 && columns.date < 0)
  ) {
    throw new Error("Planilha incompatível. Use o XLSX de movimentações exportado pelo Contabtech.");
  }

  const forcedYear = year ? Number(year) : null;
  if (
    (forcedYear && (!Number.isInteger(forcedYear) || forcedYear < 1900 || forcedYear > 2200)) ||
    (!forcedYear && (!Number.isInteger(Number(fallbackYear)) || fallbackYear < 1900 || fallbackYear > 2200))
  ) {
    throw new Error("Não foi possível identificar um ano válido para importar a planilha.");
  }
  const imported = { transactions: [], invoices: [], debts: [], goals: {} };
  const yearCounts = new Map();
  let undatedRecords = 0;

  for (const [index, values] of rows.slice(1).entries()) {
    const valueAt = (column) => (column < 0 ? "" : values[column] ?? "");
    const type = normalizeHeader(valueAt(columns.type));
    const kind =
      type === "entrada" ? "income" :
      type === "investimento" ? "investment" :
      type === "saida fixa" || type === "despesa fixa" ? "expense" :
      type === "saida variavel" || type === "despesa variavel" || type === "saida" || type === "despesa" ? "expense" :
      null;
    if (!kind) throw new Error(`Tipo de movimentação inválido na linha ${index + 2} da planilha.`);

    const amountCents = parseXlsxAmount(valueAt(columns.amount));
    if (amountCents === null || amountCents <= 0) {
      throw new Error(`Valor inválido na linha ${index + 2} da planilha.`);
    }
    const dateValue = valueAt(columns.date);
    const rawDate = parseExcelDate(dateValue);
    if (String(dateValue).trim() && !rawDate) {
      throw new Error(`Data inválida na linha ${index + 2} da planilha.`);
    }
    const rawPeriod = String(valueAt(columns.period)).trim();
    const periodMatch = rawPeriod.match(/^(\d{4})-(\d{2})$/);
    if (rawPeriod && (!periodMatch || Number(periodMatch[2]) < 1 || Number(periodMatch[2]) > 12)) {
      throw new Error(`Período inválido na linha ${index + 2} da planilha.`);
    }
    const sourceYear = rawDate
      ? Number(rawDate.slice(0, 4))
      : periodMatch
        ? Number(periodMatch[1])
        : forcedYear || Number(fallbackYear);
    if (sourceYear < 1900 || sourceYear > 2200) {
      throw new Error(`Ano inválido na linha ${index + 2} da planilha.`);
    }
    yearCounts.set(sourceYear, (yearCounts.get(sourceYear) || 0) + 1);
    const effectiveYear = forcedYear || sourceYear;
    const date = rawDate ? withYear(rawDate, effectiveYear) : "";
    const period = date
      ? date.slice(0, 7)
      : `${effectiveYear}-${String(Number(periodMatch?.[2]) || 1).padStart(2, "0")}`;
    const fixedValue = valueAt(columns.fixed);
    const fixed = kind === "expense" && (String(fixedValue).trim() ? spreadsheetBoolean(fixedValue) : type.includes("fixa"));
    const category = String(valueAt(columns.category)).trim();
    const description = String(valueAt(columns.description)).trim();
    const paymentMethod = String(valueAt(columns.paymentMethod)).trim();

    imported.transactions.push({
      kind,
      period,
      date,
      amountCents,
      ...(kind === "income" ? { incomeType: category } : {}),
      ...(kind === "expense" ? { category, fixed, paid: spreadsheetBoolean(valueAt(columns.paid)) } : {}),
      ...(kind === "investment" ? {} : { paymentMethod }),
      description,
    });
    if (!date) undatedRecords += 1;
  }

  const inferredYear =
    forcedYear ||
    [...yearCounts].sort((left, right) => right[1] - left[1])[0]?.[0] ||
    Number(fallbackYear);
  return {
    ...imported,
    taxonomy: emptyState().taxonomy,
    report: {
      sourceFormat: "XLSX",
      inferredYear,
      yearCounts: Object.fromEntries(yearCounts),
      undatedRecords,
      skippedExamples: 0,
      skippedSmallInvoicePlaceholders: 0,
      comparisons: [],
      unavailableComparisons: [],
      transactionCount: imported.transactions.length,
      invoiceCount: 0,
      debtCount: 0,
      mismatchedPeriods: [],
    },
  };
}

export async function createXlsxBackup(data) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Contabtech";
  const sheet = workbook.addWorksheet("Backup");
  const serialized = JSON.stringify(data);
  sheet.addRow(["Contabtech Finanças XLSX", 1, new Date().toISOString()]);
  sheet.addRow(["Dados"]);
  for (let offset = 0; offset < serialized.length; offset += 30000) {
    sheet.addRow([serialized.slice(offset, offset + 30000)]);
  }
  return workbook.xlsx.writeBuffer();
}

export function parseXlsxBackup(workbook) {
  const sheet = workbook.getWorksheet("Backup");
  if (sheet?.getCell("A1").value !== "Contabtech Finanças XLSX" || Number(sheet.getCell("B1").value) !== 1) {
    throw new Error("Arquivo XLSX de backup inválido ou versão não compatível.");
  }
  const chunks = [];
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber >= 3 && row.getCell(1).value) chunks.push(String(row.getCell(1).value));
  });
  if (!chunks.length) throw new Error("O arquivo XLSX não contém os dados do backup.");
  let data;
  try {
    data = JSON.parse(chunks.join(""));
  } catch (error) {
    throw new Error("Não foi possível decodificar os dados do backup XLSX.", { cause: error });
  }
  return {
    application: "Contabtech Finanças",
    schemaVersion: 1,
    exportedAt: String(sheet.getCell("C1").value || ""),
    data,
  };
}

export async function readXlsxBackupFile(file) {
  if (!file.name.toLocaleLowerCase("pt-BR").endsWith(".xlsx")) {
    throw new Error("Selecione um arquivo no formato .xlsx.");
  }
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  return parseXlsxBackup(workbook);
}

export async function readXlsxImportFile(file, options) {
  if (!file.name.toLocaleLowerCase("pt-BR").endsWith(".xlsx")) {
    throw new Error("Selecione um arquivo no formato .xlsx.");
  }
  const bytes = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  if (workbook.getWorksheet("Backup")) {
    const backup = parseXlsxBackup(workbook);
    const data = backup.data;
    if (
      !Array.isArray(data.transactions) ||
      !Array.isArray(data.invoices) ||
      !Array.isArray(data.debts) ||
      !data.goals ||
      typeof data.goals !== "object"
    ) {
      throw new Error("Arquivo XLSX de backup inválido ou incompleto.");
    }
    const year = Number(data.transactions[0]?.period?.slice(0, 4)) || new Date().getFullYear();
    return {
      ...data,
      report: {
        sourceFormat: "XLSX Backup",
        inferredYear: year,
        yearCounts: {},
        undatedRecords: 0,
        skippedExamples: 0,
        skippedSmallInvoicePlaceholders: 0,
        comparisons: [],
        unavailableComparisons: [],
        transactionCount: data.transactions.length,
        invoiceCount: data.invoices.length,
        debtCount: data.debts.length,
        mismatchedPeriods: [],
      },
    };
  }
  const transactionsSheet = workbook.getWorksheet("Movimentações");
  if (transactionsSheet) return parseTransactionsWorksheet(transactionsSheet, options);
  return parseLegacyWorkbook(workbook, options);
}
